# PCB_BULK.py - Bulk processing multiple PCB (Pushpam) files
import pandas as pd
import re
import os
import glob
from pathlib import Path


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'US6.5', 'W') -> 'VR1943EEA-6.5WG'
    """
    base = str(base).strip() if base else ''
    item_size = str(item_size).strip() if item_size else ''
    tone = str(tone).strip().upper() if tone else ''
    if not base or base.upper() == 'NAN':
        return ''
    # Detect INCH before stripping — insert IN in suffix (e.g. 7 INCH+W -> 7INWG)
    has_inch = bool(re.search(r'\bINCH\b', item_size, flags=re.IGNORECASE))
    size_num = re.sub(r'^(?:UP|US|EU|IT|UT|TS|IS)\s*', '', item_size, flags=re.IGNORECASE).strip()
    size_num = re.sub(r'\s*INCH\s*$', '', size_num, flags=re.IGNORECASE).strip()
    try:
        f = float(size_num)
        size_num = str(int(f)) if f.is_integer() else str(f)
    except (ValueError, TypeError):
        pass
    # Normalize multi-char tones: 'YV' -> 'Y', 'WG' -> 'W', etc. Keep 'PT' as-is
    if tone and len(tone) > 1 and tone != 'PT':
        first = tone[0]
        if first in ('W', 'Y', 'P', 'R'):
            tone = first
    in_part = 'IN' if has_inch else ''
    if tone == 'PT':
        suffix = f"{size_num}{in_part}PT" if size_num else (f"{in_part}PT" if in_part else 'PT')
    elif tone in ('W', 'Y', 'P', 'R'):
        suffix = f"{size_num}{in_part}{tone}G" if size_num else f"{in_part}{tone}G"
    elif tone == 'AG':
        suffix = f"{size_num}{in_part}AG" if size_num else (f"{in_part}AG" if in_part else 'AG')
    else:
        suffix = size_num
    return f"{base}-{suffix}" if suffix else base


_ITEM_SIZE_LOOKUP = None
_CLIENT_STYLE_LIST = None


def _get_size_lookup():
    global _ITEM_SIZE_LOOKUP
    if _ITEM_SIZE_LOOKUP is not None:
        return _ITEM_SIZE_LOOKUP
    _ITEM_SIZE_LOOKUP = {}
    try:
        mst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ItemSize_Mst.xlsx')
        _df_mst = pd.read_excel(mst)
        for val in _df_mst['Item Size Code'].dropna():
            vs = str(val).strip()
            if vs and vs.upper() != 'NAN':
                k = _normalize_size_key(vs)
                if k:
                    _ITEM_SIZE_LOOKUP[k] = vs
    except Exception:
        pass
    return _ITEM_SIZE_LOOKUP


def _normalize_size_key(s):
    s = str(s).strip()
    if not s or s.upper() == 'NAN':
        return ''
    m = re.match(r'^(\d+(?:\.\d+)?)\s*INCH$', s, re.IGNORECASE)
    if m:
        return f"{float(m.group(1)):.2f}inch"
    return re.sub(r'\s+', '', s).lower()


def _map_item_size(raw):
    """Map raw ItemSize to its canonical form from ItemSize_Mst.xlsx."""
    if not raw or str(raw).strip().upper() in ('', 'NAN'):
        return raw
    lookup = _get_size_lookup()
    key = _normalize_size_key(str(raw).strip())
    return lookup.get(key, raw)


def _get_client_style_list() -> list:
    """Load and cache the approved Client Style No list from PCB_CS_100826.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    try:
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'PCB_CS_100826.xlsx')
        _df_cs = pd.read_excel(cs_path)
        col = 'Client Style No'
        if col in _df_cs.columns:
            _CLIENT_STYLE_LIST = [
                str(v).strip() for v in _df_cs[col].dropna()
                if str(v).strip() and str(v).strip().upper() != 'NAN'
            ]
    except Exception:
        pass
    return _CLIENT_STYLE_LIST


def _lookup_client_style(built_code: str) -> str:
    """
    Match the built StyleCode against PCB_CS_100826.xlsx 'Client Style No'.

    Matching priority:
      1. Exact match (case-insensitive)
      2. Master entry starts-with built_code (case-insensitive)
      3. Try inserting 'IN' before the metal/tone suffix
      4. No match -> return built_code unchanged
    """
    cs_list = _get_client_style_list()
    if not built_code or not cs_list:
        return built_code

    built_upper = built_code.upper()

    # 1. Exact match (case-insensitive)
    for cs in cs_list:
        if cs.upper() == built_upper:
            return cs

    # 2. Master entry starts with built_code (case-insensitive, extra variant suffix like XV)
    for cs in cs_list:
        if cs.upper().startswith(built_upper):
            return cs

    # 3. Try inserting 'IN' before the metal/tone suffix
    _in_pat = re.compile(
        r'^(.+-)(\d+(?:\.\d+)?)((?:WG|YG|RG|AG|PT|W|Y|R|P).*)$',
        re.IGNORECASE,
    )
    m = _in_pat.match(built_code)
    if m:
        with_in = m.group(1) + m.group(2) + 'IN' + m.group(3).upper()
        with_in_upper = with_in.upper()
        # 3a. Exact match on IN-variant (case-insensitive)
        for cs in cs_list:
            if cs.upper() == with_in_upper:
                return cs
        # 3b. Prefix match on IN-variant (case-insensitive, master also has extra suffix)
        for cs in cs_list:
            if cs.upper().startswith(with_in_upper):
                return cs

    return built_code

def process_pcb_file(input_file_path, output_folder=None, priority_value=None, skuno_value=None):
    """
    Process single PCB (Pushpam) Excel file and convert to standardized format
    
    Parameters:
    input_file_path (str): Path to input Excel file
    output_folder (str): Folder for output CSV file (optional)
    priority_value (str): Priority value for all rows
    skuno_value (str): SKUNo value for all rows
    
    Returns:
    tuple: (success_status, output_path, error_message, dataframe)
    """
    try:
        # --- Step 1: Read Excel and skip top rows ---
        df = pd.read_excel(input_file_path, skiprows=1)

        # --- Step 2: Assign meaningful column names ---
        df.columns = [
            "SrNo",
            "OrderGroup",
            "ItemPoNo",
            "StyleCode",
            "Metal",
            "ItemSize",
            "Purity",
            "Nos",
            "OrderQty",
            "GrossWt",
            "NetWt",
            "StoneWt",
            "StonePcs",
            "CustomerProductionInstruction",
            "Category"   # Special remark coln in PO
        ]

        # --- Step 3: Map Metal to Tone ---
        tone_map = {
            "AG925": "",
            "G585W": "W",
            "G585P": "P",
            "G585Y": "Y",
            "G585WZ": "W",
            "G585YZ": "Y",
            "G585PZ": "P",
            "G585W-NI1811-RHC": "W",
            "G585Y-C143GR": "Y",
            "G585W-NPF301": "W",
            "G585P-C145N": "P",
            "G587Y": "Y"
        }
        df['Tone'] = df['Metal'].map(tone_map).fillna("")

        # --- Step 4: Insert Tone after Metal ---
        cols = df.columns.tolist()
        metal_index = cols.index("Metal")
        cols.remove("Tone")
        cols.insert(metal_index + 1, "Tone")
        df = df[cols]

        # --- Step 5: Add new columns after ItemPoNo ---
        # Get index of ItemPoNo
        itempo_index = df.columns.get_loc("ItemPoNo")

        # Create new columns with default blank values
        df.insert(itempo_index + 1, "ItemRefNo", "")
        df.insert(itempo_index + 2, "StockType", "")

        # Use provided Priority value or prompt
        if priority_value is None:
            priority_value = input(f"Enter Priority value for {Path(input_file_path).name} (TMP/SMP/REPSMP/REVSMP/SDPL/CUSCAD/MSP): ")
        
        df.insert(itempo_index + 3, "Priority", priority_value)
        df.insert(itempo_index + 4, "MakeType", "")

        # --- Step 6: Select and reorder final columns ---
        selected_columns = [
            'SrNo', 'StyleCode', 'ItemSize', 'OrderQty', 'Metal', 'Tone',
            'ItemPoNo', 'ItemRefNo', 'StockType', 'Priority', 'MakeType',
            'CustomerProductionInstruction', 'OrderGroup'
        ]

        df_selected = df[selected_columns].copy()
        
        # Make NaN values in CustomerProductionInstruction blank
        df_selected['CustomerProductionInstruction'] = df_selected['CustomerProductionInstruction'].fillna("")
        
        # --- Step 7: Add SpecialRemarks column after CustomerProductionInstruction ---
        cpi_index = df_selected.columns.get_loc("CustomerProductionInstruction")
        df_selected.insert(
            cpi_index + 1,
            "SpecialRemarks",
            "CUSTOMER-PCB, PO#PCB - " + df_selected['ItemPoNo'].astype(str) + ",NEW SAMPLE ORDER, DIAMOND GRADE-CZ"
        )

        special_index = df_selected.columns.get_loc("SpecialRemarks")

        # Mapping function
        def map_design_instruction(tone):
            if tone == "W":
                return "WHITE RODIUM"
            else:  # Y, PT, or blank
                return "NO RODIUM"

        # Insert column after SpecialRemarks
        df_selected.insert(
            special_index + 1,
            "DesignProductionInstruction",
            df_selected['Tone'].apply(map_design_instruction)
        )

        # --- Add StampInstruction column after DesignProductionInstruction ---
        dpi_index = df_selected.columns.get_loc("DesignProductionInstruction")

        df_selected.insert(
            dpi_index + 1,
            "StampInstruction",
            "LT STAMP (LOGO) + " + df_selected['Metal'].astype(str)
        )

        # --- Add new columns after OrderGroup ---
        ordergroup_index = df_selected.columns.get_loc("OrderGroup")

        # Use provided SKUNo value or prompt
        if skuno_value is None:
            skuno_value = input(f"Enter SKUNo value for {Path(input_file_path).name}: ")

        # Insert columns one by one after OrderGroup
        df_selected.insert(ordergroup_index + 1, "Certificate", "")
        df_selected.insert(ordergroup_index + 2, "SKUNo", skuno_value)
        df_selected.insert(ordergroup_index + 3, "Basestoneminwt", "")
        df_selected.insert(ordergroup_index + 4, "Basestonemaxwt", "")
        df_selected.insert(ordergroup_index + 5, "Basemetalminwt", "")
        df_selected.insert(ordergroup_index + 6, "Basemetalmaxwt", "")
        df_selected.insert(ordergroup_index + 7, "Productiondeliverydate", "")
        df_selected.insert(ordergroup_index + 8, "Expecteddeliverydate", "")
        df_selected.insert(ordergroup_index + 9, "SetPrice", "")
        df_selected.insert(ordergroup_index + 10, "StoneQuality", "")

        # Build full StyleCode: base-sizeWG / base-sizePT etc.
        df_selected['ItemSize'] = df_selected['ItemSize'].apply(_map_item_size)
        df_selected['StyleCode'] = df_selected.apply(
            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
        )
        df_selected['StyleCode'] = df_selected['StyleCode'].apply(_lookup_client_style)
        df_selected.loc[df_selected['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        # --- Step 8: Generate output filename ---
        input_filename = Path(input_file_path).stem
        if output_folder:
            output_path = os.path.join(output_folder, f"PCB_FORMAT_{input_filename}.csv")
        else:
            output_path = f"PCB_FORMAT_{input_filename}.csv"
        
        # --- Step 9: Export to CSV ---
        df_selected.to_csv(output_path, index=False)
        
        return True, output_path, None, df_selected
        
    except Exception as e:
        return False, None, str(e), None

def process_multiple_files(input_folder, output_folder=None, priority_value=None, skuno_value=None):
    """
    Process all Excel files in a folder
    
    Parameters:
    input_folder (str): Folder containing input Excel files
    output_folder (str): Folder for output CSV files (optional)
    priority_value (str): Common Priority value for all files
    skuno_value (str): Common SKUNo value for all files
    
    Returns:
    list: Results for each file processed
    """
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    excel_files = glob.glob(os.path.join(input_file_folder, "*.xlsx"))
    results = []
    
    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")
        success, output_path, error, df = process_pcb_file(
            file_path, output_folder, priority_value, skuno_value
        )
        
        results.append({
            'input_file': os.path.basename(file_path),
            'success': success,
            'output_file': os.path.basename(output_path) if output_path else None,
            'output_path': output_path,
            'error': error,
            'row_count': len(df) if df is not None else 0
        })
    
    return results

def main():
    """Main function for command line usage"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Bulk process PCB (Pushpam) Excel files to standardized format')
    parser.add_argument('--input', '-i', required=True, help='Input file or folder path')
    parser.add_argument('--output', '-o', help='Output folder path (optional)')
    parser.add_argument('--batch', '-b', action='store_true', help='Process all files in input folder')
    parser.add_argument('--priority', '-p', help='Priority value (TMP/SMP/REPSMP/REVSMP/SDPL/CUSCAD/MSP)')
    parser.add_argument('--skuno', '-s', help='SKUNo value for all rows')
    
    args = parser.parse_args()
    
    if args.batch:
        # Process all files in folder
        if not os.path.isdir(args.input):
            print("❌ Input path must be a folder when using --batch")
            return
        
        # Get common values if not provided
        if not args.priority:
            args.priority = input("Enter Priority value for all files (TMP/SMP/REPSMP/REVSMP/SDPL/CUSCAD/MSP): ")
        
        if not args.skuno:
            args.skuno = input("Enter SKUNo value for all files: ")
        
        results = process_multiple_files(
            args.input, args.output, args.priority, args.skuno
        )
        
        print("\n" + "="*50)
        print("BATCH PROCESSING RESULTS")
        print("="*50)
        
        success_count = 0
        for result in results:
            status = "✅ SUCCESS" if result['success'] else "❌ FAILED"
            print(f"{status}: {result['input_file']} -> {result['output_file'] or 'N/A'}")
            print(f"   Rows: {result['row_count']}")
            if result['error']:
                print(f"   Error: {result['error']}")
            print()
            
            if result['success']:
                success_count += 1
        
        print(f"Processed: {len(results)} files | Successful: {success_count} | Failed: {len(results) - success_count}")
        
    else:
        # Process single file
        if not os.path.isfile(args.input):
            print("❌ Input path must be a file when not using --batch")
            return
        
        success, output_path, error, df = process_pcb_file(
            args.input, args.output, args.priority, args.skuno
        )
        
        if success:
            print(f"✅ SUCCESS: Processed {args.input}")
            print(f"📊 Output: {output_path}")
            print(f"📋 Rows processed: {len(df)}")
            print(f"⏱️ Priority: {df['Priority'].iloc[0]}")
            print(f"🏷️ SKUNo: {df['SKUNo'].iloc[0]}")
            print("\nFirst 5 rows preview:")
            print(df.head())
        else:
            print(f"❌ FAILED: {error}")

if __name__ == "__main__":
    main()