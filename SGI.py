# SGI_BULK.py - Bulk processing multiple SGI Customer files
import pandas as pd
import re
import os
import glob
from pathlib import Path


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'US07', 'W') -> 'VR1943EEA-7WG'
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
    """Load and cache the approved Client Style No list from SGI_CS_100826.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    try:
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'SGI_CS_100826.xlsx')
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
    Match the built StyleCode against SGI_CS_100826.xlsx 'Client Style No'.

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

def process_sgi_file(input_file_path, output_folder=None, cust_order_no=None):
    """
    Process single SGI Customer Excel file and convert to standardized format
    
    Parameters:
    input_file_path (str): Path to input Excel file
    output_folder (str): Folder for output CSV file (optional)
    cust_order_no (str): CUSTOMER ORDER NO. value
    
    Returns:
    tuple: (success_status, output_path, error_message, dataframe)
    """
    try:
        # --- Step 1: Read Excel ---
        df = pd.read_excel(input_file_path)

        # --- Step 2: Select and rename columns ---
        selected_columns = ['SR NO.', 'Vendor Item', 'Size', 'Qty', 'Metal', 'Order #', 'rhodium DETAILS ']
        
        # Check if all required columns exist
        missing_columns = [col for col in selected_columns if col not in df.columns]
        if missing_columns:
            return False, None, f"Missing columns: {', '.join(missing_columns)}", None
        
        df_selected = df[selected_columns].copy()

        df_selected.rename(columns={
            'SR NO.': 'SrNo',
            'Vendor Item': 'StyleCode',
            'Size': 'ItemSize',
            'Qty': 'OrderQty',
            'Metal': 'MetalRough',
            'Order #': 'ItemPoNo'
        }, inplace=True)

        # --- Step 3: Use provided CUSTOMER ORDER NO. or prompt ---
        if cust_order_no is None:
            cust_order_no = input(f"Enter CUSTOMER ORDER NO. value for {Path(input_file_path).name}: ")

        # --- Step 4: Clean and insert columns ---
        df_selected['ItemSize'] = df_selected['ItemSize'].fillna("")
        orderqty_index = df_selected.columns.get_loc('OrderQty')

        df_selected.insert(orderqty_index + 1, 'OrderItemPcs', 1)
        df_selected.insert(orderqty_index + 2, 'Metal', '')
        df_selected.insert(orderqty_index + 3, 'Tone', '')

        # --- Step 4B: Normalize ItemSize values ---
        def normalize_size(size):
            if pd.isna(size):
                return ''
            s = str(size).upper().strip()
            match = re.search(r'US\s*(\d+(\.\d+)?)', s)  # matches US 7 or US 6.5
            if match:
                num = match.group(1)
                if '.' in num:
                    return f"US{num.zfill(4)}"  # keeps decimals like US06.5
                else:
                    return f"US{num.zfill(2)}"   # pads to 2 digits → US07, US18
            return s

        df_selected['ItemSize'] = df_selected['ItemSize'].apply(normalize_size)

        # --- Step 5: Extract Metal (numerical part) ---
        df_selected['Metal'] = df_selected['MetalRough'].apply(
            lambda x: re.search(r'\d+', str(x)).group() if pd.notna(x) and re.search(r'\d+', str(x)) else ''
        )

        # --- Step 5B: Extract Tone (Y, W, P combinations) ---
        def extract_tone(metal_rough):
            if pd.isna(metal_rough):
                return ''
            s = str(metal_rough).upper()
            s = re.sub(r'[^A-Z]', '', s)
            tone = ''
            if 'Y' in s:
                tone += 'Y'
            if 'P' in s or 'R' in s:
                tone += 'P'
            if 'W' in s:
                tone += 'W'
            return tone

        df_selected['Tone'] = df_selected['MetalRough'].apply(extract_tone)

        # --- Step 5C: Final Metal mapping (G14Y, G18W, PC95, etc.) ---
        def map_metal(metal_rough, metal, tone):
            if pd.isna(metal_rough):
                return ''
            metal_rough = str(metal_rough).upper()
            # Platinum check first
            if 'PLATINUM' in metal_rough or 'PT' in metal_rough:
                return 'PC95'
            # Gold logic
            if metal in ['14', '18']:
                tone = tone or ''
                return f'G{metal}{tone}'
            return ''

        df_selected['Metal'] = df_selected.apply(lambda row: map_metal(row['MetalRough'], row['Metal'], row['Tone']), axis=1)

        # --- Step 6: Insert new columns after ItemPoNo ---
        new_columns = [
            'ItemRefNo', 'StockType', 'MakeType', 'CustomerProductionInstruction',
            'SpecialRemarks', 'DesignProductionInstruction', 'StampInstruction',
            'OrderGroup', 'Certificate', 'SKUNo', 'Basestoneminwt', 'Basestonemaxwt',
            'Basemetalminwt', 'Basemetalmaxwt', 'Productiondeliverydate',
            'Expecteddeliverydate', 'SetPrice', 'StoneQuality', 'CUSTOMER ORDER NO.'
        ]

        itempono_index = df_selected.columns.get_loc('ItemPoNo')
        for i, col in enumerate(new_columns, start=1):
            df_selected.insert(loc=itempono_index + i, column=col, value='')

        # Fill CUSTOMER ORDER NO.
        df_selected['CUSTOMER ORDER NO.'] = cust_order_no

        # --- Step 7: Fill CustomerProductionInstruction ---
        def get_instruction(metal_rough):
            if pd.isna(metal_rough):
                return ''
            metal_rough = str(metal_rough).upper()
            mappings = {
                '14YG': '14 K YELLOW GOLD',
                '14WG': '14 K WHITE GOLD',
                '18YG': '18 K YELLOW GOLD',
                '18WG': '18 K WHITE GOLD',
                '10YG': '10 K YELLOW GOLD',
                '10WG': '10 K WHITE GOLD',
                '14RG': '14 K ROSE GOLD',
                '18RG': '18 K ROSE GOLD',
                '10PG': '10 K ROSE GOLD',
                '18PG': '18 K ROSE GOLD',
                'PT': 'PLATINUM',
                'PLATINUM': 'PLATINUM'
            }
            for key, value in mappings.items():
                if key in metal_rough:
                    return value
            return ''

        df_selected['CustomerProductionInstruction'] = df_selected['MetalRough'].apply(get_instruction)

        # --- Step 8: Fill SpecialRemarks ---
        df_selected['SpecialRemarks'] = df_selected['CustomerProductionInstruction'].apply(
            lambda x: f"{x}, DIA QLTY CLIENT PURCHASE" if x else ''
        )

        # --- Step 9: Fill DesignProductionInstruction ---
        def get_design_instruction(rhodium):
            if pd.isna(rhodium) or not str(rhodium).strip():
                base = "NO RHODIUM"
            else:
                match = re.search(r'(WHITE|YELLOW|ROSE)?\s*RHODIUM', str(rhodium).upper())
                base = match.group() if match else str(rhodium).strip().upper()
            return f"{base} these styles will take actual diamond wt. stamp on the pieces"

        df_selected['DesignProductionInstruction'] = df_selected['rhodium DETAILS '].apply(get_design_instruction)

        # --- Step 10: Fill StampInstruction ---
        df_selected['StampInstruction'] = df_selected['Metal'].apply(
            lambda x: f"{x} + SMS + ACTUAL DIA WT" if x else ''
        )

        # Build full StyleCode: base-sizeWG / base-sizePT etc.
        df_selected['ItemSize'] = df_selected['ItemSize'].apply(_map_item_size)
        df_selected['StyleCode'] = df_selected.apply(
            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
        )
        df_selected['StyleCode'] = df_selected['StyleCode'].apply(_lookup_client_style)
        df_selected.loc[df_selected['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        # --- Step 11: Clean up ---
        df_selected.drop(columns=['MetalRough', 'rhodium DETAILS '], inplace=True)

        # --- Step 12: Generate output filename ---
        input_filename = Path(input_file_path).stem
        if output_folder:
            output_path = os.path.join(output_folder, f"SGI_FORMAT_{input_filename}.csv")
        else:
            output_path = f"SGI_FORMAT_{input_filename}.csv"
        
        # --- Step 13: Export to CSV ---
        df_selected.to_csv(output_path, index=False)
        
        return True, output_path, None, df_selected
        
    except Exception as e:
        return False, None, str(e), None

def process_multiple_files(input_folder, output_folder=None, cust_order_no=None):
    """
    Process all Excel files in a folder
    
    Parameters:
    input_folder (str): Folder containing input Excel files
    output_folder (str): Folder for output CSV files (optional)
    cust_order_no (str): Common CUSTOMER ORDER NO. value for all files
    
    Returns:
    list: Results for each file processed
    """
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx"))
    results = []
    
    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")
        success, output_path, error, df = process_sgi_file(file_path, output_folder, cust_order_no)
        
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
    
    parser = argparse.ArgumentParser(description='Bulk process SGI Customer Excel files to standardized format')
    parser.add_argument('--input', '-i', required=True, help='Input file or folder path')
    parser.add_argument('--output', '-o', help='Output folder path (optional)')
    parser.add_argument('--batch', '-b', action='store_true', help='Process all files in input folder')
    parser.add_argument('--order-no', '-n', help='CUSTOMER ORDER NO. value (optional)')
    
    args = parser.parse_args()
    
    if args.batch:
        # Process all files in folder
        if not os.path.isdir(args.input):
            print("❌ Input path must be a folder when using --batch")
            return
        
        # Get CUSTOMER ORDER NO. if not provided
        if not args.order_no:
            args.order_no = input("Enter CUSTOMER ORDER NO. value for all files: ")
        
        results = process_multiple_files(args.input, args.output, args.order_no)
        
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
        
        success, output_path, error, df = process_sgi_file(args.input, args.output, args.order_no)
        
        if success:
            print(f"✅ SUCCESS: Processed {args.input}")
            print(f"📊 Output: {output_path}")
            print(f"📋 Rows processed: {len(df)}")
            print(f"📦 CUSTOMER ORDER NO.: {df['CUSTOMER ORDER NO.'].iloc[0]}")
            print("\nFirst 5 rows preview:")
            print(df.head())
        else:
            print(f"❌ FAILED: {error}")

if __name__ == "__main__":
    main()