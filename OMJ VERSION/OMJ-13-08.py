# OMJ_BULK.py - Bulk processing multiple OMJ CASTING PO files
import pandas as pd
import re
import os
import glob
from pathlib import Path


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'UP6.5', 'W') -> 'VR1943EEA-6.5WG'
         ('RG0002939A', 'EU56', 'W')  -> 'RG0002939A-56WG'
         ('AB123', '', 'PT')          -> 'AB123-PT'
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
    """Load and cache the approved Client Style No list from OMJ_100826.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    try:
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'OMJ_100826.xlsx')
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
    Match the built StyleCode against OMJ_100826.xlsx 'Client Style No'.

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

def process_omj_file(input_file_path, output_folder=None):
    """
    Process single OMJ CASTING PO Excel file and convert to standardized format
    
    Parameters:
    input_file_path (str): Path to input Excel file
    output_folder (str): Folder for output CSV file (optional)
    
    Returns:
    tuple: (success_status, output_path, error_message, dataframe)
    """
    try:
        # Step 1: Read the actual data starting from row 11 (skip first 10 rows)
        df = pd.read_excel(input_file_path, skiprows=10)

        # Step 2: Select required columns
        selected_columns = [
            'OMJ Style # ', 'Elegant Jewelry Style #', 'Quantity ',
            'Metal Type', 'Metal Color ', 'SIZE',
            'Shank Thickness', 'Shank Width', 'PO # '
        ]
        
        # Check if all required columns exist
        missing_columns = [col for col in selected_columns if col not in df.columns]
        if missing_columns:
            return False, None, f"Missing columns: {', '.join(missing_columns)}", None
        
        df_selected = df[selected_columns]

        # Step 3: Drop rows where 'OMJ Style # ' is missing
        df_cleaned = df_selected.dropna(subset=['OMJ Style # ']).copy()

        # Step 4: Convert numeric columns safely
        df_cleaned['Quantity '] = pd.to_numeric(df_cleaned['Quantity '], errors='coerce')
        df_cleaned = df_cleaned[df_cleaned['Quantity '].notna()]
        df_cleaned['Quantity '] = df_cleaned['Quantity '].astype(int)

        df_cleaned['SIZE'] = pd.to_numeric(df_cleaned['SIZE'], errors='coerce')
        df_cleaned = df_cleaned[df_cleaned['SIZE'].notna()]
        df_cleaned['SIZE'] = df_cleaned['SIZE'].astype(float)

        df_cleaned['Shank Thickness'] = pd.to_numeric(df_cleaned['Shank Thickness'], errors='coerce')
        df_cleaned = df_cleaned[df_cleaned['Shank Thickness'].notna()]
        df_cleaned['Shank Thickness'] = df_cleaned['Shank Thickness'].astype(float)

        df_cleaned['Shank Width'] = pd.to_numeric(df_cleaned['Shank Width'], errors='coerce')
        df_cleaned = df_cleaned[df_cleaned['Shank Width'].notna()]
        df_cleaned['Shank Width'] = df_cleaned['Shank Width'].astype(float)

        # Step 5: Rename columns
        df_cleaned.rename(columns={
            'OMJ Style # ': 'SKUNo',
            'Elegant Jewelry Style #': 'StyleCoderough',
            'Quantity ': 'OrderQty',
            'SIZE': 'ItemSize',
            'PO # ': 'ItemPoNo'
        }, inplace=True)

        # Step 6: Add SrNo. as the first column
        df_cleaned.insert(loc=0, column='SrNo.', value=range(1, len(df_cleaned) + 1))

        # Step 7: Add StyleCode column
        df_cleaned.insert(loc=1, column='StyleCode', value='')

        # Step 8: Create ItemSizeCopy with int if whole number else float rounded to 2 decimals
        def convert_item_size_copy(x):
            return int(x) if float(x).is_integer() else round(float(x), 2)

        df_cleaned.insert(
            loc=df_cleaned.columns.get_loc('ItemSize') + 1,
            column='ItemSizeCopy',
            value=df_cleaned['ItemSize'].map(convert_item_size_copy)
        )

        # Step 9: Format ItemSize as string with 'UP' prefix and leading zero if whole number
        def format_item_size(x):
            if float(x).is_integer():
                return f'UP{int(float(x)):02d}'
            else:
                return f'UP{float(x)}'

        df_cleaned['ItemSize'] = df_cleaned['ItemSize'].map(format_item_size)

        # Step 10: Reorder columns: StyleCode, ItemSize, ItemSizeCopy, OrderQty
        cols = df_cleaned.columns.tolist()
        cols.insert(2, cols.pop(cols.index('ItemSize')))
        cols.insert(3, cols.pop(cols.index('ItemSizeCopy')))
        cols.insert(4, cols.pop(cols.index('OrderQty')))
        df_cleaned = df_cleaned[cols]

        # Step 11: Add OrderItemPcs, Metal, Tone columns
        orderqty_index = df_cleaned.columns.get_loc('OrderQty')
        df_cleaned.insert(loc=orderqty_index + 1, column='OrderItemPcs', value=1)
        df_cleaned.insert(loc=orderqty_index + 2, column='Metal', value='')
        df_cleaned.insert(loc=orderqty_index + 3, column='Tone', value='')

        # Step 12: Fill 'Tone' column with first letter of 'Metal Color '
        df_cleaned['Tone'] = df_cleaned['Metal Color '].astype(str).str.strip().str[0]

        # Step 13: Fill 'Metal' as 'G' + digits before 'K' in Metal Type + Tone
        metal_karat = df_cleaned['Metal Type'].astype(str).str.extract(r'(?i)(\d+)(?=k)', expand=False).fillna('')
        df_cleaned['Metal'] = 'G' + metal_karat + df_cleaned['Tone']

        # Step 14: Move 'ItemPoNo' column to right after 'Tone'
        cols = df_cleaned.columns.tolist()
        tone_index = cols.index('Tone')
        itempono_col = cols.pop(cols.index('ItemPoNo'))
        cols.insert(tone_index + 1, itempono_col)
        df_cleaned = df_cleaned[cols]

        # Step 15: Add additional columns after ItemPoNo
        itempono_index = df_cleaned.columns.get_loc('ItemPoNo')
        additional_after_pono = [
            'ItemRefNo', 'StockType', 'MakeType', 'CustomerProductionInstruction',
            'SpecialRemarks', 'DesignProductionInstruction', 'StampInstruction',
            'OrderGroup', 'Certificate'
        ]
        for i, col in enumerate(additional_after_pono):
            df_cleaned.insert(loc=itempono_index + 1 + i, column=col, value='')

        # Step 16: Add columns after SKUNo
        sku_index = df_cleaned.columns.get_loc('SKUNo')
        additional_after_sku = [
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate', 'SetPrice', 'StoneQuality'
        ]
        for i, col in enumerate(additional_after_sku):
            df_cleaned.insert(loc=sku_index + 1 + i, column=col, value='')

        # Format ItemSizeCopy for output (remove trailing .00)
        def format_no_trailing_zeros(x):
            if float(x).is_integer():
                return str(int(x))
            else:
                return str(x)

        df_cleaned['ItemSizeCopy'] = df_cleaned['ItemSizeCopy'].map(format_no_trailing_zeros)

        # Step 17: Create StyleCode  e.g. 'YR4172SA-7YG'
        df_cleaned['ItemSizeCopy'] = df_cleaned['ItemSizeCopy'].apply(_map_item_size)
        df_cleaned['StyleCode'] = df_cleaned.apply(
            lambda row: _build_style_code(
                str(row['StyleCoderough']).split('-')[0].strip(),
                str(row['ItemSizeCopy']),
                'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])
            ), axis=1
        )
        df_cleaned['StyleCode'] = df_cleaned['StyleCode'].apply(_lookup_client_style)
        df_cleaned.loc[df_cleaned['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''
        df_cleaned.drop(columns=['StyleCoderough'], inplace=True)

        # Step 18: Fill 'CustomerProductionInstruction'
        def generate_instruction(thickness, width):
            return (
                "PLEASE MAKE SURE THERE IS NO POROSITY , PLEASE MAKE SURE THERE IS NO BLACK NO MILKY STONES  "
                "NO COLOR TING IN THE DIAMONDS, THE DIAMOND SHANK THICKNESS "
                f"{round(thickness, 2)} & {round(width, 2)} WIDTH WITH +/-0.1MM TOLERANCE"
            )

        df_cleaned['CustomerProductionInstruction'] = df_cleaned.apply(
            lambda row: generate_instruction(row['Shank Thickness'], row['Shank Width']),
            axis=1
        )

        # Step 19: Fill 'SpecialRemarks'
        def generate_special_remarks(sku, metal_type, metal_color, size_copy):
            sku = str(sku).strip()
            metal_type = str(metal_type).strip().upper()
            metal_color = str(metal_color).strip().upper()
            size_copy = str(size_copy).strip()
            return f"{sku}, {metal_type}, {metal_color}, Size {size_copy}, DIA QLTY"

        df_cleaned['SpecialRemarks'] = df_cleaned.apply(
            lambda row: generate_special_remarks(
                row['SKUNo'], row['Metal Type'], row['Metal Color '], row['ItemSizeCopy']
            ),
            axis=1
        )

        # Step 20: Fill 'StampInstruction'
        def generate_stamp_instruction(metal_type, size_copy, sku):
            metal_type = str(metal_type).strip().upper()
            size_copy = str(size_copy).strip()
            sku = str(sku).strip().upper()
            return f"OMJ LOGO, {metal_type}, {size_copy}, {sku}, E"

        df_cleaned['StampInstruction'] = df_cleaned.apply(
            lambda row: generate_stamp_instruction(
                row['Metal Type'], row['ItemSizeCopy'], row['SKUNo']
            ),
            axis=1
        )

        # Step 21: Drop unnecessary columns before export
        df_cleaned.drop(
            columns=['ItemSizeCopy', 'Metal Type', 'Metal Color ', 'Shank Thickness', 'Shank Width'],
            inplace=True
        )

        # Step 22: Generate output filename
        input_filename = Path(input_file_path).stem
        if output_folder:
            output_path = os.path.join(output_folder, f"OMJ_FORMAT_{input_filename}.xlsx")
        else:
            output_path = f"OMJ_FORMAT_{input_filename}.xlsx"
        
        # Step 23: Export to CSV
        df_cleaned.to_excel(output_path, index=False)
        
        return True, output_path, None, df_cleaned
        
    except Exception as e:
        return False, None, str(e), None

def process_multiple_files(input_folder, output_folder=None):
    """
    Process all Excel files in a folder
    
    Parameters:
    input_folder (str): Folder containing input Excel files
    output_folder (str): Folder for output CSV files (optional)
    
    Returns:
    list: Results for each file processed
    """
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx"))
    results = []
    
    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")
        success, output_path, error, df = process_omj_file(file_path, output_folder)
        
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
    
    parser = argparse.ArgumentParser(description='Bulk process OMJ CASTING PO Excel files to standardized format')
    parser.add_argument('--input', '-i', required=True, help='Input file or folder path')
    parser.add_argument('--output', '-o', help='Output folder path (optional)')
    parser.add_argument('--batch', '-b', action='store_true', help='Process all files in input folder')
    
    args = parser.parse_args()
    
    if args.batch:
        # Process all files in folder
        if not os.path.isdir(args.input):
            print("❌ Input path must be a folder when using --batch")
            return
        
        results = process_multiple_files(args.input, args.output)
        
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
        
        success, output_path, error, df = process_omj_file(args.input, args.output)
        
        if success:
            print(f"✅ SUCCESS: Processed {args.input}")
            print(f"📊 Output: {output_path}")
            print(f"📋 Rows processed: {len(df)}")
            print("\nFirst 5 rows preview:")
            print(df.head())
        else:
            print(f"❌ FAILED: {error}")

if __name__ == "__main__":
    main()