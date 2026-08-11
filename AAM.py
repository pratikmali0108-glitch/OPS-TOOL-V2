# AAM_BULK.py - Bulk processing multiple AAM (Aurum Adams) PO files
import pandas as pd
import re
import os
import glob
from pathlib import Path


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'UP6.5', 'W') -> 'VR1943EEA-6.5WG'
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

def process_aam_file(input_file_path, output_folder=None, priority_value=None):
    """
    Process single AAM (Aurum Adams) Excel file and convert to standardized format
    
    Parameters:
    input_file_path (str): Path to input Excel file
    output_folder (str): Folder for output Excel file (optional)
    priority_value (str): Priority value for all rows
    
    Returns:
    tuple: (success_status, output_path, error_message, dataframe)
    """
    try:
        # --- Step 1: Read Excel and skip top 7 rows ---
        df = pd.read_excel(input_file_path, skiprows=7)

        # --- Step 1a: Read cell B2 for ItemPoNo ---
        item_po_no = pd.read_excel(input_file_path, header=None, nrows=2).iloc[1, 1]  # B2

        # --- Step 2: Select specific columns ---
        selected_columns = ['AURUM Code', 'REF SHIMAYRA', 'KT', 'SIZE', 'COLOUR', 'QUALITY', 'QTY']
        
        # Check if all required columns exist
        missing_columns = [col for col in selected_columns if col not in df.columns]
        if missing_columns:
            return False, None, f"Missing columns: {', '.join(missing_columns)}", None
        
        df_selected = df[selected_columns]

        # --- Step 3: Drop rows with NaN and make a copy ---
        df_cleaned = df_selected.dropna().copy()

        # --- Step 4: Rename columns ---
        df_cleaned.rename(columns={
            'AURUM Code': 'SKUNo',
            'REF SHIMAYRA': 'StyleCode',
            'KT': 'Metal',
            'SIZE': 'ItemSize',
            'COLOUR': 'Tone',
            'QTY': 'OrderQty',
        }, inplace=True)

        # --- Step 4b: Clean StyleCode ---
        def clean_style_code(text):
            if pd.isna(text):
                return ''
            text = str(text).strip().upper()
            text = re.split(r'\s*\+\s*|\s+CHAIN|\s+YG|\s+WG|\s+RG', text)[0]
            text = re.sub(r'[-_/]+$', '', text)
            return text.strip()

        df_cleaned['StyleCode'] = df_cleaned['StyleCode'].apply(clean_style_code)

        # --- Step 5: Transform ItemSize values ---
        def transform_itemsize(x):
            if pd.isna(x):
                return ''
            x = str(x).strip().upper()
            if x == 'NS':
                return ''
            if 'MM' in x:
                num_part = ''.join([ch for ch in x if ch.isdigit()])
                return f'UP{num_part}' if num_part else ''
            return x

        df_cleaned['ItemSize'] = df_cleaned['ItemSize'].apply(transform_itemsize)

        # --- Step 6: Metal mapping logic ---
        def map_metal_tone(metal, tone):
            if pd.isna(metal) or pd.isna(tone):
                return ''
            metal = str(metal).upper()
            tone = str(tone).upper()

            if '18' in metal:
                if 'YELLOW' in tone:
                    return 'G750Y'
                elif 'WHITE' in tone:
                    return 'G750W'
                elif 'ROSE PINK' in tone or 'PINK' in tone:
                    return 'G750P'

            if '14' in metal:
                if 'YELLOW' in tone:
                    return 'G585Y'
                elif 'WHITE' in tone:
                    return 'G585W'
                elif 'ROSE PINK' in tone or 'PINK' in tone:
                    return 'G585P'

            if 'PLATINUM' in metal:
                return 'PC95'

            return metal

        df_cleaned['Metal'] = df_cleaned.apply(lambda row: map_metal_tone(row['Metal'], row['Tone']), axis=1)

        # --- Step 7: Update Tone column based on Metal postfix ---
        def update_tone_from_metal(metal):
            if pd.isna(metal):
                return ''
            metal = str(metal).upper()
            if metal.endswith('Y'):
                return 'Y'
            elif metal.endswith('W'):
                return 'W'
            elif metal.endswith('P'):
                return 'P'
            elif 'PC95' in metal:
                return 'PT'
            return ''

        df_cleaned['Tone'] = df_cleaned['Metal'].apply(update_tone_from_metal)

        # --- Step 8: Add OrderItemPcs after OrderQty ---
        df_cleaned['OrderItemPcs'] = ''

        # --- Step 9: Add ItemPoNo ---
        df_cleaned['ItemPoNo'] = item_po_no

        # --- Step 10: Add ItemRefNo, StockType, Priority, MakeType ---
        if priority_value is None:
            priority_value = input(f"Enter Priority for {Path(input_file_path).name} (e.g., URGENT, NORMAL): ")
        
        df_cleaned['ItemRefNo'] = ''
        df_cleaned['StockType'] = ''
        df_cleaned['Priority'] = priority_value
        df_cleaned['MakeType'] = ''

        # --- Step 11: CustomerProductionInstruction based on Metal ---
        def map_make_type(metal):
            mapping = {
                'G750W': '18K WHITE GOLD(OR JALINE 750)',
                'G750Y': '18K YELLOW GOLD(OR JALINE 750)',
                'G750P': '18K PINK GOLD(OR JALINE 750)',
                'G585W': '14K WHITE GOLD(OR JALINE 585)',
                'G585Y': '14K YELLOW GOLD(OR JALINE 585)',
                'G585P': '14K PINK GOLD(OR JALINE 585)',
                'PC95': 'PLATINUM(OR JALINE 95)'
            }
            return mapping.get(metal, '')

        df_cleaned['CustomerProductionInstruction'] = df_cleaned['Metal'].apply(map_make_type)

        # --- Step 12: SpecialRemarks ---
        def create_special_remarks(row):
            sku = row['SKUNo']
            metal_desc_map = {
                'G750W': '750 WHITE GOLD',
                'G750Y': '750 YELLOW GOLD',
                'G750P': '750 PINK GOLD',
                'G585W': '585 WHITE GOLD',
                'G585Y': '585 YELLOW GOLD',
                'G585P': '585 PINK GOLD',
                'PC95': '95 PLATINUM'
            }
            metal_desc = metal_desc_map.get(row['Metal'], '')

            size_desc = ''
            if str(row['ItemSize']).upper().startswith('UP'):
                num = ''.join([c for c in row['ItemSize'] if c.isdigit()])
                if num:
                    size_desc = f'SZ-{num}'

            quality = row['QUALITY'] if pd.notna(row['QUALITY']) else ''
            parts = [sku]
            if metal_desc:
                parts.append(metal_desc)
            if size_desc:
                parts.append(size_desc)
            if quality:
                parts.append(f'DIA QLTY: {quality}')
            return ', '.join(parts)

        df_cleaned['SpecialRemarks'] = df_cleaned.apply(create_special_remarks, axis=1)

        # --- Step 13: DesignProductionInstruction ---
        def map_design_instruction(tone):
            if pd.isna(tone):
                return ''
            tone = str(tone).upper()
            if tone == 'W':
                return 'WHITE RODIUM'
            elif tone in ['Y', 'P', 'PT']:
                return 'NO RODIUM'
            return ''

        df_cleaned['DesignProductionInstruction'] = df_cleaned['Tone'].apply(map_design_instruction)

        # --- Step 14: Stamp Instruction ---
        def create_stamp_instruction(row):
            special = str(row['SpecialRemarks']).upper()
            if '750' in special:
                metal_number = '750'
            elif '585' in special:
                metal_number = '585'
            elif '95' in special:
                metal_number = '95'
            else:
                metal_number = ''
            
            if metal_number:
                return f'CUSTOMER LOGO & {metal_number}'
            return ''

        df_cleaned['StampInstruction'] = df_cleaned.apply(create_stamp_instruction, axis=1)

        # --- Step 15: Add OrderGroup, Certificate ---
        df_cleaned['OrderGroup'] = ''
        df_cleaned['Certificate'] = ''

        # --- Step 16: Add new columns after SKUNo ---
        new_cols_after_sku = [
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 
            'Basemetalmaxwt', 'Productiondeliverydate', 'Expecteddeliverydate', 'ClientStyleNo.'
        ]
        for col in new_cols_after_sku:
            df_cleaned[col] = ''

        # Build full StyleCode: base-sizeWG / base-sizePT etc.
        df_cleaned['ItemSize'] = df_cleaned['ItemSize'].apply(_map_item_size)
        df_cleaned['StyleCode'] = df_cleaned.apply(
            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
        )
        df_cleaned.loc[df_cleaned['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        # --- Step 17: ClientStyleNo. mirrors StyleCode (already full format) ---
        df_cleaned['ClientStyleNo.'] = df_cleaned['StyleCode']

        # --- Step 18: Add SrNo in incremental order by StyleCode ---
        df_cleaned = df_cleaned.sort_values(by=['StyleCode']).reset_index(drop=True)
        df_cleaned.insert(0, 'SrNo', range(1, len(df_cleaned) + 1))

        # --- Step 19: Reorder columns for final export ---
        column_order = [
            'SrNo', 'StyleCode', 'ItemSize', 'OrderQty', 'OrderItemPcs',
            'Metal', 'Tone', 'ItemPoNo', 'ItemRefNo', 'StockType', 'Priority',
            'MakeType', 'CustomerProductionInstruction', 'SpecialRemarks', 
            'DesignProductionInstruction', 'StampInstruction', 'OrderGroup', 'Certificate',
            'SKUNo', 
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 
            'Basemetalmaxwt', 'Productiondeliverydate', 'Expecteddeliverydate',
            'QUALITY', 'ClientStyleNo.'
        ]

        df_cleaned = df_cleaned[column_order]
        df_cleaned.drop(columns=['QUALITY'], inplace=True)

        # --- Step 20: Generate output filename ---
        input_filename = Path(input_file_path).stem
        if output_folder:
            output_path = os.path.join(output_folder, f"AAM_FORMAT_{input_filename}.xlsx")
        else:
            output_path = f"AAM_FORMAT_{input_filename}.xlsx"
        
        # --- Step 21: Export to Excel ---
        df_cleaned.to_excel(output_path, index=False)
        
        return True, output_path, None, df_cleaned
        
    except Exception as e:
        return False, None, str(e), None

def process_multiple_files(input_folder, output_folder=None, priority_value=None):
    """
    Process all Excel files in a folder
    
    Parameters:
    input_folder (str): Folder containing input Excel files
    output_folder (str): Folder for output Excel files (optional)
    priority_value (str): Common Priority value for all files
    
    Returns:
    list: Results for each file processed
    """
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx")) + glob.glob(os.path.join(input_folder, "*.xls"))
    results = []
    
    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")
        success, output_path, error, df = process_aam_file(
            file_path, output_folder, priority_value
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
    
    parser = argparse.ArgumentParser(description='Bulk process AAM (Aurum Adams) PO Excel files to standardized format')
    parser.add_argument('--input', '-i', required=True, help='Input file or folder path')
    parser.add_argument('--output', '-o', help='Output folder path (optional)')
    parser.add_argument('--batch', '-b', action='store_true', help='Process all files in input folder')
    parser.add_argument('--priority', '-p', help='Priority value (e.g., URGENT, NORMAL)')
    
    args = parser.parse_args()
    
    if args.batch:
        # Process all files in folder
        if not os.path.isdir(args.input):
            print("❌ Input path must be a folder when using --batch")
            return
        
        # Get common Priority value if not provided
        if not args.priority:
            args.priority = input("Enter Priority value for all files (e.g., URGENT, NORMAL): ")
        
        results = process_multiple_files(args.input, args.output, args.priority)
        
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
        
        success, output_path, error, df = process_aam_file(args.input, args.output, args.priority)
        
        if success:
            print(f"✅ SUCCESS: Processed {args.input}")
            print(f"📊 Output: {output_path}")
            print(f"📋 Rows processed: {len(df)}")
            print(f"🏷️ ItemPoNo: {df['ItemPoNo'].iloc[0]}")
            print(f"⏱️ Priority: {df['Priority'].iloc[0]}")
            print("\nFirst 5 rows preview:")
            print(df.head())
        else:
            print(f"❌ FAILED: {error}")

if __name__ == "__main__":
    main()