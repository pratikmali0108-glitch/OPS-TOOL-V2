# VIMCO_BULK.py - Bulk processing multiple VIMCO PO files
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

def process_vimco_file(input_file_path, output_folder=None, item_po_no=None, order_group=None, priority_value=None):
    """
    Process single VIMCO PO Excel file and convert to standardized format
    
    Parameters:
    input_file_path (str): Path to input Excel file
    output_folder (str): Folder for output CSV file (optional)
    item_po_no (str): ItemPoNo number
    order_group (str): OrderGroup value
    priority_value (str): Priority value (default: "5 day")
    
    Returns:
    tuple: (success_status, output_path, error_message, dataframe)
    """
    try:
        # --- Step 1: Read Excel ---
        df = pd.read_excel(input_file_path)

        # --- Step 2: Select relevant columns ---
        selected_columns = ['#', 'Item #', 'Vendor Item #', 'Description', 'Size', 'Qty.', 'Dia Qlty']
        
        # Check if all required columns exist
        missing_columns = [col for col in selected_columns if col not in df.columns]
        if missing_columns:
            return False, None, f"Missing columns: {', '.join(missing_columns)}", None
        
        df = df[selected_columns].copy()

        # --- Step 3: Rename columns ---
        df.rename(columns={
            '#': 'SrNo',
            'Item #': 'SKUNo',
            'Vendor Item #': 'StyleCode',
            'Description': 'CustomerProductionInstruction',
            'Size': 'ItemSize',
            'Qty.': 'OrderQty',
            'Dia Qlty': 'DiaQuality'
        }, inplace=True)

        # --- Clean up StyleCode to remove everything after '-' ---
        df['StyleCode'] = df['StyleCode'].astype(str).str.split('-').str[0].str.strip()

        # --- Step 4: Get user input if not provided ---
        if item_po_no is None:
            item_po_no = input(f"Enter ItemPoNo number for {Path(input_file_path).name}: ")
        
        if order_group is None:
            order_group = input(f"Enter OrderGroup value for {Path(input_file_path).name}: ")
        
        if priority_value is None:
            user_priority = input(f"Enter Priority for {Path(input_file_path).name} (press Enter to use default '5 day'): ")
            priority_value = user_priority if user_priority.strip() != "" else "5 day"

        # --- Step 5: Functions for calculated columns ---
        def map_metal(text):
            if pd.isna(text):
                return ''
            text = str(text).upper()
            if '14KY' in text:
                return 'G14Y'
            elif '14KW' in text:
                return 'G14W'
            elif '18KY' in text:
                return 'G18Y'
            elif '18KW' in text:
                return 'G18W'
            elif '10KY' in text:
                return 'G10Y'
            elif '10KW' in text:
                return 'G10W'
            elif 'PT' in text:
                return 'PC95'
            else:
                return ''

        def map_tone(metal):
            if metal == 'PC95':
                return 'PT'
            elif metal:
                return metal[-1]
            else:
                return ''

        def metal_text(metal):
            mapping = {
                "G14Y": "14K YELLOW GOLD",
                "G14W": "14K WHITE GOLD",
                "G18Y": "18K YELLOW GOLD",
                "G18W": "18K WHITE GOLD",
                "G10Y": "10K YELLOW GOLD",
                "G10W": "10K WHITE GOLD",
                "PC95": "PLATINUM"
            }
            return mapping.get(metal, "")

        def special_remarks(row):
            metal_descr = metal_text(row['Metal'])
            dia_quality = str(row['DiaQuality']).strip().upper() if pd.notna(row['DiaQuality']) else ""
            return f"{row['OrderGroup']},{row['SKUNo']},{metal_descr}, DIA QLTY - {dia_quality}"

        def design_production_instruction(row):
            tone = str(row['Tone']).upper()
            cpi = str(row['CustomerProductionInstruction']).upper() if pd.notna(row['CustomerProductionInstruction']) else ''
            semi_present = "SEMI" in cpi

            if tone == "W" and semi_present:
                return "SEMI MOUNT, WHITE RODIUM"
            elif tone == "Y" and semi_present:
                return "SEMI MOUNT, NO RODIUM"
            elif tone == "W" and not semi_present:
                return "WHITE RODIUM"
            elif tone == "Y" and not semi_present:
                return "NO RODIUM"
            elif tone == "PT" and semi_present:
                return "SEMI MOUNT, NO RODIUM"
            elif tone == "PT" and not semi_present:
                return "NO RODIUM"
            else:
                return ""

        # --- New helper functions for StampInstruction ---
        def extract_stone_weight(text):
            if pd.isna(text):
                return ""
            text = str(text)
            match = re.search(r'(\d+\.\d+|\d+)\s*(CT|CARAT|CTS|CTW|A|PCS)?', text, re.IGNORECASE)
            if match:
                return match.group(1)
            return ""

        def metal_stamp_text(metal):
            mapping = {
                "G14Y": "14K",
                "G14W": "14K",
                "G18Y": "18K",
                "G18W": "18K",
                "G10Y": "10K",
                "G10W": "10K",
                "PC95": "PT950"
            }
            return mapping.get(metal, "")

        def generate_stamp_instruction(row):
            metal_text_val = metal_stamp_text(row['Metal'])
            stone_weight = extract_stone_weight(row['CustomerProductionInstruction'])
            if metal_text_val and stone_weight:
                return f"{metal_text_val} V ON ONE SIDE AND {stone_weight} A ON OTHER SIDE"
            elif metal_text_val:
                return f"{metal_text_val} V ON ONE SIDE"
            else:
                return ""

        # --- Step 6: Create all new columns ---
        df['OrderItemPcs'] = 1
        df['Metal'] = df['CustomerProductionInstruction'].apply(map_metal)
        df['Tone'] = df['Metal'].apply(map_tone)
        df['ItemPoNo'] = item_po_no
        df['ItemRefNo'] = ''
        df['StockType'] = ''
        df['Priority'] = priority_value
        df['MakeType'] = ''
        df['StampInstruction'] = df.apply(generate_stamp_instruction, axis=1)
        df['OrderGroup'] = order_group
        df['Certificate'] = '' 
        df['SpecialRemarks'] = df.apply(special_remarks, axis=1)
        df['DesignProductionInstruction'] = df.apply(design_production_instruction, axis=1)

        # --- Step 6c: Add new columns after SKU ---
        new_cols_after_sku = [
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate', 'SetPrice', 'StoneQuality'
        ]
        for col in new_cols_after_sku:
            df[col] = ''

        # --- ZR style default ItemSize = 7 ---
        df.loc[df['StyleCode'].astype(str).str.upper().str.startswith('ZR'), 'ItemSize'] = 7

        # --- Extract size from SKU if it ends with SZ9, SZ6.5, etc. ---
        def extract_size_from_sku(sku):
            if pd.isna(sku):
                return None
            sku = str(sku).upper()
            match = re.search(r'SZ(\d+(\.\d+)?)$', sku)
            if match:
                return match.group(1)
            return None

        df['ExtractedSize'] = df['SKUNo'].apply(extract_size_from_sku)
        # Fill missing ItemSize with extracted size
        mask = df['ItemSize'].isna() & df['ExtractedSize'].notna()
        df.loc[mask, 'ItemSize'] = df.loc[mask, 'ExtractedSize']
        
        # Format ItemSize with US prefix
        df['ItemSize'] = df['ItemSize'].apply(lambda x: f"US{x}" if pd.notna(x) and str(x).strip() != "" else "")

        df.drop(columns=['ExtractedSize'], inplace=True)

        # Build full StyleCode: base-sizeWG / base-sizePT etc.
        df['ItemSize'] = df['ItemSize'].apply(_map_item_size)
        df['StyleCode'] = df.apply(
            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
        )
        df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        # --- Step 7: Reorder columns ---
        final_columns = [
            'SrNo', 'StyleCode', 'ItemSize', 'OrderQty', 'OrderItemPcs', 'Metal', 'Tone',
            'ItemPoNo', 'ItemRefNo', 'StockType', 'Priority', 'MakeType',
            'CustomerProductionInstruction', 'SpecialRemarks', 'DesignProductionInstruction',
            'StampInstruction', 'OrderGroup', 'Certificate', 'SKUNo'
        ] + new_cols_after_sku
        df = df[final_columns]

        # --- Step 8: Generate output filename ---
        input_filename = Path(input_file_path).stem
        if output_folder:
            output_path = os.path.join(output_folder, f"VIMCO_FORMAT_{input_filename}.csv")
        else:
            output_path = f"VIMCO_FORMAT_{input_filename}.csv"
        
        # --- Step 9: Export to CSV ---
        df.to_csv(output_path, index=False)
        
        return True, output_path, None, df
        
    except Exception as e:
        return False, None, str(e), None

def process_multiple_files(input_folder, output_folder=None, item_po_no=None, order_group=None, priority_value=None):
    """
    Process all Excel files in a folder
    
    Parameters:
    input_folder (str): Folder containing input Excel files
    output_folder (str): Folder for output CSV files (optional)
    item_po_no (str): Common ItemPoNo for all files
    order_group (str): Common OrderGroup for all files
    priority_value (str): Common Priority value for all files
    
    Returns:
    list: Results for each file processed
    """
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx"))
    results = []
    
    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")
        success, output_path, error, df = process_vimco_file(
            file_path, output_folder, item_po_no, order_group, priority_value
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
    
    parser = argparse.ArgumentParser(description='Bulk process VIMCO PO Excel files to standardized format')
    parser.add_argument('--input', '-i', required=True, help='Input file or folder path')
    parser.add_argument('--output', '-o', help='Output folder path (optional)')
    parser.add_argument('--batch', '-b', action='store_true', help='Process all files in input folder')
    parser.add_argument('--item-po-no', '-p', help='ItemPoNo number (optional)')
    parser.add_argument('--order-group', '-g', help='OrderGroup value (optional)')
    parser.add_argument('--priority', '-r', help='Priority value (default: "5 day")')
    
    args = parser.parse_args()
    
    if args.batch:
        # Process all files in folder
        if not os.path.isdir(args.input):
            print("❌ Input path must be a folder when using --batch")
            return
        
        # Get common values if not provided
        if not args.item_po_no:
            args.item_po_no = input("Enter ItemPoNo number for all files: ")
        
        if not args.order_group:
            args.order_group = input("Enter OrderGroup value for all files: ")
        
        if not args.priority:
            user_priority = input("Enter Priority for all files (press Enter to use default '5 day'): ")
            args.priority = user_priority if user_priority.strip() != "" else "5 day"
        
        results = process_multiple_files(
            args.input, args.output, 
            args.item_po_no, args.order_group, args.priority
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
        
        success, output_path, error, df = process_vimco_file(
            args.input, args.output, 
            args.item_po_no, args.order_group, args.priority
        )
        
        if success:
            print(f"✅ SUCCESS: Processed {args.input}")
            print(f"📊 Output: {output_path}")
            print(f"📋 Rows processed: {len(df)}")
            print(f"📦 ItemPoNo: {df['ItemPoNo'].iloc[0]}")
            print(f"🏷️ OrderGroup: {df['OrderGroup'].iloc[0]}")
            print(f"⏱️ Priority: {df['Priority'].iloc[0]}")
            print("\nFirst 5 rows preview:")
            print(df.head())
        else:
            print(f"❌ FAILED: {error}")

if __name__ == "__main__":
    main()