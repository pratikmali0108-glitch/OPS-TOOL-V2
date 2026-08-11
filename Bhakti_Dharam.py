# BHAKTI_DHARM_BULK.py - Bulk processing multiple Bhakti & Dharm PO files
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

def process_bhakti_dharm_file(input_file_path, output_folder=None, item_po_no=None, stamp_instruction=None, 
                             order_group=None, priority_value=None, po_no_value=None, size_prefix=None):
    """
    Process single Bhakti & Dharm PO Excel file and convert to standardized format
    
    Parameters:
    input_file_path (str): Path to input Excel file
    output_folder (str): Folder for output Excel file (optional)
    item_po_no (str): ItemPoNo number
    stamp_instruction (str): StampInstruction value
    order_group (str): OrderGroup value
    priority_value (str): Priority value (default: "5 day")
    po_no_value (str): PO NO. value
    size_prefix (str): Size prefix (e.g., US, UP, TS) - default: "US"
    
    Returns:
    tuple: (success_status, output_path, error_message, dataframe)
    """
    try:
        # --- Step 1: Read Excel ---
        df = pd.read_excel(input_file_path)

        # --- Step 2: Select relevant columns ---
        selected_columns = ['#', 'Item #', 'Vendor Item #', 'Description', 'Size', 'Quantity']
        
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
            'Quantity': 'OrderQty',
        }, inplace=True)

        # --- Step 4: Get user inputs if not provided ---
        if item_po_no is None:
            item_po_no = input(f"Enter ItemPoNo number for {Path(input_file_path).name}: ")
        
        if stamp_instruction is None:
            stamp_instruction = input(f"Enter StampInstruction value for {Path(input_file_path).name}: ")
        
        if order_group is None:
            order_group = input(f"Enter OrderGroup value for {Path(input_file_path).name}: ")
        
        if priority_value is None:
            user_priority = input(f"Enter Priority for {Path(input_file_path).name} (press Enter to use default '5 day'): ")
            priority_value = user_priority if user_priority.strip() != "" else "5 day"
        
        if po_no_value is None:
            po_no_value = input(f"Enter PO NO. value for {Path(input_file_path).name}: ")
        
        if size_prefix is None:
            user_prefix = input(f"Enter size prefix for {Path(input_file_path).name} (e.g., US, UP, TS - press Enter for 'US'): ").strip().upper()
            size_prefix = user_prefix if user_prefix != "" else "US"

        # --- Step 5: Helper functions ---
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

        def extract_dia_quality(text):
            if pd.isna(text):
                return ""
            text = str(text).upper()
            if "LGD" in text:
                after_lgd = text.split("LGD", 1)[1].strip()
                words = after_lgd.split()
                return ' '.join(words[:2])
            return ""

        def special_remarks(row):
            metal_descr = metal_text(row['Metal'])
            dia_quality = extract_dia_quality(row['CustomerProductionInstruction'])
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

        def clean_style_code(text):
            if pd.isna(text):
                return ''
            text = str(text).strip().upper()
            text = re.split(r'\s*\+.*| - .*|\s+CHAIN|\s+YG|\s+WG|\s+RG', text)[0]
            return text.strip()

        # --- Step 6: Create / compute new columns ---
        df['StyleCode'] = df['StyleCode'].apply(clean_style_code)
        df['OrderItemPcs'] = 1
        df['Metal'] = df['CustomerProductionInstruction'].apply(map_metal)
        df['Tone'] = df['Metal'].apply(map_tone)
        df['ItemPoNo'] = item_po_no
        df['ItemRefNo'] = ''
        df['StockType'] = ''
        df['Priority'] = priority_value
        df['MakeType'] = ''
        df['StampInstruction'] = stamp_instruction
        df['OrderGroup'] = order_group
        df['Certificate'] = '' 
        df['SpecialRemarks'] = df.apply(special_remarks, axis=1)
        df['DesignProductionInstruction'] = df.apply(design_production_instruction, axis=1)

        # --- Step 6c: Add new empty columns ---
        new_cols_after_sku = [
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate', 'SetPrice', 'StoneQuality', 'PO NO.'
        ]
        for col in new_cols_after_sku:
            df[col] = '' if col != 'PO NO.' else po_no_value

        # --- Step 6d: Size extraction & formatting ---

        # Default ring size 7 for ZR styles
        df.loc[df['StyleCode'].astype(str).str.upper().str.startswith('ZR'), 'ItemSize'] = 7

        def extract_size_from_sku(sku):
            """Extract numeric size like 9 or 9.5 from SKU ending with SZ9"""
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
        df.drop(columns=['ExtractedSize'], inplace=True)

        def format_item_size(size):
            """Format numeric ItemSize with user-defined prefix (US/UP/TS)."""
            if pd.isna(size):
                return size
            size_str = str(size).strip().upper()
            if re.fullmatch(r'\d+(\.\d+)?', size_str):
                if '.' in size_str:
                    size_str = size_str.replace('.', '')
                elif len(size_str) == 1:
                    size_str = '0' + size_str
                return f"{size_prefix}{size_str}"
            return size_str

        df['ItemSize'] = df['ItemSize'].apply(format_item_size)

        # Build full StyleCode: base-sizeWG / base-sizePT etc.
        df['ItemSize'] = df['ItemSize'].apply(_map_item_size)
        df['StyleCode'] = df.apply(
            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
        )
        df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        # --- Step 7: Reorder final columns ---
        final_columns = [
            'SrNo', 'StyleCode', 'ItemSize', 'OrderQty', 'OrderItemPcs', 'Metal', 'Tone',
            'ItemPoNo', 'ItemRefNo', 'StockType', 'Priority', 'MakeType',
            'CustomerProductionInstruction', 'SpecialRemarks', 'DesignProductionInstruction',
            'StampInstruction', 'OrderGroup', 'Certificate', 'SKUNo',
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate', 'SetPrice', 'StoneQuality', 'PO NO.'
        ]
        df = df[final_columns]

        # --- Step 8: Generate output filename ---
        input_filename = Path(input_file_path).stem
        if output_folder:
            output_path = os.path.join(output_folder, f"BHAKTI_DHARM_FORMAT_{input_filename}.xlsx")
        else:
            output_path = f"BHAKTI_DHARM_FORMAT_{input_filename}.xlsx"
        
        # --- Step 9: Export to Excel ---
        df.to_excel(output_path, index=False)
        
        return True, output_path, None, df
        
    except Exception as e:
        return False, None, str(e), None

def process_multiple_files(input_folder, output_folder=None, item_po_no=None, stamp_instruction=None, 
                          order_group=None, priority_value=None, po_no_value=None, size_prefix=None):
    """
    Process all Excel files in a folder
    
    Parameters:
    input_folder (str): Folder containing input Excel files
    output_folder (str): Folder for output Excel files (optional)
    item_po_no (str): Common ItemPoNo for all files
    stamp_instruction (str): Common StampInstruction for all files
    order_group (str): Common OrderGroup for all files
    priority_value (str): Common Priority value for all files
    po_no_value (str): Common PO NO. value for all files
    size_prefix (str): Common size prefix for all files
    
    Returns:
    list: Results for each file processed
    """
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx"))
    results = []
    
    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")
        success, output_path, error, df = process_bhakti_dharm_file(
            file_path, output_folder, item_po_no, stamp_instruction, 
            order_group, priority_value, po_no_value, size_prefix
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
    
    parser = argparse.ArgumentParser(description='Bulk process Bhakti & Dharm PO Excel files to standardized format')
    parser.add_argument('--input', '-i', required=True, help='Input file or folder path')
    parser.add_argument('--output', '-o', help='Output folder path (optional)')
    parser.add_argument('--batch', '-b', action='store_true', help='Process all files in input folder')
    parser.add_argument('--item-po-no', '-p', help='ItemPoNo number')
    parser.add_argument('--stamp-instruction', '-s', help='StampInstruction value')
    parser.add_argument('--order-group', '-g', help='OrderGroup value')
    parser.add_argument('--priority', '-r', help='Priority value (default: "5 day")')
    parser.add_argument('--po-no', '-n', help='PO NO. value')
    parser.add_argument('--size-prefix', '-x', help='Size prefix (e.g., US, UP, TS - default: "US")')
    
    args = parser.parse_args()
    
    if args.batch:
        # Process all files in folder
        if not os.path.isdir(args.input):
            print("❌ Input path must be a folder when using --batch")
            return
        
        # Get common values if not provided
        if not args.item_po_no:
            args.item_po_no = input("Enter ItemPoNo number for all files: ")
        
        if not args.stamp_instruction:
            args.stamp_instruction = input("Enter StampInstruction value for all files: ")
        
        if not args.order_group:
            args.order_group = input("Enter OrderGroup value for all files: ")
        
        if not args.priority:
            user_priority = input("Enter Priority for all files (press Enter to use default '5 day'): ")
            args.priority = user_priority if user_priority.strip() != "" else "5 day"
        
        if not args.po_no:
            args.po_no = input("Enter PO NO. value for all files: ")
        
        if not args.size_prefix:
            user_prefix = input("Enter size prefix for all files (e.g., US, UP, TS - press Enter for 'US'): ").strip().upper()
            args.size_prefix = user_prefix if user_prefix != "" else "US"
        
        results = process_multiple_files(
            args.input, args.output, args.item_po_no, args.stamp_instruction,
            args.order_group, args.priority, args.po_no, args.size_prefix
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
        
        success, output_path, error, df = process_bhakti_dharm_file(
            args.input, args.output, args.item_po_no, args.stamp_instruction,
            args.order_group, args.priority, args.po_no, args.size_prefix
        )
        
        if success:
            print(f"✅ SUCCESS: Processed {args.input}")
            print(f"📊 Output: {output_path}")
            print(f"📋 Rows processed: {len(df)}")
            print(f"🏷️ ItemPoNo: {df['ItemPoNo'].iloc[0]}")
            print(f"📝 StampInstruction: {df['StampInstruction'].iloc[0]}")
            print(f"🏷️ OrderGroup: {df['OrderGroup'].iloc[0]}")
            print(f"⏱️ Priority: {df['Priority'].iloc[0]}")
            print(f"📄 PO NO.: {df['PO NO.'].iloc[0]}")
            print(f"📏 Size Prefix: {args.size_prefix if args.size_prefix else 'US'}")
            print("\nFirst 5 rows preview:")
            print(df.head())
        else:
            print(f"❌ FAILED: {error}")

if __name__ == "__main__":
    main()