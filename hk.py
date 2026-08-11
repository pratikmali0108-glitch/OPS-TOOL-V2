import os
import re
import tempfile
import pandas as pd
import pdfplumber
from PyPDF2 import PdfReader, PdfWriter


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


def _get_client_style_list() -> list:
    """Load and cache the approved Client Style No list from HKD_CS.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    try:
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'HKD_CS_100826.xlsx')
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
    Match the built StyleCode against HKD_CS.xlsx 'Client Style No'.

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


def rotate_pdf_left(input_path):
    """Rotate PDF pages 90 degrees left."""
    reader = PdfReader(input_path)
    writer = PdfWriter()
    for page in reader.pages:
        page.rotate(90)
        writer.add_page(page)
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    with open(temp_file.name, "wb") as f_out:
        writer.write(f_out)
    return temp_file.name


def extract_raw_text_from_pdf(pdf_path):
    """Extract text from PDF after rotating."""
    rotated_pdf = rotate_pdf_left(pdf_path)
    full_text = ""
    with pdfplumber.open(rotated_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    return full_text.strip()


_TONE_SUFFIX_MAP = {
    "W": "-INWG",
    "Y": "-INYG",
    "P": "-INPG",
    "PT": "-INPT",
    "AG": "-INAG",
}

def remap_stylecode(style_code, tone):
    """Append tone-based suffix to every StyleCode."""
    sc = str(style_code).strip() if style_code else ""
    t  = str(tone).strip().upper() if tone else ""
    suffix = _TONE_SUFFIX_MAP.get(t)
    if sc and suffix:
        return f"{sc}{suffix}"
    return style_code


def clean_item_size(size_raw, prefix):
    """Clean and format item size with prefix."""
    size_raw = size_raw.strip()
    if not size_raw:
        return ""
    if re.match(r'^\d+\.\d+$', size_raw):
        size_clean = str(float(size_raw)).rstrip('0').rstrip('.')
    else:
        size_clean = size_raw
    if re.match(r'^\d+$', size_clean) and len(size_clean) == 1:
        size_clean = f"0{size_clean}"
    return f"{prefix}{size_clean}" if prefix else size_clean


def parse_purchase_order_data(full_text, size_prefix: str = "US", default_priority: str = "REG"):
    """Parse PDF text and extract purchase order data."""
    if not full_text.strip():
        return pd.DataFrame()

    po_match = re.search(r'PO\s*#\s*[:]*\s*(\d+)', full_text)
    item_po_no = po_match.group(1) if po_match else ""

    item_blocks = re.findall(
        r'(\d+)\.\s*'                          # Sr No.
        r'(\d+/\d+)\s+'                        # Order Code
        r'([\d.]+)\s+'                          # Order Qty
        r'(\S+)\s+'                             # Style Code
        r'(\S+)\s+'                             # Vendor Style
        r'(\S+)\s+'                             # SKU No
        r'(18K[T]?|14K[T]?)\s+'                  # Metal KT
        r'([YW])'                                 # Tone
        r'(?:\s+([\d.]+))?'                     # Optional Item Size
        r'[\s\S]*?Stamping Instructions:\s*([^\n]+)',  # Stamping instructions
        full_text
    )

    if not item_blocks:
        return pd.DataFrame()

    data = []
    for i, block in enumerate(item_blocks, start=1):
        (
            sr_no, order_code, order_qty, style_code,
            vendor_style, sku_no, metal_kt, tone,
            item_size, stamping_instr
        ) = block

        priority = default_priority
        formatted_size = clean_item_size(item_size or "", size_prefix)

        formatted_size = _map_item_size(formatted_size or '')
        style_code = _build_style_code(style_code, formatted_size, tone)
        style_code = _lookup_client_style(style_code)
        metal = f"G{metal_kt.replace('KT', '').replace('K', '')}{tone}"
        tone_full = "Yellow Gold" if tone == "Y" else "White Gold"
        desc_match = re.search(r'(BRACELET|EARRING|RING)', full_text[full_text.find(style_code):], re.IGNORECASE)
        desc = desc_match.group(1).capitalize() if desc_match else "Item"
        desc_full = f"{metal_kt} {tone_full} {desc} 1.00 CTW"

        if formatted_size:
            special_remarks = (
                f"HK DESIGNS,{order_code}, {style_code},{vendor_style}, "
                f"{sku_no},SZ-{formatted_size}, {metal_kt} {tone_full.upper()}"
            )
        else:
            special_remarks = (
                f"HK DESIGNS,{order_code}, {style_code},{vendor_style}, "
                f"{sku_no}, {metal_kt} {tone_full.upper()}"
            )

        design_prod_instr = "White Rodium" if tone == "W" else "No Rodium"

        data.append({
            "SrNo": i,
            "StyleCode": style_code,
            "ItemSize": formatted_size,
            "OrderQty": order_qty,
            "OrderItemPcs": 1,
            "Metal": metal,
            "Tone": tone,
            "ItemPoNo": item_po_no,
            "ItemRefNo": "",
            "StockType": "",
            "Priority": priority,
            "MakeType": "",
            "CustomerProductionInstruction": desc_full,
            "SpecialRemarks": special_remarks,
            "DesignProductionInstruction": design_prod_instr,
            "StampInstruction": stamping_instr.strip(),
            "OrderGroup": "HK DESIGNS",
            "Certificate": "",
            "Basestoneminwt": "",
            "Basestonemaxwt": "",
            "Basemetalminwt": "",
            "Basemetalmaxwt": "",
            "Productiondeliverydate": "",
            "Expecteddeliverydate": "",
            "SetPrice": "",
            "StoneQuality": "",
            "SKUNo": sku_no
        })

    return pd.DataFrame(data)


def _process_excel_hk(input_path: str, size_prefix: str = "US", default_priority: str = "REG") -> pd.DataFrame:
    """
    Process HK Designs Excel file with proper column mapping.
    Based on logic from Jupyter_Notebooks/HK_Excel.ipynb.
    """
    # Step 1: Read Excel and skip top 5 rows
    df = pd.read_excel(input_path, skiprows=5)
    
    # Step 2: Use the 7th row as header
    df.columns = df.iloc[0]
    df = df[1:].reset_index(drop=True)
    
    # Step 3: Clean current column names
    df.columns = df.columns.astype(str).str.strip()
    
    # Step 4: Rename ALL columns manually
    new_column_names = [
        "SrNo", "Item#", "VendorItem#", "StyleCode", "SKUNo", "ItemSize",
        "OrderQty", "DeliveryDate", "Hallmark", "GCAL", "not needed", "Metal",
        "Tone", "StampInstr", "Customer", "b", "PO#", "Remark",
        "MinMax", "Color", "Quality", "StoneShape", "CustomerProductionInstruction"
    ]
    df.columns = new_column_names[:len(df.columns)]
    
    # Step 5: Select relevant columns
    selected_columns = [
        'SrNo', 'StyleCode', 'SKUNo',
        'ItemSize', 'OrderQty', 'Metal', 'Tone', 'PO#',
        'MinMax', 'Quality', 'CustomerProductionInstruction'
    ]
    df_selected = df[selected_columns].copy()
    
    # Step 5A: Clean and format ItemSize
    def format_item_size(size):
        if pd.isna(size):
            return ""
        s = str(size).strip().upper()
        if s.startswith("US"):
            num_part = s.replace("US", "").strip()
            try:
                if float(num_part).is_integer():
                    return f"US{int(float(num_part)):02d}"
                else:
                    return f"US{float(num_part)}"
            except:
                return s
        else:
            try:
                if float(s).is_integer():
                    return f"US{int(float(s)):02d}"
                else:
                    return f"US{float(s)}"
            except:
                return s
    
    df_selected["ItemSize"] = df_selected["ItemSize"].apply(format_item_size)
    
    # Step 6: Reorder columns
    item_size = df_selected.pop('ItemSize')
    df_selected.insert(df_selected.columns.get_loc('StyleCode') + 1, 'ItemSize', item_size)
    
    order_qty = df_selected.pop('OrderQty')
    df_selected.insert(df_selected.columns.get_loc('ItemSize') + 1, 'OrderQty', order_qty)
    
    df_selected.insert(df_selected.columns.get_loc('OrderQty') + 1, 'OrderItemPcs', value=1)
    
    metal = df_selected.pop('Metal')
    df_selected.insert(df_selected.columns.get_loc('OrderItemPcs') + 1, 'Metal', metal)
    
    tone = df_selected.pop('Tone')
    df_selected.insert(df_selected.columns.get_loc('Metal') + 1, 'Tone', tone)
    
    # Step 8: Map Metal column (default to non-recycled)
    def map_metal(value):
        if pd.isna(value):
            return value
        val = str(value).upper().strip()
        if "PLATINUM" in val:
            return "PC95"  # Default to non-recycled
        return val
    
    df_selected["Metal"] = df_selected["Metal"].apply(map_metal)
    
    # Step 9: Map Tone column based on Metal
    def map_tone(metal_value, tone_value):
        if pd.isna(metal_value):
            return tone_value
        val = str(metal_value).upper().strip()
        
        if val == "PC95" or val == "PC95Z":
            return "PT"
        
        gold_map = {
            "G14W": "W", "G14Y": "Y", "G14P": "P",
            "G18W": "W", "G18Y": "Y", "G18P": "P",
            "G10W": "W", "G10Y": "Y", "G10P": "P",
            "G14WZ": "W", "G14YZ": "Y", "G14PZ": "P",
            "G18WZ": "W", "G18YZ": "Y", "G18PZ": "P",
            "G10WZ": "W", "G10YZ": "Y", "G10PZ": "P"
        }
        
        if val in gold_map:
            return gold_map[val]
        
        return tone_value
    
    df_selected["Tone"] = df_selected.apply(lambda row: map_tone(row["Metal"], row["Tone"]), axis=1)

    # Build full StyleCode: base-sizeWG / base-sizePT etc.
    df_selected["ItemSize"] = df_selected["ItemSize"].apply(_map_item_size)
    df_selected["StyleCode"] = df_selected.apply(
        lambda row: _build_style_code(row["StyleCode"], row["ItemSize"], 'AG' if str(row.get("Metal", "")).upper() == 'AG925' else str(row["Tone"])), axis=1
    )
    df_selected["StyleCode"] = df_selected["StyleCode"].apply(_lookup_client_style)
    df_selected.loc[df_selected["Metal"].astype(str).str.upper() == 'AG925', "Tone"] = ''

    # Step 10: Read cell C3 for ItemPoNo
    po_value = pd.read_excel(input_path, header=None, engine='openpyxl').iloc[2, 2]
    df_selected.insert(df_selected.columns.get_loc('Tone') + 1, 'ItemPoNo', po_value)
    
    # Step 11: Add columns after ItemPoNo
    priority_value = default_priority.upper()
    new_cols = {
        "ItemRefNo": "",
        "StockType": "",
        "Priority": priority_value,
        "MakeType": ""
    }
    
    insert_pos = df_selected.columns.get_loc("ItemPoNo") + 1
    for col_name, col_value in new_cols.items():
        df_selected.insert(insert_pos, col_name, col_value)
        insert_pos += 1
    
    cust_prod = df_selected.pop('CustomerProductionInstruction')
    df_selected.insert(df_selected.columns.get_loc('MakeType') + 1, 'CustomerProductionInstruction', cust_prod)
    
    # Step 12: Add OrderGroup column (default to empty, can be set via UI)
    order_group_value = ""
    df_selected.insert(
        df_selected.columns.get_loc("CustomerProductionInstruction") + 1,
        "OrderGroup",
        order_group_value
    )
    
    # Step 13: Generate SpecialRemarks
    def generate_special_remarks(row):
        metal = str(row["Metal"]).upper()
        tone_code = ""
        
        if "PC95" in metal:
            tone_code = "PW"
        elif metal.startswith(("G10", "G14", "G18")):
            if "W" in metal:
                tone_code = "GW"
            elif "Y" in metal:
                tone_code = "GY"
            elif "P" in metal:
                tone_code = "GP"
        
        po_val = str(row["PO#"]) if "PO#" in row else ""
        if "/" in po_val:
            po_suffix = po_val.split("/", 1)[1]
        else:
            po_suffix = ""
        
        order_group = str(row['OrderGroup']) if row['OrderGroup'] else ""
        sku = str(row['SKUNo']) if pd.notna(row['SKUNo']) else ""
        style = str(row['StyleCode']) if pd.notna(row['StyleCode']) else ""
        
        return f"{order_group},{sku},{style}-{tone_code}{po_suffix}"
    
    df_selected.insert(
        df_selected.columns.get_loc("CustomerProductionInstruction") + 1,
        "SpecialRemarks",
        df_selected.apply(generate_special_remarks, axis=1)
    )
    
    # Step 14: DesignProductionInstruction based on Tone
    def generate_design_instruction(tone_value):
        if pd.isna(tone_value):
            return ""
        tone = str(tone_value).upper().strip()
        if tone == "W":
            return "WHITE RODIUM"
        elif tone in ["Y", "P", "PT"]:
            return "NO RODIUM"
        else:
            return ""
    
    df_selected.insert(
        df_selected.columns.get_loc("SpecialRemarks") + 1,
        "DesignProductionInstruction",
        df_selected["Tone"].apply(generate_design_instruction)
    )
    
    # Step 15: StampInstruction
    def generate_stamp_instruction(row):
        metal = str(row["Metal"]).upper() if pd.notna(row["Metal"]) else ""
        minmax = str(row["MinMax"]).upper() if pd.notna(row["MinMax"]) else ""
        quality = str(row["Quality"]).upper() if pd.notna(row["Quality"]) else ""
        
        gcal_text = "GCAL cert inscription" if "GCAL" in minmax else ""
        
        parts = [metal, "HK LOGO"]
        if gcal_text:
            parts.append(gcal_text)
        if quality:
            parts.append(quality)
        
        return " + ".join(parts)
    
    df_selected.insert(
        df_selected.columns.get_loc("DesignProductionInstruction") + 1,
        "StampInstruction",
        df_selected.apply(generate_stamp_instruction, axis=1)
    )
    
    # Step 16: Certificate column after OrderGroup
    df_selected.insert(
        df_selected.columns.get_loc("OrderGroup") + 1,
        "Certificate",
        ""
    )
    
    # Step 18: Add multiple columns after Certificate (before SKUNo)
    new_columns_after_certificate = [
        "Basestoneminwt", "Basestonemaxwt", "Basemetalminwt", "Basemetalmaxwt",
        "Productiondeliverydate", "Expecteddeliverydate", "SetPrice", "StoneQuality"
    ]
    
    # Move SKUNo to after these new columns
    sku_no = df_selected.pop('SKUNo')
    
    insert_pos = df_selected.columns.get_loc("Certificate") + 1
    for col in new_columns_after_certificate:
        df_selected.insert(insert_pos, col, "")
        insert_pos += 1
    
    # Now insert SKUNo after all the new columns
    df_selected.insert(insert_pos, 'SKUNo', sku_no)
    
    # Step 19: Drop unnecessary columns
    df_selected.drop(columns=['PO#', 'MinMax', 'Quality'], inplace=True, errors='ignore')
    
    return df_selected


def process_hk_file(input_path: str, output_dir: str, size_prefix: str = "US", default_priority: str = "REG"):
    """Process HK file - PDF or Excel."""
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()
        
        if ext == '.pdf':
            full_text = extract_raw_text_from_pdf(input_path)
            df = parse_purchase_order_data(full_text, size_prefix=size_prefix or "US", default_priority=(default_priority or "REG").upper())
            if df is None or df.empty:
                return False, None, "No structured data extracted from PDF", None
            df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''
            output_path = os.path.join(output_dir, f"{base_name}_HK_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        elif ext in ['.xlsx', '.xls']:
            df = _process_excel_hk(input_path, size_prefix=size_prefix or "US", default_priority=(default_priority or "REG").upper())
            output_path = os.path.join(output_dir, f"{base_name}_HK_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None
