import os
import re
import pandas as pd
import pdfplumber


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
    """Load and cache the approved Client Style No list from ANAYA_CS.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    try:
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'ANAYA_CS_100826.xlsx')
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
    Match the built StyleCode against ANAYA_CS.xlsx 'Client Style No'.

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


def _convert_size(size):
    if pd.isna(size):
        return ''
    size_str = str(size).strip().upper()
    if 'SIZE-' in size_str:
        try:
            num = int(size_str.split('-')[-1])
            return f'UP{num:02d}'
        except Exception:
            return size_str
    return size_str


def _metal_code(metal_str, user_tone=None):
    """
    Generate metal code from metal string.
    If user_tone is provided, use it to determine metal type.
    """
    metal_str = str(metal_str).upper()
    
    # If user explicitly selected AG (Silver), return AG925
    if user_tone and user_tone.upper() == 'AG':
        return 'AG925'
    
    # If user explicitly selected PT (Platinum), return PC95
    if user_tone and user_tone.upper() == 'PT':
        return 'PC95'
    
    # Check for platinum in description (only if user didn't select a gold tone)
    if not user_tone or user_tone.upper() not in ('W', 'Y', 'P', 'R'):
        if re.search(r'\b(?:PLATINUM|PT(?:950|900)?)\b', metal_str):
            return 'PC95'
    
    # Check for silver in description (only if user didn't select a gold tone)
    if not user_tone or user_tone.upper() not in ('W', 'Y', 'P', 'R'):
        if re.search(r'\b(?:SILVER|AG925|925)\b', metal_str):
            return 'AG925'
    
    # Extract karat for gold
    if re.search(r'\b14K(?:T)?\b', metal_str):
        karat = '14'
    elif re.search(r'\b18K(?:T)?\b', metal_str):
        karat = '18'
    elif re.search(r'\b10K(?:T)?\b', metal_str):
        karat = '10'
    else:
        karat = 'XX'
    
    # Use user-provided tone if available, otherwise extract from metal_str
    if user_tone and user_tone.upper() in ('W', 'Y', 'P', 'R'):
        tone = user_tone.upper()
    else:
        if 'WHITE' in metal_str:
            tone = 'W'
        elif 'YELLOW' in metal_str:
            tone = 'Y'
        elif 'PINK' in metal_str or 'ROSE' in metal_str:
            tone = 'P'
        else:
            tone = 'X'
    
    return f'G{karat}{tone}'


def _extract_text_from_pdf(pdf_path: str) -> str:
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            if page_text:
                text += page_text + "\n"
    return text


def _parse_pdf_to_df(full_text: str, tone: str = "Y") -> pd.DataFrame:
    data = []
    # PO number
    po_match = re.search(r"PO\s*#\s*(\d+)", full_text)
    po_no = po_match.group(1) if po_match else ""

    # Determine karat from text
    # Look for patterns like "14KT" or "18KT" etc.
    karat = None
    
    # Check for karat in the text
    metal_text = full_text.upper()
    if re.search(r'\b18K(?:T)?\b', metal_text):
        karat = '18'
    elif re.search(r'\b14K(?:T)?\b', metal_text):
        karat = '14'
    elif re.search(r'\b10K(?:T)?\b', metal_text):
        karat = '10'

    # Determine base metal from text using user-selected tone
    # Use word boundaries to avoid false matches
    if re.search(r'\b(?:PLATINUM|PT(?:950|900)?)\b', metal_text) and not karat:
        base_metal = "PC95"
        tone = "PT"
    elif re.search(r'\b(?:SILVER|AG925|925)\b', metal_text):
        base_metal = "AG925"
        tone = "AG"
    elif karat:
        # Use detected karat with user-selected tone
        base_metal = f"G{karat}{tone}"
    else:
        base_metal = f"GXX{tone}"

    # Item size: try patterns like Size-7 → "US07" or extract from style code
    size = ""
    m_size = re.search(r"Size[-\s]?(\d+(?:\.\d+)?)", full_text, re.IGNORECASE)
    if m_size:
        raw = float(m_size.group(1))
        # For sizes >= 14, treat as INCH; otherwise as ring size
        if raw >= 14:
            size = f"{int(raw)} INCH" if raw.is_integer() else f"{raw} INCH"
        else:
            size = f"US{int(raw):02d}" if raw.is_integer() else f"US{raw}"

    # Description and remarks (optional)
    special_remarks = "Need Hallmark & Trademark on every piece."
    customer_instr = ""
    m_desc = re.search(r"(\d+\.\d+\s*CT\s*TW[^\n]+)", full_text, re.IGNORECASE)
    if m_desc:
        customer_instr = m_desc.group(1)
        if size:
            customer_instr = customer_instr.replace("Size-", "SZ-")

    # Extract rows like: 1 QR0350H-I1/7 70.00 or 1 QR0350H 70.00
    # Pattern explanation: sr_no style_code (with optional -suffix/number) quantity
    for m in re.finditer(r"^(\d+)\s+([A-Z0-9]+)(?:-[A-Z0-9]+)?(?:/(\d+))?\s+(\d+(?:\.\d+)?)\b", full_text, re.MULTILINE):
        sr_no = m.group(1)
        style_code_base = m.group(2)  # Extract base code like QR0350H
        size_from_code = m.group(3)    # Extract size like 7 from /7
        order_qty = m.group(4).split('.')[0]
        
        # If size was extracted from the style code pattern (e.g., QR0350H-I1/7 -> size=7)
        if size_from_code and not size:
            raw = float(size_from_code)
            if raw >= 14:
                size = f"{int(raw)} INCH" if raw.is_integer() else f"{raw} INCH"
            else:
                size = f"US{int(raw):02d}" if raw.is_integer() else f"US{raw}"
        
        data.append({
            'SrNo': sr_no,
            'StyleCode': style_code_base,
            'ItemSize': size,
            'OrderQty': order_qty,
            'OrderItemPcs': 1,
            'Metal': base_metal,
            'Tone': tone,
            'ItemPoNo': po_no,
            'ItemRefNo': '',
            'StockType': '',
            'MakeType': '',
            'CustomerProductionInstruction': customer_instr,
            'SpecialRemarks': special_remarks,
            'DesignProductionInstruction': '',
            'StampInstruction': "'A' on one side and metal KT on other side of the ring",
            'OrderGroup': '',
            'Certificate': '',
            'SKUNo': '',
            'Basestoneminwt': '',
            'Basestonemaxwt': '',
            'Basemetalminwt': '',
            'Basemetalmaxwt': '',
            'Productiondeliverydate': '',
            'Expecteddeliverydate': '',
            'Blank_Column': '',
            'SetPrice': '',
            'StoneQuality': ''
        })

    if not data:
        return pd.DataFrame()
    df_pdf = pd.DataFrame(data)
    df_pdf['ItemSize'] = df_pdf['ItemSize'].apply(_map_item_size)
    df_pdf['StyleCode'] = df_pdf.apply(
        lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1
    )
    df_pdf['StyleCode'] = df_pdf['StyleCode'].apply(_lookup_client_style)
    return df_pdf


def process_anaya_file(input_path: str, output_dir: str, tone: str = "Y"):
    """
    Implements the Excel-to-Excel mapping logic from EDA_Anaya_excel.ipynb (FINAL CODE section).
    Returns (success: bool, output_path: str|None, error_message: str|None, df: pd.DataFrame|None)
    """
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()
        if ext == '.pdf':
            text = _extract_text_from_pdf(input_path)
            df = _parse_pdf_to_df(text, tone=(tone or 'Y').upper())
            if df is None or df.empty:
                return False, None, 'No structured rows found in PDF', None
            output_path = os.path.join(output_dir, f"{base_name}_ANAYA_CLEANED.csv")
            df.to_csv(output_path, index=False)
            return True, output_path, None, df

        if ext not in ['.xlsx', '.xls']:
            return False, None, f'Unsupported file type for Anaya: {ext}', None

        # Step 1: Read Excel (data table starts after 9 header rows)
        df = pd.read_excel(input_path, skiprows=9)

        # Step 2: Select required columns
        selected_columns = ['Serial No', 'Style No', 'Description', 'Diamonds', 'Qty', 'Sizes']
        missing = [c for c in selected_columns if c not in df.columns]
        if missing:
            return False, None, f'Missing expected columns: {", ".join(missing)}', None
        df_selected = df[selected_columns].copy()

        # Step 3: Rename
        df_selected.rename(columns={
            'Serial No': 'SrNo',
            'Style No': 'StyleCode',
            'Description': 'MetalR',
            'Diamonds': 'CustomerProductionInstruction',
            'Qty': 'OrderQty',
            'Sizes': 'ItemSize'
        }, inplace=True)

        # Step 4: Clean StyleCode
        df_selected.dropna(subset=['StyleCode'], inplace=True)
        df_selected['StyleCode'] = (
            df_selected['StyleCode']
            .astype(str)
            .str.split('-').str[0]
            .str.replace(r'[_\s]+', '', regex=True)
            .str.strip()
        )

        # Step 5: Clean and transform ItemSize
        df_selected['ItemSize'] = df_selected['ItemSize'].apply(_convert_size)
        itemsize = df_selected.pop('ItemSize')
        df_selected.insert(df_selected.columns.get_loc('StyleCode') + 1, 'ItemSize', itemsize)
        orderqty = df_selected.pop('OrderQty')
        df_selected.insert(df_selected.columns.get_loc('ItemSize') + 1, 'OrderQty', orderqty)

        # Step 6: Clean MetalR
        df_selected['MetalR'] = (
            df_selected['MetalR']
            .astype(str)
            .str.replace('\n', ' ', regex=False)
            .str.strip()
        )

        # Step 7: Create Metal (pass user-selected tone)
        df_selected.insert(df_selected.columns.get_loc('MetalR'), 'Metal', 
                          df_selected['MetalR'].apply(lambda x: _metal_code(x, user_tone=tone)))

        # Step 8: Tone from user selection or Metal code
        # If user selected gold tone (W/Y/P/R), use it; if AG/PT, extract from Metal
        def _get_tone(metal_code):
            if tone and tone.upper() in ('W', 'Y', 'P', 'R'):
                return tone.upper()
            elif str(metal_code).upper() == 'AG925':
                return 'AG'
            elif str(metal_code).upper() == 'PC95':
                return 'PT'
            else:
                return str(metal_code)[-1] if metal_code else ''
        
        df_selected.insert(df_selected.columns.get_loc('Metal') + 1, 'Tone', 
                          df_selected['Metal'].apply(_get_tone))

        # Step 9: Extract ItemPoNo from G5
        try:
            item_po_no = pd.read_excel(input_path, header=None, usecols="G", nrows=5).iloc[4, 0]
        except Exception:
            item_po_no = ''
        df_selected.insert(df_selected.columns.get_loc('Tone') + 1, 'ItemPoNo', item_po_no)

        # Step 10: Add ItemRefNo, StockType, MakeType
        additional_cols = ['ItemRefNo', 'StockType', 'MakeType']
        pos = df_selected.columns.get_loc('ItemPoNo') + 1
        for col in additional_cols:
            df_selected.insert(pos, col, '')
            pos += 1

        # Step 11: SpecialRemarks
        df_selected.insert(
            df_selected.columns.get_loc('CustomerProductionInstruction') + 1,
            'SpecialRemarks',
            'Need Hallmark "A" and Trademark on Every piece'
        )

        # Step 12: DesignProductionInstruction
        df_selected.insert(
            df_selected.columns.get_loc('SpecialRemarks') + 1,
            'DesignProductionInstruction',
            ''
        )

        # Step 13: StampInstruction
        df_selected.insert(
            df_selected.columns.get_loc('DesignProductionInstruction') + 1,
            'StampInstruction',
            "'A' on one side and metal KT on other side of the ring"
        )

        # Step 14: Add extra columns
        new_columns = [
            'OrderGroup', 'Certificate', 'SKUNo', 'Basestoneminwt', 'Basestonemaxwt',
            'Basemetalminwt', 'Basemetalmaxwt', 'Productiondeliverydate',
            'Expecteddeliverydate', 'Blank_Column', 'SetPrice', 'StoneQuality'
        ]
        pos = df_selected.columns.get_loc('StampInstruction') + 1
        for col in new_columns:
            df_selected.insert(pos, col, '')
            pos += 1

        # Step 15: OrderItemPcs
        df_selected.insert(df_selected.columns.get_loc('OrderQty') + 1, 'OrderItemPcs', value=1)

        # Step 16: Cleanup
        df_selected.drop(columns=['MetalR'], inplace=True)

        # Build full StyleCode: base-sizeYG / base-sizeWG / base-sizePT etc.
        df_selected['ItemSize'] = df_selected['ItemSize'].apply(_map_item_size)
        df_selected['StyleCode'] = df_selected.apply(
            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
        )
        df_selected['StyleCode'] = df_selected['StyleCode'].apply(_lookup_client_style)
        df_selected.loc[df_selected['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        # Step 17: Export
        output_path = os.path.join(output_dir, f"{base_name}_ANAYA_CLEANED.csv")
        df_selected.to_csv(output_path, index=False)

        return True, output_path, None, df_selected
    except Exception as e:
        return False, None, str(e), None


