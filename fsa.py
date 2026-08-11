import os
import re
import pandas as pd
import pdfplumber


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'EU58', 'W') -> 'VR1943EEA-58WG'
    """
    base = str(base).strip() if base else ''
    item_size = str(item_size).strip() if item_size else ''
    tone = str(tone).strip().upper() if tone else ''
    if not base or base.upper() == 'NAN':
        return ''
    
    has_inch = bool(re.search(r'\bINCH\b', item_size, flags=re.IGNORECASE))
    size_num = re.sub(r'^(?:UP|US|EU|IT|UT|TS|IS)\s*', '', item_size, flags=re.IGNORECASE).strip()
    size_num = re.sub(r'\s*INCH\s*$', '', size_num, flags=re.IGNORECASE).strip()
    try:
        f = float(size_num)
        size_num = str(int(f)) if f.is_integer() else str(f)
    except (ValueError, TypeError):
        pass
        
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
    if not raw or str(raw).strip().upper() in ('', 'NAN'):
        return raw
    lookup = _get_size_lookup()
    key = _normalize_size_key(str(raw).strip())
    return lookup.get(key, raw)


def extract_full_text_from_pdf(pdf_path: str) -> str:
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text:
                full_text += text + "\n"
    return full_text

def find_po_number(text: str) -> str:
    m = re.search(r"Order\s+(\d+)", text, flags=re.IGNORECASE)
    return m.group(1) if m else ""


def _format_pdf_date(day: str, month_name: str, year: str) -> str:
    try:
        dt = pd.to_datetime(f"{day} {month_name} {year}", dayfirst=True)
        return dt.strftime("%d-%m-%Y")
    except (ValueError, TypeError):
        return f"{day} {month_name} {year}".strip()


def find_order_date(text: str) -> str:
    m = re.search(r"Order\s+Date[\s\-:]*(\d{1,2})[\.\s\-]*([A-Za-z]+)[\.\s\-]*(\d{4})", text, flags=re.IGNORECASE)
    if not m:
        return ""
    return _format_pdf_date(m.group(1), m.group(2), m.group(3))


def find_requested_delivery_date(text: str) -> str:
    m = re.search(r"Requested\s+Delivery\s+Date[\s\-:]*(\d{1,2})[\.\s\-]*([A-Za-z]+)[\.\s\-]*(\d{4})", text, flags=re.IGNORECASE)
    if not m:
        return ""
    return _format_pdf_date(m.group(1), m.group(2), m.group(3))



# Pattern A: supplier_ref [tone] SKU size open_qty total_qty Piece ...price
# Supplier ref may be multiple tokens separated by spaces/hyphens (e.g. "AR-82551J- LGD-")
# so we allow one or more word-chunks before the 6-10 digit SKU number.
HEADER_PAT_A = re.compile(
    r"^\s*([A-Z0-9][A-Z0-9\-]*(?:\s+[A-Z0-9][A-Z0-9\-]*)*?)\s+"
    r"(?:(?:[YWR]G)\s*750\s+)?(\d{6,10})\s+(STA|\d+)\s+(\d+)\s+(\d+)\s+(?:Piece|Pieces)\b",
    flags=re.IGNORECASE,
)
# Pattern B: supplier_ref size open_qty total_qty Piece (no SKU on this line)
HEADER_PAT_B = re.compile(
    r"^\s*([A-Z][A-Z0-9\-]{2,})\s+(STA|\d+)\s+(\d+)\s+(\d+)\s+(?:Piece|Pieces)\b",
    flags=re.IGNORECASE,
)
# Pattern C: SKU-first line (no supplier ref on this line)
HEADER_PAT_C = re.compile(
    r"^\s*(\d{6,10})\s+(STA|\d+)\s+(\d+)\s+(\d+)\s+(?:Piece|Pieces)\b",
    flags=re.IGNORECASE,
)
TOTAL_SKU_PAT = re.compile(r"^\s*TOTAL\s+(\d{6,10})\b", flags=re.IGNORECASE)
# Style token: starts with letter, may have hyphens, ends with digit OR letter (e.g. AR-82551L)
STYLE_TOKEN_PAT = re.compile(r"\b([A-Z][A-Z0-9\-]{3,})\b")
EXCLUDE_STYLE_TOKENS = {"YG750", "WG750", "RG750", "YG", "WG", "RG", "LGD", "HPHT", "STA", "TOTAL"}
# Price pattern: a decimal number that appears after "Piece" on the header line
PRICE_PAT = re.compile(r"(?:Piece|Pieces)\s+[\d.]+\s+[\d.]+\s+([\d.]+)", flags=re.IGNORECASE)



def is_item_header_v2(line: str) -> bool:
    return bool(HEADER_PAT_A.search(line) or HEADER_PAT_B.search(line) or HEADER_PAT_C.search(line))


def find_style_in_block(block_lines: list[str]) -> str:
    for ln in block_lines[1:]:
        if re.search(r"^\s*TOTAL\b", ln, flags=re.IGNORECASE):
            break
        m = STYLE_TOKEN_PAT.search(ln)
        if m:
            token = m.group(1)
            if token not in EXCLUDE_STYLE_TOKENS and not token.isdigit() and len(token) >= 5:
                return token
    return ""


def parse_items_v2(text: str, default_priority: str = "REG", default_stamp_var: str = "") -> list[dict]:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    items: list[dict] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if not is_item_header_v2(line):
            i += 1
            continue

        style_code = ""
        sku_no = ""
        size_token = ""
        order_qty = ""
        set_price = ""

        mA = HEADER_PAT_A.search(line)
        mB = HEADER_PAT_B.search(line) if not mA else None
        mC = HEADER_PAT_C.search(line) if not (mA or mB) else None

        if mA:
            # groups: (1)supplier_ref (2)sku (3)size (4)open_qty (5)total_qty
            # supplier_ref may be multi-token (e.g. "AR-82551J- LGD-"); take only the first token
            raw_ref = mA.group(1).strip()
            style_code = raw_ref.split()[0].rstrip('-')
            sku_no = mA.group(2)
            size_token = mA.group(3)
            order_qty = mA.group(4)
        elif mB:
            style_code = mB.group(1)
            size_token = mB.group(2)
            order_qty = mB.group(3)
        elif mC:
            sku_no = mC.group(1)
            size_token = mC.group(2)
            order_qty = mC.group(3)
        else:
            i += 1
            continue

        # Extract unit price from the header line (appears after "Piece metal_wt total_wt price")
        pm = PRICE_PAT.search(line)
        if pm:
            set_price = pm.group(1)

        block_lines = [line]
        j = i + 1
        sku_from_total = ""
        while j < len(lines):
            nxt = lines[j]
            if is_item_header_v2(nxt):
                break
            block_lines.append(nxt)
            t = TOTAL_SKU_PAT.search(nxt)
            if t:
                sku_from_total = t.group(1)
            j += 1
        i = j

        if not sku_no:
            sku_no = sku_from_total

        if not style_code:
            style_code = find_style_in_block(block_lines)

        tone = ''
        joined = " ".join(block_lines)
        if re.search(r"\bWG\s*750|WG750", joined):
            tone = 'W'
        elif re.search(r"\bRG\s*750|RG750", joined):
            tone = 'R'
        elif re.search(r"\bYG\s*750|YG750", joined):
            tone = 'Y'
        else:
            if re.search(r"WHITE", joined, re.IGNORECASE):
                tone = 'W'
            elif re.search(r"ROSE", joined, re.IGNORECASE):
                tone = 'R'
            elif re.search(r"YELLOW", joined, re.IGNORECASE):
                tone = 'Y'

        fineness = '750' if (re.search(r"\b18\s*CARA?\b", joined, re.IGNORECASE) or re.search(r"\b750\b", joined)) else ''

        diamond_quality = ''
        for bl in block_lines:
            m = re.search(r"\b([A-Z]{1,2}-?SI\d|[A-Z]{1,2}-?VS\d?|[A-Z]{1,2}-?VVS\d?|[A-Z]{1,2}-?I\d)\b", bl)
            if m:
                diamond_quality = m.group(1)
                break

        carat_line = ''
        for bl in block_lines:
            m = re.search(r"\b18\s*CARA?\s*-\s*750\b", bl, re.IGNORECASE)
            if m:
                carat_line = m.group(0)
                break
        if not carat_line and fineness == '750':
            carat_line = '18 CARA - 750'

        if size_token and size_token.upper() != 'STA':
            item_size = f"EU{size_token}"
        else:
            item_size = ""

        metal = f"G{fineness}{tone}" if fineness and tone else (f"G{fineness}" if fineness else "")
        stamp_variable_text = 'lgd' if (default_stamp_var or '').lower() == 'lgd' else ''

        # Detect LGD from block text
        if re.search(r"\bLGD\b", joined, re.IGNORECASE):
            stamp_variable_text = 'lgd'

        tone_to_desc = {'Y': 'YELLOW GOLD', 'W': 'WHITE GOLD', 'R': 'ROSE GOLD'}
        tone_desc = tone_to_desc.get(tone, '')
        parts = []
        if sku_no:
            parts.append(sku_no)
        if fineness or tone_desc:
            txt = " ".join([p for p in [fineness, tone_desc] if p]).strip()
            if txt:
                parts.append(txt)
        if diamond_quality:
            parts.append(f"DIA QLTY: {diamond_quality}")
        special_remarks = ",".join(parts)

        common_sentence = "Polishing and setting must be very well done."
        customer_prod_instruction = f"{carat_line}, {common_sentence}" if carat_line else common_sentence
        design_prod_instruction = "white rodium" if tone == 'W' else "no rodoium"

        items.append({
            'SrNo': len(items) + 1,
            'StyleCode': style_code,
            'ItemSize': item_size,
            'OrderQty': order_qty,
            'OrderItemPcs': 1,
            'Metal': metal,
            'Tone': tone,
            'ItemPoNo': '',
            'ItemRefNo': '',
            'StockType': '',
            'Priority': default_priority,
            'MakeType': '',
            'CustomerProductionInstruction': customer_prod_instruction,
            'SpecialRemarks': special_remarks,
            'DesignProductionInstruction': design_prod_instruction,
            'StampInstruction': f"750+customer logo+{stamp_variable_text}".rstrip('+'),
            'OrderGroup': '',
            'Certificate': '',
            'SKUNo': sku_no,
            'Basestoneminwt': '',
            'Basestonemaxwt': '',
            'Basemetalminwt': '',
            'Basemetalmaxwt': '',
            'Productiondeliverydate': '',
            'Expecteddeliverydate': '',
            'SetPrice': set_price,
            'StoneQuality': '',
        })

    return items


def _verify_and_blank_itemsize_from_cs(df: pd.DataFrame) -> pd.DataFrame:
    """
    Verifies StyleCode and ItemSize against CS file, checks Category,
    and blanks ItemSize for EARRING, NECK PIECE, and PENDANT categories.
    
    Logic:
    1. Extract base style from generated StyleCode (part before '-')
    2. Look up in CS file using 'Client Style No' column
    3. Check the Category
    4. If Category contains EARRING, NECK PIECE, or PENDANT, blank ItemSize and rebuild StyleCode
    """
    try:
        cs_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 
                                     'CS_220526', 'FSA_CS_070826.xlsx')
        
        if not os.path.exists(cs_file_path):
            # If CS file not found, return df as is
            return df
        
        # Load CS file
        cs_df = pd.read_excel(cs_file_path)
        
        # Create lookup dictionary: Client Style No -> Category
        # We'll need to match both full StyleCode and base style
        cs_lookup = {}
        for _, row in cs_df.iterrows():
            client_style = str(row.get('Client Style No', '')).strip()
            category = str(row.get('Category', '')).strip().upper()
            
            if client_style and client_style.upper() != 'NAN':
                # Store full client style
                cs_lookup[client_style.upper()] = category
                
                # Also store base style (part before '-')
                base_style = client_style.split('-')[0].strip()
                if base_style:
                    cs_lookup[base_style.upper()] = category
        
        # Process each row in the output dataframe
        for idx, row in df.iterrows():
            style_code = str(row.get('StyleCode', '')).strip()
            item_size = str(row.get('ItemSize', '')).strip()
            
            if not style_code or style_code.upper() == 'NAN':
                continue
            
            # Extract base style (part before '-' if exists)
            base_style = style_code.split('-')[0] if '-' in style_code else style_code
            
            # Try to find category by matching StyleCode or base style
            category = None
            if style_code.upper() in cs_lookup:
                category = cs_lookup[style_code.upper()]
            elif base_style.upper() in cs_lookup:
                category = cs_lookup[base_style.upper()]
            
            # If category found and matches EARRING, PENDANT, or NECK PIECE
            if category:
                keywords = ['EARRING', 'NECK PIECE', 'NECKPIECE', 'PENDANT']
                if any(keyword in category for keyword in keywords):
                    # Blank out ItemSize for these categories
                    df.at[idx, 'ItemSize'] = ''
                    
                    # Rebuild StyleCode without size
                    tone = str(row.get('Tone', '')).strip()
                    metal = str(row.get('Metal', '')).strip()
                    
                    # Rebuild StyleCode with blank ItemSize
                    df.at[idx, 'StyleCode'] = _build_style_code(
                        base_style, 
                        '',  # blank ItemSize
                        'AG' if metal.upper() == 'AG925' else tone
                    )
        
        return df
        
    except Exception as e:
        # If any error occurs, return df as is
        print(f"Warning: Could not verify CS file: {e}")
        return df


def process_fsa_file(input_path: str, output_dir: str, default_priority: str = "REG", default_stamp_var: str = ""):
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()

        # Sanitize fallback logic for priority coming from standard text forms
        cleaned_priority = str(default_priority).strip().upper() if (default_priority and str(default_priority).strip()) else "REG"

        if ext == '.pdf':
            text = extract_full_text_from_pdf(input_path)
            po_no = find_po_number(text)
            po_date = find_order_date(text)
            e_del_date = find_requested_delivery_date(text)
            
            items = parse_items_v2(text, default_priority=cleaned_priority, default_stamp_var=(default_stamp_var or ""))
            for it in items:
                it['ItemPoNo'] = po_no
                it['Date'] = po_date  # Date column for GATI sheet
                it['Podate'] = po_date  # PoDate column for GATI sheet
                it['E Del Date'] = e_del_date
            
            # 'Priority' added inside requested_columns list below
            requested_columns = [
                'SrNo', 'StyleCode', 'ItemSize', 'OrderQty', 'OrderItemPcs', 'Metal', 'Tone',
                'ItemPoNo', 'ItemRefNo', 'StockType', 'Priority', 'MakeType', 'CustomerProductionInstruction',
                'SpecialRemarks', 'DesignProductionInstruction', 'StampInstruction', 'OrderGroup',
                'Certificate', 'SKUNo', 'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt',
                'Basemetalmaxwt', 'Productiondeliverydate', 'Expecteddeliverydate', 'BlankColumn',
                'SetPrice', 'StoneQuality', 'Date', 'Podate', 'E Del Date',
            ]
            df = pd.DataFrame(items)
            for col in requested_columns:
                if col not in df.columns:
                    df[col] = ''
            df = df[requested_columns]
            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)
            df['StyleCode'] = df.apply(
                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
            )
            df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''
            
            # Verify and blank ItemSize for EARRING, NECK PIECE, PENDANT categories
            df = _verify_and_blank_itemsize_from_cs(df)
            
            output_path = os.path.join(output_dir, f"{base_name}_FSA_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
            
        elif ext in ['.xlsx', '.xls', '.csv']:
            df = pd.read_excel(input_path) if ext in ['.xlsx', '.xls'] else pd.read_csv(input_path)
            
            # Ensure custom dynamic priority values flow down to Passthrough Excel files as well
            if 'Priority' in df.columns:
                df['Priority'] = cleaned_priority
            else:
                df.insert(loc=min(10, len(df.columns)), column='Priority', value=cleaned_priority)

            output_path = os.path.join(output_dir, f"{base_name}_FSA_PASSTHROUGH.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None