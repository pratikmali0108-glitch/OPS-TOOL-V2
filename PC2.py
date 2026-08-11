import os
import re
from pathlib import Path
import pandas as pd
import pdfplumber


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'IT12', 'W') -> 'VR1943EEA-12WG'
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
    """Load and cache the approved Client Style No list from PC2_CS_100826.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    try:
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'PC2_CS_100826.xlsx')
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
    Match the built StyleCode against PC2_CS_100826.xlsx 'Client Style No'.

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


def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Read a PDF and return full extracted text across all pages.
    """
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text += page_text + "\n"
    return full_text


def parse_pc2_items_from_text(extracted_text: str) -> list[dict]:
    """
    Parse the PC2 Damiani multi-item text into a list of dict rows.
    Logic adapted from the notebook's 3rd cell.
    """
    data: list[dict] = []
    sr_no = 1

    # Find item sections ending at the legal notice line
    item_sections = re.findall(
        r'U\.M\.\s+Quantità\s+Importo\n(.*?)ulteriori misure previste dalla legge\.',
        extracted_text,
        re.DOTALL
    )

    for section in item_sections:
        lines = section.split('\n')
        current_item_lines: list[str] = []

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Detect line with item info
            item_pattern = r'^[A-Z0-9\-]+\s+\d+\s+.*?ORO\s+\w+\s+PZ\s+\d+'
            if re.match(item_pattern, line):
                if current_item_lines:
                    item_data = _parse_single_pc2_item(current_item_lines, sr_no)
                    if item_data:
                        data.append(item_data)
                        sr_no += 1
                current_item_lines = [line]
            elif current_item_lines:
                current_item_lines.append(line)

        # last in this section
        if current_item_lines:
            item_data = _parse_single_pc2_item(current_item_lines, sr_no)
            if item_data:
                data.append(item_data)
                sr_no += 1

    return data


def _parse_single_pc2_item(item_lines: list[str], sr_no: int) -> dict | None:
    item_text = ' '.join(item_lines)
    words = item_text.split()
    if len(words) < 2:
        return None

    raw_style = words[0]
    style_code = raw_style.split('-')[0].strip()
    item_ref_no = words[1]

    size_match = re.search(r'mis\.(\d+)', item_text)
    if size_match:
        size_num = size_match.group(1).zfill(2)
        item_size = f"IT{size_num}"
    else:
        item_size = ""

    tone_match = re.search(r'(ORO\s+\w+)\s+PZ', item_text)
    tone = ""
    if tone_match:
        oro_type = tone_match.group(1)
        if "ORO BIANCO" in oro_type:
            tone = "WG"
        elif "ORO ROSE" in oro_type:
            tone = "PG"
        elif "ORO GIALLO" in oro_type:
            tone = "YG"

    tone_display = {"WG": "W", "YG": "Y", "PG": "P"}.get(tone, "")

    qty_match = re.search(r'PZ\s+(\d+)', item_text)
    order_qty = qty_match.group(1) if qty_match else ""

    customer_production_instruction = {
        "WG": "750 WHITE GOLD",
        "PG": "750 PINK GOLD",
        "YG": "750 YELLOW GOLD"
    }.get(tone, "")

    design_production_instruction = {
        "WG": "WHITE RODIUM",
        "PG": "NO RODIUM",
        "YG": "NO RODIUM"
    }.get(tone, "")

    size_for_remarks = item_size.replace("IT", "").strip()
    special_remarks = f"SVP - {item_ref_no}, {customer_production_instruction}, {design_production_instruction}, ITALIAN SIZE - {size_for_remarks}"

    return {
        'SrNo': sr_no,
        'StyleCode': style_code,
        'ItemSize': item_size,
        'OrderQty': order_qty,
        'OrderItemPcs': 1,
        'Metal': "G750",
        'Tone': tone_display,
        'ItemPoNo.': "",
        'ItemRefNo': item_ref_no,
        'StockType': "",
        'Priority': "REG",
        'MakeType': "",
        'CustomerProductionInstruction': customer_production_instruction,
        'SpecialRemarks': special_remarks,
        'DesignProductionInstruction': design_production_instruction,
        'StampInstruction': "750 SALVINI",
        'OrderGroup': "",
        'Certificate': "",
        'SKUNo': "",
        'Basestoneminwt': "",
        'Basestonemaxwt': "",
        'Basemetalminwt': "",
        'Basemetalmaxwt': "",
        'Productiondeliverydate': "",
        'Expecteddeliverydate': "",
        'SetPrice': "",
        'StoneQuality': ""
    }


def process_pc2_file(pdf_file_path: str, output_folder: str | None = None, item_po_no: str | None = None):
    """
    Process a single PC2 PDF into standardized Excel, returning tuple(success, output_path, error, df)
    """
    try:
        extracted_text = extract_text_from_pdf(pdf_file_path)
        rows = parse_pc2_items_from_text(extracted_text)
        df = pd.DataFrame(rows)

        if item_po_no:
            df['ItemPoNo.'] = item_po_no

        # Fix metal code based on tone
        def update_metal(row):
            metal = row['Metal']
            tone = str(row['Tone']).upper()
            if metal == 'G750':
                if tone in ['W', 'WG']:
                    return 'G750W'
                elif tone in ['Y', 'YG']:
                    return 'G750Y'
                elif tone in ['P', 'PG']:
                    return 'G750P'
            return metal

        df['Metal'] = df.apply(update_metal, axis=1)

        # Build full StyleCode: base-sizeWG / base-sizePT etc.
        df['ItemSize'] = df['ItemSize'].apply(_map_item_size)
        df['StyleCode'] = df.apply(
            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
        )
        df['StyleCode'] = df['StyleCode'].apply(_lookup_client_style)
        df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        input_filename = Path(pdf_file_path).stem
        output_folder = output_folder or os.getcwd()
        output_path = os.path.join(output_folder, f"PC2_FORMAT_{input_filename}.xlsx")
        df.to_excel(output_path, index=False)
        return True, output_path, None, df
    except Exception as e:
        return False, None, str(e), None


