import os
import re
import pandas as pd
import pdfplumber
from datetime import datetime
from decimal import Decimal, InvalidOperation


# ─────────────────────────────────────────────
#  INTERNAL HELPERS
# ─────────────────────────────────────────────

def _find(pattern: str, text: str, default: str = "") -> str:
    m = re.search(pattern, text)
    return m.group(1).strip() if m else default


_ITEMSIZE_MST_CACHE = None


def _canon_number_key(num_str: str) -> str:
    """
    Canonicalize a numeric string so "07" -> "7", "5.50" -> "5.5".
    """
    s = (str(num_str) if num_str is not None else "").strip()
    if not s:
        return ""
    s = s.replace(",", "")
    try:
        d = Decimal(s)
    except InvalidOperation:
        return ""
    # Normalize removes trailing zeros (e.g. 5.50 -> 5.5)
    s2 = format(d.normalize(), "f")
    if "." in s2:
        s2 = s2.rstrip("0").rstrip(".")
    return s2


def _norm_item_size_code_key(size: str) -> str:
    """
    Normalize a size like "7" / "07" into master key "07".
    Supports decimal sizes too (e.g. "5.5").
    """
    s = (str(size) if size is not None else "").strip()
    if not s:
        return ""

    m = re.search(r"(\d+(?:\.\d+)?)", s)
    if not m:
        return ""
    return _canon_number_key(m.group(1))


def _load_item_size_mst() -> dict:
    """
    Load ItemSize master.

    Master file structure (based on your current Excel):
      - column: 'Item Size Code'
      - values: 'IT 07', 'US07', 'UP07', 'TS 07', ...

    Returns:
      mapping[prefix][size_key] -> exact Item Size Code string from the master.
    """
    global _ITEMSIZE_MST_CACHE
    if _ITEMSIZE_MST_CACHE is not None:
        return _ITEMSIZE_MST_CACHE

    mst_path = os.path.join(os.path.dirname(__file__), "ItemSize_Mst.xlsx")
    if not os.path.exists(mst_path):
        _ITEMSIZE_MST_CACHE = {}
        return _ITEMSIZE_MST_CACHE

    mst = pd.read_excel(mst_path, dtype=str)
    mst.columns = [str(c).strip() for c in mst.columns]

    if "Item Size Code" not in mst.columns:
        _ITEMSIZE_MST_CACHE = {}
        return _ITEMSIZE_MST_CACHE

    out = {}
    for raw in mst["Item Size Code"].dropna().tolist():
        val = str(raw).strip()
        if not val:
            continue

        # Prefix is the leading letters (e.g. "UP07", "TS 07").
        m = re.match(r"^\s*([A-Za-z]+)", val)
        if not m:
            continue
        prefix = m.group(1).upper()

        # Extract the first numeric value from the master entry (e.g. UP5.5 / TS 07 / 6.30 INCH)
        m_num = re.search(r"(\d+(?:\.\d+)?)", val)
        if not m_num:
            continue
        size_key = _canon_number_key(m_num.group(1))
        if not size_key:
            continue

        # If multiple rows collapse to the same numeric key (e.g. "UP09" and "UP 9:50" both
        # normalize to key "9"), prefer the clean canonical two-digit/integer form.
        def _master_value_is_preferred(v: str) -> bool:
            u = str(v).strip().upper()
            if ':' in u:
                return False
            if '.' in u:
                return False
            digits = re.sub(r"\D", "", u)
            return len(digits) == 2

        existing = out.get(prefix, {}).get(size_key)
        if existing is None:
            out.setdefault(prefix, {})[size_key] = val
        else:
            if _master_value_is_preferred(val) and not _master_value_is_preferred(existing):
                out.setdefault(prefix, {})[size_key] = val

    _ITEMSIZE_MST_CACHE = out
    return out


def _map_item_size_from_mst(size: str, prefix: str) -> str:
    """
    Map ItemSize using ItemSize master (Item Size Code column).
    Falls back to the existing _build_item_size logic if not found.
    """
    key = _norm_item_size_code_key(size)
    if not key:
        return _build_item_size(size, prefix)

    prefix = (prefix or "").strip().upper()
    mst_map = _load_item_size_mst()
    mapped = mst_map.get(prefix, {}).get(key, "")
    return mapped if mapped else _build_item_size(size, prefix)


_CLIENT_STYLE_DF = None
_CS_ITEMSIZE_DATA = None

# CS file path for ItemSize lookup from CS_100826 folder
CS_ITEMSIZE_FILE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "CS_100826",
    "BDL_CS_100826.xlsx"
)


def _get_client_style_df() -> pd.DataFrame:
    """Load and cache DHI_CS.xlsx as a dataframe."""
    global _CLIENT_STYLE_DF
    if _CLIENT_STYLE_DF is not None:
        return _CLIENT_STYLE_DF
    try:
        cs_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "CS_220526",
            "DHI_CS.xlsx"
        )
        _CLIENT_STYLE_DF = pd.read_excel(cs_path, dtype=str)
        # Clean up column names
        _CLIENT_STYLE_DF.columns = [str(c).strip() for c in _CLIENT_STYLE_DF.columns]
    except Exception:
        _CLIENT_STYLE_DF = pd.DataFrame()
    return _CLIENT_STYLE_DF


def _load_cs_itemsize_data() -> pd.DataFrame:
    """Load and cache CS_100826/BDL_CS_100826.xlsx for ItemSize lookup."""
    global _CS_ITEMSIZE_DATA
    if _CS_ITEMSIZE_DATA is not None:
        return _CS_ITEMSIZE_DATA
    try:
        if os.path.exists(CS_ITEMSIZE_FILE_PATH):
            _CS_ITEMSIZE_DATA = pd.read_excel(CS_ITEMSIZE_FILE_PATH, dtype=str)
            # Clean up column names
            _CS_ITEMSIZE_DATA.columns = [str(c).strip() for c in _CS_ITEMSIZE_DATA.columns]
            print(f"✅ Loaded CS ItemSize file with {len(_CS_ITEMSIZE_DATA)} records from CS_100826")
        else:
            print(f"⚠️ CS ItemSize file not found at: {CS_ITEMSIZE_FILE_PATH}")
            _CS_ITEMSIZE_DATA = pd.DataFrame()
    except Exception as e:
        print(f"⚠️ Error loading CS ItemSize file: {e}")
        _CS_ITEMSIZE_DATA = pd.DataFrame()
    return _CS_ITEMSIZE_DATA


def _lookup_itemsize_from_cs_100826(style_code: str) -> str:
    """
    Look up ItemSize from CS_100826/BDL_CS_100826.xlsx based on StyleCode.
    Matches the 'Client Style No' column and returns the 'ItemSize' value.
    
    Args:
        style_code: The StyleCode to look up (e.g., "ZR1473SG-WG")
    
    Returns:
        ItemSize string from CS file, or empty string if not found
    """
    cs_data = _load_cs_itemsize_data()
    
    if cs_data.empty:
        return ""
    
    if 'Client Style No' not in cs_data.columns or 'ItemSize' not in cs_data.columns:
        print(f"⚠️ Required columns not found in CS file. Available columns: {cs_data.columns.tolist()}")
        return ""
    
    # Look for exact match in "Client Style No" column
    matching_rows = cs_data[cs_data['Client Style No'].str.strip() == style_code]
    
    if not matching_rows.empty:
        item_size = matching_rows.iloc[0]['ItemSize']
        # Clean and format the ItemSize
        if pd.notna(item_size) and str(item_size).strip():
            item_size_str = str(item_size).strip()
            print(f"📋 Found ItemSize '{item_size_str}' for StyleCode '{style_code}' in CS_100826")
            return item_size_str
    
    print(f"⚠️ No ItemSize found in CS_100826 file for StyleCode '{style_code}'")
    return ""


def _lookup_client_style(style_no: str, size: str, metal: str) -> str:
    """
    Look up Client Style No from DHI_CS.xlsx, match by Style No, then find by size and metal!
    If no match found, returns empty string!
    """
    df = _get_client_style_df()
    if df.empty or "Client Style No" not in df.columns or "Style No" not in df.columns:
        return ""  # No file found or columns missing

    style_no = str(style_no).strip()
    metal = str(metal).strip()
    size = str(size).strip()

    # First filter by Style No
    filtered = df[df["Style No"].str.strip() == style_no]

    if filtered.empty:
        return ""  # No entries for this Style No

    # Extract size number
    size_num_str = re.search(r"(\d+(?:\.\d+)?)", size)
    size_num = size_num_str.group(1) if size_num_str else ""
    
    # Find metal suffix to look for
    if "PT" in metal or metal.startswith("PC95"):
        metal_suffix = "PT"
    elif metal.endswith("W"):
        metal_suffix = "WG"
    elif metal.endswith("Y"):
        metal_suffix = "YG"
    elif metal.endswith("P"):
        metal_suffix = "RG"
    elif "AG" in metal:
        metal_suffix = "AG"
    else:
        metal_suffix = ""
    
    # First, look for Client Style No that has BOTH the size_num AND metal_suffix in it!
    candidates = []
    for idx, row in filtered.iterrows():
        client_style = str(row["Client Style No"]).strip()
        has_size = size_num in client_style
        has_metal = metal_suffix in client_style
        if has_size and has_metal:
            candidates.append(client_style)
    
    if len(candidates) > 0:
        # If multiple candidates, pick the shortest one (least extra stuff)
        return min(candidates, key=len)
    
    # If no candidates with both, try with just size_num
    for idx, row in filtered.iterrows():
        client_style = str(row["Client Style No"]).strip()
        if size_num in client_style:
            return client_style
    
    # If that doesn't work, try with just metal_suffix
    for idx, row in filtered.iterrows():
        client_style = str(row["Client Style No"]).strip()
        if metal_suffix in client_style:
            return client_style
    
    # Last resort: return first entry
    return str(filtered.iloc[0]["Client Style No"]).strip()


def _metal_karat_type(description: str, item_num: str):
    """Return (karat_str, color_char) from description or item number."""
    src = (description + " " + item_num).upper()
    # Check for platinum first!
    if "PT" in src or "PLATINUM" in src:
        return "950", "PT"
    # Then check for gold!
    for kt in ['18K', '14K']:
        for col in ['W', 'Y', 'P']:
            if f"{kt}{col}" in src:
                return kt.replace('K', ''), col
    return '14', 'W'


def _build_style_code(item_num: str, size: str, description: str, metal: str) -> str:
    """
    Build StyleCode — e.g. ZR2740L-7WG or RG0001964QA-9PT
    ZR items always use size 7 in the StyleCode (catalog standard).
    All other items use the actual ordered size.
    Prefix (UP/US/TS) is NOT included here — it belongs in ItemSize only.
    """
    base = re.sub(r'^[\d]+K[WYPR]*\s*', '', item_num.strip(), flags=re.IGNORECASE)
    _, color = _metal_karat_type(description, item_num)
    sz = '7' if base.upper().startswith('ZR') else (size.strip() or '7')
    
    if metal.startswith("PC95") or "PT" in metal:
        metal_sfx = "PT"
    else:
        metal_sfx = f"{color}G" if color in ('W', 'Y') else 'P'
    
    return f"{base}-{sz}{metal_sfx}"


def _build_item_size(size: str, prefix: str) -> str:
    """Build ItemSize — UP07 / US07 / TS 07"""
    sz = size.strip() if size.strip() else '7'
    # Pad integer sizes to 2 digits; leave decimals as-is
    try:
        sz = str(int(float(sz))).zfill(2) if float(sz) == int(float(sz)) else sz
    except (ValueError, OverflowError):
        sz = sz.zfill(2)
    return f"TS {sz}" if prefix == 'TS' else f"{prefix}{sz}"


def _build_metal(description: str, item_num: str, recycled: bool) -> str:
    """Map description to metal code: G14W / G14WZ / G14Y / G18W …"""
    karat, color = _metal_karat_type(description, item_num)
    suffix = 'Z' if recycled else ''
    if color == "PT":
        return f"PC95{suffix}"
    if color in ('W', 'Y'):
        return f"G{karat}{color}{suffix}"
    return f"PC95{suffix}"


def _get_tone(metal: str) -> str:
    if metal.startswith('PC95') or metal.startswith('PT'):
        return 'PT'
    if 'W' in metal:
        return 'W'
    if 'Y' in metal:
        return 'Y'
    return 'P'


def _metal_to_label(metal: str) -> str:
    mapping = {
        'G14W':  '14 WHITE GOLD',  'G14WZ': '14 WHITE GOLD',
        'G14Y':  '14 YELLOW GOLD', 'G14YZ': '14 YELLOW GOLD',
        'G18W':  '18 WHITE GOLD',  'G18WZ': '18 WHITE GOLD',
        'G18Y':  '18 YELLOW GOLD', 'G18YZ': '18 YELLOW GOLD',
        'PC95':  'PLATINUM',       'PC95Z': 'PLATINUM',
        'PT':    'PLATINUM',       'PTZ':   'PLATINUM',
    }
    for k, v in mapping.items():
        if metal.startswith(k):
            return v
    return metal


def _build_special_remark(order_group: str, sku_no: str, metal: str) -> str:
    parts = [p for p in [order_group, sku_no, _metal_to_label(metal), 'DIA QLTY-LGD-GH VS'] if p]
    return ','.join(parts)


def _extract_stamp_instruction_for_size(notes_text: str, size: str) -> str:
    """
    Extract specific stamping text from Notes correlating with the item's size.
    """
    if not notes_text:
        return ""
    
    sz_norm = size.strip()
    lines = notes_text.split('\n')
    
    # 1. Attempt a line match where the specific size context and "Stamping" coexist
    for line in lines:
        if "stamping" in line.lower():
            if f"SZ{sz_norm}" in line or f"/ {sz_norm}/" in line or f"-SZ{sz_norm}" in line or line.strip().startswith(f"R63738-PT-SZ{sz_norm}"):
                m = re.search(r"Stamping\s*[:\-]?\s*([^/\n\r]+)", line, flags=re.IGNORECASE)
                if m:
                    return m.group(1).strip().replace('"', '').replace("'", "")
                    
    # 2. Fallback: Parse all stamp expressions and find the one containing or ending with the size string
    matches = re.findall(r"Stamping\s*[:\-]?\s*([^/\n\r]+)", notes_text, flags=re.IGNORECASE)
    if matches:
        for match in matches:
            cleaned = match.strip().replace('"', '').replace("'", "")
            if cleaned.endswith(sz_norm) or f" {sz_norm}" in cleaned:
                return cleaned
        # Default fallback to the very first match if layout breaks rules
        return matches[0].strip().replace('"', '').replace("'", "")
    
    return ""


# ─────────────────────────────────────────────
#  PDF READERS
# ─────────────────────────────────────────────

def _read_pdf_text(pdf_path: str) -> str:
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + "\n"
    return full_text


def _parse_header(raw_text: str) -> dict:
    return {
        "Order #":     _find(r"Order\s*#[:\s]+(\S+)", raw_text),
        "P.O. #":      _find(r"P\.O\.\s*#[:\s]+(\S+)", raw_text),
        "Date":        _find(r"Date[:\s]+([\d/]+)", raw_text),
        "Due Date":    _find(r"Due\s*Date[:\s]+([\d/]+)", raw_text),
        "Cancel Date": _find(r"Cancel\s*Date[:\s]+([\d/]+)", raw_text),
        "Reference":   _find(r"Reference[:\s]+(\S+)", raw_text),
        "Vendor #":    _find(r"Vendor\s*#[:\s]*(\S+)", raw_text),
        "Phone #":     _find(r"Phone\s*#[:\s]+(\S+)", raw_text),
        "Ship Via":    _find(r"Ship\s*Via[:\s]+(\S+)", raw_text),
    }


# ─────────────────────────────────────────────
#  WORD-POSITION LINE-ITEM PARSER
# ─────────────────────────────────────────────

_Y_TOL = 6   # pixels: group words within this vertical band into one row


def _group_into_rows(words):
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: w['top'])
    rows = []
    current_y = sorted_words[0]['top']
    current_group = []
    for w in sorted_words:
        if abs(w['top'] - current_y) <= _Y_TOL:
            current_group.append(w)
        else:
            rows.append((current_y, sorted(current_group, key=lambda w: w['x0'])))
            current_y = w['top']
            current_group = [w]
    if current_group:
        rows.append((current_y, sorted(current_group, key=lambda w: w['x0'])))
    return rows


def _find_header_row(rows):
    for i, (y, row_words) in enumerate(rows):
        texts = {w['text'] for w in row_words}
        if '#' in texts and 'Memo' in texts and 'Description' in texts:
            return i, y
    return None, None


def _col_x(header_words):
    cx = {}
    for w in header_words:
        t = w['text']
        x = w['x0']
        if t == '#' and 'hash' not in cx:            cx['hash'] = x
        elif t == 'Memo':                             cx['memo_hdr'] = x
        elif t == 'Item' and 'item_hdr' not in cx:  cx['item_hdr'] = x
        elif t == 'Vendor':                          cx['vendor_hdr'] = x
        elif t == 'Job':                             cx['job_hdr'] = x
        elif t == 'Description':                     cx['desc_hdr'] = x
        elif t == 'Size':                            cx['size_hdr'] = x
        elif t == 'Quantity':                        cx['qty_hdr'] = x
        elif t == 'Weight':                          cx['weight_hdr'] = x
        elif t == 'Amount':                          cx['amount_hdr'] = x
    return cx


def _assign_col(x, cx):
    hash_x     = cx.get('hash', 45)
    item_hdr   = cx.get('item_hdr', 101)
    vendor_hdr = cx.get('vendor_hdr', 166)
    job_hdr    = cx.get('job_hdr', 223)
    desc_hdr   = cx.get('desc_hdr', 281)
    size_hdr   = cx.get('size_hdr', 360)
    qty_hdr    = cx.get('qty_hdr', 389)
    weight_hdr = cx.get('weight_hdr', 430)
    amount_hdr = cx.get('amount_hdr', 541)

    b_line_memo  = (hash_x + item_hdr) / 2
    b_memo_item  = (item_hdr + vendor_hdr) / 2
    b_item_job   = (vendor_hdr + job_hdr) / 2
    b_job_desc   = (job_hdr + desc_hdr) / 2

    if x < b_line_memo:
        return 'line_no'
    elif x < b_memo_item:
        return 'memo'
    elif x < b_item_job:
        return 'item_no'
    elif x < b_job_desc:
        return 'job_bag'
    elif x < size_hdr:
        return 'desc_or_sku'
    elif x < qty_hdr:
        return 'size'
    elif x < weight_hdr:
        return 'qty'
    elif x < weight_hdr + 25:
        return 'weight'
    elif x < amount_hdr - 20:
        return 'unit_cost'
    else:
        return 'amount'


def _parse_line_items_by_position(pdf_path: str):
    all_items_raw = []
    all_notes = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue

            rows = _group_into_rows(words)
            hdr_idx, hdr_y = _find_header_row(rows)
            if hdr_idx is None:
                continue

            cx = _col_x(rows[hdr_idx][1])

            gt_y = None
            for (y, rw) in rows:
                if any('Grand' in w['text'] for w in rw):
                    gt_y = y
                    break

            data_rows = []
            notes_rows = []
            in_notes = False
            for (y, rw) in rows[hdr_idx + 1:]:
                if gt_y is not None and y >= gt_y:
                    break
                texts = [w['text'] for w in rw]
                assigned_cols = {_assign_col(w['x0'], cx) for w in rw}
                is_non_data = (
                    assigned_cols.issubset({'qty', 'weight', 'unit_cost', 'amount'})
                    and not any(_assign_col(w['x0'], cx) == 'line_no' for w in rw)
                    and not any(_assign_col(w['x0'], cx) == 'memo' for w in rw)
                    and not any(_assign_col(w['x0'], cx) == 'desc_or_sku' for w in rw)
                )
                if is_non_data:
                    in_notes = True
                    continue
                if in_notes:
                    notes_rows.append((y, rw))
                else:
                    data_rows.append((y, rw))

            items = []
            cur = None

            for (y, rw) in data_rows:
                row_cols = {}
                for w in rw:
                    col = _assign_col(w['x0'], cx)
                    row_cols.setdefault(col, []).append(w['text'])

                has_line_no = 'line_no' in row_cols and any(
                    t.isdigit() for t in row_cols['line_no']
                )

                if has_line_no:
                    if cur is not None:
                        items.append(cur)
                    line_no_val = next((t for t in row_cols.get('line_no', []) if t.isdigit()), '1')
                    cur = {
                        'line_no':    line_no_val,
                        'memo_parts': row_cols.get('memo', []),
                        'item_parts': row_cols.get('item_no', []),
                        'job_bag':    ' '.join(row_cols.get('job_bag', [])),
                        'sku_parts':  [],
                        'desc_parts': [],
                        'size':       ' '.join(row_cols.get('size', [])),
                        'qty':        ' '.join(row_cols.get('qty', [])),
                        'weight':     ' '.join(row_cols.get('weight', [])),
                        'unit_cost':  ' '.join(row_cols.get('unit_cost', [])),
                        'amount':     ' '.join(row_cols.get('amount', [])),
                    }
                    for t in row_cols.get('desc_or_sku', []):
                        if t.startswith('SKU#') or t.startswith('SKU '):
                            cur['sku_parts'].append(t)
                        else:
                            cur['desc_parts'].append(t)
                else:
                    if cur is None:
                        cur = {
                            'line_no': '1', 'memo_parts': [], 'item_parts': [],
                            'job_bag': '', 'sku_parts': [], 'desc_parts': [],
                            'size': '', 'qty': '', 'weight': '', 'unit_cost': '', 'amount': '',
                        }
                    cur['memo_parts'].extend(row_cols.get('memo', []))
                    cur['item_parts'].extend(row_cols.get('item_no', []))
                    if not cur['job_bag']:
                        cur['job_bag'] = ' '.join(row_cols.get('job_bag', []))
                    if not cur['size']:
                        cur['size'] = ' '.join(row_cols.get('size', []))
                    if not cur['qty']:
                        cur['qty'] = ' '.join(row_cols.get('qty', []))
                    for t in row_cols.get('desc_or_sku', []):
                        if t.startswith('SKU#') or t.startswith('SKU '):
                            cur['sku_parts'].append(t)
                        else:
                            cur['desc_parts'].append(t)

            if cur is not None:
                items.append(cur)

            for raw in items:
                memo_str = ''.join(raw['memo_parts'])
                
                item_str = ' '.join(raw['item_parts']).strip()
                if not item_str and memo_str:
                    for prefix in ['RG', 'ZR', 'PR', 'XR']:
                        if prefix in memo_str:
                            split_idx = memo_str.rfind(prefix)
                            if split_idx > 0:
                                item_str = memo_str[split_idx:]
                                memo_str = memo_str[:split_idx]
                                raw['item_parts'] = [item_str]
                                break

                size_str = raw['size'].strip()
                if not size_str:
                    m = re.search(r'[Ss][Zz]([\d.]+)', memo_str)
                    if m:
                        size_str = m.group(1)
                if not size_str:
                    size_str = '7'

                item_str = ' '.join(raw['item_parts']).strip()
                desc_str = ' '.join(raw['desc_parts']).strip()
                sku_str = ' '.join(raw['sku_parts']).strip()

                qty_raw = raw['qty'].replace('$', '').strip()
                qty = 1
                m_qty = re.search(r'\d+', qty_raw)
                if m_qty:
                    qty = int(m_qty.group())
                
                if memo_str or item_str or desc_str or sku_str:
                    all_items_raw.append({
                        'Line #':        raw['line_no'],
                        'Memo #':        memo_str,
                        'Item #':        item_str,
                        'Vendor Item #': sku_str,
                        'Job Bag #':     raw['job_bag'],
                        'Description':   desc_str,
                        'Size':          size_str,
                        'Quantity':      qty,
                    })

            for (y, rw) in notes_rows:
                line_text = ' '.join(w['text'] for w in rw)
                if 'RightClick' in line_text or 'Copyright' in line_text or line_text.startswith('2023') or line_text.startswith('2024') or line_text.startswith('2025') or line_text.startswith('2026'):
                    continue
                all_notes.append(line_text)

    notes_text = '\n'.join(all_notes).strip()
    return all_items_raw, notes_text


# ─────────────────────────────────────────────
#  PUBLIC PROCESSOR
# ─────────────────────────────────────────────

def process_bdldhi_file(
    filepath: str,
    output_dir: str,
    recycled: bool = False,
    order_group: str = '',
    priority: str = '-5',
) -> tuple:
    """
    Process a Bhakti / Dharm International LLC PDF purchase order.
    """
    try:
        raw_text = _read_pdf_text(filepath)
        header = _parse_header(raw_text)

        line_items, notes_text = _parse_line_items_by_position(filepath)

        if not line_items:
            return False, None, "Could not extract any line items from the PDF. Please verify the file is a valid SHIMAYRA VPO purchase order.", None

        item_po_no = header.get('Order #', '')
        po_no = header.get('P.O. #', '')

        rows = []
        missing_styles = []
        for sr_no, row in enumerate(line_items, start=1):
            item_num    = str(row.get('Item #', ''))
            size        = str(row.get('Size', '7'))
            desc        = str(row.get('Description', ''))
            qty         = row.get('Quantity', 1)
            vendor_item = str(row.get('Vendor Item #', ''))
            memo        = str(row.get('Memo #', ''))

            metal      = _build_metal(desc, item_num, recycled)
            tone       = _get_tone(metal)
            if not metal.startswith('G'):
                tone = 'PT' if metal.startswith('PC95') else ''
            initial_style_code = _build_style_code(item_num, size, desc, metal)
            base_style_no = re.sub(r'^[\d]+K[WYPR]*\s*', '', item_num.strip(), flags=re.IGNORECASE)
            style_code = _lookup_client_style(base_style_no, size, metal)
            
            if not style_code:
                missing_styles.append(f"Line {sr_no}: Style No {base_style_no}")
            
            base = re.sub(r'^[\d]+K[WYPR]*\s*', '', item_num.strip(), flags=re.IGNORECASE)
            effective_size = '7' if base.upper().startswith('ZR') else size
            
            # First, try to lookup ItemSize from CS_100826 based on the StyleCode
            cs_item_size = _lookup_itemsize_from_cs_100826(style_code if style_code else "")
            if cs_item_size:
                # Use the ItemSize from CS_100826 file
                item_size = cs_item_size
            else:
                # Fallback to ItemSize_Mst with default 'UP' prefix
                item_size = _map_item_size_from_mst(effective_size, 'UP')

            # Dynamically pull the exact stamp instructions matching this specific item's size
            stamp_instruction = _extract_stamp_instruction_for_size(notes_text, size)

            rows.append({
                'SrNo':                             sr_no,
                'StyleCode':                     style_code if style_code else initial_style_code,
                'ItemSize':                      item_size,
                'OrderQty':                      qty,
                'OrderItemPcs':                  1,
                'Metal':                         metal,
                'Tone':                          tone,
                'ItemPoNo':                      item_po_no,
                'ItemRefNo':                     '',
                'StockType':                     '',
                'Priority':                      priority,
                'MakeType':                      '',
                'CustomerProductionInstruction': f"{vendor_item} {desc}".strip(),
                'SpecialRemarks':                 _build_special_remark(order_group, memo, metal),
                'DesignProductionInstruction':   'WHITE RHODIUM' if tone == 'W' else 'NO RHODIUM',
                'StampInstruction':              stamp_instruction,
                'OrderGroup':                    order_group,
                'PO. No.':                       po_no,
                'Certificate':                   '',
                'SKUNo':                         memo,
                'Basestoneminwt':                '',
                'Basestonemaxwt':                '',
                'Basemetalminwt':                '',
                'Basemetalmaxwt':                '',
                'Productiondeliverydate':        '',
                'Expecteddeliverydate':          '',
                'SetPrice':                      '',
                'StoneQuality':                  '',
            })
        
        if missing_styles:
            error_msg = "Client style not present for: " + "; ".join(missing_styles)
            return False, None, error_msg, None

        df = pd.DataFrame(rows)

        base_name = re.sub(r'[^\w\-]', '_', os.path.splitext(os.path.basename(filepath))[0])
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"BDLDHI_{base_name}_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        df.to_excel(output_path, index=False)

        return True, output_path, None, df

    except Exception as exc:
        return False, None, str(exc), None