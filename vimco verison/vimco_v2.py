import os
import re
import pandas as pd
import pdfplumber
from datetime import datetime

# ─────────────────────────────────────────────
#  INTERNAL HELPERS (FROM BDLDHI.py adapted for VIMCO PDF)
# ─────────────────────────────────────────────

def _find(pattern: str, text: str, default: str = "") -> str:
    m = re.search(pattern, text)
    return m.group(1).strip() if m else default

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
    }

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
        if '#' in texts and 'Vendor' in texts and 'Description' in texts:
            return i, y
    return None, None

def _col_x(header_words):
    cx = {}
    for w in header_words:
        t = w['text']
        x = w['x0']
        if t == '#' and 'hash' not in cx:          cx['hash'] = x
        elif t == 'Item' and 'item_hdr' not in cx: cx['item_hdr'] = x
        elif t == 'Vendor':                        cx['vendor_hdr'] = x
        elif t == 'Job':                           cx['job_hdr'] = x
        elif t == 'Description':                  cx['desc_hdr'] = x
        elif t == 'Dia':                           cx['dia_hdr'] = x
        elif t == 'Metal':                         cx['metal_hdr'] = x
        elif t == 'Size':                          cx['size_hdr'] = x
        elif t == 'Qty.':                          cx['qty_hdr'] = x
        elif t == 'Weight':                        cx['weight_hdr'] = x
        elif t == 'Cost':                          cx['cost_hdr'] = x
        elif t == 'Amount':                        cx['amount_hdr'] = x
    return cx

def _assign_col(x, cx):
    hash_x     = cx.get('hash', 45)
    item_hdr   = cx.get('item_hdr', 57)
    vendor_hdr = cx.get('vendor_hdr', 122)
    job_hdr    = cx.get('job_hdr', 180)
    desc_hdr   = cx.get('desc_hdr', 223)
    dia_hdr    = cx.get('dia_hdr', 317)
    metal_hdr  = cx.get('metal_hdr', 369)
    size_hdr   = cx.get('size_hdr', 410)
    qty_hdr    = cx.get('qty_hdr', 440)
    weight_hdr = cx.get('weight_hdr', 459)
    cost_hdr   = cx.get('cost_hdr', 502)
    amount_hdr = cx.get('amount_hdr', 540)

    b_line_item   = (hash_x + item_hdr) / 2
    b_item_vendor = (item_hdr + vendor_hdr) / 2
    b_vendor_job  = (vendor_hdr + job_hdr) / 2
    b_job_desc    = (job_hdr + desc_hdr) / 2
    
    # Description can be wide, so give it space up to just before Dia
    b_desc_dia    = dia_hdr - 10 
    
    b_dia_metal   = (dia_hdr + metal_hdr) / 2
    b_metal_size  = (metal_hdr + size_hdr) / 2
    b_size_qty    = (size_hdr + qty_hdr) / 2
    b_qty_weight  = (qty_hdr + weight_hdr) / 2
    b_weight_cost = (weight_hdr + cost_hdr) / 2
    b_cost_amount = (cost_hdr + amount_hdr) / 2

    if x < b_line_item: return 'line_no'
    elif x < b_item_vendor: return 'item_no'
    elif x < b_vendor_job: return 'vendor_item_no'
    elif x < b_job_desc: return 'job_bag'
    elif x < b_desc_dia: return 'description'
    elif x < b_dia_metal: return 'dia_qlty'
    elif x < b_metal_size: return 'metal_pdf'
    elif x < b_size_qty: return 'size'
    elif x < b_qty_weight: return 'qty'
    elif x < b_weight_cost: return 'weight'
    elif x < b_cost_amount: return 'cost'
    else: return 'amount'

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
                
                assigned_cols = {_assign_col(w['x0'], cx) for w in rw}
                # Summary row has qty, weight, cost, amount but NO line_no
                is_summary = (
                    assigned_cols.issubset({'qty', 'weight', 'cost', 'amount'})
                    and any(_assign_col(w['x0'], cx) == 'qty' for w in rw)
                    and not any(_assign_col(w['x0'], cx) == 'line_no' for w in rw)
                    and data_rows
                )
                if is_summary:
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
                        'item_no':    row_cols.get('item_no', []),
                        'vendor_item_no': row_cols.get('vendor_item_no', []),
                        'job_bag':    row_cols.get('job_bag', []),
                        'description': row_cols.get('description', []),
                        'dia_qlty':   row_cols.get('dia_qlty', []),
                        'metal_pdf':  row_cols.get('metal_pdf', []),
                        'size':       row_cols.get('size', []),
                        'qty':        row_cols.get('qty', []),
                        'weight':     row_cols.get('weight', []),
                        'cost':       row_cols.get('cost', []),
                        'amount':     row_cols.get('amount', []),
                    }
                else:
                    if cur is None:
                        cur = {
                            'line_no': '1', 'item_no': [], 'vendor_item_no': [],
                            'job_bag': [], 'description': [], 'dia_qlty': [], 'metal_pdf': [],
                            'size': [], 'qty': [], 'weight': [], 'cost': [], 'amount': [],
                        }
                    cur['item_no'].extend(row_cols.get('item_no', []))
                    cur['vendor_item_no'].extend(row_cols.get('vendor_item_no', []))
                    cur['job_bag'].extend(row_cols.get('job_bag', []))
                    cur['description'].extend(row_cols.get('description', []))
                    cur['dia_qlty'].extend(row_cols.get('dia_qlty', []))
                    cur['metal_pdf'].extend(row_cols.get('metal_pdf', []))
                    cur['size'].extend(row_cols.get('size', []))
                    cur['qty'].extend(row_cols.get('qty', []))

            if cur is not None:
                items.append(cur)

            for raw in items:
                item_str = ''.join(raw['item_no']).strip()
                vendor_item_str = ''.join(raw['vendor_item_no']).strip()
                job_bag_str = ' '.join(raw['job_bag']).strip()
                desc_str = ' '.join(raw['description']).strip()
                dia_str = ' '.join(raw['dia_qlty']).strip()
                metal_str = ' '.join(raw['metal_pdf']).strip()
                size_str = ''.join(raw['size']).strip()
                
                has_inch = bool(re.search(r'INCH|\"|”|\bIN\b', size_str, flags=re.IGNORECASE))

                # Filter out stray non-numeric tokens before parsing qty
                qty_tokens = [t for t in raw['qty'] if re.search(r'\d', t)]
                qty_raw = ' '.join(qty_tokens).strip()
                qty = 1
                m_qty = re.search(r'^\s*(\d+)\s*$', qty_raw)
                if m_qty:
                    qty = int(m_qty.group(1))
                else:
                    # Fallback: qty may have been pushed into weight column by a stray token
                    weight_tokens = [t for t in raw.get('weight', []) if re.match(r'^\d+$', t.strip())]
                    if weight_tokens:
                        qty = int(weight_tokens[0].strip())

                all_items_raw.append({
                    'Line #':        raw['line_no'],
                    'Item #':        item_str,
                    'Vendor Item #': vendor_item_str,
                    'Job Bag #':     job_bag_str,
                    'Description':   desc_str,
                    'Dia Qlty':      dia_str,
                    'Metal_PDF':     metal_str,
                    'Size':          size_str,
                    'Quantity':      qty,
                    '_has_inch':     has_inch,
                })

            for (y, rw) in notes_rows:
                line_text = ' '.join(w['text'] for w in rw)
                if 'RightClick' in line_text or 'Copyright' in line_text or line_text.startswith('2023') or line_text.startswith('2024') or line_text.startswith('2025') or line_text.startswith('2026'):
                    continue
                all_notes.append(line_text)

    notes_text = '\n'.join(all_notes).strip()
    return all_items_raw, notes_text


# ─────────────────────────────────────────────
#  VIMCO LOGIC (FROM VIMCO.py)
# ─────────────────────────────────────────────

def _build_style_code(base, item_size, tone, cpi="", has_inch=False):
    base = str(base).strip() if base else ''
    item_size = str(item_size).strip() if item_size else ''
    tone = str(tone).strip().upper() if tone else ''
    cpi = str(cpi).strip().upper() if cpi else ''
    if not base or base.upper() == 'NAN':
        return ''
    
    # Check if size explicitly contains an inch marker (" or ” or INCH or word boundary IN),
    # but prioritize passed has_inch flag from PDF extraction
    if not has_inch:
        has_inch = bool(re.search(r'INCH|\"|”|\bIN\b', item_size, flags=re.IGNORECASE))
    
    size_num = re.sub(r'^(?:UP|US|EU|IT|UT|TS|IS)\s*', '', item_size, flags=re.IGNORECASE).strip()
    size_num = re.sub(r'\s*INCH\s*$', '', size_num, flags=re.IGNORECASE).strip()
    size_num = re.sub(r'[^0-9.]', '', size_num)
    # String-based normalization (safer than float to avoid rounding surprises):
    # - strip trailing zeros after a decimal point
    # - strip the decimal point itself if nothing remains after removing zeros
    if size_num and '.' in size_num:
        size_num = size_num.rstrip('0').rstrip('.')

    if tone and len(tone) > 1 and tone not in ('PT', 'AG', 'TT'):
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
    elif tone == 'TT':
        suffix = f"{size_num}{in_part}TT" if size_num else (f"{in_part}TT" if in_part else 'TT')
    else:
        suffix = size_num

    # Append SM suffix if Customer Production Instruction contains "SEMI MOUNT" or "SEMI-"
    if "SEMI MOUNT" in cpi or "SEMI-" in cpi or "SEMI" in cpi:
        suffix = f"{suffix}SM"

    return f"{base}-{suffix}" if suffix else base

_ITEM_SIZE_LOOKUP = None

def _get_size_lookup():
    global _ITEM_SIZE_LOOKUP
    if _ITEM_SIZE_LOOKUP is not None:
        return _ITEM_SIZE_LOOKUP
    _ITEM_SIZE_LOOKUP = {}
    try:
        mst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ItemSize_Mst.xlsx')
        if os.path.exists(mst):
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
    m = re.match(r'^(\d+(?:\.\d+)?)\s*(?:INCH|\"|”)$', s, re.IGNORECASE)
    if m:
        return f"{float(m.group(1)):.2f}inch"
    return re.sub(r'\s+', '', s).lower()

def _map_item_size(raw):
    if not raw or str(raw).strip().upper() in ('', 'NAN'):
        return raw
    lookup = _get_size_lookup()
    key = _normalize_size_key(str(raw).strip())
    return lookup.get(key, raw)

def map_metal(text):
    if pd.isna(text): return ''
    text = str(text).upper()
    # For two-tone items (14KTT), extract the actual color metal (e.g. 14KWY → G14Y, 14KY → G14Y)
    if '14KTT' in text:
        # Look for the secondary metal indicator like 14KWY, 14KYW, 14KY, 14KW etc.
        m = re.search(r'14K([WY])Y|14KY([W])?|14K([WY])W', text)
        if m:
            # Determine the dominant color: Y takes priority in two-tone
            if 'Y' in (m.group(1) or '') + (m.group(2) or '') + (m.group(3) or ''):
                return 'G14Y'
            else:
                return 'G14W'
        # Fallback: scan for 14KWY or 14KY explicitly
        if '14KWY' in text or '14KYW' in text:
            return 'G14Y'
        if '14KY' in text: return 'G14Y'
        if '14KW' in text: return 'G14W'
        return 'G14Y'  # default two-tone to yellow
    if '14KY' in text: return 'G14Y'
    elif '14KW' in text: return 'G14W'
    elif '18KY' in text: return 'G18Y'
    elif '18KW' in text: return 'G18W'
    elif '10KY' in text: return 'G10Y'
    elif '10KW' in text: return 'G10W'
    elif 'PT' in text: return 'PC95'
    return ''

def map_tone(metal):
    if metal == 'PC95': return 'PT'
    elif metal: return metal[-1]
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
    dia_quality = str(row.get('DiaQuality', '')).strip().upper() if pd.notna(row.get('DiaQuality', '')) else ""
    return f"{row.get('OrderGroup', '')},{row.get('SKUNo', '')},{metal_descr}, DIA QLTY - {dia_quality}"

def design_production_instruction(row):
    tone = str(row['Tone']).upper()
    cpi = str(row['CustomerProductionInstruction']).upper() if pd.notna(row['CustomerProductionInstruction']) else ''
    semi_present = "SEMI" in cpi or "SEMI-" in cpi
    if tone == "W" and semi_present: return "SEMI MOUNT, WHITE RODIUM"
    elif tone == "Y" and semi_present: return "SEMI MOUNT, NO RODIUM"
    elif tone == "W" and not semi_present: return "WHITE RODIUM"
    elif tone == "Y" and not semi_present: return "NO RODIUM"
    elif tone == "PT" and semi_present: return "SEMI MOUNT, NO RODIUM"
    elif tone == "PT" and not semi_present: return "NO RODIUM"
    return ""

def extract_stone_weight(text):
    if pd.isna(text): return ""
    match = re.search(r'(\d+\.\d+|\d+)\s*(CT|CARAT|CTS|CTW|A|PCS)?', str(text), re.IGNORECASE)
    if match: return match.group(1)
    return ""

def metal_stamp_text(metal):
    mapping = {
        "G14Y": "14K", "G14W": "14K", "G18Y": "18K", "G18W": "18K",
        "G10Y": "10K", "G10W": "10K", "PC95": "PT950"
    }
    return mapping.get(metal, "")

def generate_stamp_instruction(row):
    metal_text_val = metal_stamp_text(row['Metal'])
    stone_weight = extract_stone_weight(row['CustomerProductionInstruction'])
    if metal_text_val and stone_weight:
        return f"{metal_text_val} V ON ONE SIDE AND {stone_weight} A ON OTHER SIDE"
    elif metal_text_val:
        return f"{metal_text_val} V ON ONE SIDE"
    return ""

def extract_size_from_sku(sku):
    if pd.isna(sku): return None
    match = re.search(r'SZ(\d+(\.\d+)?)$', str(sku).upper())
    if match: return match.group(1)
    return None


def process_vimco_v2_file(filepath: str, output_dir: str, order_group: str = '', priority: str = '5 day') -> tuple:
    try:
        raw_text = _read_pdf_text(filepath)
        header = _parse_header(raw_text)

        line_items, notes_text = _parse_line_items_by_position(filepath)

        if not line_items:
            return False, None, "Could not extract any line items from the PDF.", None

        item_po_no = header.get('Order #', '')
        
        rows = []
        for sr_no, item in enumerate(line_items, start=1):
            sku_no = item.get('Item #', '')
            style_code = item.get('Vendor Item #', '')
            
            if style_code:
                style_code = str(style_code).split('-')[0].strip()

            cpi = item.get('Description', '')
            metal_pdf = item.get('Metal_PDF', '')
            size = item.get('Size', '')
            qty = item.get('Quantity', 1)
            dia_quality = item.get('Dia Qlty', '')
            has_inch = item.get('_has_inch', False)

            row = {
                'SrNo': sr_no,
                'StyleCode': style_code,
                'ItemSize': size,
                'OrderQty': qty,
                '_has_inch': has_inch,
                'CustomerProductionInstruction': cpi,
                'MetalPDF_Raw': metal_pdf,
                'SKUNo': sku_no,
                'DiaQuality': dia_quality,
                'OrderGroup': order_group,
                'ItemPoNo': item_po_no
            }
            rows.append(row)

        df = pd.DataFrame(rows)

        df['OrderItemPcs'] = 1
        # Combine CPI and MetalPDF_Raw for map_metal so it catches 14KW etc.
        df['Metal'] = (df['CustomerProductionInstruction'] + ' ' + df['MetalPDF_Raw']).apply(map_metal)
        df['Tone'] = df['Metal'].apply(map_tone)
        # For TT (two-tone) items: Tone stays as derived from Metal (e.g. 'Y' from G14Y),
        # but we flag them so _build_style_code appends 'TT' as the suffix.
        df['_is_tt'] = df['CustomerProductionInstruction'].str.upper().str.contains('14KTT', na=False)
        df['ItemRefNo'] = ''
        df['StockType'] = ''
        df['Priority'] = priority
        df['MakeType'] = ''
        df['StampInstruction'] = df.apply(generate_stamp_instruction, axis=1)
        df['Certificate'] = ''
        df['SpecialRemarks'] = df.apply(special_remarks, axis=1)
        df['DesignProductionInstruction'] = df.apply(design_production_instruction, axis=1)

        new_cols_after_sku = [
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate', 'SetPrice', 'StoneQuality'
        ]
        for col in new_cols_after_sku:
            df[col] = ''

        # Only default ZR style sizes to 7 if no actual size was extracted from PDF
        zr_mask = df['StyleCode'].astype(str).str.upper().str.startswith('ZR')
        empty_size_mask = df['ItemSize'].astype(str).str.strip() == ''
        df.loc[zr_mask & empty_size_mask, 'ItemSize'] = '7'

        df['ExtractedSize'] = df['SKUNo'].apply(extract_size_from_sku)
        mask = (df['ItemSize'] == '') & df['ExtractedSize'].notna()
        df.loc[mask, 'ItemSize'] = df.loc[mask, 'ExtractedSize']
        
        # Preserve inch indicator (") without forcing a US prefix
        def _format_raw_size(x):
            if pd.isna(x) or str(x).strip() == "":
                return ""
            x_str = str(x).strip()
            if '"' in x_str or '”' in x_str or 'INCH' in x_str.upper():
                return x_str
            return f"US{x_str}"

        df['ItemSize'] = df['ItemSize'].apply(_format_raw_size)
        df.drop(columns=['ExtractedSize'], inplace=True)

        df['ItemSize'] = df['ItemSize'].apply(_map_item_size)
        
        # Pass CustomerProductionInstruction into _build_style_code to check for "SEMI MOUNT" / "SEMI-",
        # and _has_inch flag from PDF extraction
        df['StyleCode'] = df.apply(
            lambda row: _build_style_code(
                row['StyleCode'], row['ItemSize'],
                'TT' if row.get('_is_tt') else ('AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])),
                row['CustomerProductionInstruction'],
                row.get('_has_inch', False)
            ), axis=1
        )

        # Clean up ItemSize column for final output post style-code mapping
        df['ItemSize'] = df['ItemSize'].apply(
            lambda x: re.sub(r'[^A-Za-z0-9\-.]', '', str(x)) if pd.notna(x) else ''
        )

        df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''
        df.drop(columns=['_is_tt', '_has_inch'], inplace=True)
        
        # Load VIMCO_CS.xlsx to get ItemSize from Client Style Master
        missing_style_codes = []
        try:
            cs_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_220526', 'VIMCO_CS.xlsx')
            if os.path.exists(cs_file):
                cs_df = pd.read_excel(cs_file, header=2)
                # Clean up column names
                cs_df.columns = ['Select', 'Client Style No', 'ItemSize', 'Style No', 'Client Code', 'Category']
                # Build a lookup dictionary: Client Style No -> ItemSize
                cs_lookup = {}
                for _, cs_row in cs_df.iterrows():
                    cs_style = str(cs_row['Client Style No']).strip()
                    cs_size = str(cs_row['ItemSize']).strip() if pd.notna(cs_row['ItemSize']) else ''
                    if cs_style and cs_size:
                        cs_lookup[cs_style] = cs_size
                # Update ItemSize and collect missing style codes
                def get_item_size(sc):
                    sc_str = str(sc).strip()
                    if sc_str in cs_lookup:
                        return cs_lookup[sc_str]
                    else:
                        if sc_str not in missing_style_codes:
                            missing_style_codes.append(sc_str)
                        return ''
                df['ItemSize'] = df['StyleCode'].apply(get_item_size)
        except Exception as e:
            pass  # Skip ItemSize mapping if CS file isn't accessible

        final_columns = [
            'SrNo', 'StyleCode', 'ItemSize', 'OrderQty', 'OrderItemPcs', 'Metal', 'Tone',
            'ItemPoNo', 'ItemRefNo', 'StockType', 'Priority', 'MakeType',
            'CustomerProductionInstruction', 'SpecialRemarks', 'DesignProductionInstruction',
            'StampInstruction', 'OrderGroup', 'Certificate', 'SKUNo'
        ] + new_cols_after_sku

        for col in final_columns:
            if col not in df.columns:
                df[col] = ''
        df = df[final_columns]

        base_name = re.sub(r'[^\w\-]', '_', os.path.splitext(os.path.basename(filepath))[0])
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"VIMCO_FORMAT_{base_name}_{timestamp}.xlsx"
        
        if output_dir:
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
            output_path = os.path.join(output_dir, output_filename)
        else:
            output_path = output_filename
            
        df.to_excel(output_path, index=False)
        
        return True, output_path, None, df

    except Exception as exc:
        return False, None, str(exc), None

if __name__ == "__main__":
    import argparse
    import glob
    
    parser = argparse.ArgumentParser(description='Process VIMCO PDF PO files to standardized Excel format')
    parser.add_argument('--input', '-i', required=True, help='Input PDF file or folder path')
    parser.add_argument('--output', '-o', help='Output folder path (optional)')
    parser.add_argument('--order-group', '-g', default='', help='OrderGroup value (optional)')
    parser.add_argument('--priority', '-r', default='5 day', help='Priority value (default: "5 day")')
    parser.add_argument('--batch', '-b', action='store_true', help='Process all PDF files in input folder')
    
    args = parser.parse_args()
    
    if args.batch:
        if not os.path.isdir(args.input):
            print("❌ Input path must be a folder when using --batch")
            exit(1)
        pdfs = glob.glob(os.path.join(args.input, "*.pdf"))
        for pdf in pdfs:
            print(f"Processing {pdf}...")
            success, out_path, err, df = process_vimco_v2_file(pdf, args.output, args.order_group, args.priority)
            if success:
                print(f"✅ Created {out_path}")
            else:
                print(f"❌ Error processing {pdf}: {err}")
    else:
        if not os.path.isfile(args.input):
            print("❌ Input path must be a file when not using --batch")
            exit(1)
        success, out_path, err, df = process_vimco_v2_file(args.input, args.output, args.order_group, args.priority)
        if success:
            print(f"✅ Created {out_path}")
        else:
            print(f"❌ Error: {err}")