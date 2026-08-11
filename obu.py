import os
import re
import pandas as pd
import pdfplumber


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'EU52', 'W') -> 'VR1943EEA-52WG'
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
_OBU_CS_CACHE = None
_STYLE_PATTERN_CACHE = {}

# Article-code line tokens: SKU is digit-led (9-DD038-…); style reference is letter-led (ER016475-…).
_OBU_SKU_RE = re.compile(r'^\d+-', re.IGNORECASE)
_OBU_STYLE_RE = re.compile(r'^[A-Z]{2,}\d', re.IGNORECASE)
_OBU_STYLE_REF_RE = re.compile(
    r'^(?P<base>[A-Z]{2,}\d+[A-Z]*)'
    r'(?:-(?P<suffix>WG|YG|RG|[WYPR]G|[WYPR]))?'
    r'(?:-\d+)*$',
    re.IGNORECASE,
)


def _get_obu_cs():
    global _OBU_CS_CACHE
    if _OBU_CS_CACHE is not None:
        return _OBU_CS_CACHE
    try:
        cs_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "CS_220526",
            "OBU_CS.xlsx"
        )
        if os.path.exists(cs_path):
            _OBU_CS_CACHE = pd.read_excel(cs_path, dtype=str)
            _OBU_CS_CACHE.columns = [str(c).strip() for c in _OBU_CS_CACHE.columns]
    except Exception:
        _OBU_CS_CACHE = pd.DataFrame()
    return _OBU_CS_CACHE


def _metal_matches_cs(pdf_metal: str, cs_metal: str) -> bool:
    """Return True when PDF-derived metal matches an OBU_CS Base Metal entry."""
    pdf_metal = str(pdf_metal).strip().upper()
    cs_metal = str(cs_metal).strip().upper()
    if not pdf_metal or not cs_metal:
        return False
    if pdf_metal == cs_metal:
        return True
    if pdf_metal.rstrip('Z') == cs_metal.rstrip('Z'):
        return True
    pdf_tone = _map_tone_from_metal(pdf_metal)
    cs_tone = _map_tone_from_metal(cs_metal)
    return bool(pdf_tone and pdf_tone == cs_tone)


def _lookup_obu_cs(style_no: str, item_size: str, sku: str = "", pdf_metal: str = ""):
    """
    Look up Client Style No, Base Metal, and ItemSize from OBU_CS.xlsx.
    Returns a tuple: (client_style_no, base_metal, cs_item_size)
    """
    df = _get_obu_cs()
    if df.empty:
        return (None, None, None)

    # First try matching by Style Alias No (which is exactly the SKU from PDF)
    sku = str(sku).strip()
    if sku and "Style Alias No" in df.columns:
        alias_filtered = df[df["Style Alias No"].str.strip() == sku]
        if not alias_filtered.empty:
            client_style = str(alias_filtered.iloc[0]["Client Style No"]).strip() if "Client Style No" in df.columns else None
            base_metal = str(alias_filtered.iloc[0]["Base Metal"]).strip() if "Base Metal" in df.columns else None
            cs_item_size = str(alias_filtered.iloc[0]["ItemSize"]).strip() if "ItemSize" in df.columns else None
            return (client_style, base_metal, cs_item_size)

    # If no alias match, try matching by Style No
    style_no = str(style_no).strip()
    if not style_no or "Style No" not in df.columns:
        return (None, None, None)
    filtered = df[df["Style No"].str.strip() == style_no]
    if filtered.empty:
        return (None, None, None)

    pdf_metal = str(pdf_metal).strip().upper()
    item_size = str(item_size).strip() if item_size else ''

    matched_row = None
    if "ItemSize" in df.columns and item_size:
        size_filtered = filtered[filtered["ItemSize"].str.strip() == item_size]
        if not size_filtered.empty:
            candidates = size_filtered
            if pdf_metal and "Base Metal" in df.columns:
                metal_filtered = candidates[candidates["Base Metal"].apply(lambda m: _metal_matches_cs(pdf_metal, m))]
                if not metal_filtered.empty:
                    candidates = metal_filtered
            if len(candidates) == 1:
                matched_row = candidates.iloc[0]
            elif len(candidates) > 1 and pdf_metal and "Base Metal" in df.columns:
                metal_filtered = candidates[candidates["Base Metal"].apply(lambda m: _metal_matches_cs(pdf_metal, m))]
                if len(metal_filtered) == 1:
                    matched_row = metal_filtered.iloc[0]
    elif pdf_metal and "Base Metal" in df.columns:
        metal_filtered = filtered[filtered["Base Metal"].apply(lambda m: _metal_matches_cs(pdf_metal, m))]
        if len(metal_filtered) == 1:
            matched_row = metal_filtered.iloc[0]

    if matched_row is None:
        return (None, None, None)

    client_style = str(matched_row["Client Style No"]).strip() if "Client Style No" in df.columns else None
    base_metal = str(matched_row["Base Metal"]).strip() if "Base Metal" in df.columns else None
    cs_item_size = str(matched_row["ItemSize"]).strip() if "ItemSize" in df.columns else None
    return (client_style, base_metal, cs_item_size)


def _map_tone_from_metal(metal):
    """
    Map a Metal value (from OBU_CS.xlsx) to a Tone value for output.
    """
    metal = str(metal).strip().upper()
    if metal == 'PC95':
        return 'PT'
    elif metal == 'AG925':
        return ''
    elif metal.endswith('W') or metal.endswith('WZ'):
        return 'W'
    elif metal.endswith('Y') or metal.endswith('YZ'):
        return 'Y'
    elif metal.endswith('P') or metal.endswith('PZ'):
        return 'P'
    elif 'W' in metal:
        return 'W'
    elif 'Y' in metal:
        return 'Y'
    elif 'P' in metal:
        return 'P'
    else:
        return ''


def _map_metal_from_text(text, sku=''):
    """Parse metal code from PO description text and SKU (e.g. 14 KY -> G585YZ)."""
    text_u = str(text or '').upper()
    sku_u = str(sku or '').upper()
    if not text_u and not sku_u:
        return ''

    compact = re.sub(r'[^A-Z0-9]', '', text_u)
    has_585 = bool(re.search(r'\b585\b', text_u) or 'STAMP585' in compact)

    is_yellow = bool(
        re.search(r'(?:10|14|18)KY', compact)
        or re.search(r'(?:10|14|18)K?Y(?![A-Z])', compact)
        or re.search(r'-YG(?:-|$)', sku_u)
    )
    is_white = bool(
        re.search(r'(?:10|14|18)KW', compact)
        or re.search(r'(?:10|14|18)K?W(?![A-Z])', compact)
        or re.search(r'-WG(?:-|$)', sku_u)
    )
    is_rose = bool(re.search(r'(?:10|14|18)KR', compact) or re.search(r'-RG(?:-|$)', sku_u))

    karat = 14
    if re.search(r'18K', compact):
        karat = 18
    elif re.search(r'10K', compact):
        karat = 10

    if 'PT' in compact or 'PLAT' in compact:
        return 'PC95'

    if is_yellow and not is_white:
        if has_585 and karat == 14:
            return 'G585YZ'
        return f'G{karat}Y'
    if is_white and not is_yellow:
        if has_585 and karat == 14:
            return 'G585WZ'
        return f'G{karat}W'
    if is_rose:
        if has_585 and karat == 14:
            return 'G585PZ'
        return f'G{karat}P'
    return ''


def _tone_from_sku(sku):
    """Extract tone letter from SKU tokens like YG / WG."""
    sku_u = str(sku or '').upper()
    if re.search(r'-YG(?:-|$)', sku_u):
        return 'Y'
    if re.search(r'-WG(?:-|$)', sku_u):
        return 'W'
    if re.search(r'-RG(?:-|$)', sku_u):
        return 'P'
    return ''


def _normalize_tone_letter(tone):
    """Normalize WG/YG/… tone tokens to a single tone letter."""
    tone = str(tone or '').strip().upper()
    if not tone:
        return ''
    if tone in ('PT', 'AG'):
        return tone
    if tone in ('WG', 'YG', 'RG'):
        return tone[0]
    if len(tone) > 1:
        return tone[0] if tone[0] in ('W', 'Y', 'P', 'R') else ''
    return tone if tone in ('W', 'Y', 'P', 'R') else ''


def _parse_style_reference(raw):
    """
    Parse a PDF style reference such as ER016475-WG-0 or RG058855-YG
    into (base_style_no, tone_letter).
    """
    raw = re.sub(r'[^A-Z0-9\-]', '', str(raw or '').strip().upper())
    if not raw:
        return '', ''
    match = _OBU_STYLE_REF_RE.match(raw)
    if not match:
        if _OBU_STYLE_RE.match(raw):
            return raw, ''
        return '', ''
    base = match.group('base')
    suffix = (match.group('suffix') or '').upper()
    return base, _normalize_tone_letter(suffix)


def _item_size_from_sku(sku, style_no=''):
    """
    Extract EU ring size from SKU when present (e.g. WG-80-57 -> EU57).
    Single trailing numbers (WG-30 carat weight for earrings) are ignored.
    """
    sku = str(sku or '').strip().upper()
    if not sku:
        return ''
    match = re.search(r'(?:WG|YG|RG)-(\d+(?:-\d+)*)', sku, re.IGNORECASE)
    if not match:
        return ''
    nums = re.findall(r'\d+', match.group(1))
    if len(nums) < 2:
        return ''
    size = nums[-1]
    pattern = _infer_style_suffix_pattern(style_no) if style_no else 'LEGACY'
    if pattern == 'EU_TONE_GEM' or not style_no:
        return f'EU{size}'
    return ''


def _parse_article_line(line):
    """
    Parse the PDF article-code row into SKU, style number, tone, item size,
    and the raw reference token (Your reference column).
    """
    code_token_re = re.compile(r'[A-Z0-9][A-Z0-9\-]*[A-Z0-9]')
    tokens = [
        re.sub(r'[^A-Z0-9\-]', '', token)
        for token in code_token_re.findall(str(line or ''))
    ]

    sku_full = ''
    style_raw = ''
    for token in tokens:
        if token.isdigit():
            continue
        if _OBU_STYLE_RE.match(token):
            style_raw = token
        elif _OBU_SKU_RE.match(token) and not sku_full:
            sku_full = token

    style_code, tone = _parse_style_reference(style_raw)
    if not tone:
        tone = _tone_from_sku(sku_full)
    item_size = _item_size_from_sku(sku_full, style_code)
    return sku_full, style_code, tone, item_size, style_raw


def _infer_style_suffix_pattern(style_no):
    """
    Infer how Client Style No values are formed for a Style No using OBU_CS.
    Returns one of: EU_TONE_GEM, TONE_GEM, TONE_PAIR, LEGACY.
    """
    style_no = str(style_no or '').strip()
    if not style_no:
        return 'LEGACY'
    if style_no in _STYLE_PATTERN_CACHE:
        return _STYLE_PATTERN_CACHE[style_no]

    pattern = 'LEGACY'
    df = _get_obu_cs()
    if not df.empty and 'Style No' in df.columns:
        rows = df[df['Style No'].str.strip() == style_no]
        suffixes = []
        for val in rows.get('Client Style No', pd.Series()).dropna():
            client_style = str(val).strip()
            if client_style.startswith(style_no):
                suffixes.append(client_style[len(style_no):])

        if suffixes:
            eu_gem = sum(1 for suffix in suffixes if re.match(r'-EU\d+[WYPR]GEM$', suffix, re.I))
            tone_gem = sum(1 for suffix in suffixes if re.match(r'-[WYPR]GEM$', suffix, re.I))
            tone_pair = sum(1 for suffix in suffixes if re.match(r'-(WG|YG|RG)$', suffix, re.I))
            if eu_gem and eu_gem >= max(tone_gem, tone_pair):
                pattern = 'EU_TONE_GEM'
            elif tone_pair and tone_pair >= tone_gem:
                pattern = 'TONE_PAIR'
            elif tone_gem or any(suffix == '-AGEM' for suffix in suffixes):
                pattern = 'TONE_GEM'

    _STYLE_PATTERN_CACHE[style_no] = pattern
    return pattern


def _build_obu_style_code(style_no, item_size, tone, metal=''):
    """
    Build StyleCode for OBU output using the suffix pattern from OBU_CS.
    """
    style_no = str(style_no or '').strip()
    item_size = str(item_size or '').strip()
    tone = _normalize_tone_letter(tone)
    metal = str(metal or '').strip().upper()
    if not style_no or style_no.upper() == 'NAN':
        return ''

    pattern = _infer_style_suffix_pattern(style_no)
    if pattern == 'EU_TONE_GEM' and item_size and tone in ('W', 'Y', 'P', 'R'):
        return f'{style_no}-{item_size}{tone}GEM'
    if pattern == 'TONE_GEM':
        if metal == 'AG925':
            return f'{style_no}-AGEM'
        if tone in ('W', 'Y', 'P', 'R'):
            return f'{style_no}-{tone}GEM'
    if pattern == 'TONE_PAIR' and tone in ('W', 'Y', 'P', 'R'):
        pair_map = {'W': 'WG', 'Y': 'YG', 'P': 'RG', 'R': 'RG'}
        return f'{style_no}-{pair_map[tone]}'

    return _build_style_code(style_no, item_size, tone)


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


def _read_pdf_text(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join([page.extract_text() or '' for page in pdf.pages])


def _extract_order_date(text: str) -> str:
    match = re.search(r'Order date\s*:\s*(.+)', text, re.IGNORECASE)
    return match.group(1).strip() if match else ''


def process_obu_file(input_path: str, output_dir: str):
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()
        if ext == '.pdf':
            text = _read_pdf_text(input_path)
            code_token_re = re.compile(r'[A-Z0-9][A-Z0-9\-]*[A-Z0-9]')
            po_re = re.compile(r'PO#\s*:\s*(\d+)')
            article_header_re = re.compile(r'^Article code', re.IGNORECASE)
            quantity_line_re = re.compile(r'^(\d+)\s+(\d+)$')

            po_match = po_re.search(text)
            item_po_no = po_match.group(1) if po_match else ''
            order_date = _extract_order_date(text)

            lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
            blocks, current = [], []
            for ln in lines:
                if article_header_re.match(ln):
                    if current:
                        blocks.append(current)
                        current = []
                    current.append(ln)
                elif current:
                    current.append(ln)
            if current:
                blocks.append(current)

            items = []
            for b_index, block in enumerate(blocks):
                is_last_block = (b_index == len(blocks) - 1)
                btxt = "\n".join(block)

                codes = []
                article_line = ''
                for ln in block[1:4]:
                    if _OBU_SKU_RE.search(ln) or _OBU_STYLE_RE.search(ln):
                        article_line = ln
                        break
                if article_line:
                    sku_full, style_code, tone, item_size, style_raw = _parse_article_line(article_line)
                else:
                    sku_full, style_code, tone, item_size, style_raw = '', '', '', '', ''
                item_ref_no = style_raw if style_raw and style_raw != sku_full else ''

                sr_no = ''
                order_qty = ''
                order_item_pcs = ''
                try:
                    desc_idx = next(i for i, ln in enumerate(block) if ln.lower().startswith('description'))
                except StopIteration:
                    desc_idx = None
                if desc_idx is not None:
                    for ln in block[desc_idx:desc_idx+5]:
                        qm = quantity_line_re.match(ln)
                        if qm:
                            sr_no = qm.group(1)
                            order_qty = qm.group(2)
                            order_item_pcs = order_qty
                            break

                desc_lines = []
                for ln in block:
                    if article_header_re.match(ln) or quantity_line_re.match(ln) or ln.lower().startswith('description'):
                        continue
                    if code_token_re.fullmatch(ln.replace(' ', '')):
                        continue
                    desc_lines.append(ln)
                full_desc = ' '.join(desc_lines)
                if is_last_block:
                    pot_idx = re.search(r'Purchase order Total', full_desc, flags=re.IGNORECASE)
                    if pot_idx:
                        full_desc = full_desc[:pot_idx.start()].strip()
                split_match = re.search(r'\b(stamp\b.*)', full_desc, flags=re.IGNORECASE)
                if split_match:
                    customer_instr = full_desc[:split_match.start()].strip()
                    stamp_instr = full_desc[split_match.start():].strip()
                else:
                    customer_instr = full_desc
                    stamp_instr = ''
                customer_instr = re.sub(r'\s*\band\s*$', '', customer_instr, flags=re.IGNORECASE).strip()

                certificate = ''
                if sku_full:
                    nums = re.findall(r'\d+', sku_full)
                    if nums and nums[-1] == '100':
                        certificate = 'IGI Certified'

                sku_no = ''
                if sku_full:
                    sku_no = sku_full

                pdf_metal = _map_metal_from_text(full_desc, sku_full)
                if not tone:
                    tone = _tone_from_sku(sku_full)
                if not tone and pdf_metal:
                    tone = _map_tone_from_metal(pdf_metal)

                client_style, metal, cs_item_size = _lookup_obu_cs(
                    style_code, item_size, sku_full, pdf_metal=pdf_metal
                )
                final_style_code = client_style if client_style else style_code
                final_item_size = cs_item_size if cs_item_size else item_size
                use_client_style = client_style is not None
                use_cs_item_size = cs_item_size is not None

                if not metal and pdf_metal:
                    metal = pdf_metal
                if metal:
                    tone = _map_tone_from_metal(metal)
                elif pdf_metal:
                    metal = pdf_metal
                    tone = _map_tone_from_metal(pdf_metal)
                
                items.append({
                    'SrNo': sr_no,
                    'StyleCode': final_style_code,
                    'ItemSize': final_item_size,
                    'OrderQty': order_qty,
                    'OrderItemPcs': 1,
                    'Metal': metal,
                    'Tone': tone,
                    'ItemPoNo': item_po_no,
                    'ItemRefNo': item_ref_no,
                    'StockType': '',
                    'MakeType': '',
                    'CustomerProductionInstruction': customer_instr,
                    'SpecialRemarks': '',
                    'DesignProductionInstruction': '',
                    'StampInstruction': stamp_instr,
                    'OrderGroup': '',
                    'Certificate': certificate,
                    'SKUNo': sku_no,
                    'Basestoneminwt': '',
                    'Basestonemaxwt': '',
                    'Basemetalminwt': '',
                    'Basemetalmaxwt': '',
                    'Productiondeliverydate': '',
                    'Expecteddeliverydate': '',
                    'Blank': '',
                    'SetPrice': '',
                    'StoneQuality': 'VVS+' if re.search(r'\bVVS\+\b', btxt) else '',
                    'Date': order_date,
                    '_use_client_style': use_client_style,
                    '_use_cs_item_size': use_cs_item_size,
                })

            columns_order = [
                'SrNo','StyleCode','ItemSize','OrderQty','OrderItemPcs','Metal','Tone','ItemPoNo','ItemRefNo',
                'StockType','MakeType','CustomerProductionInstruction','SpecialRemarks','DesignProductionInstruction',
                'StampInstruction','OrderGroup','Certificate','SKUNo','Basestoneminwt','Basestonemaxwt','Basemetalminwt',
                'Basemetalmaxwt','Productiondeliverydate','Expecteddeliverydate','Blank', 'SetPrice','StoneQuality',
                'Date'
            ]
            df = pd.DataFrame(items, columns=columns_order + ['_use_client_style', '_use_cs_item_size'])
            
            # Only apply _map_item_size to rows that didn't use an ItemSize from CS
            mask_size = ~df['_use_cs_item_size']
            df.loc[mask_size, 'ItemSize'] = df.loc[mask_size, 'ItemSize'].apply(_map_item_size)
            
            # Only apply _build_obu_style_code to rows that didn't use a Client Style No from CS
            mask_style = ~df['_use_client_style']
            df.loc[mask_style, 'StyleCode'] = df.loc[mask_style].apply(
                lambda row: _build_obu_style_code(
                    row['StyleCode'],
                    row['ItemSize'],
                    'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone']),
                    str(row.get('Metal', '')),
                ),
                axis=1,
            )
            
            # Drop the temporary columns
            df.drop(['_use_client_style', '_use_cs_item_size'], axis=1, inplace=True)
            df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''
            output_path = os.path.join(output_dir, f"{base_name}_OBU_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        elif ext in ['.xlsx', '.xls', '.csv']:
            df = pd.read_excel(input_path) if ext in ['.xlsx', '.xls'] else pd.read_csv(input_path)
            output_path = os.path.join(output_dir, f"{base_name}_OBU_PASSTHROUGH.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None


