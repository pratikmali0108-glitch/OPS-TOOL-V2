import os
import re
import pandas as pd
import pdfplumber
from typing import List, Dict, Optional


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
_CLIENT_STYLE_DF = None


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


def _get_client_style_df() -> pd.DataFrame:
    """Load and cache the full Client Style dataframe from AMBITION_ANERI_CS_100826.xlsx, ANAYA_CS_100826.xlsx, or fallback."""
    global _CLIENT_STYLE_DF
    if _CLIENT_STYLE_DF is not None:
        return _CLIENT_STYLE_DF
    _CLIENT_STYLE_DF = pd.DataFrame()
    try:
        # Try CS_100826 folder first - AMBITION file
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'AMBITION_ANERI_CS_100826.xlsx')
        if not os.path.exists(cs_path):
            # Try ANAYA file (some POs may reference ANAYA styles)
            cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'ANAYA_CS_100826.xlsx')
        if not os.path.exists(cs_path):
            # Fallback to old location
            cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Ambition_CS.xlsx')
        _CLIENT_STYLE_DF = pd.read_excel(cs_path)
    except Exception:
        pass
    return _CLIENT_STYLE_DF


def _get_client_style_list() -> List[str]:
    """Load and cache the approved Client Style No list from AMBITION_ANERI_CS_100826.xlsx, ANAYA_CS_100826.xlsx, or fallback to Ambition_CS.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    df = _get_client_style_df()
    if df is not None and not df.empty:
        col = 'Client Style No'
        if col in df.columns:
            _CLIENT_STYLE_LIST = [
                str(v).strip() for v in df[col].dropna()
                if str(v).strip() and str(v).strip().upper() != 'NAN'
            ]
    return _CLIENT_STYLE_LIST


def _reset_client_style_cache():
    """Invalidate the cached client style list and dataframe so next call reloads from disk."""
    global _CLIENT_STYLE_LIST, _CLIENT_STYLE_DF
    _CLIENT_STYLE_LIST = None
    _CLIENT_STYLE_DF = None


def _lookup_client_style_with_size(base_code: str, item_size: str, tone: str) -> tuple[str, str]:
    """
    Enhanced lookup that matches StyleCode AND ItemSize against the master CS file.
    
    Returns: (matched_style_code, matched_item_size)
    
    Strategy:
      1. Build the expected style code from base + size + tone
      2. Look for exact match in master list
      3. If ItemSize doesn't match master, try alternative size formats (US07, UP07, etc.)
      4. Return the matched style code and its corresponding ItemSize from master
    """
    df = _get_client_style_df()
    if df is None or df.empty or not base_code:
        built_code = _build_style_code(base_code, item_size, tone)
        return built_code, item_size
    
    # Check if required columns exist
    if 'Client Style No' not in df.columns or 'Style No' not in df.columns:
        built_code = _build_style_code(base_code, item_size, tone)
        return built_code, item_size
    
    # Filter to rows matching the base style code
    base_upper = base_code.upper()
    style_matches = df[df['Style No'].astype(str).str.upper() == base_upper].copy()
    
    if style_matches.empty:
        # No match found for base code, return built code
        built_code = _build_style_code(base_code, item_size, tone)
        return built_code, item_size
    
    # Build the target style code
    built_code = _build_style_code(base_code, item_size, tone)
    built_upper = built_code.upper()
    
    # Try to find exact match
    exact_match = style_matches[style_matches['Client Style No'].astype(str).str.upper() == built_upper]
    if not exact_match.empty:
        matched_row = exact_match.iloc[0]
        matched_style = str(matched_row['Client Style No']).strip()
        matched_size = str(matched_row['ItemSize']).strip() if 'ItemSize' in matched_row and pd.notna(matched_row['ItemSize']) else item_size
        return matched_style, matched_size
    
    # If no exact match, try different ItemSize prefix variations (US07, UP07, etc.)
    if item_size and 'ItemSize' in df.columns:
        # Extract numeric part from item_size
        size_num_match = re.search(r'(\d+(?:\.\d+)?)', str(item_size))
        if size_num_match:
            size_num = size_num_match.group(1)
            try:
                f = float(size_num)
                # Generate alternative prefixes
                prefixes = ['US', 'UP', 'EU', 'IT', 'UT', 'TS', 'IS', '']
                for prefix in prefixes:
                    if f.is_integer():
                        test_sizes = [f"{prefix}{int(f):02d}", f"{prefix}{int(f)}"]
                    else:
                        test_sizes = [f"{prefix}{f}"]
                    
                    for test_size in test_sizes:
                        test_code = _build_style_code(base_code, test_size, tone)
                        test_upper = test_code.upper()
                        
                        test_match = style_matches[style_matches['Client Style No'].astype(str).str.upper() == test_upper]
                        if not test_match.empty:
                            matched_row = test_match.iloc[0]
                            matched_style = str(matched_row['Client Style No']).strip()
                            matched_size = str(matched_row['ItemSize']).strip() if pd.notna(matched_row['ItemSize']) else test_size
                            return matched_style, matched_size
            except (ValueError, TypeError):
                pass
    
    # No match found with any size variation, return original built code
    return built_code, item_size


def _lookup_client_style(built_code: str) -> str:
    """
    Match the built StyleCode against Ambition_CS.xlsx 'Client Style No'.

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


def _map_item_size_with_prefix(size_val: str, prefix: str) -> str:
    """
    Map a numeric size to its canonical ItemSize using a sizing-system prefix.
    """
    if not size_val:
        return size_val
    if re.search(r'\bINCH\b', size_val, re.IGNORECASE):
        return _map_item_size(size_val)
    if not prefix:
        return _map_item_size(size_val)

    lookup = _get_size_lookup()
    try:
        f = float(size_val)
        if f.is_integer():
            candidates = [
                f"{prefix}{int(f):02d}",   # UP07
                f"{prefix}{int(f)}",        # UP7
                f"{prefix} {int(f):02d}",   # UP 07
                f"{prefix} {int(f)}",       # UP 7
            ]
        else:
            candidates = [
                f"{prefix}{f}",             # UP7.5
                f"{prefix} {f}",            # UP 7.5
                f"{prefix}{size_val}",      # UP7.50
                f"{prefix} {size_val}",     # UP 7.50
            ]
        for cand in candidates:
            key = _normalize_size_key(cand)
            if key in lookup:
                return lookup[key]
    except (ValueError, TypeError):
        pass

    return _map_item_size(size_val)


def _read_pdf_text(pdf_path: str) -> str:
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text += page_text + "\n"
    return full_text


def _extract_po_number(full_text: str) -> Optional[str]:
    m = re.search(r"PO\s*#\s*:\s*(\d+)", full_text, flags=re.IGNORECASE)
    return m.group(1) if m else None


def _split_items(full_text: str) -> List[str]:
    lines = [ln.strip() for ln in full_text.splitlines()]
    items: List[str] = []
    current: List[str] = []
    inside_items = False
    for ln in lines:
        if re.match(r"^\d+\.", ln):
            inside_items = True
            if current:
                items.append("\n".join(current).strip())
            current = [ln]
            continue
        if inside_items:
            current.append(ln)
            if "**FOR SHIMAYRA**" in ln or "FOR SHIMAYRA" in ln:
                items.append("\n".join(current).strip())
                current = []
                inside_items = False
    if current:
        items.append("\n".join(current).strip())
    return [it for it in items if it]


_METAL_PAT = re.compile(
    r'(?:'
    r'SILV(?:ER)?(?=\s|$)|'
    r'\b(?:14KT|18KT|10KT|14K|18K|GOLD|PLAT(?:INUM)?|PT950)\b'
    r')',
    re.IGNORECASE,
)


def _find_item_size_and_qty(line: str) -> tuple[str, str]:
    tokens = re.findall(r"\d+\.\d+|\d+", line)
    if not tokens:
        return "", ""
    size_idx = next((i for i, tok in enumerate(tokens) if re.match(r"^\d+\.\d+$", tok)), 0)
    size_val = tokens[size_idx]
    qty_val = tokens[size_idx + 1] if size_idx + 1 < len(tokens) else ""
    return size_val, qty_val


def _parse_po_item_line(line: str) -> dict:
    out: dict = {
        'sr_no': '', 'item_ref_no': '', 'sku_no': '',
        'description': '', 'metal': '', 'tone': '',
        'item_size': '', 'order_qty': '', 'vendor_style': '',
    }

    m = re.match(r'^(\d+)\.\s+', line)
    if not m:
        return out
    out['sr_no'] = m.group(1)
    pos = m.end()

    m2 = re.match(r'(\d+/\d+)\s+', line[pos:])
    if m2:
        out['item_ref_no'] = m2.group(1)
        pos += m2.end()

    m3 = re.match(r'([A-Za-z]?\d+)\s+', line[pos:])
    if m3:
        pos += m3.end()

    m4 = re.match(r'([A-Za-z0-9]+(?:-[A-Za-z0-9]+)+)\s+', line[pos:])
    if m4:
        out['sku_no'] = m4.group(1)
        pos += m4.end()
    else:
        m4b = re.match(r'(\S+)\s+', line[pos:])
        if m4b:
            out['sku_no'] = m4b.group(1)
            pos += m4b.end()

    rest = line[pos:]

    mm = _METAL_PAT.search(rest)
    if mm:
        out['description'] = rest[:mm.start()].strip()
        mk = mm.group(0).upper()
        if mk in ('SILV', 'SILVER'):
            out['metal'] = 'AG925'
        elif mk in ('14KT', '14K'):
            out['metal'] = 'G14'
        elif mk in ('18KT', '18K'):
            out['metal'] = 'G18'
        elif mk == '10KT':
            out['metal'] = 'G10'
        elif mk == 'GOLD':
            out['metal'] = 'G14'
        elif mk in ('PLAT', 'PLATINUM', 'PT950'):
            out['metal'] = 'PT950'

        after_metal = rest[mm.end():]

        mc = re.match(r'\s+([A-Z]{1,3})\b', after_metal)
        if mc:
            raw_tone = mc.group(1).upper()
            if len(raw_tone) > 1 and raw_tone != 'PT':
                first = raw_tone[0]
                if first in ('W', 'Y', 'R', 'P'):
                    raw_tone = first
            out['tone'] = raw_tone
            after_metal = after_metal[mc.end():]

        ms = re.match(r'\s+(\d+\.\d+)', after_metal)
        if ms:
            out['item_size'] = ms.group(1)
            after_size = after_metal[ms.end():]
            mq = re.match(r'\s+(\d+)', after_size)
            if mq:
                out['order_qty'] = mq.group(1)

    if not out['item_size'] or not out['order_qty']:
        fb_size, fb_qty = _find_item_size_and_qty(line)
        if not out['item_size']:
            out['item_size'] = fb_size
        if not out['order_qty']:
            out['order_qty'] = fb_qty

    tokens = line.split()
    while tokens and re.fullmatch(r'\d{7,}', tokens[-1]):
        tokens.pop()
    for tok in reversed(tokens):
        clean = re.sub(r'^[A-Za-z]{3}/\d{1,2}/\d{4}', '', tok)
        if clean and re.search(r'[A-Za-z]', clean) and re.search(r'\d', clean):
            out['vendor_style'] = clean
            break

    return out


def _format_item_size(size_val: str) -> str:
    if not size_val:
        return ''
    try:
        f = float(size_val)
        if f >= 14:
            return f"{int(f)} INCH" if f.is_integer() else f"{f} INCH"
        return str(int(f)) if f.is_integer() else str(f)
    except (ValueError, TypeError):
        return size_val


def _find_item_ref_no(line: str) -> Optional[str]:
    m = re.search(r"(\d{5,}/\d+)", line)
    return m.group(1) if m else None


def _extract_design_instructions(block_text: str) -> Optional[str]:
    phrases = re.findall(r"\*\*([^*]+)\*\*", block_text)
    return " ".join(p.strip() for p in phrases) if phrases else None


def _extract_stamp_instruction(block_text: str) -> Optional[str]:
    m = re.search(r"Special\s+Inst\.[^\n]*?STAMP\s+([^,\n]+)", block_text, flags=re.IGNORECASE)
    return m.group(1).strip() if m else None


def _extract_special_remarks(block_text: str) -> Optional[str]:
    for ln in block_text.splitlines():
        if ln.upper().startswith("SPECIAL INST."):
            return ln.split("Special Inst.", 1)[-1].strip()
    return None


def _extract_stone_quality(special_remarks_text: str) -> str:
    if not special_remarks_text:
        return ""
    m = re.search(r'\bDIA[M]?\s+([A-Z0-9\s]+?)(?:,|\s+GROSS|\s+$)', special_remarks_text, re.IGNORECASE)
    if m:
        return f"DIA {m.group(1).strip()}"
    return ""


def _build_special_remarks(item_ref_no: str, sku_no: str, metal: str, item_size: str, 
                           special_remarks_r: str) -> str:
    parts = []
    
    if item_ref_no:
        parts.append(item_ref_no)
    
    if sku_no:
        parts.append(sku_no)
    
    if metal and metal.upper().startswith('AG925'):
        parts.append("SILV")
    
    if item_size:
        parts.append(f"SZ-{item_size}")
    
    stone_quality = _extract_stone_quality(special_remarks_r)
    if stone_quality:
        parts.append(stone_quality)
    
    return ", ".join(parts)


def _extract_size_from_stylecode(style_code: str) -> Optional[str]:
    if not style_code:
        return None
    
    m = re.search(r'(\d+(?:\.\d+)?)IN', style_code, re.IGNORECASE)
    if m:
        size_num = m.group(1)
        try:
            f = float(size_num)
            return f"{f:.2f} INCH" if not f.is_integer() else f"{int(f):.2f} INCH"
        except (ValueError, TypeError):
            pass
    
    return None


def process_ambition_file(input_path: str, output_dir: str, size_prefix: str = 'UP'):
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()
        if ext == '.pdf':
            full_text = _read_pdf_text(input_path)
            item_po_no = _extract_po_number(full_text) or ""
            blocks = _split_items(full_text)
            rows: List[Dict[str, str]] = []
            for blk in blocks:
                first_line = next((ln for ln in blk.splitlines() if re.match(r"^\d+\.\s", ln)),
                                  blk.splitlines()[0] if blk.splitlines() else "")

                parsed = _parse_po_item_line(first_line)

                sr_no = parsed['sr_no'] or (re.match(r"^(\d+)\.", first_line.strip()) or type('', (), {'group': lambda s, n: ''})()).group(1)
                item_ref_no = parsed['item_ref_no'] or _find_item_ref_no(first_line) or ""
                sku_no = parsed['sku_no'] or ""
                metal = parsed['metal'] or ""
                tone = parsed['tone'] or ""
                order_qty = parsed['order_qty'] or ""
                vendor_style = parsed['vendor_style'] or ""
                cust_instr = parsed['description'] or ""

                _SKIP_PREFIXES = re.compile(
                    r'^(special\s+inst|comment|total\s*:|page\s+\d|\d+\.)',
                    re.IGNORECASE,
                )
                blk_lines = [ln.strip() for ln in blk.splitlines() if ln.strip()]
                first_idx = next(
                    (i for i, ln in enumerate(blk_lines) if re.match(r"^\d+\.\s", ln)), 0
                )
                if first_idx + 1 < len(blk_lines):
                    desc_line = blk_lines[first_idx + 1]
                    if desc_line and not _SKIP_PREFIXES.match(desc_line):
                        cust_instr = (cust_instr + " " + desc_line).strip()

                item_size = _format_item_size(parsed['item_size'])

                raw_design_instr = _extract_design_instructions(blk) or ""
                stamp_instr = _extract_stamp_instruction(blk) or ""
                special_remarks = _extract_special_remarks(blk) or ""

                # Extract tone from special remarks if "14K ,A" pattern exists
                if not tone and special_remarks:
                    # Look for patterns like "14K ,A" or "14K, A" or "18K ,A"
                    tone_match = re.search(r'\b(?:14K|18K|10K)\s*,\s*([A-Z])\b', special_remarks, re.IGNORECASE)
                    if tone_match:
                        tone_letter = tone_match.group(1).upper()
                        # Map A -> W (White)
                        if tone_letter == 'A':
                            tone = 'W'
                        elif tone_letter in ('W', 'Y', 'P', 'R'):
                            tone = tone_letter

                style_code_base = vendor_style
                if style_code_base:
                    mt = re.search(r'-([A-Z]+)$', style_code_base)
                    if mt:
                        tone_full = mt.group(1)
                        if not tone:
                            tone = (tone_full[0] if tone_full else '').replace('V', 'W')
                        style_code_base = style_code_base[:mt.start()]

                item_size = _map_item_size_with_prefix(item_size or '', size_prefix or '')

                # Build Metal column: metal + tone (e.g., G14W, G18Y)
                _gold_bases = ('G14', 'G18', 'G10', 'G9')
                if metal.upper().startswith(_gold_bases) and tone:
                    metal_col = metal + tone.upper()
                else:
                    metal_col = metal
                
                # Tone column logic:
                # - If AG925 (silver): Tone = '' (blank)
                # - Otherwise: Tone = the detected tone (W, Y, P, R, etc.)
                if metal.upper() == 'AG925':
                    tone_col = ''
                else:
                    tone_col = tone.upper() if tone else ''

                # DesignProductionInstruction: prefix based on Tone
                if tone_col == 'W':
                    design_instr = f"White Rhodium {raw_design_instr}".strip()
                elif tone_col in ('Y', 'P', 'R', 'PT'):
                    design_instr = f"No Rhodium {raw_design_instr}".strip()
                else:
                    design_instr = raw_design_instr.strip()

                _sc_tone = 'AG' if metal.upper() == 'AG925' else (tone or '')
                # Use enhanced lookup that matches both StyleCode AND ItemSize
                style_code, item_size = _lookup_client_style_with_size(style_code_base, item_size or "", _sc_tone)
                
                extracted_size = _extract_size_from_stylecode(style_code)
                if extracted_size:
                    item_size = _map_item_size(extracted_size)
                
                special_remarks_combined = _build_special_remarks(
                    item_ref_no, sku_no, metal, item_size, special_remarks
                )

                rows.append({
                    "SrNo": sr_no,
                    "StyleCode": style_code,
                    "ItemSize": item_size or "",
                    "OrderQty": order_qty or "",
                    "OrderItemPcs": 1,
                    "Metal": metal_col or "",
                    "Tone": tone_col,
                    "ItemPoNo": item_po_no,
                    "ItemRefNo": item_ref_no,
                    "StockType": "",
                    "MakeType": "",
                    "CustomerProductionInstruction": cust_instr,
                    "SpeacialRemarksR": special_remarks,
                    "DesignProductionInstruction": design_instr,
                    "StampInstruction": stamp_instr,
                    "OrderGroup": "",
                    "Certificate": "",
                    "SKUNo": sku_no,
                    "Basestoneminwt": "",
                    "Basestonemaxwt": "",
                    "Basemetalminwt": "",
                    "Basemetalmaxwt": "",
                    "Productiondeliverydate": "",
                    "Expecteddeliverydate": "",
                    "SetPrice": "",
                    "StoneQuality": "",
                    "SpecialRemarks": special_remarks_combined,
                })

            columns = [
                "SrNo","StyleCode","ItemSize","OrderQty","OrderItemPcs","Metal","Tone","ItemPoNo","ItemRefNo","StockType","MakeType","CustomerProductionInstruction","SpeacialRemarksR","DesignProductionInstruction","StampInstruction","OrderGroup","Certificate","SKUNo","Basestoneminwt","Basestonemaxwt","Basemetalminwt","Basemetalmaxwt","Productiondeliverydate","Expecteddeliverydate","SetPrice","StoneQuality","SpecialRemarks"
            ]
            df = pd.DataFrame(rows, columns=columns)
            output_path = os.path.join(output_dir, f"{base_name}_AMBITION_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        elif ext in ['.xlsx', '.xls', '.csv']:
            df = pd.read_excel(input_path) if ext in ['.xlsx', '.xls'] else pd.read_csv(input_path)
            
            # Apply rhodium prefix check if 'Tone' and 'DesignProductionInstruction' columns exist
            if 'Tone' in df.columns and 'DesignProductionInstruction' in df.columns:
                def update_design_inst(row):
                    val = str(row['DesignProductionInstruction']) if pd.notna(row['DesignProductionInstruction']) else ""
                    tone_val = str(row['Tone']).strip().upper() if pd.notna(row['Tone']) else ""
                    prefix = "White Rhodium" if tone_val == 'W' else "No Rhodium"
                    return f"{prefix} {val}".strip()
                df['DesignProductionInstruction'] = df.apply(update_design_inst, axis=1)

            output_path = os.path.join(output_dir, f"{base_name}_AMBITION_PASSTHROUGH.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None