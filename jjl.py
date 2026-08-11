import os
import re
import pandas as pd
import pdfplumber


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'EU56', 'W') -> 'VR1943EEA-56WG'
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


def _read_pdf_text(pdf_path: str) -> str:
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text += page_text + "\n"
    return full_text


def _map_eu(size):
    try:
        base = 40
        return f"EU{int(size)+base}"
    except Exception:
        return ""


def process_jjl_file(input_path: str, output_dir: str, default_priority: str = "REG", default_diamond_quality: str = "REG"):
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()
        if ext == '.pdf':
            text = _read_pdf_text(input_path)
            ITEM_PO_NO = re.search(r"\b\d{6}\b", text)
            item_po_no = ITEM_PO_NO.group(0) if ITEM_PO_NO else ""
            pattern = re.compile(
                r"([A-Z0-9\-]+)-O[^\n]*?(?:N-(\d+))?\n(\d+)\s+[A-Z0-9]+\n([A-Z ]+CT\..*?)(?=(?:\n[A-Z0-9\-]+-O|$))",
                re.DOTALL
            )
            matches = list(pattern.finditer(text))
            items = []
            for i, match in enumerate(matches, start=1):
                style_code = match.group(1).strip()
                size = match.group(2).strip() if match.group(2) else ""
                order_qty = match.group(3).strip()
                desc = match.group(4).strip()
                if i == len(matches):
                    cut = desc.find("Polígono")
                    if cut != -1:
                        desc = desc[:cut].rstrip()
                tone = "W" if "WHITE" in desc.upper() else ("Y" if "YELLOW" in desc.upper() else "")
                metal = f"G750{tone}" if tone else "G750"
                design_instr = "White Rodium" if "WHITE" in desc.upper() else "No Rodium"
                eu_size = _map_eu(size) if size else ""
                priority = (default_priority or "REG")
                dia_qlty = (default_diamond_quality or "REG")
                special_remarks = f"{metal}"
                if eu_size:
                    special_remarks += f",{eu_size}"
                special_remarks += f",DIA QUALITY: {dia_qlty}"
                item = {
                    "SrNO": i,
                    "StyleCode": style_code,
                    "ItemSize": eu_size,
                    "OrderQty": order_qty,
                    "OrderItemPcs": 1,
                    "Metal": metal,
                    "Tone": tone,
                    "ItemPoNo": item_po_no,
                    "ItemRefNo": "",
                    "StockType": "",
                    "Priority": priority,
                    "MakeType": "",
                    "CustomerProductionInstruction": desc,
                    "SpecialRemarks": special_remarks,
                    "DesignProductionInstruction": design_instr,
                    "StampInstruction": "750 +logo",
                    "OrderGroup": "JJL",
                    "Certificate": "",
                    "SKUNo": "",
                    "Basestoneminwt": "",
                    "Basestonemaxwt": "",
                    "Basemetalminwt": "",
                    "Basemetalmaxwt": "",
                    "Productiondeliverydate": "",
                    "Expecteddeliverydate": "",
                    "SetPrice": "",
                    "StoneQuality": "",
                }
                items.append(item)
            df = pd.DataFrame(items)
            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)
            df['StyleCode'] = df.apply(
                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
            )
            df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''
            output_path = os.path.join(output_dir, f"{base_name}_JJL_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        elif ext in ['.xlsx', '.xls', '.csv']:
            df = pd.read_excel(input_path) if ext in ['.xlsx', '.xls'] else pd.read_csv(input_path)
            output_path = os.path.join(output_dir, f"{base_name}_JJL_PASSTHROUGH.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None


