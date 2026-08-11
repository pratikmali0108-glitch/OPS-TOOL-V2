import os
import re
import pandas as pd


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', '6.5', 'W') -> 'VR1943EEA-6.5WG'
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
    """Load and cache the approved Client Style No list from RBL_CS_100826.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    try:
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'RBL_CS_100826.xlsx')
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
    Match the built StyleCode against RBL_CS_100826.xlsx 'Client Style No'.

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


def process_rbl_file(input_path: str, output_dir: str, end_customer_name: str = "", priority_value: str = ""):
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(input_path, skiprows=3)
            df_cleaned = df.dropna(axis=1, how='all')
            if df_cleaned.empty or df_cleaned.shape[1] < 2:
                return False, None, 'Unexpected RBL sheet structure', None
            key_col = df_cleaned.columns[0]
            val_col = df_cleaned.columns[1]
            meta_rows = df_cleaned[[key_col, val_col]].copy()
            meta_rows.columns = ["Key", "Value"]

            def get_value(prefix: str, default: str = "") -> str:
                mask = meta_rows["Key"].astype(str).str.strip().str.upper().str.startswith(prefix.upper())
                vals = meta_rows.loc[mask, "Value"].dropna().astype(str).str.strip()
                return vals.iloc[0] if not vals.empty else default

            first_cell = str(df_cleaned.iloc[0, 0]) if not df_cleaned.empty else ""
            ct_match = re.search(r"(\d+(?:\.\d+)?)\s*CT", first_cell, flags=re.IGNORECASE)
            ct_value = None
            if ct_match:
                try:
                    ct_numeric = float(ct_match.group(1))
                    ct_value = f"{ct_numeric:.2f} CT"
                except ValueError:
                    ct_value = ct_match.group(1).strip() + " CT"
            if not ct_value:
                stone_weight = get_value("STONE WEIGHT", "")
                ct_match = re.search(r"(\d+(?:\.\d+)?)\s*CT", str(stone_weight), flags=re.IGNORECASE)
                if ct_match:
                    try:
                        ct_numeric = float(ct_match.group(1))
                        ct_value = f"{ct_numeric:.2f} CT"
                    except ValueError:
                        ct_value = ct_match.group(1).strip() + " CT"

            style_code = get_value("STYLE #", "").strip()
            if not style_code:
                paren = re.search(r"\(([^)]+)\)", first_cell)
                if paren:
                    style_code = paren.group(1).strip()

            raw_size = get_value("SIZE", "").strip()
            item_size = "" if raw_size in {"", "~", "-", "NA", "NaN"} else raw_size

            order_qty_str = get_value("ORDER QTY", "")
            order_qty_match = re.search(r"\d+", str(order_qty_str))
            order_qty = int(order_qty_match.group(0)) if order_qty_match else None

            po_date_raw = get_value("PO DATE", "")
            po_date_fmt = ""
            if po_date_raw:
                try:
                    po_dt = pd.to_datetime(po_date_raw)
                    po_date_fmt = po_dt.strftime("%d-%m-%Y")
                except Exception:
                    try:
                        parts = str(po_date_raw).split(" ")[0].split("-")
                        if len(parts) == 3:
                            y, m, d = parts
                            po_date_fmt = f"{d}-{m}-{y}"
                    except Exception:
                        po_date_fmt = str(po_date_raw)
            item_po_no = f"email dated as on {po_date_fmt}" if po_date_fmt else ""

            metal_karat_raw = get_value("METAL", "").upper().replace("KARAT", "KT").strip()
            metal_color_raw = get_value("METAL COLOR", "").upper().strip()
            karat_num_match = re.search(r"(8|9|10|14|18|22|24)\s*KT", metal_karat_raw)
            karat_num = karat_num_match.group(1) if karat_num_match else ""

            def normalize_color(color: str) -> str:
                color = color.upper()
                if "WHITE" in color:
                    return "W"
                if "YELLOW" in color:
                    return "Y"
                if "ROSE" in color or "PINK" in color:
                    return "R"
                return ""

            tone_char = normalize_color(metal_color_raw)
            metal_code = f"G{karat_num}{tone_char}" if karat_num and tone_char else ""
            tone_value = tone_char.upper()

            size_fragment = f", SZ-{item_size}" if item_size else ""
            karat_k = (karat_num + "k") if karat_num else ""
            special_remarks = ""
            if end_customer_name or karat_num or metal_color_raw:
                special_remarks = f"{end_customer_name}, {karat_k} {metal_color_raw}{size_fragment}".strip().strip(", ")

            design_prod_instr = "White Rodium" if tone_char == "W" else "No Rodium"
            stamp_instruction = f"{karat_k} +RB LOGO+{ct_value}".strip() if karat_k or ct_value else "RB LOGO"

            columns = [
                "SrNo","StyleCode","ItemSize","OrderQty","OrderItemPcs","Metal","Tone","ItemPoNo","ItemRefNo","StockType","Priority","MakeType","CustomerProductionInstruction","SpecialRemarks","DesignProductionInstruction","StampInstruction","OrderGroup","SKUNo","Basestoneminwt","Basestonemaxwt","Basemetalminwt","Basemetalmaxwt","Productiondeliverydate","Expecteddeliverydate","Blank","SetPrice","StoneQuality",
            ]
            item_size = _map_item_size(item_size or '')
            style_code = _build_style_code(style_code, item_size, tone_value)
            style_code = _lookup_client_style(style_code)
            row = {
                "SrNo": 1,
                "StyleCode": style_code,
                "ItemSize": item_size,
                "OrderQty": order_qty,
                "OrderItemPcs": "",
                "Metal": metal_code,
                "Tone": tone_value,
                "ItemPoNo": item_po_no,
                "ItemRefNo": "",
                "StockType": "",
                "Priority": (priority_value or "").upper(),
                "MakeType": "",
                "CustomerProductionInstruction": "",
                "SpecialRemarks": special_remarks,
                "DesignProductionInstruction": design_prod_instr,
                "StampInstruction": stamp_instruction,
                "OrderGroup": end_customer_name,
                "SKUNo": "",
                "Basestoneminwt": "",
                "Basestonemaxwt": "",
                "Basemetalminwt": "",
                "Basemetalmaxwt": "",
                "Productiondeliverydate": "",
                "Expecteddeliverydate": "",
                "Blank": "",
                "SetPrice": "",
                "StoneQuality": "",
            }
            df_out = pd.DataFrame([row], columns=columns)
            if "Tone" in df_out.columns:
                df_out["Tone"] = df_out["Tone"].astype(str).str.upper()
            output_path = os.path.join(output_dir, f"{base_name}_RBL_MAPPED.xlsx")
            df_out.to_excel(output_path, index=False)
            return True, output_path, None, df_out
        elif ext == '.pdf':
            return False, None, 'RBL expects Excel input as per notebook', None
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None


