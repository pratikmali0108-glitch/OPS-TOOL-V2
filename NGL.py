import os
import re
from pathlib import Path
import pandas as pd


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
    """Load and cache the approved Client Style No list from NGL_CS_100926.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    try:
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'NGL_CS_100926.xlsx')
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
    Match the built StyleCode against NGL_CS_100926.xlsx 'Client Style No'.

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


def process_ngl_file(
    input_file_path: str,
    output_folder: str | None = None,
    order_qty: str | int = "",
    item_po_no: str = "",
    priority: str = "",
    additional_after_dia: str = ""
):
    """
    Process a single NGL Excel file (per EDA_NGL.ipynb logic) and export CSV.

    Returns: (success: bool, output_path: str|None, error: str|None, df: pd.DataFrame|None)
    """
    try:
        # Read Excel (skip first row as in notebook)
        df = pd.read_excel(input_file_path, skiprows=1)

        # Rename columns per notebook mapping
        rename_map = {
            'Unnamed: 0': 'SrNo',
            'Unnamed: 1': 'ItemSize',
            'Unnamed: 2': 'StyleCode',
            'Unnamed: 3': 'Metal',
            'Unnamed: 4': 'Gold Rate',
            'Unnamed: 5': 'Gold Wt',
            'Unnamed: 6': 'Stone Wt',
            'Unnamed: 7': 'Metal Wt',
            'Unnamed: 8': 'Dia Pcs',
            'Set Rate': 'SetRate',
            'Set Val': 'SetValue',
            'Unnamed: 19': 'LabourRate',
            'Total Labour': 'TotalLabour',
            'Unnamed: 21': 'OtherCharges',
            'Unnamed: 22': 'MakingCharges',
            'Unnamed: 23': 'Wastage',
            'Unnamed: 24': 'FinalValue',
            'Unnamed: 25': 'Discount',
            'Unnamed: 26': 'NetValue'
        }
        df.rename(columns=rename_map, inplace=True)

        # Keep rows where SrNo exists
        if 'SrNo' in df.columns:
            df = df[df['SrNo'].notna()].reset_index(drop=True)

        # Columns we use
        required_cols = ['SrNo', 'StyleCode', 'ItemSize', 'Metal', 'Dia Qlty']
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            return False, None, f"Missing columns: {', '.join(missing)}", None

        df_selected = df[required_cols].copy()

        # Insert OrderQty and OrderItemPcs after ItemSize
        itemsize_pos = df_selected.columns.get_loc('ItemSize') + 1
        df_selected.insert(itemsize_pos, 'OrderQty', order_qty)
        df_selected.insert(itemsize_pos + 1, 'OrderItemPcs', '')

        # Map Metal to codes
        metal_mapping = {
            '585/W': 'G585W',
            '585/Y': 'G585Y',
            '585/R': 'G585P',
            '585/P': 'G585P',
            '585/RW': 'G585PW',
            '585/PW': 'G585PW',
            '750/W': 'G750W',
            '750/Y': 'G750Y',
            '750/R': 'G750P',
            '750/P': 'G750P',
            '750/RW': 'G750PW',
            '750/PW': 'G750PW'
        }
        df_selected['Metal'] = df_selected['Metal'].astype(str).str.strip().map(metal_mapping).fillna('')

        # Tone from Metal suffix
        df_selected['Tone'] = df_selected['Metal'].str.extract(r'(W|Y|P)$')[0].fillna('')
        tone_col = df_selected.pop('Tone')
        df_selected.insert(df_selected.columns.get_loc('Metal') + 1, 'Tone', tone_col)

        # Add columns after Tone
        tone_pos = df_selected.columns.get_loc('Tone') + 1
        for col_name, value in [
            ('ItemPoNo', item_po_no),
            ('ItemRefNo', ''),
            ('StockType', ''),
            ('Priority', priority),
            ('MakeType', ''),
            ('CustomerProductionInstruction', '')
        ]:
            df_selected.insert(tone_pos, col_name, value)
            tone_pos += 1

        # SpecialRemarks based on Dia Qlty and additional text
        def create_special_remarks(row):
            dia = row['Dia Qlty'] if pd.notna(row['Dia Qlty']) else ''
            return f"DIA QLTY-{dia}{additional_after_dia}"

        cpi_pos = df_selected.columns.get_loc('CustomerProductionInstruction') + 1
        df_selected.insert(cpi_pos, 'SpecialRemarks', df_selected.apply(create_special_remarks, axis=1))

        # DesignProductionInstruction from Tone
        def map_design_instruction(tone: str):
            if pd.isna(tone):
                return ''
            tone = str(tone).strip().upper()
            if tone == 'W':
                return 'WHITE RODIUM'
            elif tone in ['Y', 'PW', 'PY']:
                return 'NO RODIUM'
            return ''

        dpi_pos = df_selected.columns.get_loc('SpecialRemarks') + 1
        df_selected.insert(dpi_pos, 'DesignProductionInstruction', df_selected['Tone'].apply(map_design_instruction))

        # StampInstruction as Metal + ' + LOGO'
        stamp_pos = df_selected.columns.get_loc('DesignProductionInstruction') + 1
        df_selected.insert(
            stamp_pos,
            'StampInstruction',
            df_selected['Metal'].apply(lambda x: f"{x}+ LOGO" if isinstance(x, str) and x != '' else '')
        )
        stamp_pos = df_selected.columns.get_loc('StampInstruction') + 1

        # Additional trailing columns
        new_columns = [
            'OrderGroup', 'Certificate', 'SKUNo',
            'Basestoneminwt', 'Basestonemaxwt',
            'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate'
        ]
        for col_name in new_columns:
            df_selected.insert(stamp_pos, col_name, '')
            stamp_pos += 1

        # Drop Dia Qlty from output
        if 'Dia Qlty' in df_selected.columns:
            df_selected.drop(columns=['Dia Qlty'], inplace=True)

        # Build full StyleCode: base-sizeWG / base-sizePT etc.
        df_selected['ItemSize'] = df_selected['ItemSize'].apply(_map_item_size)
        df_selected['StyleCode'] = df_selected.apply(
            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
        )
        df_selected['StyleCode'] = df_selected['StyleCode'].apply(_lookup_client_style)
        df_selected.loc[df_selected['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        # Output path
        input_filename = Path(input_file_path).stem
        output_folder = output_folder or os.getcwd()
        output_path = os.path.join(output_folder, f"NGL_FORMAT_{input_filename}.csv")
        df_selected.to_csv(output_path, index=False)

        return True, output_path, None, df_selected
    except Exception as e:
        return False, None, str(e), None


