import os
import re
from pathlib import Path
import pandas as pd


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
    """Load and cache the approved Client Style No list from DCT_CS.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    try:
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'DCT_100826.xlsx')
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
    Match the built StyleCode against DCT_CS.xlsx 'Client Style No'.

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


def process_dct_file(
    input_file_path: str,
    output_folder: str | None = None,
    priority: str = ""
):
    """
    Process a single DCT Excel file based on EDA_DCT.ipynb logic and export CSV.

    Returns: (success: bool, output_path: str|None, error: str|None, df: pd.DataFrame|None)
    """
    try:
        # Read Excel skipping first row
        df = pd.read_excel(input_file_path, skiprows=1)

        # Keep non-null 'Sr No.' rows
        if 'Sr No.' not in df.columns:
            return False, None, "Missing required column: 'Sr No.'", None
        df_clean = df[df['Sr No.'].notna()].reset_index(drop=True)
        df_copy = df_clean.copy()

        # Select expected columns
        required_cols = [
            'Sr No.','Po #','Ring Size','Unnamed: 2','SKU #','Gold Karat','Dia Qlty','Price of a single item USD.1'
        ]
        missing = [c for c in required_cols if c not in df_copy.columns]
        if missing:
            return False, None, f"Missing columns: {', '.join(missing)}", None

        df_selected = df_copy[required_cols].copy()

        # Rename
        df_selected.rename(columns={
            'Sr No.': 'SrNo',
            'SKU #': 'StyleCode',
            'Ring Size': 'ItemSize',
            'Price of a single item USD.1': 'OrderQty',
            'Gold Karat': 'Metal',
            'Po #': 'ItemPoNo',
            'Unnamed: 2': 'SKUNo',
        }, inplace=True)

        # Format ItemSize -> EU<int>
        df_selected['ItemSize'] = df_selected['ItemSize'].apply(
            lambda x: f"EU{int(x)}" if pd.notna(x) else ''
        )

        # Reorder: StyleCode → ItemSize → OrderQty
        cols = df_selected.columns.tolist()
        cols.insert(1, cols.pop(cols.index('StyleCode')))
        cols.insert(2, cols.pop(cols.index('ItemSize')))
        cols.insert(3, cols.pop(cols.index('OrderQty')))
        df_selected = df_selected[cols]

        # Add OrderItemPcs after OrderQty
        df_selected.insert(df_selected.columns.get_loc('OrderQty') + 1, 'OrderItemPcs', value='')

        # Move Metal after OrderItemPcs
        metal_col = df_selected.pop('Metal')
        df_selected.insert(df_selected.columns.get_loc('OrderItemPcs') + 1, 'Metal', metal_col)

        # Map Metal
        metal_mapping = {
            '14KT RW': 'G585PW',
            '14KT R': 'G585P',
            '14KT Y': 'G585Y',
            '14KT W': 'G585W',
            '18KT RW': 'G750PW',
            '18KT R': 'G750P',
            '18KT Y': 'G750Y',
            '18KT W': 'G750W'
        }
        df_selected['Metal'] = df_selected['Metal'].map(metal_mapping).fillna('')

        # Tone from Metal
        df_selected.insert(
            df_selected.columns.get_loc('Metal') + 1,
            'Tone',
            df_selected['Metal'].str.extract(r'(PW|P|Y|W)$')[0].fillna('')
        )

        # Insert columns after ItemPoNo
        insert_pos = df_selected.columns.get_loc('ItemPoNo') + 1
        for col_name in ['ItemRefNo','StockType', 'Priority', 'MakeType','CustomerProductionInstruction']:
            df_selected.insert(insert_pos, col_name, '')
            insert_pos += 1

        # Fill Priority from parameter
        if priority:
            df_selected['Priority'] = priority

        # SpecialRemarks after CustomerProductionInstruction
        special_remarks_pos = df_selected.columns.get_loc('CustomerProductionInstruction') + 1

        def create_special_remarks(row):
            size = row['ItemSize'].replace('EU','') if row['ItemSize'] else ''
            dia = row['Dia Qlty'] if pd.notna(row['Dia Qlty']) else ''
            return f"{row['SKUNo']} {row['Metal']} SZ-{size} DIA: {dia}"

        df_selected.insert(special_remarks_pos, 'SpecialRemarks', df_selected.apply(create_special_remarks, axis=1))

        # DesignProductionInstruction from Tone
        dpi_pos = df_selected.columns.get_loc('SpecialRemarks') + 1

        def map_design_instruction(tone):
            if pd.isna(tone):
                return ''
            tone = str(tone).strip().upper()
            if tone == 'W':
                return 'WHITE RODIUM'
            elif tone in ['Y', 'PW', 'PY','P']:
                return 'NO RODIUM'
            return ''

        df_selected.insert(dpi_pos, 'DesignProductionInstruction', df_selected['Tone'].apply(map_design_instruction))

        # StampInstruction from numeric part of Metal (585/750)
        stamp_pos = df_selected.columns.get_loc('DesignProductionInstruction') + 1
        df_selected.insert(
            stamp_pos,
            'StampInstruction',
            df_selected['Metal'].str.extract(r'(\d{3})')[0].fillna('')
        )

        # OrderGroup, Certificate after StampInstruction
        ordergroup_pos = df_selected.columns.get_loc('StampInstruction') + 1
        for col_name in ['OrderGroup', 'Certificate']:
            df_selected.insert(ordergroup_pos, col_name, '')
            ordergroup_pos += 1

        # Additional columns after SKUNo
        sku_pos = df_selected.columns.get_loc('SKUNo') + 1
        trailing_cols = [
            'Basestoneminwt', 'Basestonemaxwt',
            'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate'
        ]
        for col_name in trailing_cols:
            df_selected.insert(sku_pos, col_name, '')
            sku_pos += 1

        # Drop Dia Qlty
        if 'Dia Qlty' in df_selected.columns:
            df_selected.drop(columns=['Dia Qlty'], inplace=True)

        # Build full StyleCode: base-sizeWG / base-sizePT etc.
        df_selected['ItemSize'] = df_selected['ItemSize'].apply(_map_item_size)
        df_selected['StyleCode'] = df_selected.apply(
            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
        )
        df_selected['StyleCode'] = df_selected['StyleCode'].apply(_lookup_client_style)
        df_selected.loc[df_selected['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        # Output
        input_filename = Path(input_file_path).stem
        output_folder = output_folder or os.getcwd()
        output_path = os.path.join(output_folder, f"DCT_FORMAT_{input_filename}.csv")
        df_selected.to_csv(output_path, index=False)

        return True, output_path, None, df_selected
    except Exception as e:
        return False, None, str(e), None


