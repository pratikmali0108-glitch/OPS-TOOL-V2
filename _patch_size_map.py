"""
Adds _map_item_size helper to all 19 processor files and applies it before
every _build_style_code call so that ItemSize values are matched against
ItemSize_Mst.xlsx before use.
"""
import os, sys

# ── Helper block to insert after _build_style_code in every file ─────────────
HELPER_BLOCK = (
    "\n\n"
    "_ITEM_SIZE_LOOKUP = None\n"
    "\n"
    "\n"
    "def _get_size_lookup():\n"
    "    global _ITEM_SIZE_LOOKUP\n"
    "    if _ITEM_SIZE_LOOKUP is not None:\n"
    "        return _ITEM_SIZE_LOOKUP\n"
    "    _ITEM_SIZE_LOOKUP = {}\n"
    "    try:\n"
    "        mst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ItemSize_Mst.xlsx')\n"
    "        _df_mst = pd.read_excel(mst)\n"
    "        for val in _df_mst['Item Size Code'].dropna():\n"
    "            vs = str(val).strip()\n"
    "            if vs and vs.upper() != 'NAN':\n"
    "                k = _normalize_size_key(vs)\n"
    "                if k:\n"
    "                    _ITEM_SIZE_LOOKUP[k] = vs\n"
    "    except Exception:\n"
    "        pass\n"
    "    return _ITEM_SIZE_LOOKUP\n"
    "\n"
    "\n"
    "def _normalize_size_key(s):\n"
    "    s = str(s).strip()\n"
    "    if not s or s.upper() == 'NAN':\n"
    "        return ''\n"
    "    m = re.match(r'^(\\d+(?:\\.\\d+)?)\\s*INCH$', s, re.IGNORECASE)\n"
    "    if m:\n"
    "        return f\"{float(m.group(1)):.2f}inch\"\n"
    "    return re.sub(r'\\s+', '', s).lower()\n"
    "\n"
    "\n"
    "def _map_item_size(raw):\n"
    "    \"\"\"Map raw ItemSize to its canonical form from ItemSize_Mst.xlsx.\"\"\"\n"
    "    if not raw or str(raw).strip().upper() in ('', 'NAN'):\n"
    "        return raw\n"
    "    lookup = _get_size_lookup()\n"
    "    key = _normalize_size_key(str(raw).strip())\n"
    "    return lookup.get(key, raw)\n"
)

ANCHOR = "    return f\"{base}-{suffix}\" if suffix else base\n"

# ── Per-file mapping-insertion specs ─────────────────────────────────────────
# (old_snippet, new_snippet) pairs to add _map_item_size at the right place.
# For DataFrame files: insert df['ItemSize']=df['ItemSize'].apply(_map_item_size)
#                      right before the df.apply(_build_style_code) line.
# For loop files: insert item_size = _map_item_size(item_size) after size is final.

PATCHES = {
    # ── DataFrame files using `df` variable ──────────────────────────────────
    'jjl.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'obu.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'uneek.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'ju_pendant.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'AAM.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'Bhakti_Dharam.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'DCT.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'MOR.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'NGL.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'PC2.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'PCB.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'SGI.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    'VIMCO.py': (
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['StyleCode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    # ── fsa.py uses 'Stylecode' (lowercase c) ────────────────────────────────
    'fsa.py': (
        "            df['Stylecode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['Stylecode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
        "            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)\n"
        "            df['Stylecode'] = df.apply(\n"
        "                lambda row: _build_style_code(row['Stylecode'], row['ItemSize'], row['Tone']), axis=1\n"
        "            )\n",
    ),
    # ── anaya.py — two locations (df_pdf and df_selected) ───────────────────
    'anaya.py': [
        (
            "    df_pdf['StyleCode'] = df_pdf.apply(\n"
            "        lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
            "    )\n",
            "    df_pdf['ItemSize'] = df_pdf['ItemSize'].apply(_map_item_size)\n"
            "    df_pdf['StyleCode'] = df_pdf.apply(\n"
            "        lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
            "    )\n",
        ),
        (
            "        df_selected['StyleCode'] = df_selected.apply(\n"
            "            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
            "        )\n",
            "        df_selected['ItemSize'] = df_selected['ItemSize'].apply(_map_item_size)\n"
            "        df_selected['StyleCode'] = df_selected.apply(\n"
            "            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], row['Tone']), axis=1\n"
            "        )\n",
        ),
    ],
    # ── hk.py — loop path (formatted_size) + Excel df_selected path ─────────
    'hk.py': [
        (
            "        style_code = _build_style_code(style_code, formatted_size, tone)\n",
            "        formatted_size = _map_item_size(formatted_size or '')\n"
            "        style_code = _build_style_code(style_code, formatted_size, tone)\n",
        ),
        (
            '    df_selected["StyleCode"] = df_selected.apply(\n'
            '        lambda row: _build_style_code(row["StyleCode"], row["ItemSize"], row["Tone"]), axis=1\n',
            '    df_selected["ItemSize"] = df_selected["ItemSize"].apply(_map_item_size)\n'
            '    df_selected["StyleCode"] = df_selected.apply(\n'
            '        lambda row: _build_style_code(row["StyleCode"], row["ItemSize"], row["Tone"]), axis=1\n',
        ),
    ],
    # ── rbl.py — loop path ───────────────────────────────────────────────────
    'rbl.py': (
        "            style_code = _build_style_code(style_code, item_size, tone_value)\n",
        "            item_size = _map_item_size(item_size or '')\n"
        "            style_code = _build_style_code(style_code, item_size, tone_value)\n",
    ),
    # ── ambition.py — loop path (after INCH conversion) ──────────────────────
    'ambition.py': (
        "                # Build StyleCode with INCH-formatted size so 'IN' is inserted correctly\n"
        "                # e.g. '7 INCH' + 'W' -> 'BR0000094K-7INWG'\n"
        "                style_code = _build_style_code(style_code, item_size or \"\", tone or \"\")\n",
        "                # Build StyleCode with INCH-formatted size so 'IN' is inserted correctly\n"
        "                # e.g. '7 INCH' + 'W' -> 'BR0000094K-7INWG'\n"
        "                item_size = _map_item_size(item_size or '')\n"
        "                style_code = _build_style_code(style_code, item_size or \"\", tone or \"\")\n",
    ),
    # ── OMJ.py — uses ItemSizeCopy column ────────────────────────────────────
    'OMJ.py': (
        "        # Step 17: Create StyleCode  e.g. 'YR4172SA-7YG'\n"
        "        df_cleaned['StyleCode'] = df_cleaned.apply(\n",
        "        # Step 17: Create StyleCode  e.g. 'YR4172SA-7YG'\n"
        "        df_cleaned['ItemSizeCopy'] = df_cleaned['ItemSizeCopy'].apply(_map_item_size)\n"
        "        df_cleaned['StyleCode'] = df_cleaned.apply(\n",
    ),
}


def patch_file(fname, patches, add_helpers):
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), fname)
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    changed = False

    # 1. Add helper functions after _build_style_code (only if not present)
    if add_helpers and '_map_item_size' not in content:
        if ANCHOR in content:
            content = content.replace(ANCHOR, ANCHOR + HELPER_BLOCK, 1)
            changed = True
            print(f'  + helpers added')
        else:
            print(f'  ! anchor not found, helpers NOT added')

    # 2. Apply mapping insertions
    if isinstance(patches, tuple):
        patches = [patches]
    for old, new in patches:
        if old in content:
            content = content.replace(old, new, 1)
            changed = True
            print(f'  + map insertion applied')
        elif new in content:
            print(f'  . map insertion already present')
        else:
            print(f'  ! pattern not found: {repr(old[:60])}')

    if changed:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f'  => SAVED')
    else:
        print(f'  => no changes needed')


ALL_FILES = list(PATCHES.keys())

print('Patching files...\n')
for fname in ALL_FILES:
    print(f'--- {fname} ---')
    patch_file(fname, PATCHES[fname], add_helpers=True)
    print()

print('Done.')
