import os
import re
import pandas as pd
import pdfplumber


# ============================================================
# GLOBAL CACHES
# ============================================================

_ITEM_SIZE_LOOKUP = None
_OBU_CS_CACHE = None
_STYLE_PATTERN_CACHE = {}


# ============================================================
# BASIC STYLE CODE BUILDER
# ============================================================

def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G'
    or PT for platinum.

    Examples:
        ('VR1943EEA', 'EU52', 'W')
            -> 'VR1943EEA-52WG'

        ('BR0000367S', '6.9 INCH', 'W')
            -> 'BR0000367S-6.9INWG'
    """

    base = str(base).strip() if base else ''
    item_size = str(item_size).strip() if item_size else ''
    tone = str(tone).strip().upper() if tone else ''

    if not base or base.upper() == 'NAN':
        return ''

    # --------------------------------------------------------
    # Detect INCH before stripping
    # --------------------------------------------------------
    has_inch = bool(
        re.search(r'\bINCH\b', item_size, flags=re.IGNORECASE)
        or re.search(r'\bIN\b', item_size, flags=re.IGNORECASE)
    )

    # --------------------------------------------------------
    # Remove common size prefixes
    # --------------------------------------------------------
    size_num = re.sub(
        r'^(?:UP|US|EU|IT|UT|TS|IS)\s*',
        '',
        item_size,
        flags=re.IGNORECASE
    ).strip()

    # Remove INCH / IN
    size_num = re.sub(
        r'\s*(?:INCH|IN)\s*$',
        '',
        size_num,
        flags=re.IGNORECASE
    ).strip()

    # --------------------------------------------------------
    # Normalize numeric values
    # --------------------------------------------------------
    try:
        f = float(size_num)

        if f.is_integer():
            size_num = str(int(f))
        else:
            size_num = str(round(f, 1)).rstrip('0').rstrip('.')

    except (ValueError, TypeError):
        pass

    # --------------------------------------------------------
    # Normalize multi-character tones
    # --------------------------------------------------------
    if tone and len(tone) > 1 and tone != 'PT':
        first = tone[0]

        if first in ('W', 'Y', 'P', 'R'):
            tone = first

    in_part = 'IN' if has_inch else ''

    # --------------------------------------------------------
    # Platinum
    # --------------------------------------------------------
    if tone == 'PT':

        suffix = (
            f"{size_num}{in_part}PT"
            if size_num
            else (f"{in_part}PT" if in_part else 'PT')
        )

    # --------------------------------------------------------
    # Gold tones
    # --------------------------------------------------------
    elif tone in ('W', 'Y', 'P', 'R'):

        suffix = (
            f"{size_num}{in_part}{tone}G"
            if size_num
            else f"{in_part}{tone}G"
        )

    # --------------------------------------------------------
    # Silver
    # --------------------------------------------------------
    elif tone == 'AG':

        suffix = (
            f"{size_num}{in_part}AG"
            if size_num
            else (f"{in_part}AG" if in_part else 'AG')
        )

    else:

        suffix = size_num

    return f"{base}-{suffix}" if suffix else base


# ============================================================
# OBU PATTERNS
# ============================================================

# Article-code line tokens:
# SKU is digit-led:
#     9-LD025-WG1.5-17.5
#
# Style/reference is letter-led:
#     BR0000367S
#     RSBR3198

_OBU_SKU_RE = re.compile(
    r'^\d+-',
    re.IGNORECASE
)

_OBU_STYLE_RE = re.compile(
    r'^[A-Z]{2,}\d',
    re.IGNORECASE
)

_OBU_STYLE_REF_RE = re.compile(
    r'^(?P<base>[A-Z]{2,}\d+[A-Z]*)'
    r'(?:-(?P<suffix>WG|YG|RG|[WYPR]G|[WYPR]))?'
    r'(?:-\d+)*$',
    re.IGNORECASE
)


# ============================================================
# LOAD OBU_CS.XLSX
# ============================================================

def _get_obu_cs():
    """
    Load OBU_CS.xlsx and cache it.
    """

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

            _OBU_CS_CACHE = pd.read_excel(
                cs_path,
                dtype=str
            )

            _OBU_CS_CACHE.columns = [
                str(c).strip()
                for c in _OBU_CS_CACHE.columns
            ]

        else:

            _OBU_CS_CACHE = pd.DataFrame()

    except Exception:

        _OBU_CS_CACHE = pd.DataFrame()

    return _OBU_CS_CACHE


# ============================================================
# METAL MATCHING
# ============================================================

def _metal_matches_cs(pdf_metal: str, cs_metal: str) -> bool:
    """
    Return True when PDF-derived metal matches an
    OBU_CS Base Metal entry.
    """

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

    return bool(
        pdf_tone and
        pdf_tone == cs_tone
    )


# ============================================================
# FIND REFERENCE IN OBU_CS
# ============================================================

def _find_reference_in_obu_cs(reference):
    """
    Determine whether a PDF reference exists in OBU_CS.xlsx.

    Checks:
        - Style No
        - Client Style No
        - Style Alias No

    Returns:
        (style_no, client_style_no)

    Example:

        BR0000367S

    could return:

        ('BR0000367S', 'BR0000367S-6.9INWG')
    """

    reference = str(reference or '').strip()

    if not reference:
        return None, None

    df = _get_obu_cs()

    if df.empty:
        return None, None

    reference_upper = reference.upper()

    # --------------------------------------------------------
    # Check Style No
    # --------------------------------------------------------

    if 'Style No' in df.columns:

        matches = df[
            df['Style No']
            .fillna('')
            .astype(str)
            .str.strip()
            .str.upper()
            == reference_upper
        ]

        if not matches.empty:

            row = matches.iloc[0]

            style_no = (
                str(row['Style No']).strip()
                if 'Style No' in df.columns
                else reference
            )

            client_style = (
                str(row['Client Style No']).strip()
                if 'Client Style No' in df.columns
                else None
            )

            return style_no, client_style

    # --------------------------------------------------------
    # Check Client Style No
    # --------------------------------------------------------

    if 'Client Style No' in df.columns:

        matches = df[
            df['Client Style No']
            .fillna('')
            .astype(str)
            .str.strip()
            .str.upper()
            == reference_upper
        ]

        if not matches.empty:

            row = matches.iloc[0]

            style_no = (
                str(row['Style No']).strip()
                if 'Style No' in df.columns
                else reference
            )

            client_style = (
                str(row['Client Style No']).strip()
                if 'Client Style No' in df.columns
                else reference
            )

            return style_no, client_style

    # --------------------------------------------------------
    # Check Style Alias No
    # --------------------------------------------------------

    if 'Style Alias No' in df.columns:

        matches = df[
            df['Style Alias No']
            .fillna('')
            .astype(str)
            .str.strip()
            .str.upper()
            == reference_upper
        ]

        if not matches.empty:

            row = matches.iloc[0]

            style_no = (
                str(row['Style No']).strip()
                if 'Style No' in df.columns
                else None
            )

            client_style = (
                str(row['Client Style No']).strip()
                if 'Client Style No' in df.columns
                else None
            )

            return style_no, client_style

    return None, None


# ============================================================
# FIND CORRECT PDF REFERENCE
# ============================================================

def _choose_obu_reference(reference_text):
    """
    PDF 'Your reference' can contain multiple references.

    Example:
        BR0000367S, RSBR3198

    The correct reference is the one that matches OBU_CS.xlsx.

    Therefore:

        BR0000367S, RSBR3198
                |
                +--> BR0000367S  <-- selected if found in OBU_CS

    Returns:
        selected_reference,
        matching_style_no,
        matching_client_style
    """

    reference_text = str(reference_text or '').strip()

    if not reference_text:
        return '', None, None

    # --------------------------------------------------------
    # Split comma-separated references
    # --------------------------------------------------------

    candidates = [
        x.strip()
        for x in re.split(r'[,;/|]+', reference_text)
        if x.strip()
    ]

    # --------------------------------------------------------
    # First priority:
    # exact match against OBU_CS
    # --------------------------------------------------------

    for candidate in candidates:

        style_no, client_style = _find_reference_in_obu_cs(
            candidate
        )

        if style_no or client_style:

            return (
                candidate,
                style_no,
                client_style
            )

    # --------------------------------------------------------
    # Second priority:
    # references beginning with BR
    #
    # This is useful when OBU_CS is temporarily unavailable.
    # --------------------------------------------------------

    for candidate in candidates:

        if re.match(
            r'^BR\d+[A-Z]*$',
            candidate,
            re.IGNORECASE
        ):

            return candidate, candidate, None

    # --------------------------------------------------------
    # Last fallback:
    # first reference
    # --------------------------------------------------------

    if candidates:

        return candidates[0], candidates[0], None

    return '', None, None


# ============================================================
# OBU LOOKUP
# ============================================================

def _lookup_obu_cs(
    style_no: str,
    item_size: str,
    sku: str = "",
    pdf_metal: str = "",
    reference_text: str = ""
):
    """
    Look up Client Style No, Base Metal and ItemSize from OBU_CS.xlsx.

    Priority:

    1. Exact SKU / Style Alias No
    2. Correct reference from 'Your reference'
    3. Style No + ItemSize + Metal

    Returns:

        client_style_no,
        base_metal,
        cs_item_size,
        matched_style_no
    """

    df = _get_obu_cs()

    if df.empty:
        return None, None, None, None

    sku = str(sku or '').strip()
    pdf_metal = str(pdf_metal or '').strip().upper()
    item_size = str(item_size or '').strip()

    # ========================================================
    # 1. SKU / STYLE ALIAS MATCH
    # ========================================================

    if sku and 'Style Alias No' in df.columns:

        alias_filtered = df[
            df['Style Alias No']
            .fillna('')
            .astype(str)
            .str.strip()
            == sku
        ]

        if not alias_filtered.empty:

            row = alias_filtered.iloc[0]

            client_style = (
                str(row['Client Style No']).strip()
                if 'Client Style No' in df.columns
                else None
            )

            base_metal = (
                str(row['Base Metal']).strip()
                if 'Base Metal' in df.columns
                else None
            )

            cs_item_size = (
                str(row['ItemSize']).strip()
                if 'ItemSize' in df.columns
                else None
            )

            matched_style_no = (
                str(row['Style No']).strip()
                if 'Style No' in df.columns
                else None
            )

            return (
                client_style,
                base_metal,
                cs_item_size,
                matched_style_no
            )

    # ========================================================
    # 2. REFERENCE MATCH
    # ========================================================

    if reference_text:

        references = [
            x.strip()
            for x in re.split(
                r'[,;/|]+',
                str(reference_text)
            )
            if x.strip()
        ]

        for reference in references:

            # ------------------------------------------------
            # Exact Style No match
            # ------------------------------------------------

            if 'Style No' in df.columns:

                filtered = df[
                    df['Style No']
                    .fillna('')
                    .astype(str)
                    .str.strip()
                    .str.upper()
                    == reference.upper()
                ]

                if not filtered.empty:

                    # ----------------------------------------
                    # Try size + metal inside matched style
                    # ----------------------------------------

                    candidates = filtered.copy()

                    if item_size and 'ItemSize' in df.columns:

                        size_filtered = candidates[
                            candidates['ItemSize']
                            .fillna('')
                            .astype(str)
                            .str.strip()
                            == item_size
                        ]

                        if not size_filtered.empty:
                            candidates = size_filtered

                    if pdf_metal and 'Base Metal' in df.columns:

                        metal_filtered = candidates[
                            candidates['Base Metal'].apply(
                                lambda m:
                                _metal_matches_cs(
                                    pdf_metal,
                                    m
                                )
                            )
                        ]

                        if not metal_filtered.empty:
                            candidates = metal_filtered

                    row = candidates.iloc[0]

                    client_style = (
                        str(row['Client Style No']).strip()
                        if 'Client Style No' in df.columns
                        else None
                    )

                    base_metal = (
                        str(row['Base Metal']).strip()
                        if 'Base Metal' in df.columns
                        else None
                    )

                    cs_item_size = (
                        str(row['ItemSize']).strip()
                        if 'ItemSize' in df.columns
                        else None
                    )

                    matched_style_no = (
                        str(row['Style No']).strip()
                        if 'Style No' in df.columns
                        else reference
                    )

                    return (
                        client_style,
                        base_metal,
                        cs_item_size,
                        matched_style_no
                    )

            # ------------------------------------------------
            # Exact Client Style No match
            # ------------------------------------------------

            if 'Client Style No' in df.columns:

                filtered = df[
                    df['Client Style No']
                    .fillna('')
                    .astype(str)
                    .str.strip()
                    .str.upper()
                    == reference.upper()
                ]

                if not filtered.empty:

                    row = filtered.iloc[0]

                    client_style = (
                        str(row['Client Style No']).strip()
                    )

                    base_metal = (
                        str(row['Base Metal']).strip()
                        if 'Base Metal' in df.columns
                        else None
                    )

                    cs_item_size = (
                        str(row['ItemSize']).strip()
                        if 'ItemSize' in df.columns
                        else None
                    )

                    matched_style_no = (
                        str(row['Style No']).strip()
                        if 'Style No' in df.columns
                        else None
                    )

                    return (
                        client_style,
                        base_metal,
                        cs_item_size,
                        matched_style_no
                    )

    # ========================================================
    # 3. STYLE NO + SIZE + METAL
    # ========================================================

    style_no = str(style_no or '').strip()

    if not style_no or 'Style No' not in df.columns:
        return None, None, None, None

    filtered = df[
        df['Style No']
        .fillna('')
        .astype(str)
        .str.strip()
        == style_no
    ]

    if filtered.empty:
        return None, None, None, None

    matched_row = None

    if 'ItemSize' in df.columns and item_size:

        size_filtered = filtered[
            filtered['ItemSize']
            .fillna('')
            .astype(str)
            .str.strip()
            == item_size
        ]

        if not size_filtered.empty:

            candidates = size_filtered

            if pdf_metal and 'Base Metal' in df.columns:

                metal_filtered = candidates[
                    candidates['Base Metal'].apply(
                        lambda m:
                        _metal_matches_cs(
                            pdf_metal,
                            m
                        )
                    )
                ]

                if not metal_filtered.empty:
                    candidates = metal_filtered

            if len(candidates) >= 1:
                matched_row = candidates.iloc[0]

    elif pdf_metal and 'Base Metal' in df.columns:

        metal_filtered = filtered[
            filtered['Base Metal'].apply(
                lambda m:
                _metal_matches_cs(
                    pdf_metal,
                    m
                )
            )
        ]

        if not metal_filtered.empty:
            matched_row = metal_filtered.iloc[0]

    if matched_row is None:
        return None, None, None, None

    client_style = (
        str(matched_row['Client Style No']).strip()
        if 'Client Style No' in df.columns
        else None
    )

    base_metal = (
        str(matched_row['Base Metal']).strip()
        if 'Base Metal' in df.columns
        else None
    )

    cs_item_size = (
        str(matched_row['ItemSize']).strip()
        if 'ItemSize' in df.columns
        else None
    )

    matched_style_no = (
        str(matched_row['Style No']).strip()
        if 'Style No' in df.columns
        else style_no
    )

    return (
        client_style,
        base_metal,
        cs_item_size,
        matched_style_no
    )


# ============================================================
# METAL -> TONE
# ============================================================

def _map_tone_from_metal(metal):

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

    return ''


# ============================================================
# TEXT -> METAL
# ============================================================

def _map_metal_from_text(text, sku=''):

    text_u = str(text or '').upper()
    sku_u = str(sku or '').upper()

    if not text_u and not sku_u:
        return ''

    compact = re.sub(
        r'[^A-Z0-9]',
        '',
        text_u
    )

    has_585 = bool(
        re.search(r'\b585\b', text_u)
        or 'STAMP585' in compact
    )

    is_yellow = bool(
        re.search(r'(?:10|14|18)KY', compact)
        or re.search(
            r'(?:10|14|18)K?Y(?![A-Z])',
            compact
        )
        or re.search(
            r'-YG(?:-|$)',
            sku_u
        )
    )

    is_white = bool(
        re.search(r'(?:10|14|18)KW', compact)
        or re.search(
            r'(?:10|14|18)K?W(?![A-Z])',
            compact
        )
        or re.search(
            r'-WG(?:-|$)',
            sku_u
        )
    )

    is_rose = bool(
        re.search(
            r'(?:10|14|18)KR',
            compact
        )
        or re.search(
            r'-RG(?:-|$)',
            sku_u
        )
    )

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


# ============================================================
# TONE FROM SKU
# ============================================================

def _tone_from_sku(sku):

    sku_u = str(sku or '').upper()

    if re.search(r'-YG(?:-|$)', sku_u):
        return 'Y'

    if re.search(r'-WG(?:-|$)', sku_u):
        return 'W'

    if re.search(r'-RG(?:-|$)', sku_u):
        return 'P'

    return ''


# ============================================================
# NORMALIZE TONE
# ============================================================

def _normalize_tone_letter(tone):

    tone = str(tone or '').strip().upper()

    if not tone:
        return ''

    if tone in ('PT', 'AG'):
        return tone

    if tone in ('WG', 'YG', 'RG'):
        return tone[0]

    if len(tone) > 1:

        return (
            tone[0]
            if tone[0] in ('W', 'Y', 'P', 'R')
            else ''
        )

    return (
        tone
        if tone in ('W', 'Y', 'P', 'R')
        else ''
    )


# ============================================================
# PARSE STYLE REFERENCE
# ============================================================

def _parse_style_reference(raw):

    raw = re.sub(
        r'[^A-Z0-9\-]',
        '',
        str(raw or '').strip().upper()
    )

    if not raw:
        return '', ''

    match = _OBU_STYLE_REF_RE.match(raw)

    if not match:

        if _OBU_STYLE_RE.match(raw):
            return raw, ''

        return '', ''

    base = match.group('base')

    suffix = (
        match.group('suffix') or ''
    ).upper()

    return (
        base,
        _normalize_tone_letter(suffix)
    )


# ============================================================
# CM -> INCH
# ============================================================

def _cm_to_inch_size(cm_value):
    """
    Convert centimeters to inches and round to 1 decimal.

    Example:
        17.5 cm -> 6.9 INCH

    The returned value intentionally contains INCH because
    _build_style_code() converts it to IN in the StyleCode.
    """

    try:

        cm_value = float(
            str(cm_value).replace(',', '.').strip()
        )

        inches = cm_value / 2.54

        # One decimal place:
        # 6.8897 -> 6.9

        inches_rounded = round(
            inches,
            1
        )

        return f"{inches_rounded:.1f} INCH"

    except (ValueError, TypeError):

        return ''


# ============================================================
# SIZE FROM TEXT
# ============================================================

def _item_size_from_text(text):
    """
    Extract bracelet/necklace/etc. size from description text.

    Examples:

        '17,5cm'
        -> '6.9 INCH'

        '17.5 cm'
        -> '6.9 INCH'

    This is intentionally based on the description rather
    than the SKU because the SKU's final number can represent
    something else.
    """

    text = str(text or '')

    # --------------------------------------------------------
    # Find:
    #
    # 17,5cm
    # 17.5cm
    # 17,5 cm
    # 17.5 cm
    # --------------------------------------------------------

    match = re.search(
        r'(\d+(?:[.,]\d+)?)\s*CM\b',
        text,
        flags=re.IGNORECASE
    )

    if not match:
        return ''

    cm_value = match.group(1)

    return _cm_to_inch_size(
        cm_value
    )


# ============================================================
# SIZE FROM SKU
# ============================================================

def _item_size_from_sku(sku, style_no=''):

    sku = str(sku or '').strip().upper()

    if not sku:
        return ''

    match = re.search(
        r'(?:WG|YG|RG)-(\d+(?:-\d+)*)',
        sku,
        re.IGNORECASE
    )

    if not match:
        return ''

    nums = re.findall(
        r'\d+',
        match.group(1)
    )

    if len(nums) < 2:
        return ''

    size = nums[-1]

    pattern = (
        _infer_style_suffix_pattern(style_no)
        if style_no
        else 'LEGACY'
    )

    if pattern == 'EU_TONE_GEM' or not style_no:
        return f'EU{size}'

    return ''


# ============================================================
# PARSE ARTICLE LINE
# ============================================================

def _parse_article_line(line):
    """
    Parse the PDF article-code row.

    Returns:

        sku_full,
        style_code,
        tone,
        item_size,
        style_raw

    IMPORTANT:
    The Article code itself does NOT contain the BR reference.

    Example:

        9-LD025-WG1.5-17.5

    The BR reference comes from the following
    'Your reference' section and is handled separately.
    """

    code_token_re = re.compile(
        r'[A-Z0-9][A-Z0-9\-\.]*[A-Z0-9]'
    )

    tokens = [
        re.sub(
            r'[^A-Z0-9\-\.]',
            '',
            token
        )
        for token in code_token_re.findall(
            str(line or '')
        )
    ]

    sku_full = ''

    style_raw = ''

    for token in tokens:

        if _OBU_SKU_RE.match(token):

            if not sku_full:
                sku_full = token

    # --------------------------------------------------------
    # Article code contains WG/YG, so get tone from SKU
    # --------------------------------------------------------

    tone = _tone_from_sku(
        sku_full
    )

    # Do NOT use RSBR3198 from the article-code parser.
    # Correct BR reference is selected from "Your reference".

    item_size = ''

    return (
        sku_full,
        style_raw,
        tone,
        item_size,
        style_raw
    )


# ============================================================
# INFER STYLE SUFFIX PATTERN
# ============================================================

def _infer_style_suffix_pattern(style_no):

    style_no = str(style_no or '').strip()

    if not style_no:
        return 'LEGACY'

    if style_no in _STYLE_PATTERN_CACHE:
        return _STYLE_PATTERN_CACHE[style_no]

    pattern = 'LEGACY'

    df = _get_obu_cs()

    if not df.empty and 'Style No' in df.columns:

        rows = df[
            df['Style No']
            .fillna('')
            .astype(str)
            .str.strip()
            == style_no
        ]

        suffixes = []

        for val in rows.get(
            'Client Style No',
            pd.Series(dtype=str)
        ).dropna():

            client_style = str(val).strip()

            if client_style.startswith(style_no):

                suffixes.append(
                    client_style[len(style_no):]
                )

        if suffixes:

            eu_gem = sum(
                1
                for suffix in suffixes
                if re.match(
                    r'-EU\d+[WYPR]GEM$',
                    suffix,
                    re.I
                )
            )

            inch_tone_gem = sum(
                1
                for suffix in suffixes
                if re.match(
                    r'-\d+(?:\.\d+)?IN[WYPR]GEM$',
                    suffix,
                    re.I
                )
            )

            tone_gem = sum(
                1
                for suffix in suffixes
                if re.match(
                    r'-[WYPR]GEM$',
                    suffix,
                    re.I
                )
            )

            tone_pair = sum(
                1
                for suffix in suffixes
                if re.match(
                    r'-(WG|YG|RG)$',
                    suffix,
                    re.I
                )
            )

            inch_tone = sum(
                1
                for suffix in suffixes
                if re.match(
                    r'-\d+(?:\.\d+)?IN[WYPR]G$',
                    suffix,
                    re.I
                )
            )

            # ------------------------------------------------
            # IMPORTANT:
            # Support:
            #
            # BR0000367S-6.9INWG
            #
            # ------------------------------------------------

            if inch_tone >= max(
                eu_gem,
                inch_tone_gem,
                tone_gem,
                tone_pair
            ):
                pattern = 'INCH_TONE'

            elif inch_tone_gem >= max(
                eu_gem,
                tone_gem,
                tone_pair
            ):
                pattern = 'INCH_TONE_GEM'

            elif eu_gem >= max(
                tone_gem,
                tone_pair
            ):
                pattern = 'EU_TONE_GEM'

            elif tone_pair >= tone_gem:
                pattern = 'TONE_PAIR'

            elif tone_gem or any(
                suffix == '-AGEM'
                for suffix in suffixes
            ):
                pattern = 'TONE_GEM'

    _STYLE_PATTERN_CACHE[style_no] = pattern

    return pattern


# ============================================================
# BUILD OBU STYLE CODE
# ============================================================

def _build_obu_style_code(
    style_no,
    item_size,
    tone,
    metal=''
):
    """
    Build StyleCode for OBU output.

    Important examples:

        BR0000367S
        6.9 INCH
        W

        ->
        BR0000367S-6.9INWG

    Also prevents an already-complete StyleCode from
    receiving the suffix twice.
    """

    style_no = str(
        style_no or ''
    ).strip()

    item_size = str(
        item_size or ''
    ).strip()

    tone = _normalize_tone_letter(
        tone
    )

    metal = str(
        metal or ''
    ).strip().upper()

    if not style_no or style_no.upper() == 'NAN':
        return ''

    # --------------------------------------------------------
    # IMPORTANT:
    # If the style already contains the requested suffix,
    # return it as-is.
    # --------------------------------------------------------

    clean_style = style_no.rstrip()

    # Already has:
    #
    # -6.9INWG
    # -6.9ING
    # -6.9INYG
    # etc.

    if re.search(
        r'-\d+(?:\.\d+)?IN(?:WG|YG|RG|[WYPR]G)$',
        clean_style,
        re.IGNORECASE
    ):
        return clean_style

    # --------------------------------------------------------
    # If style already has a standard tone suffix, remove
    # it before rebuilding.
    # --------------------------------------------------------

    base_match = re.match(
        r'^(.*?)-(?:WG|YG|RG|[WYPR]G)$',
        clean_style,
        re.IGNORECASE
    )

    if base_match:
        clean_style = base_match.group(1)

    style_no = clean_style

    # --------------------------------------------------------
    # Normalize item size
    # --------------------------------------------------------

    if item_size:

        item_size = item_size.replace(
            ',',
            '.'
        )

    pattern = _infer_style_suffix_pattern(
        style_no
    )

    # ========================================================
    # NEW:
    # CM converted to INCH + W/Y/P/R
    #
    # BR0000367S + 6.9 INCH + W
    # =
    # BR0000367S-6.9INWG
    # ========================================================

    if (
        item_size
        and re.search(
            r'\bINCH\b|\bIN\b',
            item_size,
            re.IGNORECASE
        )
        and tone in ('W', 'Y', 'P', 'R')
    ):

        # Even if OBU_CS has a different inferred pattern,
        # the explicit PDF size should be used.

        return _build_style_code(
            style_no,
            item_size,
            tone
        )

    # ========================================================
    # Existing EU GEM pattern
    # ========================================================

    if (
        pattern == 'EU_TONE_GEM'
        and item_size
        and tone in ('W', 'Y', 'P', 'R')
    ):

        return (
            f'{style_no}-{item_size}'
            f'{tone}GEM'
        )

    # ========================================================
    # Existing INCH TONE pattern
    # ========================================================

    if (
        pattern == 'INCH_TONE'
        and item_size
        and tone in ('W', 'Y', 'P', 'R')
    ):

        return _build_style_code(
            style_no,
            item_size,
            tone
        )

    # ========================================================
    # Existing INCH TONE GEM pattern
    # ========================================================

    if (
        pattern == 'INCH_TONE_GEM'
        and item_size
        and tone in ('W', 'Y', 'P', 'R')
    ):

        size_num = re.sub(
            r'\s*(?:INCH|IN)\s*$',
            '',
            item_size,
            flags=re.IGNORECASE
        ).strip()

        return (
            f'{style_no}-'
            f'{size_num}IN'
            f'{tone}GEM'
        )

    # ========================================================
    # TONE GEM
    # ========================================================

    if pattern == 'TONE_GEM':

        if metal == 'AG925':
            return f'{style_no}-AGEM'

        if tone in ('W', 'Y', 'P', 'R'):
            return f'{style_no}-{tone}GEM'

    # ========================================================
    # TONE PAIR
    # ========================================================

    if (
        pattern == 'TONE_PAIR'
        and tone in ('W', 'Y', 'P', 'R')
    ):

        pair_map = {
            'W': 'WG',
            'Y': 'YG',
            'P': 'RG',
            'R': 'RG'
        }

        return (
            f'{style_no}-'
            f'{pair_map[tone]}'
        )

    # ========================================================
    # Default
    # ========================================================

    return _build_style_code(
        style_no,
        item_size,
        tone
    )


# ============================================================
# ITEM SIZE MASTER
# ============================================================

def _get_size_lookup():

    global _ITEM_SIZE_LOOKUP

    if _ITEM_SIZE_LOOKUP is not None:
        return _ITEM_SIZE_LOOKUP

    _ITEM_SIZE_LOOKUP = {}

    try:

        mst = os.path.join(
            os.path.dirname(
                os.path.abspath(__file__)
            ),
            'ItemSize_Mst.xlsx'
        )

        _df_mst = pd.read_excel(mst)

        for val in _df_mst[
            'Item Size Code'
        ].dropna():

            vs = str(val).strip()

            if (
                vs
                and vs.upper() != 'NAN'
            ):

                k = _normalize_size_key(
                    vs
                )

                if k:
                    _ITEM_SIZE_LOOKUP[k] = vs

    except Exception:
        pass

    return _ITEM_SIZE_LOOKUP


# ============================================================
# NORMALIZE SIZE KEY
# ============================================================

def _normalize_size_key(s):

    s = str(s).strip()

    if not s or s.upper() == 'NAN':
        return ''

    m = re.match(
        r'^(\d+(?:\.\d+)?)\s*INCH$',
        s,
        re.IGNORECASE
    )

    if m:

        return (
            f"{float(m.group(1)):.2f}inch"
        )

    return re.sub(
        r'\s+',
        '',
        s
    ).lower()


# ============================================================
# MAP ITEM SIZE
# ============================================================

def _map_item_size(raw):

    if (
        not raw
        or str(raw).strip().upper()
        in ('', 'NAN')
    ):
        return raw

    lookup = _get_size_lookup()

    key = _normalize_size_key(
        str(raw).strip()
    )

    return lookup.get(
        key,
        raw
    )


# ============================================================
# READ PDF
# ============================================================

def _read_pdf_text(pdf_path: str) -> str:

    with pdfplumber.open(pdf_path) as pdf:

        return "\n".join(
            [
                page.extract_text() or ''
                for page in pdf.pages
            ]
        )


# ============================================================
# ORDER DATE
# ============================================================

def _extract_order_date(text: str) -> str:

    match = re.search(
        r'Order date\s*:\s*(.+)',
        text,
        re.IGNORECASE
    )

    return (
        match.group(1).strip()
        if match
        else ''
    )


# ============================================================
# EXTRACT YOUR REFERENCE
# ============================================================

def _extract_your_reference(block):
    """
    Extract the 'Your reference' line from a PDF block.

    Example:

        Your reference
        BR0000367S, RSBR3198

    Returns:

        BR0000367S, RSBR3198
    """

    for i, line in enumerate(block):

        if re.match(
            r'^Your reference',
            line,
            re.IGNORECASE
        ):

            # ------------------------------------------------
            # Usually the reference is on the next line.
            # ------------------------------------------------

            if i + 1 < len(block):

                next_line = str(
                    block[i + 1]
                ).strip()

                if next_line:
                    return next_line

    return ''


# ============================================================
# PROCESS OBU FILE
# ============================================================

def process_obu_file(
    input_path: str,
    output_dir: str
):

    try:

        base_name = os.path.splitext(
            os.path.basename(input_path)
        )[0]

        ext = os.path.splitext(
            input_path
        )[1].lower()

        # ====================================================
        # PDF
        # ====================================================

        if ext == '.pdf':

            text = _read_pdf_text(
                input_path
            )

            code_token_re = re.compile(
                r'[A-Z0-9][A-Z0-9\-\.]*[A-Z0-9]'
            )

            po_re = re.compile(
                r'PO#\s*:\s*(\d+)'
            )

            article_header_re = re.compile(
                r'^Article code',
                re.IGNORECASE
            )

            quantity_line_re = re.compile(
                r'^(\d+)\s+(\d+)$'
            )

            po_match = po_re.search(
                text
            )

            item_po_no = (
                po_match.group(1)
                if po_match
                else ''
            )

            order_date = _extract_order_date(
                text
            )

            lines = [
                ln.strip()
                for ln in text.split('\n')
                if ln.strip()
            ]

            blocks = []
            current = []

            # =================================================
            # BUILD ARTICLE BLOCKS
            # =================================================

            for ln in lines:

                if article_header_re.match(ln):

                    if current:
                        blocks.append(
                            current
                        )

                        current = []

                    current.append(ln)

                elif current:

                    current.append(ln)

            if current:
                blocks.append(current)

            items = []

            # =================================================
            # PROCESS EACH ITEM
            # =================================================

            for b_index, block in enumerate(blocks):

                is_last_block = (
                    b_index == len(blocks) - 1
                )

                btxt = "\n".join(block)

                # ---------------------------------------------
                # ARTICLE CODE
                # ---------------------------------------------

                article_line = ''

                for ln in block[1:4]:

                    if (
                        _OBU_SKU_RE.search(ln)
                        or _OBU_STYLE_RE.search(ln)
                    ):

                        article_line = ln

                        break

                if article_line:

                    (
                        sku_full,
                        style_code,
                        tone,
                        item_size,
                        style_raw
                    ) = _parse_article_line(
                        article_line
                    )

                else:

                    sku_full = ''
                    style_code = ''
                    tone = ''
                    item_size = ''
                    style_raw = ''

                # ---------------------------------------------
                # YOUR REFERENCE
                # ---------------------------------------------

                your_reference = (
                    _extract_your_reference(
                        block
                    )
                )

                # ---------------------------------------------
                # CHOOSE CORRECT OBU REFERENCE
                #
                # Example:
                #
                # BR0000367S, RSBR3198
                #
                # becomes:
                #
                # BR0000367S
                # ---------------------------------------------

                (
                    selected_reference,
                    reference_style_no,
                    reference_client_style
                ) = _choose_obu_reference(
                    your_reference
                )

                # ---------------------------------------------
                # ITEM REF NO
                #
                # Keep the full reference line because this
                # is useful for tracing.
                # ---------------------------------------------

                item_ref_no = your_reference

                # ---------------------------------------------
                # QUANTITY
                # ---------------------------------------------

                sr_no = ''
                order_qty = ''
                order_item_pcs = ''

                try:

                    desc_idx = next(
                        i
                        for i, ln in enumerate(block)
                        if ln.lower().startswith(
                            'description'
                        )
                    )

                except StopIteration:

                    desc_idx = None

                if desc_idx is not None:

                    for ln in block[
                        desc_idx:
                        desc_idx + 8
                    ]:

                        qm = quantity_line_re.match(
                            ln
                        )

                        if qm:

                            sr_no = qm.group(1)

                            order_qty = qm.group(2)

                            order_item_pcs = order_qty

                            break

                # ---------------------------------------------
                # DESCRIPTION
                # ---------------------------------------------

                desc_lines = []

                for ln in block:

                    if (
                        article_header_re.match(ln)
                        or quantity_line_re.match(ln)
                        or ln.lower().startswith(
                            'description'
                        )
                        or re.match(
                            r'^Your reference',
                            ln,
                            re.IGNORECASE
                        )
                    ):
                        continue

                    # Skip pure reference/code lines
                    if code_token_re.fullmatch(
                        ln.replace(' ', '')
                    ):
                        continue

                    desc_lines.append(ln)

                full_desc = ' '.join(
                    desc_lines
                )

                # Remove PO total from final block

                if is_last_block:

                    pot_idx = re.search(
                        r'Purchase order Total',
                        full_desc,
                        flags=re.IGNORECASE
                    )

                    if pot_idx:

                        full_desc = (
                            full_desc[
                                :pot_idx.start()
                            ].strip()
                        )

                # ---------------------------------------------
                # CUSTOMER / STAMP INSTRUCTION
                # ---------------------------------------------

                split_match = re.search(
                    r'\b(stamp\b.*)',
                    full_desc,
                    flags=re.IGNORECASE
                )

                if split_match:

                    customer_instr = (
                        full_desc[
                            :split_match.start()
                        ].strip()
                    )

                    stamp_instr = (
                        full_desc[
                            split_match.start():
                        ].strip()
                    )

                else:

                    customer_instr = full_desc

                    stamp_instr = ''

                customer_instr = re.sub(
                    r'\s*\band\s*$',
                    '',
                    customer_instr,
                    flags=re.IGNORECASE
                ).strip()

                # ---------------------------------------------
                # CERTIFICATE
                # ---------------------------------------------

                certificate = ''

                if sku_full:

                    nums = re.findall(
                        r'\d+',
                        sku_full
                    )

                    if nums and nums[-1] == '100':

                        certificate = (
                            'IGI Certified'
                        )

                sku_no = (
                    sku_full
                    if sku_full
                    else ''
                )

                # ---------------------------------------------
                # METAL FROM DESCRIPTION + SKU
                # ---------------------------------------------

                pdf_metal = _map_metal_from_text(
                    full_desc,
                    sku_full
                )

                if not tone:

                    tone = _tone_from_sku(
                        sku_full
                    )

                if not tone and pdf_metal:

                    tone = _map_tone_from_metal(
                        pdf_metal
                    )

                # ---------------------------------------------
                # NEW:
                # EXTRACT SIZE FROM DESCRIPTION
                #
                # 17,5cm
                #
                # ->
                #
                # 6.9 INCH
                # ---------------------------------------------

                pdf_item_size = (
                    _item_size_from_text(
                        full_desc
                    )
                )

                # ---------------------------------------------
                # If description contains CM size, it takes
                # priority over the SKU-derived size.
                # ---------------------------------------------

                if pdf_item_size:

                    item_size = pdf_item_size

                elif not item_size:

                    item_size = _item_size_from_sku(
                        sku_full,
                        reference_style_no
                        or style_code
                    )

                # ---------------------------------------------
                # OBU CS LOOKUP
                # ---------------------------------------------

                (
                    client_style,
                    metal,
                    cs_item_size,
                    matched_style_no
                ) = _lookup_obu_cs(
                    reference_style_no
                    or selected_reference
                    or style_code,

                    item_size,

                    sku_full,

                    pdf_metal=pdf_metal,

                    reference_text=your_reference
                )

                # ---------------------------------------------
                # DETERMINE BASE STYLE
                # ---------------------------------------------

                if matched_style_no:

                    base_style_no = (
                        matched_style_no
                    )

                elif reference_style_no:

                    base_style_no = (
                        reference_style_no
                    )

                elif selected_reference:

                    base_style_no = (
                        selected_reference
                    )

                else:

                    base_style_no = style_code

                # ---------------------------------------------
                # IMPORTANT:
                #
                # We DO NOT directly use client_style here
                # because the desired output is:
                #
                # BR0000367S-6.9INWG
                #
                # rather than potentially returning an already
                # stored/old Client Style No.
                #
                # The BR style is used as the base and the
                # current PDF size + tone are appended.
                # ---------------------------------------------

                final_item_size = (
                    cs_item_size
                    if (
                        cs_item_size
                        and not pdf_item_size
                    )
                    else item_size
                )

                # ---------------------------------------------
                # Metal fallback
                # ---------------------------------------------

                if not metal and pdf_metal:

                    metal = pdf_metal

                # ---------------------------------------------
                # Tone from final metal
                # ---------------------------------------------

                if metal:

                    tone = _map_tone_from_metal(
                        metal
                    )

                elif pdf_metal:

                    metal = pdf_metal

                    tone = _map_tone_from_metal(
                        pdf_metal
                    )

                # ---------------------------------------------
                # AG925
                # ---------------------------------------------

                if (
                    str(metal).upper()
                    == 'AG925'
                ):

                    tone = 'AG'

                # ---------------------------------------------
                # BUILD FINAL STYLE CODE
                # ---------------------------------------------

                final_style_code = (
                    _build_obu_style_code(
                        base_style_no,
                        final_item_size,
                        'AG'
                        if str(metal).upper()
                        == 'AG925'
                        else tone,
                        str(metal or '')
                    )
                )

                # ---------------------------------------------
                # SAFETY:
                #
                # Remove accidental duplicate suffix:
                #
                # BR0000367S-6.9INWG-6.9INWG
                #
                # -> BR0000367S-6.9INWG
                # ---------------------------------------------

                final_style_code = re.sub(
                    r'(-\d+(?:\.\d+)?IN(?:WG|YG|RG|[WYPR]G))'
                    r'(?:\1)+$',
                    r'\1',
                    final_style_code,
                    flags=re.IGNORECASE
                )

                # ---------------------------------------------
                # ADD ITEM
                # ---------------------------------------------

                items.append({

                    'SrNo':
                        sr_no,

                    'StyleCode':
                        final_style_code,

                    'ItemSize':
                        final_item_size,

                    'OrderQty':
                        order_qty,

                    'OrderItemPcs':
                        1,

                    'Metal':
                        metal,

                    'Tone':
                        ''
                        if str(metal).upper()
                        == 'AG925'
                        else tone,

                    'ItemPoNo':
                        item_po_no,

                    'ItemRefNo':
                        item_ref_no,

                    'StockType':
                        '',

                    'MakeType':
                        '',

                    'CustomerProductionInstruction':
                        customer_instr,

                    'SpecialRemarks':
                        '',

                    'DesignProductionInstruction':
                        '',

                    'StampInstruction':
                        stamp_instr,

                    'OrderGroup':
                        '',

                    'Certificate':
                        certificate,

                    'SKUNo':
                        sku_no,

                    'Basestoneminwt':
                        '',

                    'Basestonemaxwt':
                        '',

                    'Basemetalminwt':
                        '',

                    'Basemetalmaxwt':
                        '',

                    'Productiondeliverydate':
                        '',

                    'Expecteddeliverydate':
                        '',

                    'Blank':
                        '',

                    'SetPrice':
                        '',

                    'StoneQuality':
                        'VVS+'
                        if re.search(
                            r'\bVVS\+\b',
                            btxt
                        )
                        else '',

                    'Date':
                        order_date,
                })

            # =================================================
            # OUTPUT COLUMNS
            # =================================================

            columns_order = [

                'SrNo',
                'StyleCode',
                'ItemSize',
                'OrderQty',
                'OrderItemPcs',
                'Metal',
                'Tone',
                'ItemPoNo',
                'ItemRefNo',
                'StockType',
                'MakeType',
                'CustomerProductionInstruction',
                'SpecialRemarks',
                'DesignProductionInstruction',
                'StampInstruction',
                'OrderGroup',
                'Certificate',
                'SKUNo',
                'Basestoneminwt',
                'Basestonemaxwt',
                'Basemetalminwt',
                'Basemetalmaxwt',
                'Productiondeliverydate',
                'Expecteddeliverydate',
                'Blank',
                'SetPrice',
                'StoneQuality',
                'Date'
            ]

            df = pd.DataFrame(
                items,
                columns=columns_order
            )

            # =================================================
            # OUTPUT
            # =================================================

            output_path = os.path.join(
                output_dir,
                f"{base_name}_OBU_MAPPED.xlsx"
            )

            df.to_excel(
                output_path,
                index=False
            )

            return (
                True,
                output_path,
                None,
                df
            )

        # ====================================================
        # EXCEL / CSV PASSTHROUGH
        # ====================================================

        elif ext in [
            '.xlsx',
            '.xls',
            '.csv'
        ]:

            if ext in [
                '.xlsx',
                '.xls'
            ]:

                df = pd.read_excel(
                    input_path
                )

            else:

                df = pd.read_csv(
                    input_path
                )

            output_path = os.path.join(
                output_dir,
                f"{base_name}_OBU_PASSTHROUGH.xlsx"
            )

            df.to_excel(
                output_path,
                index=False
            )

            return (
                True,
                output_path,
                None,
                df
            )

        # ====================================================
        # UNSUPPORTED
        # ====================================================

        else:

            return (
                False,
                None,
                f"Unsupported file type: {ext}",
                None
            )

    except Exception as e:

        return (
            False,
            None,
            str(e),
            None
        )