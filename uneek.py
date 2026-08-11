import os
import re
import pandas as pd
import pdfplumber


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'US06', 'W') -> 'VR1943EEA-6WG'
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
    """Load and cache the approved Client Style No list from UNEEK_CS_100826.xlsx."""
    global _CLIENT_STYLE_LIST
    if _CLIENT_STYLE_LIST is not None:
        return _CLIENT_STYLE_LIST
    _CLIENT_STYLE_LIST = []
    try:
        cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'UNEEK_CS_100826.xlsx')
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
    Match the built StyleCode against UNEEK_CS_100826.xlsx 'Client Style No'.

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


def extract_uneek_data_from_pdf(pdf_path: str, style_code: str, order_qty: str, 
                                  user_input1: str, user_input2: str, stamp_instruction: str):
    """
    Extract data from Uneek PDF and organize into 26 columns
    Based on logic from Jupyter_Notebooks/uneeek.ipynb
    """
    data = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Extract text from all pages
            full_text = ""
            for page in pdf.pages:
                full_text += page.extract_text() + "\n"
            
            # Extract PO Number
            po_match = re.search(r'Purchase Order Number:\s*(PO-\d+)', full_text)
            item_po_no = po_match.group(1) if po_match else ""
            
            # Process the extracted text line by line
            lines = full_text.split('\n')
            
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                
                # Look for item lines with pattern: Number COLVBW#### ... 
                # Example: 1 COLVBW1534 R061534 315447 26RD=0.50CTW UFJC- 6 SO- 1
                item_match = re.match(r'^(\d+)\s+(COLV[A-Z0-9]+)\s+([A-Z0-9]+)\s+(\d+)\s+(.+)', line)
                
                if item_match:
                    sr_no = item_match.group(1)
                    sku_no = item_match.group(2)
                    description_part1 = item_match.group(5).strip()
                    
                    # Look ahead to the next line for metal and size info
                    # Example next line: W 18KW SZ6 METAL- 003345
                    metal_raw = ""
                    size_from_text = ""
                    
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        
                        # Extract metal (e.g., 18KW)
                        metal_match = re.search(r'(\d+K[A-Z])', next_line)
                        if metal_match:
                            metal_raw = metal_match.group(1)
                        
                        # Extract size (e.g., SZ6)
                        size_match = re.search(r'SZ(\d+)', next_line)
                        if size_match:
                            size_from_text = size_match.group(1)
                    
                    # Build CustomerProductionInstruction
                    # Extract the stone/diamond info from description_part1
                    # Example: "26RD=0.50CTW UFJC- 6 SO- 1" -> extract "26RD=0.50CTW"
                    stone_info_match = re.match(r'^([^\s]+(?:=|CTW)[^\s]*(?:\s+[^\s]+(?:=|CTW)[^\s]*)*)', description_part1)
                    if stone_info_match:
                        stone_info = stone_info_match.group(1)
                        # Clean up: remove trailing non-alphanumeric except = and .
                        stone_info = re.sub(r'\s+[A-Z]+-\s*$', '', stone_info).strip()
                    else:
                        stone_info = description_part1.split()[0] if description_part1 else ""
                    
                    # Construct CustomerProductionInstruction: stone_info + metal + size
                    if metal_raw and size_from_text:
                        customer_production_instruction = f"{stone_info} {metal_raw} SZ{size_from_text}"
                    else:
                        customer_production_instruction = description_part1
                    
                    # Extract ItemSize
                    if size_from_text:
                        item_size = f"US{size_from_text.zfill(2)}"
                    else:
                        item_size = ""
                    
                    # Process metal and tone
                    # Convert 18KW -> G18W (not G18KW)
                    if metal_raw:
                        # Extract karat number and tone
                        # metal_raw is like "18KW", we want "G18W"
                        karat_num = metal_raw[:-1]  # "18K"
                        tone = metal_raw[-1]  # "W"
                        metal = f"G{karat_num[:-1]}{tone}"  # "G" + "18" + "W" = "G18W"
                    else:
                        metal = ""
                        tone = ""
                    
                    # Determine metal_tone for SpecialRemarks
                    if metal:
                        karat = metal[1:-1]  # Extract karat number (e.g., "18" from "G18W")
                        if tone == 'W':
                            metal_tone = f"{karat}K WHITE GOLD"
                        elif tone == 'Y':
                            metal_tone = f"{karat}K YELLOW GOLD"
                        elif tone == 'R':
                            metal_tone = f"{karat}K ROSE GOLD"
                        else:
                            metal_tone = f"{karat}K GOLD"
                    else:
                        metal_tone = ""
                    
                    # Extract size for SpecialRemarks (e.g., SZ 6)
                    if size_from_text:
                        size_display = f"SZ {size_from_text}"
                    else:
                        size_display = f"SZ {item_size}"
                    
                    # Build SpecialRemarks
                    # Format: user_input1 + metal_tone + size + Dia_qlty-abc(userinput2)
                    special_remarks = f"{user_input1},{metal_tone},{size_display},DIA QLTY-{user_input2}"
                    
                    # Determine DesignProductionInstruction based on tone
                    if tone == 'W':
                        design_production_instruction = "White Rodium"
                    else:
                        design_production_instruction = "No rodium"
                    
                    # Create row with all 26 columns
                    row = {
                        'SrNo': sr_no,
                        'StyleCode': style_code,
                        'ItemSize': item_size,
                        'OrderQty': order_qty,
                        'OrderItemPcs': '1',
                        'Metal': metal,
                        'Tone': tone,
                        'ItemPoNo': item_po_no,
                        'ItemRefNo': '',
                        'StockType': '',
                        'MakeType': '',
                        'CustomerProductionInstruction': customer_production_instruction,
                        'SpecialRemarks': special_remarks,
                        'DesignProductionInstruction': design_production_instruction,
                        'StampInstruction': stamp_instruction,
                        'OrderGroup': '',
                        'Certificate': '',
                        'SKUNo': sku_no,
                        'Basestoneminwt': '',
                        'Basestonemaxwt': '',
                        'Basemetalminwt': '',
                        'Basemetalmaxwt': '',
                        'Productiondeliverydate': '',
                        'Expecteddeliverydate': '',
                        'SetPrice': '',
                        'StoneQuality': ''
                    }
                    
                    data.append(row)
                
                i += 1
    
    except Exception as e:
        raise Exception(f"Error reading PDF file: {str(e)}")
    
    # Create DataFrame with all columns
    df = pd.DataFrame(data)
    df['ItemSize'] = df['ItemSize'].apply(_map_item_size)
    df['StyleCode'] = df.apply(
        lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
    )
    df['StyleCode'] = df['StyleCode'].apply(_lookup_client_style)
    df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''
    return df


def process_uneek_file(input_path: str, output_dir: str, style_code: str = "", order_qty: str = "1", 
                        user_input1: str = "", user_input2: str = "", stamp_instruction: str = ""):
    """
    Process UNEEK PDF file with proper column mapping.
    Based on logic from Jupyter_Notebooks/uneeek.ipynb.
    """
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()
        
        if ext != '.pdf':
            return False, None, f"Unsupported file type: {ext}. Only PDF files are supported.", None
        
        # Extract data from PDF
        df = extract_uneek_data_from_pdf(
            input_path, 
            style_code, 
            order_qty, 
            user_input1, 
            user_input2, 
            stamp_instruction
        )
        
        if df.empty:
            return False, None, "No data could be extracted from the PDF.", None
        
        # Save output
        output_path = os.path.join(output_dir, f"{base_name}_UNEEK_MAPPED.xlsx")
        df.to_excel(output_path, index=False)
        
        return True, output_path, None, df
        
    except Exception as e:
        return False, None, str(e), None
