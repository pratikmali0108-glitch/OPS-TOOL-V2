import pandas as pd
import numpy as np
import os
import re

def _get_amp_cs_path():
    """Get the path to the AMP client style master file."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_100826', 'AMP_CS_120826.xlsx')

def _load_amp_reference():
    """Load the AMP reference lookup dictionary."""
    reference_path = _get_amp_cs_path()
    if not os.path.exists(reference_path):
        return {}
    
    try:
        ref_df = pd.read_excel(reference_path)
        ref_lookup = {}
        for _, row in ref_df.iterrows():
            style_no = str(row.get('Style No', '')).strip()
            base_metal = str(row.get('Base Metal', '')).strip().upper()
            client_style = str(row.get('Client Style No', '')).strip()
            
            if style_no and base_metal and client_style:
                ref_lookup[(style_no, base_metal)] = client_style
        return ref_lookup
    except Exception:
        return {}

def _add_style_to_master(style_no, base_metal, client_style):
    """Add a new style to the AMP master file."""
    cs_path = _get_amp_cs_path()
    try:
        if os.path.exists(cs_path):
            df_cs = pd.read_excel(cs_path)
        else:
            df_cs = pd.DataFrame(columns=['Style No', 'Base Metal', 'Client Style No'])
        
        # Check if this combination already exists
        mask = (df_cs['Style No'].astype(str).str.strip() == style_no) & \
               (df_cs['Base Metal'].astype(str).str.strip().upper() == base_metal)
        
        if not mask.any():
            new_row = pd.DataFrame([{
                'Style No': style_no,
                'Base Metal': base_metal,
                'Client Style No': client_style
            }])
            df_cs = pd.concat([df_cs, new_row], ignore_index=True)
            df_cs.to_excel(cs_path, index=False)
            return True
        return False
    except Exception:
        return False

def get_metal_code(row):
    """Extract metal code from Metal Type and Tone."""
    metal_type = str(row['Metal Type']).upper()
    tone = str(row['Tone']).upper()

    if '14K' in metal_type or '14KT' in metal_type:
        if 'Y' in tone:
            return 'G14Y'
        elif 'W' in tone:
            return 'G14W'
    elif '18K' in metal_type or '18KT' in metal_type:
        if 'Y' in tone:
            return 'G18Y'
        elif 'W' in tone:
            return 'G18W'
    return None

def get_item_size(row):
    """Format item size based on Size and Category."""
    size_val = row['_Original_Size']
    category = str(row['_Original_Category']).upper()

    if pd.notna(size_val):
        category_match = any(
            kw in category
            for kw in ['BRACLET', 'BRACELET', 'NECKLACE', 'NECK PEICE', 'NECK']
        )
        if category_match:
            return f"{float(size_val):.2f} INCH"
    return size_val

def format_size_for_stylecode(size_val):
    """Format size value for StyleCode construction."""
    if pd.notna(size_val):
        fsize = float(size_val)
        if fsize == int(fsize):
            return f"{int(fsize)}IN"
        else:
            return str(fsize)
    return ""

def format_tone_variants(tone_val):
    """Get tone variants for StyleCode matching."""
    tone = str(tone_val).upper()
    variants = []
    if 'W' in tone:
        variants.extend(['WG', 'WGGD', 'W'])
    elif 'Y' in tone:
        variants.extend(['YG', 'YGGD', 'Y'])
    if not variants:
        variants = [tone]
    return variants

def get_stylecode(row, ref_lookup):
    """Build StyleCode using reference lookup or structured fallback."""
    style_r = str(row['StyleCodeR']).strip()
    metal = str(row['Metal']).strip().upper()
    
    # First, try to lookup in reference file using StyleCodeR and Metal
    lookup_key = (style_r, metal)
    if lookup_key in ref_lookup:
        return ref_lookup[lookup_key], True  # Found in master
    
    # If not found, build structured logic fallback
    size_val = row['_Original_Size']
    tone_val = row['Tone']

    size_part = format_size_for_stylecode(size_val)
    tone_variants = format_tone_variants(tone_val)

    # Create structured StyleCode as fallback
    structured_stylecode = f"{style_r}-{size_part}{tone_variants[0]}"
    
    return structured_stylecode, False  # Not found in master

def get_compact_size_for_remarks(size_val):
    """Format compact size for remarks."""
    if pd.notna(size_val):
        fsize = float(size_val)
        if fsize == int(fsize):
            return f"{int(fsize)}INCH"
        else:
            return f"{fsize}INCH"
    return ""

def get_metal_desc(metal_code):
    """Get metal description from code."""
    mapping = {
        'G14W': '14KT WHITE',
        'G14Y': '14KT YELLOW',
        'G18W': '18KT WHITE',
        'G18Y': '18KT YELLOW',
    }
    return mapping.get(str(metal_code).upper(), '')

def get_special_remarks(row):
    """Build SpecialRemarks field."""
    amipi_style = str(row['_Original_AmipiStyle']).strip()
    sz_compact = get_compact_size_for_remarks(row['_Original_Size'])
    metal_desc = get_metal_desc(row['Metal'])
    category = str(row['_Original_Category']).strip().upper()

    parts = []
    if amipi_style and amipi_style.lower() != 'nan':
        parts.append(amipi_style)
    if sz_compact:
        parts.append(f"SZ: {sz_compact}")
    if metal_desc and category and category.lower() != 'nan':
        parts.append(f"{metal_desc} {category}")
    elif metal_desc:
        parts.append(metal_desc)
    parts.append("DIAMOND QUALITY-")

    return ", ".join(parts)

def process_amp_file(input_path: str, output_dir: str):
    """
    Process AMP Excel file and return Flask-compatible results.
    
    Args:
        input_path: Path to the input Excel file
        output_dir: Directory to save the output file
        
    Returns:
        Tuple of (success: bool, output_path: str|None, error_message: str|None, df: pd.DataFrame|None)
    """
    try:
        # Load reference lookup
        ref_lookup = _load_amp_reference()
        
        # Read Excel file
        df = pd.read_excel(input_path)
        
        # Validate required columns
        required_columns = ['Metal Color', 'Description', 'Stamping Instruction', 'Vendor Style', 
                          'Order Qty', 'PURCHASE ORDER#', 'REMARK', 'Fulfil by Date',
                          'Metal Type', 'Size', 'Category']
        
        # Check for Amipi Style column (optional)
        has_amipi = 'Amipi Style' in df.columns
        
        missing = [c for c in required_columns if c not in df.columns]
        if missing:
            return False, None, f'Missing required columns: {", ".join(missing)}', None
        
        # Create SrNo column
        df['SrNo'] = df.index + 1
        
        # Save original columns before any rename
        df['_Original_Size'] = df['Size']
        df['_Original_Category'] = df['Category']
        df['_Original_AmipiStyle'] = df['Amipi Style'] if has_amipi else ''
        
        # Rename columns
        df = df.rename(columns={
            'Metal Color': 'Tone',
            'Description': 'CustomerProductionInstruction',
            'Stamping Instruction': 'StampInstruction',
            'Vendor Style': 'StyleCodeR',
            'Order Qty': 'OrderQty',
            'PURCHASE ORDER#': 'ItemPoNo',
            'REMARK': 'SpecialRemarks',
            'Fulfil by Date': 'Productiondeliverydate'
        })
        
        # Add Metal column
        df['Metal'] = df.apply(get_metal_code, axis=1)
        
        # Add DesignProductionInstruction column
        df['DesignProductionInstruction'] = np.where(
            df['Tone'].str.upper() == 'W', 'WHITE RHODIUM', 'NO RHODIUM'
        )
        
        # Create ItemSize column
        df['ItemSize'] = df.apply(get_item_size, axis=1)
        
        # Track styles not found in master and styles to add
        missing_styles = []
        
        # Create StyleCode column
        stylecodes = []
        for _, row in df.iterrows():
            stylecode, found_in_master = get_stylecode(row, ref_lookup)
            stylecodes.append(stylecode)
            
            if not found_in_master:
                style_r = str(row['StyleCodeR']).strip()
                metal = str(row['Metal']).strip().upper()
                missing_styles.append((style_r, metal, stylecode))
        
        df['StyleCode'] = stylecodes
        
        # Auto-add missing styles to master
        for style_r, metal, stylecode in missing_styles:
            _add_style_to_master(style_r, metal, stylecode)
        
        # Create SpecialRemarks column
        df['SpecialRemarks'] = df.apply(get_special_remarks, axis=1)
        
        # Create other placeholder columns
        df['OrderItemPcs'] = 1
        df['ItemRefNo'] = ''
        df['StockType'] = ''
        df['MakeType'] = ''
        df['OrderGroup'] = ''
        df['Certificate'] = ''
        df['SKUNo'] = ''
        df['Basestoneminwt'] = np.nan
        df['Basestonemaxwt'] = np.nan
        df['Basemetalminwt'] = np.nan
        df['Basemetalmaxwt'] = np.nan
        df['Expecteddeliverydate'] = ''
        
        # Select and reorder final columns
        final_columns = [
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
            'Expecteddeliverydate'
        ]
        
        for col in final_columns:
            if col not in df.columns:
                if col in ['Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 'Basemetalmaxwt']:
                    df[col] = np.nan
                else:
                    df[col] = ''
        
        df_final = df[final_columns]
        
        # Generate output filename
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"AMP_{base_name}_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        # Save to Excel
        df_final.to_excel(output_path, index=False)
        
        return True, output_path, None, df_final
        
    except Exception as e:
        return False, None, f'Error processing AMP file: {str(e)}', None
