import os
import re
import tempfile
import pandas as pd
import pdfplumber
from PyPDF2 import PdfReader, PdfWriter

# Global variable to store CS file data
CS_FILE_PATH = r"C:\Users\Omkumbhar\Desktop\OPS_Tool_03082026\CS_220526\CLG_CS_030826.xlsx"

_cs_data_cache = None


def rotate_pdf_left(input_path):
    """
    Rotate all pages in the PDF 90° counterclockwise (left)
    and return the path to the temporary rotated file.
    """
    reader = PdfReader(input_path)
    writer = PdfWriter()

    for page in reader.pages:
        page.rotate(90)
        writer.add_page(page)

    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    with open(temp_file.name, "wb") as f_out:
        writer.write(f_out)

    return temp_file.name


def extract_raw_text_from_pdf(pdf_path):
    """
    Rotates the PDF, extracts text using pdfplumber, and returns raw text.
    """
    try:
        rotated_pdf = rotate_pdf_left(pdf_path)
        full_text = ""
        
        # Try different extraction methods
        with pdfplumber.open(rotated_pdf) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # Method 1: Standard text extraction
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
                
                # Method 2: Try table extraction if standard text fails
                if not text or len(text.strip()) < 50:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            for row in table:
                                if row:
                                    full_text += " ".join([cell or "" for cell in row]) + "\n"
        
        # Clean up the text
        full_text = re.sub(r'\s+', ' ', full_text)  # Normalize whitespace

        print("\n📜 Extracted Text from Rotated PDF:")
        print("=" * 60)
        print(full_text[:1000])  # Show first 1000 chars
        print("=" * 60)

        return full_text.strip()

    except Exception as e:
        print(f"Error reading or processing PDF: {e}")
        return ""


_TONE_SUFFIX_MAP = {
    "W": "-INWG",
    "Y": "-INYG",
    "P": "-INPG",
    "PT": "-INPT",
    "AG": "-INAG",
}


def load_cs_data():
    """Load and cache CS file data."""
    global _cs_data_cache
    if _cs_data_cache is None:
        try:
            if os.path.exists(CS_FILE_PATH):
                _cs_data_cache = pd.read_excel(CS_FILE_PATH)
                print(f"✅ Loaded CS file with {len(_cs_data_cache)} records")
            else:
                print(f"⚠️ CS file not found at: {CS_FILE_PATH}")
                _cs_data_cache = pd.DataFrame()
        except Exception as e:
            print(f"⚠️ Error loading CS file: {e}")
            _cs_data_cache = pd.DataFrame()
    return _cs_data_cache


def lookup_itemsize_from_cs(style_code):
    """
    Look up ItemSize from CS file based on StyleCode.
    
    Args:
        style_code: The StyleCode to look up (e.g., "AAB04477H-7INWG")
    
    Returns:
        ItemSize string from CS file, or empty string if not found
    """
    cs_data = load_cs_data()
    
    if cs_data.empty:
        return ""
    
    # Look for exact match in "Client Style No" column
    matching_rows = cs_data[cs_data['Client Style No'] == style_code]
    
    if not matching_rows.empty:
        item_size = matching_rows.iloc[0]['ItemSize']
        # Clean and format the ItemSize
        if pd.notna(item_size):
            item_size_str = str(item_size).strip()
            print(f"📋 Found ItemSize '{item_size_str}' for StyleCode '{style_code}'")
            return item_size_str
    
    print(f"⚠️ No ItemSize found in CS file for StyleCode '{style_code}'")
    return ""

def remap_stylecode(style_code, tone):
    """Append tone-based suffix to every StyleCode."""
    sc = str(style_code).strip() if style_code else ""
    t  = str(tone).strip().upper() if tone else ""
    suffix = _TONE_SUFFIX_MAP.get(t)
    if sc and suffix:
        return f"{sc}{suffix}"
    return style_code


def clean_item_size(size_raw, prefix):
    size_raw = size_raw.strip()
    if not size_raw:
        return ""
    if re.match(r'^\d+\.\d+$', size_raw):
        size_clean = str(float(size_raw)).rstrip('0').rstrip('.')
    else:
        size_clean = size_raw
    if re.match(r'^\d+$', size_clean) and len(size_clean) == 1:
        size_clean = f"0{size_clean}"
    return f"{prefix}{size_clean}" if prefix else size_clean


def extract_item_size_from_sku(sku_no):
    """
    Derive ItemSize from CRAFT SKUNo.

    Examples:
      BE3DCLW150LC-18KW   -> 5   (no TB segment: BE3D -> 3 + 2)
      BE5D3TB170LC-18KW   -> 7   (TB170 -> last two digits / 10)
      BE5D4TB265LC-PT     -> 6.5 (TB265 -> 65 / 10)
      BE5D4TB1070LC-PT    -> 7   (TB1070 -> last two digits / 10)
    """
    sku = str(sku_no or "").strip().upper().split("-")[0]

    # Match TB followed by 3 or 4 digits to handle both TB265 and TB1070 formats
    tb_match = re.search(r"TB(\d{3,4})", sku)
    if tb_match:
        size_val = int(tb_match.group(1)[-2:]) / 10
        if float(size_val).is_integer():
            return str(int(size_val))
        return f"{size_val:.1f}".rstrip("0").rstrip(".")

    match = re.search(r"BE(\d)D(?:(\d))?", sku)
    if not match:
        return ""

    whole = int(match.group(1))
    frac_digit = match.group(2)
    if frac_digit is None:
        size_val = whole + 2
    else:
        size_val = whole + (int(frac_digit) - 1) / 2

    if float(size_val).is_integer():
        return str(int(size_val))
    return f"{size_val:.1f}".rstrip("0").rstrip(".")


def _size_numeric_for_style(size_raw):
    """Return plain numeric size string for StyleCode suffixes."""
    size_raw = str(size_raw or "").strip()
    if not size_raw:
        return ""
    numeric_match = re.search(r"(\d+(?:\.\d+)?)", size_raw)
    if not numeric_match:
        return ""
    size_numeric = numeric_match.group(1)
    if "." in size_numeric:
        return size_numeric.rstrip("0").rstrip(".")
    return size_numeric


def clean_stamping_instruction(stamp_instr):
    """
    Clean stamping instruction by removing "Page X of Y" text and other unwanted content.
    """
    if not stamp_instr:
        return ""
    
    # Remove "Page X of Y" pattern
    cleaned = re.sub(r'\s*Page\s+\d+\s+of\s+\d+\s*', '', stamp_instr, flags=re.IGNORECASE)
    
    # Remove extra whitespace
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    
    return cleaned


def parse_purchase_order_data(full_text, size_prefix: str = "US", default_priority: str = "REG"):
    """
    Parses structured data from the extracted text with multiple fallback patterns.
    """
    if not full_text.strip():
        print("⚠️ Empty text provided.")
        return pd.DataFrame()

    # Debug: Print more comprehensive information
    print(f"DEBUG: Total text length: {len(full_text)}")
    print(f"DEBUG: First 800 chars of extracted text: {full_text[:800]}")
    print(f"DEBUG: Text contains 'HKD': {'HKD' in full_text}")
    print(f"DEBUG: Text contains numbers: {bool(re.search(r'/d+', full_text))}")
    
    # More flexible PO number extraction - try multiple patterns
    po_patterns = [
        r'PO\s*#\s*[:]*\s*([0-9]+)',  # Match "PO # : 804243" format specifically
        r'HKD\s*#\s*([0-9\-]+)',      # Match "HKD # 804238-1795747" format
        r'(?:PO|P\.O\.)\s*[#:]+\s*([A-Z0-9\-]{4,})',  # Generic PO with at least 4 chars
        r'Order\s*[#:]+\s*([A-Z0-9\-]{4,})'  # Order number with at least 4 chars
    ]
    
    item_po_no = ""
    for i, pattern in enumerate(po_patterns):
        po_match = re.search(pattern, full_text, re.IGNORECASE)
        if po_match:
            item_po_no = po_match.group(1)
            print(f"DEBUG: PO pattern {i+1} matched: '{po_match.group(0)}' -> PO: '{item_po_no}'")
            break
        else:
            print(f"DEBUG: PO pattern {i+1} no match: {pattern}")
    
    print(f"DEBUG: Final PO: {item_po_no}")

    # Try multiple regex patterns to handle different formats
    item_blocks = []
    _stamp_instr_pat = (
        r'[\s\S]*?Stamping Instructions:\s*'
        r'(.*?)'
        r'(?=\s+\d+\.\s+\d+/\d+\s+|\s+Page\s+\d+\s+of\s+\d+|\s*$)'
    )

    # Pattern 1: Original strict pattern
    print("DEBUG: Trying original strict pattern...")
    item_blocks = re.findall(
        r'(\d+)\.\s*'                          # Sr No.
        r'(\d+/\d+)\s+'                        # Order Code
        r'([\d.]+)\s+'                          # Order Qty
        r'(\S+)\s+'                             # Vendor Style # (internal style code)
        r'(\S+)\s+'                             # Style # (customer reference)
        r'(\S+)\s+'                             # SKU No
        r'(18K[T]?|14K[T]?|PLAT|P)\s+'             # Metal KT (including PLAT and P)
        r'([YWPRT]?)'                            # Tone (optional, can be Y, W, P, R, T, etc.)
        r'(?:\s+([\d.]+))?'                     # Optional Item Size
        + _stamp_instr_pat,
        full_text
    )
    
    # Pattern 2: More flexible pattern
    if not item_blocks:
        print("DEBUG: Trying more flexible pattern...")
        item_blocks = re.findall(
            r'(\d+)\.?\s*'                         # Sr No (optional dot)
            r'([A-Z0-9/\-]+)\s+'                   # Order Code (more flexible)
            r'([\d.]+)\s+'                         # Order Qty
            r'([A-Z0-9\-]+)\s+'                    # Style Code
            r'([A-Z0-9\-]+)\s+'                    # Vendor Style
            r'([A-Z0-9\-]+)\s+'                    # SKU No
            r'(18K[T]?|14K[T]?|PLAT|PT|P)\s*'        # Metal KT (including PT and P)
            r'([YWPRT]?)\s*'                       # Tone
            r'(?:([\d.]+)\s*)?'                    # Optional Item Size
            + _stamp_instr_pat.replace("Stamping Instructions:", "(?:Stamp|Instructions?)[:]*"),
            full_text,
            re.IGNORECASE | re.DOTALL
        )
    
    # Pattern 3: Very flexible - split by lines and look for patterns
    if not item_blocks:
        print("DEBUG: Trying line-by-line pattern...")
        lines = [line.strip() for line in full_text.split('\n') if line.strip()]
        
        for i, line in enumerate(lines):
            # Look for lines that might contain item data
            if re.search(r'\d+.*?(18K|14K|PLAT|P)', line, re.IGNORECASE):
                print(f"DEBUG: Found potential item line: {line}")
                
                # Try to extract data from this line
                parts = re.split(r'\s+', line)
                if len(parts) >= 7:  # Minimum expected parts
                    # Look for stamping instructions in next few lines
                    stamp_instr = ""
                    for j in range(i+1, min(i+5, len(lines))):
                        if re.search(r'stamp|instruction', lines[j], re.IGNORECASE):
                            stamp_instr = lines[j]
                            break
                    
                    # Try to construct a match
                    try:
                        sr_no = parts[0].rstrip('.')
                        if sr_no.isdigit():
                            item_blocks.append((
                                sr_no,
                                parts[1] if len(parts) > 1 else "",
                                parts[2] if len(parts) > 2 else "1",
                                parts[3] if len(parts) > 3 else "",
                                parts[4] if len(parts) > 4 else "",
                                parts[5] if len(parts) > 5 else "",
                                parts[6] if len(parts) > 6 else "P",
                                parts[7] if len(parts) > 7 else "Y",
                                parts[8] if len(parts) > 8 else "",
                                clean_stamping_instruction(stamp_instr)
                            ))
                    except (IndexError, ValueError):
                        continue
    
    # Pattern 4: Most flexible - look for any numeric sequences with metal info
    if not item_blocks:
        print("DEBUG: Trying most flexible pattern...")
        # Find all potential item data by looking for metal keywords
        metal_matches = re.finditer(r'(18K|14K|PLAT|P)', full_text, re.IGNORECASE)
        
        for match in metal_matches:
            start = max(0, match.start() - 200)  # Look 200 chars before
            end = min(len(full_text), match.end() + 200)  # Look 200 chars after
            context = full_text[start:end]
            
            # Try to find structured data in this context
            numbers = re.findall(r'\d+(?:\.\d+)?', context)
            words = re.findall(r'[A-Z0-9\-]+', context)
            
            if len(numbers) >= 3 and len(words) >= 5:
                item_blocks.append((
                    numbers[0],  # Sr No
                    f"{numbers[1]}/{numbers[2]}" if len(numbers) > 2 else words[0],  # Order Code
                    numbers[1] if len(numbers) > 1 else "1",  # Order Qty
                    words[0] if words else "",  # Style Code
                    words[1] if len(words) > 1 else "",  # Vendor Style
                    words[2] if len(words) > 2 else "",  # SKU No
                    match.group(1),  # Metal KT
                    "Y",  # Default tone
                    numbers[3] if len(numbers) > 3 else "",  # Item Size
                    ""  # Stamping instructions
                ))
                break  # Take first match
    
    
    print(f"DEBUG: Found {len(item_blocks)} item blocks")
    if item_blocks:
        print(f"DEBUG: First item block: {item_blocks[0]}")

    if not item_blocks:
        print("⚠️ No item blocks found.")
        return pd.DataFrame()

    data = []
    for i, block in enumerate(item_blocks, start=1):
        (
            sr_no, order_code, order_qty, style_code,
            vendor_style, sku_no, metal_kt, tone,
            item_size, stamping_instr
        ) = block[:10] if len(block) >= 10 else block + ("",) * (10 - len(block))

        # Use default priority from parameter
        priority = default_priority

        # Item size comes from SKUNo; fall back to PDF size when SKU has no pattern
        sku_size = extract_item_size_from_sku(sku_no)
        if sku_size:
            item_size = sku_size

        # Clean & format size
        formatted_size = clean_item_size(item_size or "", size_prefix)

        # Handle different metal types and tones
        metal_kt = metal_kt.upper() if metal_kt else ""
        tone = tone.upper() if tone else "Y"  # Default to Y if empty
        
        if metal_kt in ["PLAT", "PT", "PLATINUM"]:
            metal = "PC95"
            tone = "PT"
            tone_full = "Platinum"
        elif metal_kt == "P":
            # Handle standalone "P" as 14K Pink Gold
            metal = "G14P"
            tone = "P"
            tone_full = "Pink Gold"
        else:
            # Extract numeric part from metal (14K, 18K, etc.)
            metal_num = re.sub(r'[^0-9]', '', metal_kt) or "14"  # Default to 14 if no number found
            metal = f"G{metal_num}{tone}"
            
            # Map tone to full description
            tone_map = {
                "Y": "Yellow Gold",
                "W": "White Gold", 
                "P": "Pink Gold",
                "R": "Rose Gold",
                "T": "Two Tone"
            }
            tone_full = tone_map.get(tone, "Yellow Gold")

        # Save original style code for text searching
        original_style_code = str(style_code).strip() if style_code else ""

        # Find item type first using original style code
        search_area = full_text[full_text.find(original_style_code):] if original_style_code else full_text
        desc_match = re.search(r'(BRACELET|EARRING|RING)', search_area, re.IGNORECASE)
        desc = desc_match.group(1).capitalize() if desc_match else "Item"
        is_bracelet = desc_match and desc_match.group(1).upper() == "BRACELET"
        is_earring = desc_match and desc_match.group(1).upper() == "EARRING"

        size_numeric = _size_numeric_for_style(item_size)

        # Get tone suffix first
        tone_suffix = ""
        t = str(tone).strip().upper() if tone else ""
        if t in _TONE_SUFFIX_MAP:
            tone_suffix = _TONE_SUFFIX_MAP[t]

        # Format style code with SKU-derived size appended when available
        if is_earring and original_style_code and tone_suffix:
            # FOR EARRINGS: Format strictly as Combination-WG (No numeric size indicator)
            clean_tone_suffix = tone_suffix.lstrip('-').replace('IN', '', 1)
            style_code = f"{original_style_code}-{clean_tone_suffix}"
        elif original_style_code and size_numeric and tone_suffix:
            if is_bracelet:
                style_code = f"{original_style_code}-{size_numeric}{tone_suffix.lstrip('-')}"
            else:
                clean_tone_suffix = tone_suffix.lstrip('-').replace('IN', '', 1)
                style_code = f"{original_style_code}-{size_numeric}{clean_tone_suffix}"
        elif is_bracelet and original_style_code and tone_suffix:
            style_code = remap_stylecode(original_style_code, tone)
        elif original_style_code and tone_suffix:
            clean_tone_suffix = tone_suffix.lstrip('-').replace('IN', '', 1)
            style_code = f"{original_style_code}-{clean_tone_suffix}"
        else:
            style_code = remap_stylecode(original_style_code, tone)
        
        # NOW LOOKUP ItemSize from CS file based on the formatted StyleCode
        cs_item_size = lookup_itemsize_from_cs(style_code)
        if cs_item_size:
            # Use the ItemSize from CS file
            formatted_size = cs_item_size
        # If no CS data found, keep the existing formatted_size from SKU extraction
        
        # Extract CTW value from the text around this item (use original style code)
        item_start = full_text.find(original_style_code) if original_style_code else -1
        if item_start != -1:
            # Search in a reasonable range around the item (next 500 characters)
            item_section = full_text[item_start:item_start + 500]
            ctw_match = re.search(r'(\d+\.?\d*)\s*CTW', item_section, re.IGNORECASE)
            ctw_value = ctw_match.group(1) if ctw_match else "1.00"
        else:
            ctw_value = "1.00"
        
        # Format description based on metal type
        if metal_kt.upper() == "PLAT":
            desc_full = f"PLATINUM {desc} {ctw_value} CTW"
        else:
            desc_full = f"{metal_kt} {tone_full} {desc} {ctw_value} CTW"

        if formatted_size and not is_earring:
            special_remarks = (
                f"BRILLIANT EARTH CRAFT,{order_code},{vendor_style}, "
                f"{sku_no},SZ-{formatted_size}, {metal_kt} {tone_full.upper()},COC CERTIFIED RE-CYCLE GOLD"
            )
        else:
            special_remarks = (
                f"BRILLIANT EARTH CRAFT,{order_code}, {vendor_style}, "
                f"{sku_no}, {metal_kt} {tone_full.upper()},COC CERTIFIED RE-CYCLE GOLD"
            )

        design_prod_instr = "White Rhodium" if tone == "W" else "No Rhodium"
        item_ref_no = str(vendor_style).strip() if vendor_style else ""

        data.append({
            "SrNO": i,
            "StyleCode": style_code,
            "ItemSize": formatted_size,
            "OrderQty": order_qty,
            "OrderItemPcs": 1,
            "Metal": metal,
            "Tone": tone,
            "ItemPoNo": item_po_no,
            "ItemRefNo": item_ref_no,
            "StockType": "",
            "Priority": priority,
            "MakeType": "",
            "CustomerProductionInstruction": desc_full,
            "SpecialRemarks": special_remarks,
            "DesignProductionInstruction": design_prod_instr,
            "StampInstruction": clean_stamping_instruction(stamping_instr),
            "OrderGroup": "BRILLIANT EARTH CRAFT",
            "Certificate": "",
            "SKUNo": sku_no,
            "Basestoneminwt": "",
            "Basestonemaxwt": "",
            "Basemetalminwt": "",
            "Basemetalmaxwt": "",
            "Productiondeliverydate": "",
            "Expecteddeliverydate": "",
            "SetPrice": "",
            "StoneQuality": ""
        })

    df = pd.DataFrame(data)

    # Global cleanup to keep ItemSize completely blank for Earring rows
    if not df.empty:
        df.loc[df['CustomerProductionInstruction'].str.contains('Earring', case=False, na=False), 'ItemSize'] = ""

    print(f"\n✅ {len(df)} items successfully parsed.")
    return df


def process_craft_file(input_path: str, output_dir: str, size_prefix: str = "US", default_priority: str = "REG"):
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()

        if ext == '.pdf':
            full_text = extract_raw_text_from_pdf(input_path)
            df = parse_purchase_order_data(full_text, size_prefix=size_prefix or "US", default_priority=(default_priority or "REG").upper())
            if df is None or df.empty:
                return False, None, "No structured data extracted from PDF", None
            df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''
            output_path = os.path.join(output_dir, f"{base_name}_CRAFT_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        elif ext in ['.xlsx', '.xls', '.csv']:
            df = pd.read_excel(input_path) if ext in ['.xlsx', '.xls'] else pd.read_csv(input_path)
            output_path = os.path.join(output_dir, f"{base_name}_CRAFT_PASSTHROUGH.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None


# ================== MAIN EXECUTION ==================
if __name__ == "__main__":
    # For testing - you can change this path to test specific PDFs
    pdf_file_path = r"C:\Users\Admin\Desktop\full stack Order Processing Tool_Testing\HKD#804243-1795833-Shimayra.pdf"
    print(f"\n📂 Reading and processing PDF: {pdf_file_path}")

    full_text = extract_raw_text_from_pdf(pdf_file_path)

    if full_text:
        df = parse_purchase_order_data(full_text)

        if not df.empty:
            print("\n✅ Final Structured Data:")
            print("=" * 80)
            print(df)
            print("=" * 80)
            output_path = "final_OUTPUT_CRAFT.xlsx"
            df.to_excel(output_path, index=False)
            print(f"\n💾 Data successfully saved to '{output_path}'")
        else:
            print("⚠️ No structured data could be extracted.")
    else:
        print("❌ No text extracted from the PDF.")