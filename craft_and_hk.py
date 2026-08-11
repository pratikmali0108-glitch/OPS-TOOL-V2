import os
import re
import tempfile
import pandas as pd
import pdfplumber
from PyPDF2 import PdfReader, PdfWriter


def rotate_pdf_left(input_path):
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
    rotated_pdf = rotate_pdf_left(pdf_path)
    full_text = ""
    with pdfplumber.open(rotated_pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
    return full_text.strip()


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


def parse_purchase_order_data(full_text, size_prefix: str = "US", default_priority: str = "REG"):
    if not full_text.strip():
        return pd.DataFrame()

    po_match = re.search(r'PO\s*#\s*[:]*\s*(\d+)', full_text)
    item_po_no = po_match.group(1) if po_match else ""

    item_blocks = re.findall(
        r'(\d+)\.\s*'                          # Sr No.
        r'(\d+/\d+)\s+'                        # Order Code
        r'([\d.]+)\s+'                          # Order Qty
        r'(\S+)\s+'                             # Style Code
        r'(\S+)\s+'                             # Vendor Style
        r'(\S+)\s+'                             # SKU No
        r'(18K[T]?|14K[T]?)\s+'                  # Metal KT
        r'([YW])'                                 # Tone
        r'(?:\s+([\d.]+))?'                     # Optional Item Size
        r'[\s\S]*?Stamping Instructions:\s*([^\n]+)',  # Stamping instructions
        full_text
    )

    if not item_blocks:
        return pd.DataFrame()

    data = []
    for i, block in enumerate(item_blocks, start=1):
        (
            sr_no, order_code, order_qty, style_code,
            vendor_style, sku_no, metal_kt, tone,
            item_size, stamping_instr
        ) = block

        priority = default_priority
        formatted_size = clean_item_size(item_size or "", size_prefix)

        metal = f"G{metal_kt.replace('KT', '').replace('K', '')}{tone}"
        tone_full = "Yellow Gold" if tone == "Y" else "White Gold"
        desc_match = re.search(r'(BRACELET|EARRING|RING)', full_text[full_text.find(style_code):], re.IGNORECASE)
        desc = desc_match.group(1).capitalize() if desc_match else "Item"
        desc_full = f"{metal_kt} {tone_full} {desc} 1.00 CTW"

        if formatted_size:
            special_remarks = (
                f"BRILLIANT EARTH CRAFT,{order_code}, {style_code},{vendor_style}, "
                f"{sku_no},SZ-{formatted_size}, {metal_kt} {tone_full.upper()},COC CERTIFIED RE-CYCLE GOLD"
            )
        else:
            special_remarks = (
                f"BRILLIANT EARTH CRAFT,{order_code}, {style_code},{vendor_style}, "
                f"{sku_no}, {metal_kt} {tone_full.upper()},COC CERTIFIED RE-CYCLE GOLD"
            )

        design_prod_instr = "White Rodium" if tone == "W" else "No Rodium"

        data.append({
            "SrNO": i,
            "StyleCode": style_code,
            "ItemSize": formatted_size,
            "OrderQty": order_qty,
            "OrderItemPcs": 1,
            "Metal": metal,
            "Tone": tone,
            "ItemPoNo": item_po_no,
            "ItemRefNo": "",
            "StockType": "",
            "Priority": priority,
            "MakeType": "",
            "CustomerProductionInstruction": desc_full,
            "SpecialRemarks": special_remarks,
            "DesignProductionInstruction": design_prod_instr,
            "StampInstruction": stamping_instr.strip(),
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

    return pd.DataFrame(data)


def process_craft_hk_file(input_path: str, output_dir: str, size_prefix: str = "US", default_priority: str = "REG"):
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()

        if ext == '.pdf':
            full_text = extract_raw_text_from_pdf(input_path)
            df = parse_purchase_order_data(full_text, size_prefix=size_prefix or "US", default_priority=(default_priority or "REG").upper())
            if df is None or df.empty:
                return False, None, "No structured data extracted from PDF", None
            output_path = os.path.join(output_dir, f"{base_name}_CRAFT_HK_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        elif ext in ['.xlsx', '.xls', '.csv']:
            df = pd.read_excel(input_path) if ext in ['.xlsx', '.xls'] else pd.read_csv(input_path)
            output_path = os.path.join(output_dir, f"{base_name}_CRAFT_HK_PASSTHROUGH.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None


