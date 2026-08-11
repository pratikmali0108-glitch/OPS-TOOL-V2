import pandas as pd
import re
import pdfplumber
import os


def process_ju_file(pdf_path: str, output_dir: str, priority: str = "REG"):
    """
    Process JU PDF file and extract order data.
    
    Args:
        pdf_path: Path to the input PDF file
        output_dir: Directory where output file will be saved
        priority: Priority value (default: "REG")
    
    Returns:
        tuple: (success: bool, output_path: str, error: str, df: DataFrame)
    """
    try:
        # Read PDF and extract text
        full_text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"
        
        if not full_text.strip():
            return False, None, "No text could be extracted from the PDF", None
        
        # Process the extracted text
        data = []
        lines = full_text.splitlines()
        sr_no = 1
        current_item_po_no = ""

        for i, raw_line in enumerate(lines):
            line = raw_line.strip()
            if not line:
                continue

            # Capture the current ItemPoNo from the nearest preceding Order header
            order_match = re.search(r'Order\s+\d+/\s+\w+/\s+\w+/\s+(\d+)', line)
            if order_match:
                current_item_po_no = order_match.group(1)
                continue

            # Detect item row like: "1 ABP08249C 1.52 AG925 ..."
            item_match = re.match(r'^(\d+)\s+([A-Z0-9]+)\s+([\d.]+)\s+AG925', line)
            if not item_match:
                continue

            style_code = item_match.group(2)
            metal = "AG925"
            tone = "W"

            # Extract quantity - look for pattern "number space large_number" before dates
            qty_value_match = re.search(r'\s(\d{2,4})\s+[\d,]+\.[\d]{2}\s+\d{2}/\d{2}/\d{2}', line)
            order_qty = qty_value_match.group(1) if qty_value_match else ""

            # Initialize fields
            customer_instruction = ""
            special_remarks = ""
            design_instruction = ""
            stamp_instruction = ""
            sku_no = ""
            item_type = ""

            for j in range(i + 1, min(i + 20, len(lines))):
                nxt = lines[j].strip()
                if not nxt:
                    continue

                # Customer instruction - extract everything after "Cust.Inst"
                if re.search(r'Cust\.?\s*Inst', nxt, re.IGNORECASE):
                    cust_match = re.search(r'Cust\.?\s*Inst\.?\s+(.*?)(?=\s+\d+\s+Plt\s+Rate|$)', nxt, re.IGNORECASE)
                    if cust_match:
                        customer_instruction = cust_match.group(1).strip()
                        # Extract item type
                        item_type_match = re.search(r'Fashion\s+(Pendant|Bangle|necklace|ring|earring)s?', customer_instruction, re.IGNORECASE)
                        if item_type_match:
                            item_type = item_type_match.group(1).capitalize()

                # Design/Production Instruction - extract everything after "Prd Inst."
                if re.search(r'Prd\.?\s*Inst', nxt, re.IGNORECASE):
                    prd_match = re.search(r'Prd\.?\s*Inst\.?\s+(.*?)(?=\s+SS|$)', nxt, re.IGNORECASE)
                    if prd_match:
                        design_instruction = prd_match.group(1).strip()

                # Stamp Instruction - extract everything after "Stmp Inst"
                if re.search(r'Stmp\s*Inst', nxt, re.IGNORECASE):
                    stmp_match = re.search(r'Stmp\s*Inst\.?\s+(.*?)(?=\s+Bill\s+of|$)', nxt, re.IGNORECASE)
                    if stmp_match:
                        stamp_instruction = stmp_match.group(1).strip()

                # SKU - extract everything after "SKU#"
                if "SKU#" in nxt:
                    m = re.search(r'SKU#\s+([\w\s]+?)(?=\s+Prt\s+Cd|$)', nxt)
                    if m:
                        sku_no = m.group(1).strip()

                # Special Remarks - extract everything after "Sepcial Rem"
                if re.search(r'Sepcial\s*Rem', nxt, re.IGNORECASE):
                    special_rem_match = re.search(r'Sepcial\s*Rem\.?\s+(.*?)(?=\s+Prt\s+Cd|$)', nxt, re.IGNORECASE)
                    if special_rem_match:
                        special_remarks = special_rem_match.group(1).strip()

                # Stop if next item line begins
                if re.match(r'^\d+\s+[A-Z0-9]+', nxt):
                    break

            # Format SpecialRemarks based on item type
            if item_type:
                special_remarks = f"SKU # {metal}, FASHION {item_type}"

            # Format StampInstruction based on metal type
            if metal == "AG925":
                stamp_instruction = "925,JD LOGO"
            elif metal == "14KT":
                stamp_instruction = "14,JD LOGO"
            elif metal == "18KT":
                stamp_instruction = "18,JD LOGO"
            else:
                # For any other metal format, try to extract the number
                metal_num = re.search(r'(\d+)', metal)
                if metal_num:
                    stamp_instruction = f"{metal_num.group(1)},JD LOGO"
                else:
                    stamp_instruction = "JD LOGO"

            # Set DesignProductionInstruction to "All White" if Tone is "W"
            if tone == "W":
                design_instruction = "All White"

            data.append({
                'SrNo': sr_no,
                'StyleCode': style_code,
                'ItemSize': "",
                'OrderQty': order_qty,
                'OrderItemPcs': "",
                'Metal': metal,
                'Tone': tone,
                'ItemPoNo.': current_item_po_no,
                'ItemRefNo': "",
                'StockType': "",
                'Priority': priority,
                'MakeType': "",
                'CustomerProductionInstruction': customer_instruction,
                'SpecialRemarks': special_remarks,
                'DesignProductionInstruction': design_instruction,
                'StampInstruction': stamp_instruction,
                'OrderGroup': "STERLING JEWELERS(OUTLET)",
                'Certificate': "",
                'SKUNo': sku_no,
                'Basestoneminwt': "",
                'Basestonemaxwt': "",
                'Basemetalminwt': "",
                'Basemetalmaxwt': "",
                'Productiondeliverydate': "",
                'Expecteddeliverydate': "",
                'SetPrice': "",
                'StoneQuality': ""
            })
            sr_no += 1

        # Create DataFrame
        df = pd.DataFrame(data)
        
        if df.empty:
            return False, None, "No items were extracted from the PDF. Check if the PDF format matches expected pattern.", None
        
        # Generate output filename
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_filename = f"JU_Processed_{base_name}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        
        # Save to Excel
        df.to_excel(output_path, index=False)
        
        return True, output_path, None, df
    
    except Exception as e:
        return False, None, f"Error processing file: {str(e)}", None
