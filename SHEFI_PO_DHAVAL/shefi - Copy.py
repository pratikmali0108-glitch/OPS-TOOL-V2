"""
SHEFI NEW PO - Batch PDF to Excel Converter
Processes every *.pdf found in the script's directory and writes
one Excel file per PDF into the Output/ sub-folder.
"""

import os
import re
import glob
import pdfplumber
import pandas as pd
from datetime import datetime


# ── Helpers ───────────────────────────────────────────────────────────────────

def format_date(date_str: str) -> str:
    """Convert M/D/YYYY  →  DD-Mon-YY  (e.g. 3/27/2026 → 27-Mar-26)."""
    try:
        return datetime.strptime(date_str.strip(), "%m/%d/%Y").strftime("%d-%b-%y")
    except Exception:
        return date_str.strip()


# ── Core extraction ───────────────────────────────────────────────────────────

def extract_shefi_po(pdf_path: str) -> list[dict]:
    """
    Extract all PO line items page-by-page so every item carries its own
    correct page-header values.  Returns a list of row dicts.
    """
    all_rows: list[dict] = []
    item_counter = 0

    # Item data line: row-number  CODE1  [rest…]
    # CODE1 can start with a digit (e.g. 709525D) or an uppercase letter (LGD244910E)
    ITEM_PAT = re.compile(r'^(\d+)\s+([A-Z0-9][A-Z0-9]+)\s+(.*)', re.DOTALL)

    # Category / description header line.
    # Accepts an optional leading "CODE[/NUM] [/ ]" prefix (PDF rendering artefact).
    # The prefix MUST contain at least one digit so that letter-only words like
    # "LGD" are not accidentally split off from the category name.
    # Then: "CategoryWord(s): optional_metal_desc"
    CAT_PAT = re.compile(
        r'^(?:([A-Z0-9]*\d[A-Z0-9]*(?:/\d+)?)\s+(?:/\s+)?)?([A-Za-z][A-Za-z0-9 ]+):\s*(.*)'
    )

    # First word of a matched category name that signals it is NOT a real category
    # (i.e. it is a PO-header field, column header, or footer element)
    CAT_SKIP = {
        "Order", "Page", "Vendor", "Ship", "Grand", "RightClick", "Phone",
        "Due", "Cancel", "Date", "Reference", "Fax", "Purchase", "Right",
        "Copyright", "Memo", "Job", "Bag", "Weight", "Unit", "Item",
        "Description", "Size", "Quantity", "Amount", "Cost",
    }

    # Lines to unconditionally ignore everywhere in the item area
    SKIP_MARKERS = ("Grand Total", "RightClick", "Copyright", "20180426")

    # ── inner helpers ─────────────────────────────────────────────────────────

    def is_vendor_code(word: str) -> bool:
        """
        Return True when a token looks like an alphanumeric PO code
        (Item # or Vendor Item #) rather than the start of a description.
        """
        if not re.search(r'\d', word):
            return False            # pure letters → description word (WAY, INSIDE …)
        if re.match(r'^\d+[KT]', word, re.IGNORECASE):
            return False            # metal type: 14KW, 14KY, 10KR, 14TT …
        if re.match(r'^[A-Z][a-z]', word):
            return False            # mixed-case desc word: Set, Heart, Dragonfly …
        return True

    def parse_item_rest(rest: str):
        """
        Given the portion of an item data line *after* the first code (Item #),
        return (vendor_code, inline_desc, size, qty).

        Handles four size formats (tried in order):
          1. mm size   – …desc  9.70mm  QTY  0.0000…   (pendant/earring diameter)
          2. decimal   – …desc  6.5     QTY  0.0000…   (ring size with decimal)
          3. integer   – …desc  6       QTY  0.0000…   (ring size as whole number)
          4. no size   – …desc          QTY  0.0000…
        """
        tokens = rest.split()
        vendor_code = ""
        start = 0

        if tokens and is_vendor_code(tokens[0]):
            vendor_code = tokens[0]
            start = 1

        tail = " ".join(tokens[start:])

        # 1. mm size (e.g. 9.70mm, 4.30mm, 6mm): …desc SIZE_mm QTY 0.0000…
        m = re.search(r'^(.*?)\s+(\d+(?:\.\d+)?mm)\s+(\d+)\s+0\.0000', tail)
        if m:
            return vendor_code, m.group(1).strip(), m.group(2), m.group(3)

        # 2. Decimal size without mm (e.g. 6.5): …desc SIZE QTY 0.0000…
        m = re.search(r'^(.*?)\s+(\d+\.\d+)\s+(\d+)\s+0\.0000', tail)
        if m:
            return vendor_code, m.group(1).strip(), m.group(2), m.group(3)

        # 3. Integer size (e.g. 6): two consecutive integers before 0.0000
        #    Note: "N 0.0000" alone never matches here because (\d+)\s+0\.0000
        #    would grab "0" from "0.0000" and then fail on the dot.
        m = re.search(r'^(.*?)\s+(\d+)\s+(\d+)\s+0\.0000', tail)
        if m:
            return vendor_code, m.group(1).strip(), m.group(2), m.group(3)

        # 4. No size — only qty before 0.0000
        m = re.search(r'^(.*?)\s+(\d+)\s+0\.0000', tail)
        if m:
            return vendor_code, m.group(1).strip(), "", m.group(2)

        return vendor_code, tail.strip(), "", ""

    # ── per-PDF state (shared via closure) ───────────────────────────────────
    current_item: dict | None = None

    def flush_item():
        nonlocal current_item
        if current_item:
            extras = current_item.pop("_extra", [])
            if extras:
                current_item["Description"] += " " + " ".join(extras)
            all_rows.append(current_item)
            current_item = None

    # ── iterate pages ─────────────────────────────────────────────────────────
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            lines = [ln.strip() for ln in page_text.split("\n")]

            # ── Per-page header fields ─────────────────────────────────────
            hdr = {
                "Customer":    "SHEFI DIAMONDS, INC",
                "Order#":      "",
                "Page#":       "",
                "PO#":         "",
                "Date":        "",
                "Due Date":    "",
                "Cancel Date": "",
                "Ref":         "",
                "Vendor#":     "",
                "Ship Via":    "",
            }

            for line in lines:
                m = re.search(r"Order #:\s*(\d+)", line)
                if m:
                    hdr["Order#"] = m.group(1)

                m = re.search(r"Page #:\s*(\d+\s+of\s+\d+)", line)
                if m:
                    hdr["Page#"] = m.group(1)

                m = re.search(r"P\.O\. #:\s*(\S+)", line)
                if m:
                    hdr["PO#"] = m.group(1)

                m = re.search(
                    r"Date:\s*(\d+/\d+/\d+)\s+Due Date:\s*(\d+/\d+/\d+)"
                    r"\s+Cancel Date:\s*(\d+/\d+/\d+)",
                    line,
                )
                if m:
                    hdr["Date"]        = format_date(m.group(1))
                    hdr["Due Date"]    = format_date(m.group(2))
                    hdr["Cancel Date"] = format_date(m.group(3))

                m = re.search(r"Reference:\s*(.*?)\s+Vendor\s*#:", line)
                if m:
                    hdr["Ref"] = m.group(1).strip()

                m = re.search(r"Vendor #:(\S+)", line)
                if m:
                    hdr["Vendor#"] = m.group(1)

                m = re.search(r"Ship Via:\s*(.+)", line)
                if m:
                    hdr["Ship Via"] = m.group(1).strip()

            # ── Find the start of the item table (after column header row) ─
            item_start = 0
            for i, line in enumerate(lines):
                if "Memo #" in line and "Item #" in line:
                    item_start = i + 1
                    break

            # ── Per-page item extraction ───────────────────────────────────
            current_item = None          # reset at page boundary
            current_cat: str = ""
            current_metal: str = ""
            cat_vendor_prefix: str = ""

            for line in lines[item_start:]:
                if not line:
                    continue

                # Skip weight-total summary rows ("5 0.0000") and footers
                if re.match(r'^\d+\s+\d+\.\d{4}$', line):
                    continue
                if any(kw in line for kw in SKIP_MARKERS):
                    continue

                # ── Item data line ─────────────────────────────────────────
                m_item = ITEM_PAT.match(line)
                if m_item:
                    flush_item()
                    item_counter += 1
                    first_code = m_item.group(2)
                    vendor_code, line_desc, size_val, qty_val = parse_item_rest(
                        m_item.group(3)
                    )

                    # Merge category-line prefix with item-line vendor code.
                    # e.g. prefix="203105P", vendor_code="P31719" → "203105P / P31719"
                    if cat_vendor_prefix and vendor_code:
                        vendor_code = f"{cat_vendor_prefix} / {vendor_code}"
                    elif cat_vendor_prefix:
                        vendor_code = cat_vendor_prefix

                    desc = " ".join(
                        p for p in (current_cat, current_metal, line_desc) if p
                    )

                    current_item = {
                        **hdr,
                        "#":             item_counter,
                        "Memo #":        "",
                        "Item #":        first_code,
                        "Vendor Item #": vendor_code,
                        "Description":   desc,
                        "Size":          size_val,
                        "Quantity":      qty_val,
                        "_extra":        [],
                    }
                    cat_vendor_prefix = ""
                    continue

                # ── Category / description header line ─────────────────────
                m_cat = CAT_PAT.match(line)
                if m_cat:
                    prefix   = (m_cat.group(1) or "").strip()
                    cat_name = m_cat.group(2).strip()
                    cat_rest = m_cat.group(3).strip()
                    first_word = cat_name.split()[0]
                    if first_word not in CAT_SKIP:
                        flush_item()
                        current_cat       = cat_name
                        current_metal     = cat_rest
                        cat_vendor_prefix = prefix
                        continue

                # ── Continuation / extra description ───────────────────────
                if current_item:
                    current_item["_extra"].append(line)

            # Save any item still open at end of page
            flush_item()

    return all_rows


# ── Excel writer ──────────────────────────────────────────────────────────────

COLUMNS = [
    "Customer", "Order#", "Page#", "PO#", "Date", "Due Date", "Cancel Date",
    "Ref", "Vendor#", "Ship Via", "#", "Memo #", "Item #", "Vendor Item #",
    "Description", "Size", "Quantity",
]


def save_to_excel(rows: list[dict], output_path: str) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMNS]
    df.to_excel(output_path, index=False)
    return df


# ── Flask-compatible single-file API ──────────────────────────────────────────

def process_shefi_new_file(input_path: str, output_dir: str):
    """
    Process a single SHEFI New PO PDF and write an Excel file.

    Parameters
    ----------
    input_path : str   Path to the uploaded PDF file.
    output_dir : str   Directory where the output .xlsx should be saved.

    Returns
    -------
    tuple: (success: bool, output_path: str|None, error: str|None, df: DataFrame|None)
    """
    try:
        rows = extract_shefi_po(input_path)
        if not rows:
            return False, None, "No line items could be extracted from the PDF.", None
        stem = os.path.splitext(os.path.basename(input_path))[0]
        out_file = os.path.join(output_dir, f"SHEFI_NEW_PO_{stem}.xlsx")
        df = save_to_excel(rows, out_file)
        return True, out_file, None, df
    except Exception as exc:
        return False, None, str(exc), None


# ── Batch runner ──────────────────────────────────────────────────────────────

def main() -> None:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "Output")
    os.makedirs(output_dir, exist_ok=True)

    pdf_files = sorted(glob.glob(os.path.join(script_dir, "*.pdf")))

    if not pdf_files:
        print("No PDF files found in:", script_dir)
        return

    print(f"Found {len(pdf_files)} PDF(s) to process.\n")

    for pdf_path in pdf_files:
        pdf_name   = os.path.splitext(os.path.basename(pdf_path))[0]
        excel_path = os.path.join(output_dir, f"{pdf_name}.xlsx")

        print(f"  Processing : {os.path.basename(pdf_path)}")
        try:
            rows = extract_shefi_po(pdf_path)
            df   = save_to_excel(rows, excel_path)
            print(f"  Saved      : Output/{pdf_name}.xlsx  ({len(df)} item(s) extracted)")
        except Exception as exc:
            print(f"  ERROR      : {exc}")
        print()

    print("Done.")


if __name__ == "__main__":
    main()
