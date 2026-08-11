"""
shefinew.py - SHEFI New PO (PDF-based) Processor
==================================================
Exposes:
  • process_shefi_new_file(input_path, output_dir)  – Flask / single-file API
  • main()                                           – CLI bulk-folder runner

Column specification (17 columns):
  Customer | Order# | Page# | PO# | Date | Due Date | Cancel Date |
  Ref | Vendor# | Ship Via | # | Memo # | Item # | Vendor Item # |
  Description | Size | Quantity
"""

import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber

# ── Folder configuration (for standalone CLI use) ─────────────────────────────
_BASE_DIR      = Path(__file__).parent
_INPUT_FOLDER  = _BASE_DIR / "SHEFI_NEW_PO"
_OUTPUT_FOLDER = _INPUT_FOLDER / "Output"

# ── Column order ──────────────────────────────────────────────────────────────
COLUMNS = [
    "Customer", "Order#", "Page#", "PO#", "Date", "Due Date", "Cancel Date",
    "Ref", "Vendor#", "Ship Via", "#", "Memo #", "Item #", "Vendor Item #",
    "Description", "Size", "Quantity",
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _fmt_date(date_str: str) -> str:
    """Convert M/D/YYYY  →  DD-Mon-YY  (e.g. 3/27/2026 → 27-Mar-26)."""
    try:
        return datetime.strptime(date_str.strip(), "%m/%d/%Y").strftime("%d-%b-%y")
    except Exception:
        return date_str.strip()


# ── Core extraction ───────────────────────────────────────────────────────────

def extract_shefi_po(pdf_path) -> list:
    """
    Extract all PO rows from a single SHEFI New-PO PDF.

    Processes page-by-page so each row carries its own correct Page# value.
    Returns a list of dicts (one per line item).
    """
    all_rows: list = []
    item_counter = 0          # sequential across all pages

    CAT_RE  = re.compile(r"^[A-Za-z][A-Za-z0-9 ]+:$")
    ITEM_RE = re.compile(r"^(\d+)\s+([A-Z][A-Z0-9]+)\s+(.+)")
    DESC_KW = ("Rd.", "Ctw", "Lab", "Grown", "Diamond")
    HDR_KW  = ("Order", "Page", "P.O.", "Ref", "Ship", "Grand", "Right",
                "Copy", "Reserved")

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            lines = [ln.strip() for ln in page_text.split("\n")]

            # ── Per-page header ───────────────────────────────────────────────
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
                    hdr["Date"]        = _fmt_date(m.group(1))
                    hdr["Due Date"]    = _fmt_date(m.group(2))
                    hdr["Cancel Date"] = _fmt_date(m.group(3))

                m = re.search(r"Reference:\s*(\S+)", line)
                if m:
                    hdr["Ref"] = m.group(1)

                m = re.search(r"Vendor #:(\S+)", line)
                if m:
                    hdr["Vendor#"] = m.group(1)

                m = re.search(r"Ship Via:\s*(\S+)", line)
                if m:
                    hdr["Ship Via"] = m.group(1)

            # ── Per-page item extraction ──────────────────────────────────────
            current_item = None
            current_cat  = None
            expect_item  = False

            def _save(item):
                """Finalise description, strip temp key, push to all_rows."""
                if item:
                    extras = item.pop("_extra", [])
                    if extras:
                        item["Description"] += " " + " ".join(extras)
                    all_rows.append(item)

            for line in lines:

                # ── Category header (e.g. "LGD Anniversary:") ────────────────
                if CAT_RE.match(line) and not any(kw in line for kw in HDR_KW):
                    _save(current_item)
                    current_item = None
                    current_cat  = line.rstrip(":").strip()
                    expect_item  = True
                    continue

                # ── Item data line ────────────────────────────────────────────
                if expect_item:
                    m = ITEM_RE.match(line)
                    if m:
                        item_counter += 1
                        item_no = m.group(2)
                        rest    = m.group(3)
                        words   = rest.split()

                        # Vendor Item # ----------------------------------------
                        # First word of `rest` is a vendor code when it:
                        #   • does NOT start with "14K"
                        #   • starts with 2 uppercase letters
                        #   • is at least 8 chars (e.g. RG0004161E)
                        vendor_item = ""
                        w_off = 0
                        if (words
                                and not re.match(r"^14K", words[0])
                                and re.match(r"^[A-Z]{2}", words[0])
                                and len(words[0]) >= 8):
                            vendor_item = words[0]
                            w_off = 1

                        # Metal description (e.g. "14KW Shared Prong") ---------
                        metal_desc = (
                            " ".join(words[w_off: w_off + 3])
                            if len(words) >= w_off + 3 else ""
                        )

                        # Stone count on same line (integer before decimal size)
                        stone_on_line = ""
                        idx_pm = w_off + 3
                        if len(words) > idx_pm:
                            cand = words[idx_pm]
                            if re.match(r"^\d+$", cand):
                                if (len(words) > idx_pm + 1
                                        and re.match(r"^\d+\.\d+$", words[idx_pm + 1])):
                                    stone_on_line = cand

                        # Size & Quantity – pattern: <decimal> <int> 0.0000 ----
                        sq = re.search(r"(\d+\.\d+)\s+(\d+)\s+0\.0000", rest)
                        size_val = sq.group(1) if sq else ""
                        qty_val  = sq.group(2) if sq else ""

                        # Initial description ----------------------------------
                        desc_parts = [current_cat or "", metal_desc]
                        if stone_on_line:
                            desc_parts.append(stone_on_line)
                        desc = " ".join(p for p in desc_parts if p)

                        current_item = dict(hdr)
                        current_item.update({
                            "#":             item_counter,
                            "Memo #":        "",
                            "Item #":        item_no,
                            "Vendor Item #": vendor_item,
                            "Description":   desc,
                            "Size":          size_val,
                            "Quantity":      qty_val,
                            "_extra":        [],
                        })
                        expect_item = False
                        continue

                # ── Description continuation lines ────────────────────────────
                if current_item and line:
                    if any(kw in line for kw in DESC_KW):
                        current_item["_extra"].append(line)

            _save(current_item)   # last item on page
            current_item = None

    return all_rows


def _build_df(rows: list) -> pd.DataFrame:
    """Turn extracted rows into a properly ordered DataFrame."""
    df = pd.DataFrame(rows)
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[COLUMNS]


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

        df = _build_df(rows)

        stem = Path(input_path).stem
        out_file = os.path.join(output_dir, f"SHEFI_NEW_PO_{stem}.xlsx")
        df.to_excel(out_file, index=False)

        return True, out_file, None, df

    except Exception as exc:
        return False, None, str(exc), None


# ── CLI bulk-folder runner ────────────────────────────────────────────────────

def main():
    """Process every PDF in SHEFI_NEW_PO/ and write results to SHEFI_NEW_PO/Output/."""
    _OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    pdf_files = sorted(_INPUT_FOLDER.glob("*.pdf"))

    if not pdf_files:
        print(f"No PDF files found in: {_INPUT_FOLDER}")
        sys.exit(1)

    print(f"Found {len(pdf_files)} PDF(s) in : {_INPUT_FOLDER}")
    print(f"Output folder           : {_OUTPUT_FOLDER}")
    print("=" * 70)

    combined: list = []

    for pdf_path in pdf_files:
        print(f"\nProcessing: {pdf_path.name} ...", end=" ", flush=True)
        success, out_file, err, df = process_shefi_new_file(
            str(pdf_path), str(_OUTPUT_FOLDER)
        )
        if success:
            print(f"{len(df)} item(s) -> {Path(out_file).name}")
            combined.append(df)
        else:
            print(f"ERROR - {err}")

    if combined:
        df_all = pd.concat(combined, ignore_index=True)
        combo = _OUTPUT_FOLDER / "SHEFI_ALL_POs_COMBINED.xlsx"
        df_all.to_excel(str(combo), index=False)
        print(f"\n{'='*70}")
        print(f"Combined ({len(df_all)} total items): {combo.name}")

    print("\nDone.")


if __name__ == "__main__":
    main()
