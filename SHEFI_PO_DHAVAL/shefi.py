"""
SHEFI NEW PO - Batch PDF to Excel Converter
Processes every *.pdf found in the script's directory and writes
one Excel file per PDF into the Output/ sub-folder.
"""

import os
import re
import glob
from pathlib import Path
import pdfplumber
import pandas as pd
from datetime import datetime


# ── Paths / Helpers ──────────────────────────────────────────────────────────

# Project base directory (OPS_Tool root), used to locate shared resources
BASE_DIR = Path(__file__).resolve().parent.parent

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
    # Also includes special characters like +, -, / that may appear in item codes
    ITEM_PAT = re.compile(r'^(\d+)\s+([A-Z0-9][A-Z0-9+\-/]+)\s+(.*)', re.DOTALL)

    # Category / description header line.
    # Accepts an optional leading "CODE[/NUM] [/ ] [-]" prefix (PDF rendering artefact).
    # The prefix MUST contain at least one digit so that letter-only words like
    # "LGD" are not accidentally split off from the category name.
    # Then: "CategoryWord(s): optional_metal_desc"
    # Examples: "NK0000102QK - LGD Opera: 14KW" or "ABB02091M - LGD Bracelet: 14KW"
    CAT_PAT = re.compile(
        r'^(?:([A-Z0-9]*\d[A-Z0-9]*(?:/\d+)?)\s+(?:[-/]\s+)?)?([A-Za-z][A-Za-z0-9 ]+):\s*(.*)'
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

        Handles five size formats (tried in order):
          1. inch size – …desc  17"  QTY  0.0000…   (necklace/bracelet length in inches)
          2. mm size   – …desc  9.70mm  QTY  0.0000…   (pendant/earring diameter)
          3. decimal   – …desc  6.5     QTY  0.0000…   (ring size with decimal)
          4. integer   – …desc  6       QTY  0.0000…   (ring size as whole number)
          5. no size   – …desc          QTY  0.0000…
        """
        tokens = rest.split()
        vendor_code = ""
        start = 0

        if tokens and is_vendor_code(tokens[0]):
            vendor_code = tokens[0]
            start = 1

        tail = " ".join(tokens[start:])

        # 1. inch size (e.g. 17", 18", 7"): …desc SIZE" QTY 0.0000…
        m = re.search(r'^(.*?)\s+(\d+(?:\.\d+)?")\s+(\d+)\s+0\.0000', tail)
        if m:
            return vendor_code, m.group(1).strip(), m.group(2), m.group(3)

        # 2. mm size (e.g. 9.70mm, 4.30mm, 6mm): …desc SIZE_mm QTY 0.0000…
        m = re.search(r'^(.*?)\s+(\d+(?:\.\d+)?mm)\s+(\d+)\s+0\.0000', tail)
        if m:
            return vendor_code, m.group(1).strip(), m.group(2), m.group(3)

        # 3. Decimal size without mm (e.g. 6.5): …desc SIZE QTY 0.0000…
        m = re.search(r'^(.*?)\s+(\d+\.\d+)\s+(\d+)\s+0\.0000', tail)
        if m:
            return vendor_code, m.group(1).strip(), m.group(2), m.group(3)

        # 4. Integer size (e.g. 6): two consecutive integers before 0.0000
        #    Note: "N 0.0000" alone never matches here because (\d+)\s+0\.0000
        #    would grab "0" from "0.0000" and then fail on the dot.
        m = re.search(r'^(.*?)\s+(\d+)\s+(\d+)\s+0\.0000', tail)
        if m:
            return vendor_code, m.group(1).strip(), m.group(2), m.group(3)

        # 5. No size — only qty before 0.0000
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
                    # Skip lines that are shipping instructions, price info, or footer junk
                    if not (line.startswith('>') or 
                            'SHIP ON' in line.upper() or 
                            re.search(r'JULY<', line) or
                            re.match(r'^[\d\s.]+Q\s+\$', line) or
                            re.match(r'^\d+\s+0\.0000\s+Q', line)):
                        current_item["_extra"].append(line)

            # Save any item still open at end of page
            flush_item()

    return all_rows


# ── Excel writer ──────────────────────────────────────────────────────────────

COLUMNS = [
    "Customer", "Order#", "Page#", "PO#", "Date", "Due Date", "Cancel Date",
    "Ref", "Vendor#", "Ship Via", "#", "Memo #", "Item #", "Vendor Item #",
    "Description", "Size", "Quantity", "Gold Karat", "Metal Color", 
    "present_in_csm", "csm_style",
]


def save_to_excel(rows: list[dict], output_path: str) -> pd.DataFrame:
    df = pd.DataFrame(rows)

    # Ensure all expected columns exist
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # Extract Gold Karat and Metal Color from Description column
    def extract_gold_karat(description: str) -> str:
        """Extract gold karat or platinum info from description."""
        desc = str(description).upper()
        
        # Check for platinum/PT first
        if re.search(r'\bPT\b|\bPLATINUM\b|\bPC95\b', desc):
            if re.search(r'\bPC95\b', desc):
                return 'PC95'
            elif re.search(r'\bPT\b', desc):
                return 'PT'
            else:
                return 'PLATINUM'
        
        # Pattern 1: Number + K + Letter (e.g., 14KW, 18KY, 10KR, 14TT)
        match = re.search(r'(\d{1,2})K([A-Z])', desc)
        if match:
            return f"{match.group(1)}K{match.group(2)}"
        
        # Pattern 2: Number + TT (two-tone like 14TT, 10TT)
        match = re.search(r'(\d{1,2})TT', desc)
        if match:
            return f"{match.group(1)}TT"
        
        # Pattern 3: Number + K + space + Letter (e.g., "14K W")
        match = re.search(r'(\d{1,2})K\s+([A-Z])', desc)
        if match:
            return f"{match.group(1)}K{match.group(2)}"
        
        # Pattern 4: Just number + K (e.g., 14K, 18K, 10K)
        match = re.search(r'(\d{1,2})K\b', desc)
        if match:
            return f"{match.group(1)}K"
        
        # Pattern 5: Number + space + "PLATINUM" (e.g., "10 PLATINUM")
        match = re.search(r'\b(\d{1,2})\s+PLATINUM\b', desc)
        if match:
            return f"{match.group(1)}PLATINUM"
        
        return ''
    
    def determine_metal_color(gold_karat: str) -> str:
        """Determine metal color based on gold karat notation."""
        if not gold_karat:
            return ''
        
        karat_upper = gold_karat.upper()
        
        # White indicator (W suffix)
        if karat_upper.endswith('W'):
            return 'White'
        
        # Yellow indicator (Y suffix)
        if karat_upper.endswith('Y'):
            return 'Yellow'
        
        # Rose/Red indicator (R suffix)
        if karat_upper.endswith('R'):
            return 'Rose'
        
        # Two-Tone indicator (TT suffix)
        if karat_upper.endswith('TT') or 'TT' in karat_upper:
            return 'Two-Tone'
         
        # Platinum/PT/PC95 should have blank Metal Color
        if any(pt in karat_upper for pt in ['PT', 'PC95', 'PLATINUM']):
            return ''
        
        # For just "14K", "18K" etc. without color suffix, leave blank
        return ''
    
    # Apply extraction to Description column
    if 'Description' in df.columns:
        df['Gold Karat'] = df['Description'].apply(extract_gold_karat)
        df['Metal Color'] = df['Gold Karat'].apply(determine_metal_color)

    # SHEFI New PO extras:
    # 1) If "Vendor Item #" contains multiple styles like "P33164 / 204128P",
    #    and exactly one of those tokens exists in shefi_cs.xlsx (Style_No),
    #    replace "Vendor Item #" with the matching token.
    # 2) Add "present_in_csm" as yes/no depending on whether Vendor Item # is
    #    present in shefi_cs.xlsx (Style_No).
    def _split_vendor_tokens(v: object) -> list[str]:
        v_str = str(v).strip()
        if not v_str:
            return []
        return [t.strip() for t in re.split(r"\s*/\s*", v_str) if t.strip()]

    def _vendor_prefix(v: object) -> str:
        tokens = _split_vendor_tokens(v)
        return tokens[0] if tokens else str(v).strip()

    def _parse_metal_from_description(desc_upper: str) -> tuple[str, int | None, bool]:
        """
        Returns (tone_letter, karat, include_G).

        Examples in description:
          - "14KY" -> tone="Y", karat=14, include_G=True
          - "10TT" -> tone="T", karat=10, include_G=True
          - "14TT" -> tone="T", karat=14, include_G=False  (exception: no 'G')
        """
        # Exception / special TT tokens first
        m = re.search(r"(14TT)", desc_upper)
        if m:
            return "T", 14, False
        m = re.search(r"(10TT)", desc_upper)
        if m:
            return "T", 10, True

        # Normal Kx tokens like 14KY / 14KW / 10KR
        m = re.search(r"(14)K([A-Z])", desc_upper)
        if m:
            return m.group(2), 14, True
        m = re.search(r"(10)K([A-Z])", desc_upper)
        if m:
            return m.group(2), 10, True

        # Fallback: if we only find "14K" or "10K" without trailing letter
        m = re.search(r"(14)K", desc_upper)
        if m:
            return "", 14, True
        m = re.search(r"(10)K", desc_upper)
        if m:
            return "", 10, True

        return "", None, False

    def _build_csm_style(vendor_item: object, description: object, size_val: object) -> str:
        prefix = _vendor_prefix(vendor_item)
        if not prefix:
            return ""

        desc_upper = str(description or "").upper()
        size_str = str(size_val or "").strip()

        tone, karat, include_g = _parse_metal_from_description(desc_upper)
        if karat is None:
            # Unknown metal => keep only the vendor style prefix
            return prefix

        # X always present; V only for 14K. G present for 10K/14K except 14TT.
        xv = "XV" if karat == 14 else "X"
        g_part = "G" if include_g else ""
        metal_part = f"{tone}{g_part}{xv}"

        is_ring = "RING" in desc_upper
        is_bracelet = "BRACELET" in desc_upper

        # Only append size for Rings/Bracelets (avoid earrings/pendants, etc.)
        if (is_ring or is_bracelet) and size_str and size_str.lower() not in ("nan", ""):
            in_suffix = "IN" if is_bracelet else ""
            return f"{prefix}-{size_str}{in_suffix}{metal_part}"

        # No size: examples like "204210P-YGXV"
        return f"{prefix}-{metal_part}"

    def _normalise_size(v: object) -> str:
        """
        Keep size when it is valid:
        - Numeric values within [4, 20] (ring sizes)
        - Inch measurements like 17", 7" (necklaces/bracelets)
        - mm measurements like 9.70mm, 6mm (pendants/earrings)
        Otherwise return blank.
        """
        s = str(v or "").strip()
        if not s or s.lower() == "nan":
            return ""
        
        # Keep inch measurements (e.g., 17", 7.5")
        if s.endswith('"'):
            return s
        
        # Keep mm measurements (e.g., 9.70mm, 6mm)
        if s.endswith('mm'):
            return s
        
        # Numeric ring sizes - only keep if in valid range [4, 20]
        try:
            n = float(s)
            if 4.0 <= n <= 20.0:
                return s
        except Exception:
            pass
        
        return ""

    shefi_cs_path = BASE_DIR / "shefi_cs.xlsx"
    if shefi_cs_path.exists():
        try:
            shefi_cs_df = pd.read_excel(shefi_cs_path)
            # Normalise the master style list
            style_no_list = (
                shefi_cs_df["Style_No"]
                .astype(str)
                .str.strip()
                .tolist()
            )
            style_no_set = set(style_no_list)

            def _pick_present_style(v: object) -> str:
                v_str = str(v).strip()
                if not v_str:
                    return v_str

                tokens = [t.strip() for t in re.split(r"\s*/\s*", v_str) if t.strip()]
                if len(tokens) < 2:
                    return v_str

                present = [t for t in tokens if t in style_no_set]
                if len(present) == 1:
                    return present[0]
                return v_str

            df["Vendor Item #"] = df["Vendor Item #"].apply(_pick_present_style)
            df["present_in_csm"] = df["Vendor Item #"].apply(
                lambda x: "yes" if str(x).strip() in style_no_set else "no"
            )
        except Exception:
            # On any issue with reading / parsing the master, default to "no"
            df["present_in_csm"] = "no"
    else:
        # Master file missing – still return a valid column
        df["present_in_csm"] = "no"

    # Validate size range (only keep sizes 4..20).
    df["Size"] = df["Size"].apply(_normalise_size)

    # csm_style rules are based on description data and size.
    df["csm_style"] = df.apply(
        lambda r: _build_csm_style(
            r.get("Vendor Item #", ""),
            r.get("Description", ""),
            r.get("Size", ""),
        ),
        axis=1,
    )

    df = df[COLUMNS]
    
    # Replace any remaining NaN values with empty strings
    df = df.fillna('')
    
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
