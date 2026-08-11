"""
extract_sku.py
--------------
Extracts SKU values from all CSV and XLS files found in subdirectories.

Handles two file types:
  - CSV  : looks for a column named  'SKU #', 'SKU#', or 'SKU'  (case-insensitive)
  - XLS  : these are HTML-disguised-as-XLS files; parses the embedded PODETAIL
            table and extracts the 'SKU' column  (the one whose header appears
            at cell B12 when opened in Excel)

Output  : sku_output.csv  written to the same folder as this script.
"""

import os
import re
import csv
import io
import pandas as pd
from bs4 import BeautifulSoup


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

SKU_COLUMN_ALIASES = {"sku #", "sku#", "sku"}      # normalised (lower-stripped)


def _normalise(name: str) -> str:
    return name.strip().lower()


# ---------------------------------------------------------------------------
# CSV  extraction
# ---------------------------------------------------------------------------

def extract_sku_from_csv(filepath: str) -> list[dict]:
    """
    Reads a CSV/TSV file, finds the SKU column (SKU #, SKU#, or SKU),
    and returns a list of dicts with keys: file, folder, sku.
    """
    rows = []
    encodings = ["utf-8", "windows-1252", "latin-1"]

    raw = None
    for enc in encodings:
        try:
            with open(filepath, encoding=enc) as fh:
                raw = fh.read()
            break
        except UnicodeDecodeError:
            continue

    if raw is None:
        print(f"  [WARN] Could not decode {filepath}")
        return rows

    # strip the Excel "sep=<char>" hint line if present
    lines = raw.splitlines()
    sep = ","
    start = 0
    if lines and lines[0].startswith("sep="):
        # do NOT strip — the separator may itself be whitespace (e.g. a tab)
        sep_char = lines[0][4:]
        if sep_char:
            sep = sep_char
        start = 1
        raw = "\n".join(lines[start:])

    try:
        df = pd.read_csv(io.StringIO(raw), sep=sep, dtype=str, low_memory=False)
    except Exception as exc:
        print(f"  [WARN] pandas failed on {filepath}: {exc}")
        return rows

    # find matching SKU column
    sku_col = None
    for col in df.columns:
        if _normalise(col) in SKU_COLUMN_ALIASES:
            sku_col = col
            break

    if sku_col is None:
        print(f"  [WARN] No SKU column found in {filepath}  (cols: {list(df.columns)[:8]})")
        return rows

    folder = os.path.basename(os.path.dirname(filepath))
    fname  = os.path.basename(filepath)

    for val in df[sku_col].dropna():
        val = str(val).strip()
        if val:
            rows.append({"folder": folder, "file": fname, "sku": val})

    return rows


# ---------------------------------------------------------------------------
# XLS (HTML) extraction
# ---------------------------------------------------------------------------

def _is_html_xls(filepath: str) -> bool:
    """Returns True when the .xls file is actually an HTML document."""
    try:
        with open(filepath, "rb") as fh:
            header = fh.read(16)
        return header.lstrip().startswith(b"<")
    except OSError:
        return False


def extract_sku_from_html_xls(filepath: str) -> list[dict]:
    """
    Parses an HTML-disguised XLS file, locates the PODETAIL table
    (the one that contains a 'SKU' column header), and extracts SKU values.
    """
    rows = []
    encodings = ["windows-1252", "utf-8", "latin-1"]

    soup = None
    for enc in encodings:
        try:
            with open(filepath, encoding=enc) as fh:
                content = fh.read()
            soup = BeautifulSoup(content, "html.parser")
            break
        except (UnicodeDecodeError, Exception):
            continue

    if soup is None:
        print(f"  [WARN] Could not parse {filepath}")
        return rows

    folder = os.path.basename(os.path.dirname(filepath))
    fname  = os.path.basename(filepath)

    for table in soup.find_all("table"):
        table_rows = table.find_all("tr")
        if not table_rows:
            continue

        # look for a header row that contains a 'SKU' cell
        header_idx = None
        sku_col_idx = None
        for i, tr in enumerate(table_rows):
            cells = tr.find_all("td")
            texts = [c.get_text(strip=True) for c in cells]
            for j, t in enumerate(texts):
                if _normalise(t) in SKU_COLUMN_ALIASES:
                    header_idx  = i
                    sku_col_idx = j
                    break
            if header_idx is not None:
                break

        if header_idx is None or sku_col_idx is None:
            continue   # this table has no SKU column

        # extract data rows below the header
        for tr in table_rows[header_idx + 1:]:
            cells = tr.find_all("td")
            if sku_col_idx >= len(cells):
                continue
            val = cells[sku_col_idx].get_text(strip=True)
            # skip blank, total/summary rows
            if not val or not re.match(r"^\d+", val):
                continue
            rows.append({"folder": folder, "file": fname, "sku": val})

    if not rows:
        print(f"  [WARN] No SKU data found in {filepath}")

    return rows


# ---------------------------------------------------------------------------
# XLS (real binary Excel) extraction  — fallback via xlrd
# ---------------------------------------------------------------------------

def extract_sku_from_binary_xls(filepath: str) -> list[dict]:
    try:
        import xlrd
    except ImportError:
        print("  [WARN] xlrd not installed; cannot read binary XLS.")
        return []

    rows = []
    folder = os.path.basename(os.path.dirname(filepath))
    fname  = os.path.basename(filepath)

    try:
        wb = xlrd.open_workbook(filepath)
    except Exception as exc:
        print(f"  [WARN] xlrd failed on {filepath}: {exc}")
        return rows

    for sheet in wb.sheets():
        # scan every row for a header containing SKU alias
        header_row_idx = None
        sku_col_idx = None
        for r in range(min(sheet.nrows, 30)):
            vals = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
            for c, v in enumerate(vals):
                if _normalise(v) in SKU_COLUMN_ALIASES:
                    header_row_idx = r
                    sku_col_idx    = c
                    break
            if header_row_idx is not None:
                break

        if header_row_idx is None:
            continue

        for r in range(header_row_idx + 1, sheet.nrows):
            val = str(sheet.cell_value(r, sku_col_idx)).strip()
            if val and val not in ("", "0.0", "0"):
                rows.append({"folder": folder, "file": fname, "sku": val})

    return rows


# ---------------------------------------------------------------------------
# dispatcher
# ---------------------------------------------------------------------------

def extract_sku(filepath: str) -> list[dict]:
    ext = os.path.splitext(filepath)[1].lower()

    if ext in (".csv",):
        return extract_sku_from_csv(filepath)

    if ext in (".xls", ".xlsx", ".xlsm"):
        if _is_html_xls(filepath):
            return extract_sku_from_html_xls(filepath)
        else:
            return extract_sku_from_binary_xls(filepath)

    return []


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    base_dir    = os.path.dirname(os.path.abspath(__file__))
    output_file = os.path.join(base_dir, "sku_output.csv")

    all_rows: list[dict] = []

    output_basename = os.path.basename(output_file).lower()

    for root, _dirs, files in os.walk(base_dir):
        for filename in files:
            # skip this script and the output file itself
            if filename.lower() in (os.path.basename(__file__).lower(), output_basename):
                continue
            ext = os.path.splitext(filename)[1].lower()
            if ext not in (".csv", ".xls", ".xlsx", ".xlsm"):
                continue

            filepath = os.path.join(root, filename)
            rel      = os.path.relpath(filepath, base_dir)
            print(f"Processing: {rel}")

            extracted = extract_sku(filepath)
            print(f"  -> {len(extracted)} SKU(s) found")
            all_rows.extend(extracted)

    if not all_rows:
        print("\nNo SKU values extracted from any file.")
        return

    with open(output_file, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=["file", "sku"])
        writer.writeheader()
        writer.writerows([{"file": r["file"], "sku": r["sku"]} for r in all_rows])

    print(f"\nDone.  {len(all_rows)} total SKU row(s) written to:")
    print(f"  {output_file}")


if __name__ == "__main__":
    main()
