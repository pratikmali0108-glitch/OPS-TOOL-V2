"""
JU.py  —  Process one or more PO Excel files into a standardised output.

Usage:
    python JU.py                        # auto-detect all .xlsx in current folder
    python JU.py file1.xlsx file2.xlsx  # process specific files

Web usage:
    from JU import process_ju_excel_file
    success, output_path, error, df = process_ju_excel_file(path, out_dir,
                                                             item_po_no='PO123',
                                                             priority='REG')
"""

import pandas as pd
import re
import os
import sys
import glob

FINAL_COLUMNS = [
    "SrNo",
    "StyleCode",
    "ItemSize",
    "OrderQty",
    "OrderItemPcs",
    "Metal",
    "Tone",
    "ItemPoNo",
    "ItemRefNo",
    "StockType",
    "Priority",
    "MakeType",
    "CustomerProductionInstruction",
    "SpecialRemarks",
    "DesignProductionInstruction",
    "StampInstruction",
    "OrderGroup",
    "Certificate",
    "SKUNo",
    "Basestoneminwt",
    "Basestonemaxwt",
    "Basemetalminwt",
    "Basemetalmaxwt",
    "Productiondeliverydate",
    "Expecteddeliverydate",
    "SetPrice",
    "StoneQuality",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def map_style_and_size(style_code, _original_size):
    """
    Split StyleCode into (base, ItemSize).

    The last '-' segment is the size segment.  Within that segment the size
    code is always 2 uppercase letters + 2 digits (trailing '0' stripped).
    Any characters before the size code in that segment are a prefix that
    gets reattached to the base with a '-'.

    Examples
    --------
    'FR1500JE-SM-US070'       → ('FR1500JE-SM',          'US07')
    'ABR03466LE-YGGDUP090'    → ('ABR03466LE-YGGD',       'UP09')
    'SR0153NG-YGGD.UP100'     → ('SR0153NG-YGGD',         'UP10')
    'ZR7800F-WGUP070'         → ('ZR7800F-WG',            'UP07')
    'AAR02510YE-YGC4CZUP070'  → ('AAR02510YE-YGC4CZ',     'UP07')
    'ZR10567H-WGGDUP070'      → ('ZR10567H-WGGD',         'UP07')
    'TR1058HA-WGE3-UP070'     → ('TR1058HA-WGE3',         'UP07')
    """
    if not isinstance(style_code, str):
        return style_code, ""

    parts = style_code.split("-")
    if len(parts) > 1:
        size_raw = parts[-1]
        base     = "-".join(parts[:-1])
    else:
        size_raw = parts[0]
        base     = parts[0]

    # Normalise alternate separators (e.g. "YGGD.UP100" → "YGGDUP100")
    size_raw = size_raw.replace(".", "")

    # Strip trailing '0' padding (e.g. "UP070" → "UP07")
    stripped = size_raw[:-1] if size_raw.endswith("0") and len(size_raw) > 1 else size_raw

    # Peel off the 2-letter + 2-digit size code from the end.
    # Any leading characters are a prefix belonging to the base.
    m = re.match(r'^(.*?)([A-Z]{2}\d{2})$', stripped, re.IGNORECASE)
    if m:
        prefix = m.group(1)
        size   = m.group(2).upper()
        if prefix:
            base = base + "-" + prefix.upper()
    else:
        size = stripped

    return base, size


def resolve_metal_tone(mitemcode):
    """
    Return (Metal, Tone) for a raw MItemCode value.

    General rule for GM{karat}{tone}CS codes
    (e.g. GM10WCS, GM10YCS, GM14WCS, GM14YCS):
        Metal = G{karat}{tone}   e.g. "G10W"
        Tone  = {tone}           e.g. "W"
    Falls back to parse_metal_tone for everything else.
    """
    if isinstance(mitemcode, str):
        key = mitemcode.strip()
        m = re.match(r'^GM(\d+)([A-Z])CS$', key, re.IGNORECASE)
        if m:
            karat = m.group(1)
            tone  = m.group(2).upper()
            return (f"G{karat}{tone}", tone)
    return parse_metal_tone(mitemcode)


def parse_metal_tone(mitemcode):
    """
    Split MItemCode into (Metal, Tone).

    'AG925'    → ('AG925', '')
    'PC95'     → ('PC95',  '')
    'G14W'     → ('G14',   'W')
    'G10W/P/Y' → ('G10',   'W/P/Y')
    'GM14YCS'  → ('GM14',  'YCS')

    Rule: leading letters+digits = Metal; remaining letters/slashes = Tone.
    """
    if not isinstance(mitemcode, str) or not mitemcode.strip():
        return ("", "")

    mitemcode = mitemcode.strip()
    m = re.match(r'^([A-Z]+\d+)([A-Z/]*)$', mitemcode, re.IGNORECASE)
    if m:
        metal = m.group(1).upper()
        tone  = m.group(2).upper().strip("/")
        return (metal, tone)

    return (mitemcode, "")


def extract_sku(special_remarks):
    """
    Extract 'SKU#XXXXXXX' (with optional space after #) from SpecialRemarks.
    'SKU#1715093,...'   → 'SKU#1715093'
    'SKU# 49208821,...' → 'SKU#49208821'
    Returns '' if not found.
    """
    if not isinstance(special_remarks, str):
        return ""
    m = re.search(r'SKU#\s*([^,\s]+)', special_remarks, re.IGNORECASE)
    return f"SKU#{m.group(1).strip()}" if m else ""


# ---------------------------------------------------------------------------
# Core transformation (shared by CLI and web)
# ---------------------------------------------------------------------------

def _build_output_df(df, item_po_no, priority):
    """
    Transform a raw source DataFrame into the standardised output DataFrame.
    Returns (out_df, error_str).  error_str is None on success.
    """
    if "SrNo" not in df.columns:
        return None, "'SrNo' column not found in source file."

    df = df[df["SrNo"] == 1].copy().reset_index(drop=True)

    if df.empty:
        return None, "No rows with SrNo = 1 found in source file."

    out = pd.DataFrame()
    out["SrNo"] = range(1, len(df) + 1)

    style_mapped, size_mapped = zip(
        *df.apply(
            lambda r: map_style_and_size(r.get("StyleCode"), r.get("ItemSize")),
            axis=1,
        )
    )
    out["StyleCode"] = list(style_mapped)
    out["ItemSize"]  = list(size_mapped)

    out["OrderQty"]     = df.get("InwardQty", "")
    out["OrderItemPcs"] = df.get("ItemPcs", "")

    metal_tone   = df["MItemCode"].apply(resolve_metal_tone)
    out["Metal"] = metal_tone.apply(lambda x: x[0])
    out["Tone"]  = metal_tone.apply(lambda x: x[1])

    out["ItemPoNo"]  = item_po_no
    out["ItemRefNo"] = ""

    out["StockType"] = df.get("StockType", "")
    out["Priority"]  = priority
    out["MakeType"]  = df.get("MakeType", "")

    out["CustomerProductionInstruction"] = df.get("CustomerProductionInstruction", "")
    out["SpecialRemarks"]                = df.get("SpecialRemarks", "")
    out["DesignProductionInstruction"]   = df.get("DesignProductionInstruction", "")
    out["StampInstruction"]              = df.get("StampingInstruction", "")

    out["OrderGroup"]  = "STERLING JEWELERS(OUTLET)"
    out["Certificate"] = ""

    out["SKUNo"] = df["SpecialRemarks"].apply(extract_sku)

    for col in ["Basestoneminwt", "Basestonemaxwt", "Basemetalminwt", "Basemetalmaxwt",
                "Productiondeliverydate", "Expecteddeliverydate", "SetPrice", "StoneQuality"]:
        out[col] = ""

    return out[FINAL_COLUMNS], None


# ---------------------------------------------------------------------------
# Web-facing processor (returns (success, output_path, error, df))
# ---------------------------------------------------------------------------

def process_ju_excel_file(filepath: str, output_dir: str,
                          item_po_no: str = "", priority: str = "REG"):
    """
    Process a single JU Excel PO file.

    Returns:
        (success: bool, output_path: str | None, error: str | None, df: DataFrame | None)
    """
    try:
        raw_df = pd.read_excel(filepath)
        out_df, err = _build_output_df(raw_df, item_po_no, priority)
        if err:
            return False, None, err, None

        base_name      = os.path.splitext(os.path.basename(filepath))[0]
        output_filename = f"JU_{base_name}_Output.xlsx"
        output_path     = os.path.join(output_dir, output_filename)
        out_df.to_excel(output_path, index=False)
        return True, output_path, None, out_df
    except Exception as exc:
        return False, None, str(exc), None


# ---------------------------------------------------------------------------
# CLI per-file processing
# ---------------------------------------------------------------------------

def process_file(input_file, item_po_no, priority):
    print(f"\nReading '{input_file}' ...")
    try:
        raw_df = pd.read_excel(input_file)
    except Exception as exc:
        print(f"  ERROR reading file: {exc}")
        return

    out_df, err = _build_output_df(raw_df, item_po_no, priority)
    if err:
        print(f"  WARNING: {err} — skipping.")
        return

    print(f"  {len(out_df)} StyleCode group(s) found (rows where SrNo = 1)")
    output_file = os.path.splitext(input_file)[0] + "_Output.xlsx"
    out_df.to_excel(output_file, index=False)
    print(f"  Saved: '{output_file}'  ({len(out_df)} rows x {len(out_df.columns)} cols)")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    # Resolve input files
    if len(sys.argv) > 1:
        input_files = sys.argv[1:]
    else:
        # Auto-detect all .xlsx files in current directory, skip _Output files
        input_files = [
            f for f in glob.glob("*.xlsx")
            if not f.endswith("_Output.xlsx")
        ]
        if not input_files:
            print("No .xlsx files found in the current directory.")
            return
        print(f"Found {len(input_files)} Excel file(s): {input_files}")

    # Validate all files exist before asking for inputs
    missing = [f for f in input_files if not os.path.exists(f)]
    if missing:
        for f in missing:
            print(f"Error: file not found — '{f}'")
        return

    # Process each file with its own ItemPoNo / Priority
    for input_file in input_files:
        print(f"\n--- {input_file} ---")
        item_po_no = input(f"  Enter ItemPoNo for '{input_file}' : ").strip()
        priority   = input(f"  Enter Priority  for '{input_file}' : ").strip()
        process_file(input_file, item_po_no, priority)

    print("\nAll files processed.")


if __name__ == "__main__":
    main()
