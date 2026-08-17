import pandas as pd


def find_header_row(filepath, key_column_name):
    """Reads the raw Excel sheet without headers and searches for the row that

    contains the specified key column name.
    """
    raw_df = pd.read_excel(filepath, header=None)

    for row_idx, row in raw_df.iterrows():
        row_values = row.astype(str).str.strip().values
        if key_column_name in row_values:
            return row_idx

    raise ValueError(
        f"Could not find header '{key_column_name}' in file: {filepath}"
    )


# --- Strict Cleaning Helpers ---
def clean_style(series):
    """Trims whitespace and converts style numbers to uppercase."""
    return series.astype(str).str.strip().str.upper()


def clean_size(series):
    """Trims whitespace and converts float strings like '7.0' to '7'."""
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.lower()
    )


def clean_color(series):
    """Standardizes color names and abbreviations.

    If two colors are detected (e.g., 'yellow white gold', 'white/yellow', 'two
    tone'), maps to 'tt'.
    """

    def parse_color_value(val):
        val = str(val).strip().lower()

        # 1. Direct two-tone / multi-color checks
        if any(
            kw in val
            for kw in ["two tone", "two-tone", "tt", "two color", "two-color"]
        ):
            return "tt"

        # 2. Check for presence of multiple distinct color keywords
        color_families = {
            "w": ["white", "wg"],
            "y": ["yellow", "yg"],
            "p": ["pink", "rose", "pg", "rg"],
            "pt": ["platinum", "plat"],
        }

        found_families = set()
        for family_code, keywords in color_families.items():
            for kw in keywords:
                if kw in val:
                    found_families.add(family_code)
                    break

        if len(found_families) >= 2:
            return "tt"

        # 3. Standard single-color mapping dictionary
        color_map = {
            "white gold": "w",
            "white": "w",
            "wg": "w",
            "w": "w",
            "yellow gold": "y",
            "yellow": "y",
            "yg": "y",
            "y": "y",
            "pink gold": "p",
            "rose gold": "p",
            "pg": "p",
            "rg": "p",
            "p": "p",
            "platinum": "pt",
            "plat": "pt",
            "pt": "pt",
            "alloy": "al",
            "al": "al",
        }

        return color_map.get(val, val)

    return series.apply(parse_color_value)


def has_sm_suffix(series):
    """Checks if the string ends with 'SM' (case-insensitive)."""
    return series.astype(str).str.strip().str.upper().str.endswith("SM")


def find_quantity_column(columns):
    """Dynamically finds the quantity column name regardless of casing/naming."""
    common_names = ["quantity", "qty", "order qty", "pcs", "total qty"]
    for col in columns:
        if col.strip().lower() in common_names:
            return col

    for col in columns:
        if "qty" in col.lower() or "quantity" in col.lower():
            return col
    return None


def process_jewelry_files(input_filepath, master_filepath, item_po_no=""):
    # 1. Dynamically locate header rows
    input_header_idx = find_header_row(
        input_filepath, "Elegant Jewelry Style #"
    )
    master_header_idx = find_header_row(master_filepath, "Style No")

    # 2. Load the Excel files using detected header rows
    input_df = pd.read_excel(input_filepath, header=input_header_idx)
    master_df = pd.read_excel(master_filepath, header=master_header_idx)

    # 3. Clean column headers (strip spaces)
    input_df.columns = input_df.columns.astype(str).str.strip()
    master_df.columns = master_df.columns.astype(str).str.strip()

    # 4. Find the Quantity column in input file
    qty_col = find_quantity_column(input_df.columns)
    if not qty_col:
        raise KeyError(
            "Could not identify a 'Quantity' or 'QTY' column in the input file."
        )

    # Assign SKUNo from 'OMJ Style #'
    if "OMJ Style #" in input_df.columns:
        input_df["SKUNo"] = input_df["OMJ Style #"].astype(str).str.strip()
    else:
        input_df["SKUNo"] = ""

    # Set default OrderItemPcs to '1'
    input_df["OrderItemPcs"] = "1"

    # 5. Strictly normalize Input DataFrame keys
    input_df["match_style"] = clean_style(input_df["Elegant Jewelry Style #"])
    input_df["match_color"] = clean_color(input_df["Metal Color"])
    input_df["match_size"] = clean_size(input_df["Size"])

    if "OMJ Style #" in input_df.columns:
        input_df["match_sm"] = has_sm_suffix(input_df["OMJ Style #"])
    else:
        input_df["match_sm"] = False

    # 6. Strictly normalize Master DataFrame keys
    master_df["match_style"] = clean_style(master_df["Style No"])
    master_df["match_color"] = clean_color(master_df["COLOR"])
    master_df["match_size"] = clean_size(master_df["SIZE"])
    master_df["match_sm"] = has_sm_suffix(master_df["Client Style No"])

    # 7. List of required Master columns to bring over
    master_requested_cols = [
        "Client Style No",
        "COLOR",
        "ItemSize",
        "Base Metal",
        "SpecialRemarks",
        "CustomerProductionInstruction",
        "DesignProductionInstruction",
        "Stamping Instruction",
    ]

    available_master_cols = [
        col for col in master_requested_cols if col in master_df.columns
    ]
    merge_keys = ["match_style", "match_color", "match_size", "match_sm"]

    master_cleaned = master_df[
        merge_keys + available_master_cols
    ].drop_duplicates(subset=merge_keys, keep="first")

    # 8. Merge on all keys (Style, Color, Size, SM condition)
    merged_df = pd.merge(input_df, master_cleaned, on=merge_keys, how="left")

    # Drop missing matches
    df = merged_df.dropna(subset=["Client Style No"]).copy()

    # 9. Map/Rename existing columns
    df["StyleCode"] = df["Client Style No"]
    df["OrderQty"] = df[qty_col]
    df["Metal"] = df["Base Metal"] if "Base Metal" in df.columns else ""
    df["Tone"] = df["COLOR"] if "COLOR" in df.columns else ""
    df["StampInstruction"] = (
        df["Stamping Instruction"]
        if "Stamping Instruction" in df.columns
        else ""
    )
    df["ItemPoNo"] = item_po_no

    # Ensure optional master columns exist if missing from sheet
    for col in [
        "ItemSize",
        "CustomerProductionInstruction",
        "SpecialRemarks",
        "DesignProductionInstruction",
    ]:
        if col not in df.columns:
            df[col] = ""

    # 10. Generate Auto-incremental SrNo
    df["SrNo"] = range(1, len(df) + 1)

    # 11. Add Blank Columns
    blank_columns = [
        "ItemRefNo",
        "StockType",
        "MakeType",
        "OrderGroup",
        "Certificate",
        "Basestoneminwt",
        "Basestonemaxwt",
        "Basemetalminwt",
        "Basemetalmaxwt",
        "Productiondeliverydate",
        "Expecteddeliverydate",
        "BlankColumn",
        "SetPrice",
        "StoneQuality",
        "Date",
        "PoDate",
        "E Del Date",
    ]
    for col in blank_columns:
        df[col] = ""

    # 12. Final Ordered Schema
    final_ordered_columns = [
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
        "BlankColumn",
        "SetPrice",
        "StoneQuality",
        "Date",
        "PoDate",
        "E Del Date",
    ]

    return df[final_ordered_columns]


# Example Usage
if __name__ == "__main__":
    po_input = input("Enter Item PO Number: ") or "PO-998877"

    result_df = process_jewelry_files(
        input_filepath=r"D:\latest\omj\Copy of PO# VAL81126.xlsx",
        master_filepath=r"D:\latest\OPS-TOOL-V2\CS_100826\OMJ_CS_1408.xlsx",
        item_po_no=po_input,
    )

    print(result_df)
    result_df.to_excel('output3.xlsx')