# OMJ - Process OMJ CASTING PO files using master style lookup
import glob
import os
from pathlib import Path

import pandas as pd


def _get_omj_master_path():
    """Return the constant path to the OMJ client style master file."""
    return os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "CS_100826",
        "OMJ_CS_1408.xlsx",
    )


def find_header_row(filepath, key_column_name):
    """Find the row index that contains the specified key column name."""
    raw_df = pd.read_excel(filepath, header=None)

    for row_idx, row in raw_df.iterrows():
        row_values = row.astype(str).str.strip().values
        if key_column_name in row_values:
            return row_idx

    raise ValueError(
        f"Could not find header '{key_column_name}' in file: {filepath}"
    )


def clean_style(series):
    """Trim whitespace, convert style numbers to uppercase, and remove common suffixes."""
    return (
        series.astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"-GD$", "", regex=True)
        .str.replace(r"-GLD$", "", regex=True)
        .str.replace(r"-GOLD$", "", regex=True)
    )


def extract_style_suffix(series):
    """Extract suffix like -GD, -GLD, etc. from style numbers."""

    def get_suffix(val):
        val_str = str(val).strip().upper()
        if val_str.endswith("-GD"):
            return "GD"
        elif val_str.endswith("-GLD"):
            return "GD"  # Normalize GLD to GD
        elif val_str.endswith("-GOLD"):
            return "GD"  # Normalize GOLD to GD
        return ""

    return series.apply(get_suffix)


def clean_size(series):
    """Trim whitespace and convert float strings like '7.0' to '7'."""
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.lower()
    )


def clean_color(series):
    """Standardize color names and abbreviations."""

    def parse_color_value(val):
        val = str(val).strip().lower()

        if any(
            kw in val
            for kw in ["two tone", "two-tone", "tt", "two color", "two-color"]
        ):
            return "tt"

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


def get_color_tag(color_code):
    """Map cleaned color code ('w', 'y', 'p', 'pt', 'al', 'tt') to standard tag ('WG', 'YG', 'PG', 'PT', 'AL', 'TT')."""
    mapping = {
        "w": "WG",
        "y": "YG",
        "p": "PG",
        "pt": "PT",
        "al": "AL",
        "tt": "TT",
    }
    return mapping.get(str(color_code).lower(), str(color_code).upper())


def has_sm_suffix(series):
    """Check if the string ends with 'SM' (case-insensitive)."""
    return series.astype(str).str.strip().str.upper().str.endswith("SM")


def find_column(columns, exact_names=None, contains_any=None):
    """Find a column by exact name(s) or substring match (case-insensitive)."""
    cols = [str(c).strip() for c in columns]
    normalized = {c: c.lower() for c in cols}

    for name in exact_names or []:
        target = str(name).strip().lower()
        for col, norm in normalized.items():
            if norm == target:
                return col

    for col, norm in normalized.items():
        for token in contains_any or []:
            if token.lower() in norm:
                return col
    return None


def find_quantity_column(columns):
    """Dynamically find the quantity column name."""
    return find_column(
        columns,
        exact_names=["quantity", "qty", "order qty", "pcs", "total qty"],
        contains_any=["quantity", "qty"],
    )


def process_jewelry_files(input_filepath, master_filepath, item_po_no=""):
    """Core OMJ processing: search master using constructed Client Style No and pull linked attributes."""
    input_header_idx = find_header_row(
        input_filepath, "Elegant Jewelry Style #"
    )
    master_header_idx = find_header_row(master_filepath, "Style No")

    input_df = pd.read_excel(input_filepath, header=input_header_idx)
    master_df = pd.read_excel(master_filepath, header=master_header_idx)

    input_df.columns = input_df.columns.astype(str).str.strip()
    master_df.columns = master_df.columns.astype(str).str.strip()

    qty_col = find_quantity_column(input_df.columns)
    if not qty_col:
        raise KeyError(
            "Could not identify a 'Quantity' or 'QTY' column in the input file."
        )

    style_col = find_column(
        input_df.columns,
        exact_names=["Elegant Jewelry Style #"],
        contains_any=["elegant jewelry style"],
    )
    if not style_col:
        raise KeyError(
            "Could not identify 'Elegant Jewelry Style #' in the input file."
        )

    color_col = find_column(
        input_df.columns,
        exact_names=["Metal Color", "Metal Color ", "COLOR", "Color"],
        contains_any=["metal color", "color"],
    )
    if not color_col:
        raise KeyError(
            "Could not identify 'Metal Color' in the input file."
        )

    size_col = find_column(
        input_df.columns,
        exact_names=["Size", "SIZE", "Item Size", "ItemSize"],
        contains_any=["size"],
    )
    if not size_col:
        raise KeyError(
            "Could not identify a 'Size' column in the input file."
        )

    sku_col = find_column(
        input_df.columns,
        exact_names=["OMJ Style #", "OMJ Style # "],
        contains_any=["omj style"],
    )

    if sku_col:
        input_df["SKUNo"] = input_df[sku_col].astype(str).str.strip()
    else:
        input_df["SKUNo"] = ""

    input_df["OrderItemPcs"] = "1"

    # Cleaning helper fields
    input_df["match_style"] = clean_style(input_df[style_col])
    input_df["match_color"] = clean_color(input_df[color_col])
    input_df["match_size"] = clean_size(input_df[size_col])
    input_df["style_suffix"] = extract_style_suffix(input_df[style_col])
    input_df["match_sm"] = (
        has_sm_suffix(input_df["SKUNo"]) if sku_col else False
    )

    # Construct the target Client Style No to search in the master file
    def build_target_client_style(row):
        st = str(row["match_style"]).strip()
        sz = str(row["match_size"]).strip()
        color_tag = get_color_tag(row["match_color"])

        target = f"{st}-{sz}{color_tag}" if sz else f"{st}-{color_tag}"
        if row["match_sm"]:
            target += "SM"
        return target.upper()

    input_df["target_client_style"] = input_df.apply(
        build_target_client_style, axis=1
    )

    # Master cleaning & indexing
    master_df["norm_client_style"] = (
        master_df["Client Style No"].astype(str).str.strip().str.upper()
    )
    master_df["match_style"] = clean_style(master_df["Style No"])
    master_df["match_color"] = clean_color(master_df["COLOR"])
    master_df["match_size"] = clean_size(master_df["SIZE"])
    master_df["match_sm"] = has_sm_suffix(master_df["Client Style No"])

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

    # Primary Lookup: Direct match by constructed target_client_style
    master_primary = master_df[
        ["norm_client_style"] + available_master_cols
    ].drop_duplicates(subset=["norm_client_style"], keep="first")

    merged_df = pd.merge(
        input_df,
        master_primary,
        left_on="target_client_style",
        right_on="norm_client_style",
        how="left",
    )

    # Secondary Lookup Fallback: Match by style, color, size, SM if primary missed
    unmatched_mask = merged_df["Client Style No"].isna()
    if unmatched_mask.any():
        merge_keys = ["match_style", "match_color", "match_size", "match_sm"]
        master_secondary = master_df[
            merge_keys + available_master_cols
        ].drop_duplicates(subset=merge_keys, keep="first")

        fallback_rows = pd.merge(
            input_df[unmatched_mask].drop(
                columns=available_master_cols, errors="ignore"
            ),
            master_secondary,
            on=merge_keys,
            how="left",
        )

        for col in available_master_cols:
            merged_df.loc[unmatched_mask, col] = fallback_rows[col].values

    df = merged_df.dropna(subset=["Client Style No"]).copy()

    # Append original suffix (-GD) to the matched Client Style No
    def create_style_code(row):
        client_style = str(row["Client Style No"])
        suffix = str(row["style_suffix"]).strip()

        if suffix and suffix != "" and suffix != "nan":
            return f"{client_style}{suffix}"
        return client_style

    df["StyleCode"] = df.apply(create_style_code, axis=1)

    # --- Re-fetch instruction fields based on final generated StyleCode ---
    instruction_cols = [
        "CustomerProductionInstruction",
        "SpecialRemarks",
        "DesignProductionInstruction",
        "Stamping Instruction",
    ]
    available_inst_cols = [
        c for c in instruction_cols if c in master_df.columns
    ]

    master_inst_lookup = (
        master_df.drop_duplicates(subset=["norm_client_style"], keep="first")
        .set_index("norm_client_style")[available_inst_cols]
    )

    df["norm_style_code"] = df["StyleCode"].astype(str).str.strip().str.upper()

    if "CustomerProductionInstruction" in master_inst_lookup.columns:
        df["CustomerProductionInstruction"] = (
            df["norm_style_code"]
            .map(master_inst_lookup["CustomerProductionInstruction"])
            .fillna(df["CustomerProductionInstruction"] if "CustomerProductionInstruction" in df.columns else "")
        )
    else:
        df["CustomerProductionInstruction"] = ""

    if "SpecialRemarks" in master_inst_lookup.columns:
        df["SpecialRemarks"] = (
            df["norm_style_code"]
            .map(master_inst_lookup["SpecialRemarks"])
            .fillna(df["SpecialRemarks"] if "SpecialRemarks" in df.columns else "")
        )
    else:
        df["SpecialRemarks"] = ""

    if "DesignProductionInstruction" in master_inst_lookup.columns:
        df["DesignProductionInstruction"] = (
            df["norm_style_code"]
            .map(master_inst_lookup["DesignProductionInstruction"])
            .fillna(df["DesignProductionInstruction"] if "DesignProductionInstruction" in df.columns else "")
        )
    else:
        df["DesignProductionInstruction"] = ""

    if "Stamping Instruction" in master_inst_lookup.columns:
        df["StampInstruction"] = (
            df["norm_style_code"]
            .map(master_inst_lookup["Stamping Instruction"])
            .fillna(df["StampInstruction"] if "StampInstruction" in df.columns else "")
        )
    else:
        df["StampInstruction"] = ""

    df["OrderQty"] = df[qty_col]
    df["Metal"] = df["Base Metal"] if "Base Metal" in df.columns else ""
    df["Tone"] = df["COLOR"] if "COLOR" in df.columns else ""
    df["ItemPoNo"] = item_po_no

    for col in [
        "ItemSize",
        "CustomerProductionInstruction",
        "SpecialRemarks",
        "DesignProductionInstruction",
    ]:
        if col not in df.columns:
            df[col] = ""

    df["SrNo"] = range(1, len(df) + 1)

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


def process_omj_file(
    input_filepath,
    output_folder=None,
    master_filepath=None,
    item_po_no="",
):
    """
    Process a single OMJ CASTING PO Excel file and convert to standardized format.

    Parameters:
        input_filepath (str): Path to uploaded input Excel file
        output_folder (str): Folder for output Excel file (optional)
        master_filepath (str): Path to style master file (defaults to CS_100826/OMJ_CS_1408.xlsx)
        item_po_no (str): Item PO number supplied from the frontend

    Returns:
        tuple: (success_status, output_path, error_message, dataframe)
    """
    try:
        if not master_filepath:
            master_filepath = _get_omj_master_path()

        if not os.path.exists(master_filepath):
            return (
                False,
                None,
                f"Master file not found: {master_filepath}",
                None,
            )

        df = process_jewelry_files(
            input_filepath=input_filepath,
            master_filepath=master_filepath,
            item_po_no=item_po_no,
        )

        if df.empty:
            return (
                False,
                None,
                "No rows matched the style master. Check style, color, size, and SM suffix.",
                None,
            )

        input_filename = Path(input_filepath).stem
        if output_folder:
            output_path = os.path.join(
                output_folder, f"OMJ_FORMAT_{input_filename}.xlsx"
            )
        else:
            output_path = f"OMJ_FORMAT_{input_filename}.xlsx"

        df.to_excel(output_path, index=False)
        return True, output_path, None, df

    except Exception as e:
        return False, None, str(e), None


def process_multiple_files(input_folder, output_folder=None, item_po_no=""):
    """Process all Excel files in a folder."""
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)

    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx"))
    results = []

    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")
        success, output_path, error, df = process_omj_file(
            file_path, output_folder, item_po_no=item_po_no
        )

        results.append(
            {
                "input_file": os.path.basename(file_path),
                "success": success,
                "output_file": os.path.basename(output_path)
                if output_path
                else None,
                "output_path": output_path,
                "error": error,
                "row_count": len(df) if df is not None else 0,
            }
        )

    return results


def main():
    """Main function for command line usage."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Process OMJ CASTING PO Excel files to standardized format"
    )
    parser.add_argument(
        "--input", "-i", required=True, help="Input file or folder path"
    )
    parser.add_argument(
        "--output", "-o", help="Output folder path (optional)"
    )
    parser.add_argument(
        "--item-po-no", "-p", default="", help="Item PO number (optional)"
    )
    parser.add_argument(
        "--batch",
        "-b",
        action="store_true",
        help="Process all files in input folder",
    )

    args = parser.parse_args()

    if args.batch:
        if not os.path.isdir(args.input):
            print("Input path must be a folder when using --batch")
            return

        results = process_multiple_files(
            args.input, args.output, item_po_no=args.item_po_no
        )

        print("\n" + "=" * 50)
        print("BATCH PROCESSING RESULTS")
        print("=" * 50)

        success_count = 0
        for result in results:
            status = "SUCCESS" if result["success"] else "FAILED"
            print(
                f"{status}: {result['input_file']} -> {result['output_file'] or 'N/A'}"
            )
            print(f"   Rows: {result['row_count']}")
            if result["error"]:
                print(f"   Error: {result['error']}")
            print()

            if result["success"]:
                success_count += 1

        print(
            f"Processed: {len(results)} files | "
            f"Successful: {success_count} | "
            f"Failed: {len(results) - success_count}"
        )
    else:
        if not os.path.isfile(args.input):
            print("Input path must be a file when not using --batch")
            return

        success, output_path, error, df = process_omj_file(
            args.input, args.output, item_po_no=args.item_po_no
        )

        if success:
            print(f"SUCCESS: Processed {args.input}")
            print(f"Output: {output_path}")
            print(f"Rows processed: {len(df)}")
            print("\nFirst 5 rows preview:")
            print(df.head())
        else:
            print(f"FAILED: {error}")


if __name__ == "__main__":
    main()