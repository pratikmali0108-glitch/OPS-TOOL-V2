import pandas as pd

# Clean functions
def clean_style(series):
    return series.astype(str).str.strip().str.upper()

def clean_size(series):
    return (
        series.astype(str)
        .str.strip()
        .str.replace(r"\.0+$", "", regex=True)
        .str.lower()
    )

def clean_color(series):
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

# Read input file
input_df = pd.read_excel(r'd:\latest\OPS-TOOL-V2\PO-LEA8142026.xlsx', header=3)
input_df.columns = input_df.columns.astype(str).str.strip()

# Get the YR4171X-GD row
test_row = input_df[input_df['Elegant Jewelry Style #'].str.contains('YR4171X', na=False)].iloc[0]

print("===== INPUT ROW =====")
print(f"Style: {test_row['Elegant Jewelry Style #']}")
print(f"Size: {test_row['Size']}")
print(f"Metal Color: {test_row['Metal Color']}")

# Extract base style (remove -GD suffix)
input_style = str(test_row['Elegant Jewelry Style #']).strip().upper()
base_style = input_style.replace('-GD', '')
has_gd = '-GD' in input_style

print(f"\nBase style: {base_style}")
print(f"Has -GD suffix: {has_gd}")

# Apply cleaning
clean_input_style = base_style
clean_input_size = clean_size(pd.Series([test_row['Size']])).iloc[0]
clean_input_color = clean_color(pd.Series([test_row['Metal Color']])).iloc[0]

print(f"\nCleaned input style: {clean_input_style}")
print(f"Cleaned input size: {clean_input_size}")
print(f"Cleaned input color: {clean_input_color}")

# Read master file
master_df = pd.read_excel(r'd:\latest\OPS-TOOL-V2\CS_100826\OMJ_CS_1408.xlsx', header=0)
master_df.columns = master_df.columns.astype(str).str.strip()

# Apply cleaning to master
master_df['match_style'] = clean_style(master_df['Style No'])
master_df['match_size'] = clean_size(master_df['SIZE'])
master_df['match_color'] = clean_color(master_df['COLOR'])

print("\n===== MASTER FILE ENTRIES FOR YR4171X =====")
yr4171x_entries = master_df[master_df['match_style'] == 'YR4171X']
print(yr4171x_entries[['Style No', 'COLOR', 'SIZE', 'match_style', 'match_color', 'match_size', 'Client Style No']].to_string())

print("\n===== MATCHING LOGIC =====")
print(f"Looking for: style={clean_input_style}, color={clean_input_color}, size={clean_input_size}")

# Try to match
match = master_df[
    (master_df['match_style'] == clean_input_style) &
    (master_df['match_color'] == clean_input_color) &
    (master_df['match_size'] == clean_input_size)
]

if not match.empty:
    print(f"\n✓ MATCH FOUND!")
    print(match[['Style No', 'COLOR', 'SIZE', 'Client Style No']].to_string())
    client_style = match.iloc[0]['Client Style No']
    if has_gd:
        final_style = client_style + 'GD'
    else:
        final_style = client_style
    print(f"\nFinal output style: {final_style}")
else:
    print("\n✗ NO MATCH FOUND")
