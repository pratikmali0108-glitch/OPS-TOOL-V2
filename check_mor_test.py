
import pandas as pd
import os

test_path = r"D:\OPS_Tool_01042026\mor_test.xlsx"
print(f"Checking file: {test_path}")
print(f"File exists? {os.path.exists(test_path)}")

# Read without headers to see all raw rows
print("\n=== Reading without headers, all rows (raw) ===")
df_raw = pd.read_excel(test_path, header=None)
print(f"Number of rows: {len(df_raw)}")
print("\nFirst 6 raw rows:")
for i in range(min(6, len(df_raw))):
    print(f"\n--- Row {i} ---")
    print([str(x) for x in df_raw.iloc[i].values])
