import pandas as pd
import os

cs_dir = 'CS_220526'
if os.path.exists(cs_dir):
    files = [f for f in os.listdir(cs_dir) if f.endswith('CS.xlsx') and not f.startswith('~')]
    print('CS files found:')
    print(files)
    print()
    for f in files:
        try:
            df = pd.read_excel(os.path.join(cs_dir, f))
            print(f'{f}:')
            print(f'  Columns: {df.columns.tolist()}')
            if not df.empty:
                print(f'  First row: {df.head(1).to_dict("records")}')
            print()
        except Exception as e:
            print(f'Error reading {f}: {e}')
            print()
else:
    print(f'Directory {cs_dir} not found!')
