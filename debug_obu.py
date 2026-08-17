import sys
sys.path.insert(0, 'd:/latest/OPS-TOOL-V2')
import obu, re

pdf = 'd:/latest/obu/Purchase order-7362.pdf'
ok, path, err, df = obu.process_obu_file(pdf, 'd:/latest/obu')
print("ok:", ok, "err:", err)
if df is not None:
    print(df[['StyleCode','ItemRefNo','ItemSize','Metal','Tone','CustomerProductionInstruction']].to_string())
