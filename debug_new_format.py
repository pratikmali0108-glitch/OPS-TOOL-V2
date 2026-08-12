import os
import sys
sys.path.insert(0, r'd:\latest\OPS-TOOL-V2')

import re

TEST_TEXT = """Order To
SHIMAYRA JEWELLERY
PLOT NO: 62, SEEPZ ANDHERI (E)
MUMBAI MH 400096 INDIA
NEW YORK NY 10036
50 WEST, 47 STREET, SUITE 2100
DHARM INTERNATIONAL LLC
Ship To
57133 / US981361
8/5/2026
8/22/2026
P.O. #:
Ship Via:
Date:
Due Date:
Vendor #:
0095S
912269629949
Phone #:
#
Item #
Quantity
Vendor Item #
Cost
Amount
Cancel Date:
8/22/2026
Reference:
Weight
Unit
Job Bag #
Size
Description
Memo #
1
108
$0.00
$0.00
SKU:1507287 14KW
1.40CTW I VS Round
Diamond Studs w/
Gurdian Back (Only for
labor)
0.0000
Q
ER140-14KW-SEMI
108
0.0000
All findings and diamond will be supplied by Dharm, and complete
production is required from Shimayra.
Purchase Order
DHARM INTERNATIONAL LLC
50 WEST, 47 STREET, SUITE 2100
NEW YORK NY 10036
(212) 398-7777  Fax: (212) 398-7775
WWW.DDDPL.COM
1 of 1
Order #:
5981
Page #:
Grand Total:
$0.00
20230926.151356[s]; 20230926.151356[s]
RightClick® Copyright © 2026 CFI/Wise Choice Software® All Rights Reserved
"""

# Let's find the combined header line by joining nearby lines:
lines = [l.strip() for l in TEST_TEXT.split("\n")]

print("=" * 80)
print("CHECKING HEADER DETECTION (MULTI-LINE)")
print("=" * 80)
for idx in range(len(lines)):
    window = lines[idx:min(idx+15, len(lines))]
    combined = " ".join([w for w in window if w])
    has_memo = bool(re.search(r"Memo", combined, re.IGNORECASE))
    has_item = bool(re.search(r"Item\s*#", combined, re.IGNORECASE))
    has_desc = bool(re.search(r"Description", combined, re.IGNORECASE))
    has_size = bool(re.search(r"\bSize\b", combined, re.IGNORECASE))
    has_qty  = bool(re.search(r"Quantity", combined, re.IGNORECASE))
    score = sum([has_memo,has_item,has_desc,has_size,has_qty])
    if score >= 3:
        print(f"  idx={idx} score={score}: {combined[:150]}")

# Show raw tokens line by line after header area
print("\n" + "=" * 80)
print("KEY LINES with tokens (idx 30-55):")
print("=" * 80)
for idx in range(30, min(60, len(lines))):
    line = lines[idx]
    if line:
        print(f"  [{idx:3d}] {repr(line)}")

# Style code check
print("\n" + "=" * 80)
print("STYLE CODE DETECTION TESTS:")
print("=" * 80)

style_patterns = [
    ("OLD strict 4+ digits", r"[A-Z]{2}\d{4,}[A-Z0-9]*"),
    ("NEW flexible 2+ letters any digits", r"[A-Z]{2,}\d+[A-Z0-9\-]*"),
]
test_styles = ["ER140-14KW-SEMI", "RG0003054K", "SKU:1507287 14KW", "ER1000"]
for pname, pat in style_patterns:
    print(f"\n  Pattern: {pname}")
    for ts in test_styles:
        m = re.search(pat, ts)
        print(f"    {ts:30s} -> {m.group() if m else None}")

# Item line detection
print("\n" + "=" * 80)
print("ITEM LINE DETECTION (line starts with number + qty + money):")
print("=" * 80)

test_lines = [
    "1 108 $0.00 $0.00",
    "2 50 $123.00 $123.00",
    "1 SZ5 RG0003054K CST 81RD=2.00CTW 5 25 0.0000 Q $0.00 $0.00",
    "ER140-14KW-SEMI 108 0.0000",
    "SKU:1507287 14KW",
]
from BDLDHI import is_item_row
for tl in test_lines:
    print(f"  is_item_row('{tl[:60]}') = {is_item_row(tl)}")

# Show metal detection test
print("\n" + "=" * 80)
print("METAL DETECTION:")
print("=" * 80)

from BDLDHI import parse_metal_from_description

test_descs = [
    "SKU:1507287 14KW 1.40CTW I VS Round Diamond Studs",
    "SKU#1922398 PT CST 81RD=2.00CTW",
    "ER140-14KW-SEMI",
    "SKU:1507287 14KW 1.40CTW I VS Round Diamond Studs w/ Gurdian Back (Only for labor) ER140-14KW-SEMI 108 0.0000",
]
for td in test_descs:
    mi = parse_metal_from_description(td)
    print(f"  desc: {td[:60]}...")
    print(f"    -> tone={mi['tone_suffix']:4s} metal={mi['metal']:6s} tone_col={mi['tone']} name={mi['metal_name']}")
    print()
