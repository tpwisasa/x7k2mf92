#!/usr/bin/env python3
"""
Reads Tracker.numbers and rewrites the SNAPSHOTS data in index.html.
Run locally or via GitHub Actions.
"""
import re
from numbers_parser import Document

NUMBERS = "Tracker.numbers"
HTML = "index.html"

MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

def n(v):
    try:
        val = float(v or 0)
        return int(round(val)) if val == val else 0
    except (TypeError, ValueError):
        return 0

doc = Document(NUMBERS)
table = doc.sheets[0].tables[0]

# Collect data rows (skip 3 header rows, stop at empty YRM)
rows = []
for row in table.iter_rows(min_row=3):
    vals = [c.value for c in row]
    if not vals[0] or not str(vals[0]).strip():
        break
    rows.append(vals)

# Build SNAPSHOTS lines
# Excel cols (0-indexed): 0=YRM, 1=UOB, 2=KBANK, 3=KBANKTravel, 4=CIMB,
#   5=SCB, 6=SCBHome, 7=BAY_cash, 8=BAY_stock, 9=TISCO,
#   10=Jitta_VN, 11=Jitta_World, 12=Jitta_Thematic, 13=Crypto,
#   14=Tax_UOB, 15=Tax_Inno, 16=NonTax, 17=SharePlan, 18=LTI,
#   19=Prov_Pao, 20=Prov_Toon, 21=Gold_Physical,
#   22=Sum_Gold, 23=Sum_Cash, 24=Sum_Investment, 25=Sum_Total
snap_lines = []
for r in rows:
    vals = [
        n(r[0]),   # YRM
        n(r[25]),  # Total
        n(r[23]),  # Cash (summary)
        n(r[24]),  # Investment (summary)
        n(r[1]),   # UOB
        n(r[2]),   # KBANK
        n(r[3]),   # KBANK Travel
        n(r[4]),   # CIMB
        n(r[5]),   # SCB
        n(r[6]),   # SCB Home
        n(r[7]),   # BAY cash
        n(r[8]),   # Thai BAY
        n(r[9]),   # Thai TISCO
        n(r[10]),  # Jitta Vietnam
        n(r[11]),  # Jitta World
        n(r[12]),  # Jitta Thematic
        n(r[13]),  # Crypto
        n(r[14]),  # Tax UOB
        n(r[15]),  # Tax InnovestX
        n(r[16]),  # Non-Tax InnovestX
        n(r[17]),  # Company Stock Share Plan
        n(r[18]),  # Company Stock LTI
        n(r[19]),  # Provident Fund Pao
        n(r[20]),  # Provident Fund Toon
        n(r[21]),  # Gold Physical
    ]
    snap_lines.append("  [" + ", ".join(str(v) for v in vals) + "],")

snapshots_js = "const SNAPSHOTS = [\n" + "\n".join(snap_lines) + "\n];"

# Date label from latest YRM
yrm = str(n(rows[-1][0]))
mon = MONTHS[int(yrm[4:6]) - 1]
yr  = yrm[:4]
date_label = f"{mon} {yr}"

# Read HTML
with open(HTML, "r", encoding="utf-8") as f:
    html = f.read()

# Replace SNAPSHOTS block
html = re.sub(
    r'const SNAPSHOTS = \[[\s\S]*?\];',
    snapshots_js,
    html, count=1
)

# Replace "Last updated" text
html = re.sub(
    r'Last updated: [A-Za-z]+ \d{4}',
    f'Last updated: {date_label}',
    html
)

with open(HTML, "w", encoding="utf-8") as f:
    f.write(html)

print(f"Done — updated index.html with {len(rows)} rows, latest: {date_label}")
