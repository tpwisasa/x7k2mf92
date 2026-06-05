#!/usr/bin/env python3
"""
Reads Tracker.xlsx and rewrites the data section of index.html.
Run locally or via GitHub Actions.
"""
import re
import openpyxl

XLSX = "Tracker.xlsx"
HTML = "index.html"

MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
ALLOC_COLORS = {
    'Cash':            '#888780',
    'Company Stock':   '#c8a96e',
    'Tax Saving Fund': '#7eb89a',
    'Jitta Funds':     '#e0c07a',
    'Thai Stocks':     '#a0c4b8',
    'Provident Fund':  '#b09070',
    'Gold':            '#d4b04a',
    'Non-Tax Saving':  '#6a8f80',
    'Crypto':          '#e07b5a',
}

def n(v):
    try:
        val = float(v or 0)
        return int(round(val)) if val == val else 0  # handle NaN
    except (TypeError, ValueError):
        return 0

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.active

# Collect data rows (skip 3 header rows, stop at empty YRM)
rows = []
for row in ws.iter_rows(min_row=4, values_only=True):
    if not row[0] or not str(row[0]).strip():
        break
    rows.append(row)

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

# Latest snapshot for ALLOC
s = [
    n(rows[-1][0]),   # YRM
    n(rows[-1][25]),  # Total
    n(rows[-1][23]),  # Cash
    n(rows[-1][24]),  # Investment
    n(rows[-1][1]),   # UOB
    n(rows[-1][2]),   # KBANK
    n(rows[-1][3]),   # KBANKTravel
    n(rows[-1][4]),   # CIMB
    n(rows[-1][5]),   # SCB
    n(rows[-1][6]),   # SCBHome
    n(rows[-1][7]),   # BAY_cash
    n(rows[-1][8]),   # BAY_stock
    n(rows[-1][9]),   # TISCO
    n(rows[-1][10]),  # Jitta_VN
    n(rows[-1][11]),  # Jitta_World
    n(rows[-1][12]),  # Jitta_Thematic
    n(rows[-1][13]),  # Crypto
    n(rows[-1][14]),  # Tax_UOB
    n(rows[-1][15]),  # Tax_Inno
    n(rows[-1][16]),  # NonTax
    n(rows[-1][17]),  # SharePlan
    n(rows[-1][18]),  # LTI
    n(rows[-1][19]),  # Prov_Pao
    n(rows[-1][20]),  # Prov_Toon
    n(rows[-1][21]),  # Gold_Physical
]

alloc = [
    ('Cash',            s[2]),
    ('Company Stock',   s[20] + s[21]),
    ('Tax Saving Fund', s[17] + s[18]),
    ('Jitta Funds',     s[13] + s[14] + s[15]),
    ('Thai Stocks',     s[11] + s[12]),
    ('Provident Fund',  s[22] + s[23]),
    ('Gold',            s[24]),
    ('Non-Tax Saving',  s[19]),
    ('Crypto',          s[16]),
]

alloc_lines = []
for name, val in alloc:
    color = ALLOC_COLORS[name]
    alloc_lines.append(f"  {{ name: '{name}', color: '{color}', val: {val} }},")

alloc_js = "const ALLOC_CATEGORIES = [\n" + "\n".join(alloc_lines) + "\n];"

# Date label from latest YRM
yrm = str(s[0])
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

# Replace ALLOC_CATEGORIES block
html = re.sub(
    r'const ALLOC_CATEGORIES = \[[\s\S]*?\];',
    alloc_js,
    html, count=1
)

# Replace "Last updated" text
html = re.sub(
    r'Last updated: [A-Za-z]+ \d{4}',
    f'Last updated: {date_label}',
    html
)

# Replace allocation chart title
html = re.sub(
    r'Asset Allocation — [A-Za-z]+ \d{4}',
    f'Asset Allocation — {date_label}',
    html
)

with open(HTML, "w", encoding="utf-8") as f:
    f.write(html)

print(f"Done — updated index.html with {len(rows)} rows, latest: {date_label}")
