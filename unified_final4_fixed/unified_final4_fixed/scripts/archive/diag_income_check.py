"""Cross-check income & dividends per analyst against the raw trade tape."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd
from openpyxl import load_workbook

WB = r"C:\Users\kgontar\v2\oefof_pl\out\OAKS_EM_Stock_PL_Report.xlsx"
wb = load_workbook(WB, read_only=True, data_only=True)

# 1) Pull per-analyst totals from the Summary sheet
print("=== Summary sheet — By Analyst ===")
ws = wb["Summary"]
hdr_row = None
for r_i, row in enumerate(ws.iter_rows(values_only=True), start=1):
    if row and row[0] == "Analyst" and row[1] == "Positions":
        hdr_row = r_i
        cols = list(row)
        print(f"  hdr@{r_i}: {[c for c in cols if c]}")
        break
if hdr_row:
    idx = {h: i for i, h in enumerate(cols) if h}
    for r in ws.iter_rows(min_row=hdr_row+1, values_only=True):
        if not r or not r[0]: continue
        if "TOTAL" in str(r[0]).upper(): break
        print(f"  {r[0]:<10} pos={r[idx['Positions']]:>4} "
              f"Realised={r[idx.get('Realised (EUR)', -1)] or 0:>14,.0f} "
              f"Unrealised={r[idx.get('Unrealised (EUR)', -1)] or 0:>14,.0f} "
              f"Income&Fin={r[idx.get('Income & Financing (EUR)', -1)] or 0:>14,.0f} "
              f"Total={r[idx['Total P&L (EUR)']] or 0:>14,.0f}")

# 2) Pull from Income Detail per analyst
print("\n=== Income Detail by analyst ===")
ws = wb["Income Detail"]
hdr_row = None
for r_i, row in enumerate(ws.iter_rows(values_only=True), start=1):
    if row and row[0] == "ISIN":
        hdr_row = r_i
        cols = list(row)
        break
print(f"  hdr cols: {[c for c in cols if c]}")
idx = {h: i for i, h in enumerate(cols) if h}
data = []
for r in ws.iter_rows(min_row=hdr_row+1, values_only=True):
    if not r or not r[0]: continue
    if str(r[0]).startswith("Total"): continue
    data.append(r)
df = pd.DataFrame(data, columns=cols)
df = df.loc[:, df.columns.notna()]
print(f"  rows: {len(df)}")
ana_col = next((c for c in df.columns if "Analyst" in str(c)), None)
type_col = next((c for c in df.columns if "Type" in str(c)), None)
inc_col = next((c for c in df.columns if "Income" in str(c) and "EUR" in str(c)), None)
print(f"  using: analyst={ana_col!r} type={type_col!r} income={inc_col!r}")
if ana_col and inc_col:
    df[inc_col] = pd.to_numeric(df[inc_col], errors="coerce").fillna(0)
    g = df.groupby(ana_col)[inc_col].sum().sort_index()
    print("  Income & Financing by analyst:")
    for a, v in g.items():
        print(f"    {a!s:<10} {v:>14,.2f}")
    if type_col:
        df[type_col] = df[type_col].astype(str).str.upper()
        gt = df.groupby([ana_col, type_col])[inc_col].sum().unstack(fill_value=0)
        print("\n  By type:")
        print(gt.round(2).to_string())

# 3) Sum SB-specific from the SB analyst sheet itself
print("\n=== SB analyst sheet ===")
ws = wb["SB"]
for r_i, row in enumerate(ws.iter_rows(values_only=True), start=1):
    if r_i <= 8:
        print(f"  row {r_i}: {row}")
