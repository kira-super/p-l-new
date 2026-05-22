"""Verify whether SB/KX/HK genuinely had zero income in the period by
sampling raw trades for each position with T='I'."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd
from openpyxl import load_workbook
from oefof_pl.config import load_default_config

cfg = load_default_config()

# Load the parquet of computed positions to get analyst per ISIN
pl = pd.read_parquet(r"C:\Users\kgontar\v2\oefof_pl\out\pl_results.parquet")
print("Analysts in pl_df:", sorted(pl["Analyst"].dropna().unique().tolist()))
isin_to_analyst = dict(zip(pl["ISIN"].astype(str), pl["Analyst"].astype(str).str.upper()))

# Pull all income trades from SQL for the period
import pyodbc
cn = pyodbc.connect(cfg.sql_conn_str, timeout=60)
sql = """
SELECT t.PCODE, s.ISIN, s.SHORT_NAME AS SNAME, s.CCY, t.T, t.CDATE,
       t.UNITS, t.NSETTLE_AMOUNT, t.NINCOME, t.BCODE
FROM ccl.dbo.tTRANS t
LEFT JOIN ccl.dbo.tSecurity s ON s.SCODE = t.SCODE
WHERE UPPER(t.PCODE) IN ('OEFOF','OEFOGSSC','OEFOGSSW','OEFOHFSW','OEFOHFSC')
  AND t.T = 'I'
  AND (t.D IS NULL OR t.D <> 'Y')
"""
df = pd.read_sql(sql, cn)
# Filter to the actual period the pipeline ran on
df["CDATE"] = pd.to_datetime(df["CDATE"])
PERIOD_START = pd.Timestamp("2025-12-31")
PERIOD_END = pd.Timestamp("2026-05-04")
print(f"Total income trades (all time): {len(df)}")
df = df[(df["CDATE"] >= PERIOD_START) & (df["CDATE"] <= PERIOD_END)]
print(f"Income trades in period [{PERIOD_START.date()} -> {PERIOD_END.date()}]: {len(df)}")
print(f"\nIncome trades from SQL: {len(df)} rows")
df["analyst"] = df["ISIN"].map(isin_to_analyst).fillna("(unmapped)")
print("\nIncome trades by analyst (count):")
print(df.groupby("analyst").size().to_string())
print("\nIncome trades by analyst (sum NINCOME):")
print(df.groupby(["analyst", "CCY"])["NINCOME"].sum().round(0).to_string())

print("\nSpecifically SB / KX / HK income trades:")
sub = df[df["analyst"].isin(["SB", "KX", "HK"])]
if sub.empty:
    print("  (none — confirms zero income for these analysts in the period)")
else:
    print(sub[["ISIN", "SNAME", "CCY", "CDATE", "NINCOME", "BCODE"]].to_string(index=False))

# Show ISINs held by SB
print("\nSB held ISINs:")
sb = pl[pl["Analyst"].str.upper() == "SB"][["ISIN", "Stock Name", "Income (EUR)", "Total P&L (EUR)"]]
print(sb.to_string(index=False))
