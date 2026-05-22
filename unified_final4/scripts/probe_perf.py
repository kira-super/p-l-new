"""Inspect vw_RPT_VAL_PERF_CURRENT for realised P&L cross-check."""
import pyodbc
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()
cur.execute("SELECT TOP 1 * FROM CCL.dbo.vw_RPT_VAL_PERF_CURRENT WHERE UPPER(PCODE)='OEFOF'")
cols = [c[0] for c in cur.description]
print("Columns:")
for c in cols:
    print(f"  {c}")

# Sample top by abs PERF
try:
    cur.execute("""SELECT TOP 5 PCODE, SCODE, SNAME, CCY, UNITS, PERF_PCCY, PERF_SCCCY, PERF_CRYST_PCCY
                   FROM CCL.dbo.vw_RPT_VAL_PERF_CURRENT
                   WHERE UPPER(PCODE)='OEFOF' ORDER BY ABS(PERF_PCCY) DESC""")
    print("\nTop 5 by |PERF_PCCY|:")
    for r in cur.fetchall():
        print(f"  {r[1]:<10} {(r[2] or '')[:25]:<25} {r[3]:<4} units={float(r[4]):>15,.0f} PERF={float(r[5]):>15,.2f} CRYST={float(r[7] or 0):>15,.2f}")
except Exception as e:
    print(f"\nQuery failed: {e}")

# Look for historical equivalent
print("\n\nHistorical equivalent? vw_RPT_VAL_PERF_HIST:")
try:
    cur.execute("""SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
                   WHERE TABLE_NAME='vw_RPT_VAL_PERF_HIST' ORDER BY ORDINAL_POSITION""")
    for r in cur.fetchall():
        print(f"  {r[0]}")
except Exception as e:
    print(f"  {e}")

# Check tH_PERF table
print("\nLooking for tH_PERF / tH_VAL_PERF...")
cur.execute("""SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES
               WHERE TABLE_NAME LIKE '%PERF%' OR TABLE_NAME LIKE '%CRYST%' OR TABLE_NAME LIKE '%REAL%'""")
for r in cur.fetchall():
    print(f"  {r[0]}")
