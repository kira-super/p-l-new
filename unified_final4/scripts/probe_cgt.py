"""Inspect vw_RPT_VAL_CGT_DETAILS schema and sample for OEFOF."""
import pyodbc
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()
cur.execute("""
SELECT TOP 1 * FROM CCL.dbo.vw_RPT_VAL_CGT_DETAILS WHERE UPPER(PCODE)='OEFOF'
""")
cols = [c[0] for c in cur.description]
print("Columns:")
for c in cols:
    print(f"  {c}")
print()

cur.execute("""
SELECT TOP 5 * FROM CCL.dbo.vw_RPT_VAL_CGT_DETAILS
WHERE UPPER(PCODE)='OEFOF' ORDER BY 1 DESC
""")
rows = cur.fetchall()
print(f"\nSample {len(rows)} rows:")
for r in rows:
    print()
    for c, v in zip(cols, r):
        print(f"  {c:<25} = {v}")
