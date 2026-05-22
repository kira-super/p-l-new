"""Inspect tH_VAL columns and sample data for OEFOF end_vdate."""
import pyodbc
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()

cur.execute("""
SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME='tH_VAL' ORDER BY ORDINAL_POSITION
""")
print("tH_VAL columns:")
for r in cur.fetchall():
    print(f"  {r[0]:<25} {r[1]}")

# Find latest VDATE for any OEFOF PCODE
cur.execute("""
SELECT TOP 1 PCODE, VDATE FROM CCL.dbo.tH_VAL
WHERE UPPER(PCODE) LIKE 'OEFOF%' ORDER BY VDATE DESC
""")
r = cur.fetchone()
print(f"\nMost recent OEFOF row in tH_VAL: PCODE={r[0]}, VDATE={r[1]}")
latest_pcode, latest_vdate = r

# Sample 5 rows for that VDATE
cur.execute(f"""
SELECT TOP 5 * FROM CCL.dbo.tH_VAL
WHERE PCODE='{latest_pcode}' AND VDATE='{latest_vdate}' ORDER BY UNITS DESC
""")
cols = [c[0] for c in cur.description]
rows = cur.fetchall()
print(f"\nSample 5 rows for {latest_pcode}@{latest_vdate}:")
for r in rows:
    print()
    for c, v in zip(cols, r):
        print(f"  {c:<25} = {v}")
