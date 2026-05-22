"""List STAMP/TAX columns on tTRANS."""
import pyodbc
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()
cur.execute("""
SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME='tTRANS' AND (COLUMN_NAME LIKE '%STAMP%' OR COLUMN_NAME LIKE '%TAX%' OR COLUMN_NAME LIKE '%FEE%' OR COLUMN_NAME LIKE '%EXP%' OR COLUMN_NAME LIKE '%COMM%')
ORDER BY ORDINAL_POSITION
""")
for r in cur.fetchall():
    print(r[0])
print("\nNon-zero stamp/tax in OEFOF current period:")
# Try common candidates
for candidate in ("STAMP_DUTY", "STAMP", "NSTAMP", "STAMPDUTY"):
    try:
        cur.execute(f"SELECT COUNT(*), SUM(ISNULL({candidate},0)) FROM CCL.dbo.tTRANS WHERE UPPER(PCODE)='OEFOF' AND CDATE>='2025-12-31' AND ISNULL({candidate},0)<>0")
        n, s = cur.fetchone()
        print(f"  {candidate}: rows={n} total={s}")
    except pyodbc.Error as e:
        print(f"  {candidate}: NOT A COLUMN")
