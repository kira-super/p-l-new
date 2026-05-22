"""Inspect tSHRS_ADJ - candidate CA / shares-adjustment table."""
import pyodbc
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()

cur.execute(
    "SELECT COLUMN_NAME, DATA_TYPE FROM CCL.INFORMATION_SCHEMA.COLUMNS "
    "WHERE TABLE_NAME='tSHRS_ADJ' ORDER BY ORDINAL_POSITION"
)
print("tSHRS_ADJ columns:")
for r in cur.fetchall():
    print(f"  {r[0]:<30} {r[1]}")
print()

cur.execute("SELECT COUNT(*) FROM CCL.dbo.tSHRS_ADJ")
print("Row count:", cur.fetchone()[0])
print()

cur.execute("SELECT TOP 15 * FROM CCL.dbo.tSHRS_ADJ")
cols = [c[0] for c in cur.description]
print(" | ".join(cols))
for r in cur.fetchall():
    print(" | ".join(str(v) for v in r))
print()

# Look for our VPS Securities ISIN specifically (joined to tSecurity)
print("Looking up VPS Securities (VN000000VCK5):")
cur.execute("""
    SELECT a.*, s.ISIN, s.SHORT_NAME
    FROM CCL.dbo.tSHRS_ADJ a
    LEFT JOIN CCL.dbo.tSecurity s ON s.SCODE = a.SCODE
    WHERE s.ISIN = ?
    ORDER BY 1 DESC
""", "VN000000VCK5")
cols = [c[0] for c in cur.description]
print(" | ".join(cols))
rows = cur.fetchall()
for r in rows:
    print(" | ".join(str(v) for v in r))
print(f"({len(rows)} rows)")
