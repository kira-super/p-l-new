"""Final check: look at trade-type codes in tTRANS & inspect tDATA_EQUITY for shares-outstanding history."""
import pyodbc
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()

# 1) All distinct T (transaction-type) codes in tTRANS for OEFOF
print("--- tTRANS distinct transaction-type codes (T) for OEFOF ---")
cur.execute("""
    SELECT T, COUNT(*) AS n
    FROM CCL.dbo.tTRANS
    WHERE UPPER(PCODE)='OEFOF'
      AND (D IS NULL OR D <> 'Y')
    GROUP BY T
    ORDER BY n DESC
""")
for r in cur.fetchall():
    print(f"  T='{r[0]}'  count={r[1]}")
print()

# 2) For VPS Securities specifically -- what trade rows exist?
print("--- All tTRANS rows for VN000000VCK5 (VPS Securities) in OEFOF ---")
cur.execute("""
    SELECT t.ID, t.T, t.CDATE, t.UNITS, t.NUPRICEG, t.NSETTLE_AMOUNT,
           t.BCODE, t.D, s.SHORT_NAME
    FROM CCL.dbo.tTRANS t
    LEFT JOIN CCL.dbo.tSecurity s ON s.SCODE = t.SCODE
    WHERE UPPER(t.PCODE)='OEFOF' AND s.ISIN='VN000000VCK5'
    ORDER BY t.CDATE
""")
cols = [c[0] for c in cur.description]
print(" | ".join(cols))
for r in cur.fetchall():
    print(" | ".join(str(v)[:25] for v in r))
print()

# 3) tDATA_EQUITY columns
print("--- CCL.tDATA_EQUITY columns ---")
cur.execute("""
    SELECT COLUMN_NAME, DATA_TYPE
    FROM CCL.INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME='tDATA_EQUITY'
    ORDER BY ORDINAL_POSITION
""")
for r in cur.fetchall():
    print(f"  {r[0]:<32} {r[1]}")

# 4) Search ALL DBs for objects whose NAME starts with 'CA' or contains 'CACTION'/'CAEVT'
print("\n--- Server-wide hunt for CA-event style names (sys.objects) ---")
cur.execute("""
    SELECT 'CCL' AS db, s.name AS sch, o.name AS obj, o.type_desc
    FROM CCL.sys.objects o JOIN CCL.sys.schemas s ON s.schema_id = o.schema_id
    WHERE o.type IN ('U','V')
      AND (o.name LIKE 'CA%' OR o.name LIKE '%CACTION%'
           OR o.name LIKE '%CAEVT%' OR o.name LIKE '%CAEVENT%'
           OR o.name LIKE 'tBONUS%' OR o.name LIKE 'tSPLIT%'
           OR o.name LIKE 'tRIGHTS%' OR o.name LIKE 'tCA%'
           OR o.name LIKE 'tRG%' OR o.name LIKE '%STOCK_DIV%'
           OR o.name LIKE '%STKDV%')
    ORDER BY o.name
""")
for r in cur.fetchall():
    print(f"  {r[0]}.{r[1]}.{r[2]}  ({r[3]})")
