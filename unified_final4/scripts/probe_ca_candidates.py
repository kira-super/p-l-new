"""Check remaining CA candidates: NAV.t_DIV, CCL.tSecurity columns, vw_DF_DVD_OUT."""
import pyodbc
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()

def show_columns(db: str, table: str) -> None:
    print(f"\n--- {db}.{table} columns ---")
    cur.execute(
        f"SELECT COLUMN_NAME, DATA_TYPE FROM {db}.INFORMATION_SCHEMA.COLUMNS "
        f"WHERE TABLE_NAME=? ORDER BY ORDINAL_POSITION", table)
    for r in cur.fetchall():
        print(f"  {r[0]:<32} {r[1]}")

def show_top(db: str, table: str, n: int = 5, where: str = "") -> None:
    print(f"\n--- {db}.dbo.{table} top {n} {where} ---")
    sql = f"SELECT TOP {n} * FROM {db}.dbo.{table} {where}"
    try:
        cur.execute(sql)
        cols = [c[0] for c in cur.description]
        print(" | ".join(cols))
        for r in cur.fetchall():
            print(" | ".join(str(v)[:40] for v in r))
    except Exception as e:
        print(f"  ERR: {e}")

# 1) NAV.t_DIV (probably dividend table)
show_columns("NAV", "t_DIV")
show_top("NAV", "t_DIV", 5)

# 2) CCL.vw_DF_DVD_OUT (had NDIVIDEND column)
show_columns("CCL", "vw_DF_DVD_OUT")
show_top("CCL", "vw_DF_DVD_OUT", 5)

# 3) tSecurity columns - any CA fields?
print("\n--- CCL.tSecurity column names containing CA-keywords ---")
cur.execute("""
    SELECT COLUMN_NAME, DATA_TYPE
    FROM CCL.INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME='tSecurity'
    ORDER BY ORDINAL_POSITION
""")
all_cols = cur.fetchall()
print(f"({len(all_cols)} columns total)")
kws = ("CA", "DIV", "BONUS", "SPLIT", "RIGHT", "ADJ", "EVENT", "SUSP", "FREEZE", "CORP")
for n, t in all_cols:
    if any(k in n.upper() for k in kws):
        print(f"  {n:<32} {t}")

# 4) Look in CCL for any table with both ISIN-like and units/shares/quantity columns
print("\n--- CCL tables with both 'SCODE' and 'UNITS|SHARES|QTY|RATIO' columns ---")
cur.execute("""
    SELECT t1.TABLE_NAME
    FROM CCL.INFORMATION_SCHEMA.COLUMNS t1
    WHERE t1.COLUMN_NAME='SCODE'
      AND EXISTS (
        SELECT 1 FROM CCL.INFORMATION_SCHEMA.COLUMNS t2
        WHERE t2.TABLE_NAME = t1.TABLE_NAME
          AND (t2.COLUMN_NAME LIKE '%UNIT%' OR t2.COLUMN_NAME LIKE '%SHARE%'
               OR t2.COLUMN_NAME LIKE '%QTY%' OR t2.COLUMN_NAME LIKE '%RATIO%'
               OR t2.COLUMN_NAME LIKE '%FACTOR%')
      )
    GROUP BY t1.TABLE_NAME
""")
for r in cur.fetchall():
    print(f"  {r[0]}")
