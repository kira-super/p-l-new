import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pyodbc
from oefof_pl.config import load_default_config, TARGET_PNAME

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()

cur.execute("SELECT DISTINCT PNAME FROM dbo.vw_RPT_VAL WHERE PNAME LIKE '%OAKS%' OR PNAME LIKE '%OEFOF%' OR PNAME LIKE '%Emerging%' ORDER BY PNAME")
print("PNAMES matching OAKS/OEFOF/Emerging:")
for r in cur.fetchall():
    print(f"  {r[0]!r}")

print()
print("TARGET_PNAME =", repr(TARGET_PNAME))

cur.execute("SELECT COUNT(*) FROM dbo.vw_RPT_VAL WHERE PNAME = ? AND SORT1 = 'A'", TARGET_PNAME)
print("Rows for TARGET, SORT1=A:", cur.fetchone()[0])

cur.execute("SELECT TOP 5 PNAME, SORT1, SORT2, SNAME, FMCID FROM dbo.vw_RPT_VAL WHERE PNAME = ?", TARGET_PNAME)
print("Sample rows for TARGET:")
for r in cur.fetchall():
    print(" ", r)
