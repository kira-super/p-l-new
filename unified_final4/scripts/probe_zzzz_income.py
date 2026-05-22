"""Quantify ZZZZ-stripped income for OEFOF in the current period."""
import pyodbc
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()

cur.execute("""
SELECT s.ISIN, s.SHORT_NAME, s.CCY, t.CDATE, t.UNITS, t.NUPRICEG,
       t.NSETTLE_AMOUNT, t.NINCOME, t.NTAX, t.BCODE, t.T
FROM CCL.dbo.tTRANS t LEFT JOIN CCL.dbo.tSecurity s ON s.SCODE=t.SCODE
WHERE UPPER(t.PCODE)='OEFOF' AND (t.D IS NULL OR t.D<>'Y')
  AND UPPER(t.BCODE)='ZZZZ' AND t.T='I' AND t.CDATE>='2025-12-31'
ORDER BY t.CDATE
""")
rows = cur.fetchall()
print(f"{len(rows)} ZZZZ income rows since 2025-12-31:")
print(f"{'ISIN':<14} {'SNAME':<25} {'CCY':<4} {'DATE':<11} {'UNITS':>15} {'PRICE':>10} {'SETTLE':>18} {'INCOME':>18}")
for r in rows:
    print(f"{r[0] or '':<14} {(r[1] or '')[:25]:<25} {r[2] or '':<4} {r[3]:%Y-%m-%d}  {float(r[4]):>15,.0f} {float(r[5]):>10.4f} {float(r[6]):>18,.2f} {float(r[7]):>18,.2f}")
print()

# Compare with what compute.pl currently sees (post-strip) vs included:
# These rows currently flow into INCOME_LOCAL via t.NINCOME ... but our SQL loader
# selects them only for T='I' implicitly (NINCOME is set on income lines).
# After compute.pl strips BCODE='ZZZZ', they vanish.
print("\nIf we keep these income rows, the additional INCOME_LOCAL per ISIN is:")
cur.execute("""
SELECT s.ISIN, s.SHORT_NAME, s.CCY, SUM(t.NINCOME) AS income_total
FROM CCL.dbo.tTRANS t LEFT JOIN CCL.dbo.tSecurity s ON s.SCODE=t.SCODE
WHERE UPPER(t.PCODE)='OEFOF' AND (t.D IS NULL OR t.D<>'Y')
  AND UPPER(t.BCODE)='ZZZZ' AND t.T='I' AND t.CDATE>='2025-12-31'
GROUP BY s.ISIN, s.SHORT_NAME, s.CCY
ORDER BY ABS(SUM(t.NINCOME)) DESC
""")
for r in cur.fetchall():
    print(f"  {r[0] or '':<14} {(r[1] or '')[:30]:<30} {r[2] or '':<4} {float(r[3]):>20,.2f}")
