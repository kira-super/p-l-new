"""Audit BCODE='CORP' and BCODE='RESETS' for OEFOF in current period.

CORP: candidate corporate-action placeholder; need to verify whether to
strip like ZZZZ or keep (real broker name).
RESETS: CFD financing/reset entries; expected to net to zero.
"""
import pyodbc
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()

def audit(bcode: str) -> None:
    print(f"\n=== BCODE='{bcode}' ===")
    cur.execute("""
SELECT t.T, COUNT(*), SUM(ISNULL(t.UNITS,0)), SUM(ISNULL(t.NSETTLE_AMOUNT,0)),
       SUM(ISNULL(t.NINCOME,0)),
       SUM(CASE WHEN ISNULL(t.NUPRICEG,0)=0 THEN 1 ELSE 0 END) AS zero_px
FROM CCL.dbo.tTRANS t
WHERE UPPER(t.PCODE)='OEFOF' AND (t.D IS NULL OR t.D<>'Y')
  AND UPPER(t.BCODE)=? AND t.CDATE>='2025-12-31'
GROUP BY t.T
""", bcode)
    rows = cur.fetchall()
    if not rows:
        print(f"  (no rows in current period)")
        return
    print(f"  {'T':<3} {'rows':>6} {'units':>20} {'settle':>20} {'income':>20} {'zero_px':>8}")
    for r in rows:
        print(f"  {r[0]:<3} {r[1]:>6} {float(r[2] or 0):>20,.0f} {float(r[3] or 0):>20,.2f} {float(r[4] or 0):>20,.2f} {r[5]:>8}")

    # Sample 10 rows
    cur.execute("""
SELECT TOP 10 s.ISIN, s.SHORT_NAME, t.CDATE, t.T, t.UNITS, t.NUPRICEG, t.NSETTLE_AMOUNT, t.NINCOME
FROM CCL.dbo.tTRANS t LEFT JOIN CCL.dbo.tSecurity s ON s.SCODE=t.SCODE
WHERE UPPER(t.PCODE)='OEFOF' AND (t.D IS NULL OR t.D<>'Y')
  AND UPPER(t.BCODE)=? AND t.CDATE>='2025-12-31'
ORDER BY t.CDATE
""", bcode)
    print("  Sample:")
    for r in cur.fetchall():
        print(f"    {r[0] or '':<14} {(r[1] or '')[:20]:<20} {r[2]:%Y-%m-%d} {r[3]:<2} u={float(r[4]):>15,.0f} px={float(r[5]):>10.4f} settle={float(r[6]):>18,.2f} inc={float(r[7]):>15,.2f}")

audit("CORP")
audit("RESETS")
