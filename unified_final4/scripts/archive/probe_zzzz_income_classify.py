"""For each I+ZZZZ row, compare UNITS to prior-day holding to determine if it's
a cash dividend (UNITS = holding, price = DPS) or a stock dividend (UNITS = bonus shares)."""
import pyodbc
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()

# Pull all OEFOF trades for ISINs that have any I+ZZZZ since 2025-12-31
cur.execute("""
SELECT DISTINCT s.ISIN
FROM CCL.dbo.tTRANS t LEFT JOIN CCL.dbo.tSecurity s ON s.SCODE=t.SCODE
WHERE UPPER(t.PCODE)='OEFOF' AND (t.D IS NULL OR t.D<>'Y')
  AND UPPER(t.BCODE)='ZZZZ' AND t.T='I' AND t.CDATE>='2025-12-31'
""")
isins = [r[0] for r in cur.fetchall() if r[0]]

results = {"cash_div_match": 0, "no_match": 0, "no_holding": 0}
detail = []
for isin in isins:
    cur.execute("""SELECT t.CDATE,t.T,t.BCODE,t.UNITS,t.NUPRICEG,t.NINCOME,s.SHORT_NAME
                   FROM CCL.dbo.tTRANS t LEFT JOIN CCL.dbo.tSecurity s ON s.SCODE=t.SCODE
                   WHERE UPPER(t.PCODE)='OEFOF' AND (t.D IS NULL OR t.D<>'Y')
                   AND s.ISIN=? AND t.CDATE<='2026-05-15' ORDER BY t.CDATE""", isin)
    rows = cur.fetchall()
    pos = 0.0
    sname = ""
    for r in rows:
        sname = r[6] or sname
        if r[1] == 'I' and (r[2] or '').upper() == 'ZZZZ':
            u = float(r[3])
            if r[0].year >= 2026 or (r[0].year == 2025 and r[0].month == 12 and r[0].day >= 31):
                if pos == 0:
                    cat = "no_holding"
                elif abs(pos - u) < max(1.0, abs(pos) * 0.001):
                    cat = "cash_div_match"
                else:
                    cat = "no_match"
                results[cat] += 1
                detail.append((isin, sname[:25], r[0], u, float(r[4]), float(r[5]), pos, cat))
        else:
            sgn = 1 if r[1] == 'P' else (-1 if r[1] == 'S' else 0)
            pos += sgn * float(r[3])

print(f"Summary across {len(isins)} ISINs with I+ZZZZ rows:")
for k, v in results.items():
    print(f"  {k:<20} {v}")
print()
print(f"{'ISIN':<14} {'SNAME':<25} {'DATE':<11} {'UNITS':>15} {'PX':>10} {'INCOME':>18} {'PRIOR_HOLD':>15} {'CAT':<15}")
for d in detail:
    print(f"{d[0]:<14} {d[1]:<25} {d[2]:%Y-%m-%d}  {d[3]:>15,.0f} {d[4]:>10.4f} {d[5]:>18,.2f} {d[6]:>15,.0f} {d[7]:<15}")
