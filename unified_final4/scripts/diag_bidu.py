"""Diagnostic: BIDU pair (long BAIDU HK + short BAIDU ADR/CFD)."""
import pyodbc
from oefof_pl.config import load_default_config

PCODES = ('OEFOF', 'OEFOGSSC', 'OEFOGSSW', 'OEFOHFSW', 'OEFOHFSC')

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()

print("=== tSecurity any BIDU/BAIDU rows ===")
cur.execute(
    "SELECT SCODE, SHORT_NAME, ISIN, CCY, CAT, EX FROM dbo.tSecurity "
    "WHERE SHORT_NAME LIKE '%BIDU%' OR SHORT_NAME LIKE '%BAIDU%'"
)
for r in cur.fetchall():
    print(r)

print()
print("=== vw_RPT_VAL @ start (2025-12-31) ===")
cur.execute(
    "SELECT SNAME, ISIN, FMCID, CAT, LS, UNITS, PTVALUE FROM dbo.vw_RPT_VAL "
    "WHERE PNAME='OAKS Emerging and Frontier Fund' AND VDATE='2025-12-31' "
    "AND (SNAME LIKE '%BIDU%' OR SNAME LIKE '%BAIDU%')"
)
for r in cur.fetchall():
    print(r)

print()
print("=== vw_RPT_VAL @ end (2026-05-04) ===")
cur.execute(
    "SELECT SNAME, ISIN, FMCID, CAT, LS, UNITS, PTVALUE FROM dbo.vw_RPT_VAL "
    "WHERE PNAME='OAKS Emerging and Frontier Fund' AND VDATE='2026-05-04' "
    "AND (SNAME LIKE '%BIDU%' OR SNAME LIKE '%BAIDU%')"
)
for r in cur.fetchall():
    print(r)

print()
print("=== tH_VAL OEFOF family BIDU/BAIDU since 2025-12-30 ===")
placeholders = ",".join("?" for _ in PCODES)
cur.execute(
    "SELECT h.VDATE, h.PCODE, h.SCODE, s.SHORT_NAME, h.UNITS, h.PTVALUE "
    "FROM ccl.dbo.tH_VAL h JOIN ccl.dbo.tSecurity s ON s.SCODE=h.SCODE "
    "WHERE (s.SHORT_NAME LIKE '%BIDU%' OR s.SHORT_NAME LIKE '%BAIDU%') "
    f"AND UPPER(h.PCODE) IN ({placeholders}) "
    "AND h.VDATE >= '2025-12-30' "
    "ORDER BY h.VDATE, s.SHORT_NAME",
    *PCODES,
)
rows = cur.fetchall()
print(f"rows={len(rows)}")
for r in rows[:80]:
    print(r)
if len(rows) > 80:
    print(f"... +{len(rows)-80} more")

print()
print("=== tTRANS BIDU/BAIDU in period ===")
cur.execute(
    "SELECT t.PCODE, s.SHORT_NAME, t.T, COUNT(*) as N, "
    "SUM(t.UNITS) as NET_UNITS, SUM(t.NSETTLE_AMOUNT) as NET_CASH, "
    "SUM(t.NINCOME) as NINCOME "
    "FROM ccl.dbo.tTRANS t LEFT JOIN ccl.dbo.tSecurity s ON s.SCODE=t.SCODE "
    "WHERE (s.SHORT_NAME LIKE '%BIDU%' OR s.SHORT_NAME LIKE '%BAIDU%') "
    f"AND UPPER(t.PCODE) IN ({placeholders}) "
    "AND (t.D IS NULL OR t.D <> 'Y') "
    "AND t.CDATE BETWEEN '2025-12-31' AND '2026-05-04' "
    "GROUP BY t.PCODE, s.SHORT_NAME, t.T ORDER BY s.SHORT_NAME, t.T",
    *PCODES,
)
for r in cur.fetchall():
    print(r)

print()
print("=== Any vw_RPT_VAL row mentioning BIDU/BAIDU on end VDATE (any PNAME) ===")
cur.execute(
    "SELECT PNAME, SNAME, ISIN, FMCID, CAT, LS, UNITS, PTVALUE FROM dbo.vw_RPT_VAL "
    "WHERE VDATE='2026-05-04' AND (SNAME LIKE '%BIDU%' OR SNAME LIKE '%BAIDU%' OR FMCID LIKE '%BIDU%')"
)
for r in cur.fetchall():
    print(r)

print()
print("=== Any tH_VAL row in OEFOF family with SHORT (LS='S') and SNAME containing CFD on the BAIDU end date ===")
cur.execute(
    "SELECT TOP 30 h.VDATE, h.PCODE, s.SHORT_NAME, h.UNITS, h.PTVALUE "
    "FROM ccl.dbo.tH_VAL h JOIN ccl.dbo.tSecurity s ON s.SCODE=h.SCODE "
    "WHERE h.VDATE='2026-05-04' AND h.UNITS < 0 "
    f"AND UPPER(h.PCODE) IN ({placeholders}) ORDER BY s.SHORT_NAME",
    *PCODES,
)
for r in cur.fetchall():
    print(r)
