"""Diagnostic: find the NOK XRATE issue.

Run from the oefof_pl directory:
    python diag_nok.py

Queries vw_RPT_VAL and tH_VAL directly to identify which tables have
NULL XRATE for NOK positions and where a fallback rate can be found.
"""
import sys
import pandas as pd

sys.path.insert(0, ".")
from oefof_pl.config import load_default_config

cfg = load_default_config()

try:
    import pyodbc
except ImportError:
    print("ERROR: pyodbc not installed")
    sys.exit(1)

START_DATE = "2025-12-31"
END_DATE = "2026-05-18"
PCODES = cfg.oefof_pcodes

print(f"\n{'='*60}")
print(f"NOK XRATE diagnostic")
print(f"Start: {START_DATE}  End: {END_DATE}")
print(f"{'='*60}\n")

with pyodbc.connect(cfg.sql_conn_str, timeout=30) as cn:
    cur = cn.cursor()

    # 1. Check vw_RPT_VAL start snapshot for NOK rows
    print("=== 1. vw_RPT_VAL Dec-31 — rows with CCY=NOK ===")
    placeholders = ",".join("?" for _ in PCODES)
    cur.execute(
        f"SELECT PCODE_ORIG, ISIN, SNAME, CCY, XRATE, UNITS, PTVALUE "
        f"FROM ccl.dbo.vw_RPT_VAL "
        f"WHERE CAST(VDATE AS date)=CAST(? AS date) "
        f"AND CCY='NOK' "
        f"AND UPPER(PCODE) IN ({placeholders})",
        START_DATE, *PCODES
    )
    rows = cur.fetchall()
    if rows:
        for r in rows:
            print(f"  PCODE={r[0]}  ISIN={r[1]}  SNAME={r[2]}  CCY={r[3]}  XRATE={r[4]}  UNITS={r[5]}  PTVALUE={r[6]}")
    else:
        print("  (no rows) — NOK position not in start snapshot at all")

    # 2. Check tH_VAL Dec-31 for NOK
    print("\n=== 2. tH_VAL Dec-31 — rows with CCY=NOK ===")
    cur.execute(
        f"SELECT h.PCODE_ORIG, s.ISIN, h.SNAME, h.CCY, h.XRATE, h.UNITS "
        f"FROM ccl.dbo.tH_VAL h "
        f"LEFT JOIN ccl.dbo.tSecurity s ON s.SCODE=h.SCODE "
        f"WHERE CAST(h.VDATE AS date)=CAST(? AS date) "
        f"AND h.CCY='NOK' "
        f"AND UPPER(h.PCODE) IN ({placeholders})",
        START_DATE, *PCODES
    )
    rows = cur.fetchall()
    if rows:
        for r in rows:
            print(f"  PCODE={r[0]}  ISIN={r[1]}  SNAME={r[2]}  CCY={r[3]}  XRATE={r[4]}  UNITS={r[5]}")
    else:
        print("  (no rows) — NOK not in tH_VAL Dec-31 for OEFOF PCODEs")

    # 3. Find the most recent valid NOK XRATE in vw_RPT_VAL (any PCODE, any date)
    print("\n=== 3. Most recent valid NOK XRATE anywhere in vw_RPT_VAL ===")
    cur.execute(
        "SELECT TOP 5 VDATE, CCY, XRATE, PCODE_ORIG "
        "FROM ccl.dbo.vw_RPT_VAL "
        "WHERE CCY='NOK' AND XRATE IS NOT NULL AND XRATE > 0 "
        "ORDER BY VDATE DESC"
    )
    rows = cur.fetchall()
    if rows:
        for r in rows:
            print(f"  VDATE={r[0]}  CCY={r[1]}  XRATE={r[2]}  PCODE={r[3]}")
    else:
        print("  (no valid NOK XRATE found anywhere)")

    # 4. Check tTRANS for NOK trades - does it have XRATE column?
    print("\n=== 4. Check tTRANS schema for XRATE column ===")
    cur.execute(
        "SELECT TOP 1 * FROM ccl.dbo.tTRANS WHERE 1=0"
    )
    ttrans_cols = [c[0] for c in cur.description]
    if "XRATE" in ttrans_cols:
        print("  tTRANS HAS XRATE column ✓")
        # Get most recent NOK XRATE from tTRANS
        placeholders = ",".join("?" for _ in PCODES)
        cur.execute(
            f"SELECT TOP 5 t.CDATE, s.CCY, t.XRATE, t.PCODE "
            f"FROM ccl.dbo.tTRANS t "
            f"LEFT JOIN ccl.dbo.tSecurity s ON s.SCODE=t.SCODE "
            f"WHERE s.CCY='NOK' AND t.XRATE IS NOT NULL AND t.XRATE > 0 "
            f"AND UPPER(t.PCODE) IN ({placeholders}) "
            f"ORDER BY t.CDATE DESC",
            *PCODES
        )
        rows = cur.fetchall()
        if rows:
            for r in rows:
                print(f"  CDATE={r[0]}  CCY={r[1]}  XRATE={r[2]}  PCODE={r[3]}")
        else:
            print("  (no NOK trades with valid XRATE)")
    else:
        print(f"  tTRANS does NOT have XRATE column")
        print(f"  Available cols: {[c for c in ttrans_cols if 'RATE' in c.upper() or 'FX' in c.upper() or 'EX' in c.upper()]}")

    # 5. What positions have NOK in the start-of-period trades (YTD 2026)?
    print("\n=== 5. NOK trades in period (Jan-May 2026) ===")
    cur.execute(
        f"SELECT TOP 10 t.CDATE, s.ISIN, s.SHORT_NAME, s.CCY, t.T, t.UNITS, t.NUPRICEG "
        f"FROM ccl.dbo.tTRANS t "
        f"LEFT JOIN ccl.dbo.tSecurity s ON s.SCODE=t.SCODE "
        f"WHERE s.CCY='NOK' "
        f"AND UPPER(t.PCODE) IN ({placeholders}) "
        f"AND CAST(t.CDATE AS date) >= '2026-01-01' "
        f"AND (t.D IS NULL OR t.D <> 'Y') "
        f"ORDER BY t.CDATE DESC",
        *PCODES
    )
    rows = cur.fetchall()
    if rows:
        for r in rows:
            print(f"  CDATE={r[0]}  ISIN={r[1]}  SNAME={r[2]}  CCY={r[3]}  T={r[4]}  UNITS={r[5]}  PRICE={r[6]}")
    else:
        print("  (no NOK trades in 2026)")

print("\nDone.")
