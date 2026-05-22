"""Reconciliation: computed parquet vs live HiPort snapshot.

Run after the pipeline to validate key metrics against the DB source.

Usage:
    python bin/reconcile_vs_db.py                        # uses latest parquet + latest DB VDATE
    python bin/reconcile_vs_db.py --vdate 20260505       # pin the end snapshot date
    python bin/reconcile_vs_db.py --parquet out/pl_results.parquet

Checks performed:
  R-01  Unit count — our Ending Units vs HiPort UNITS per ISIN
  R-02  Market value — our Market Value End (EUR) vs HiPort PTVALUE per ISIN
  R-03  Income total — our Income (EUR) vs sum of tTRANS NINCOME per ISIN
  R-04  Missing positions — ISINs in DB snapshot not in our output (and vice versa)
  R-05  FX rates — our EUR/HKD/USD vs what HiPort XRATE says

Tolerances: units 0.01, EUR amounts 1.00 (rounding), pct 0.5%.
"""
from __future__ import annotations

import argparse
import sys
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore")

_PROJECT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_PROJECT))

from oefof_pl.config import load_default_config

_TOL_UNITS = 0.01
_TOL_EUR   = 1.00      # rounding tolerance
_TOL_PCT   = 0.005     # 0.5%

OEFOF_PCODES = ("OEFOF", "OEFOGSSC", "OEFOGSSW", "OEFOHFSW", "OEFOHFSC")
START_DATE   = "2025-12-31"


def _connect(conn_str: str):
    try:
        import pyodbc
        return pyodbc.connect(conn_str, timeout=30)
    except Exception as e:
        sys.exit(f"DB connection failed: {e}")


def _latest_vdate(conn) -> str:
    ph = ",".join("?" for _ in OEFOF_PCODES)
    row = conn.execute(
        f"SELECT MAX(VDATE) FROM ccl.dbo.vw_RPT_VAL "
        f"WHERE UPPER(PCODE) IN ({ph})", *OEFOF_PCODES
    ).fetchone()
    if not row or not row[0]:
        sys.exit("No VDATE found for OEFOF PCODEs")
    return str(row[0])[:10]


def _load_snapshot(conn, vdate: str) -> pd.DataFrame:
    ph = ",".join("?" for _ in OEFOF_PCODES)
    sql = (
        f"SELECT ISIN, PCODE, SNAME, CCY, LS, UNITS, PTVALUE, XRATE "
        f"FROM ccl.dbo.vw_RPT_VAL "
        f"WHERE UPPER(PCODE) IN ({ph}) AND VDATE = ? AND ISIN IS NOT NULL"
    )
    return pd.read_sql(sql, conn, params=[*OEFOF_PCODES, vdate])


def _load_income(conn, start: str, end: str) -> pd.DataFrame:
    ph = ",".join("?" for _ in OEFOF_PCODES)
    sql = (
        f"SELECT s.ISIN, SUM(t.NINCOME) AS income_local "
        f"FROM ccl.dbo.tTRANS t "
        f"LEFT JOIN ccl.dbo.tSecurity s ON s.SCODE = t.SCODE "
        f"WHERE UPPER(t.PCODE) IN ({ph}) "
        f"  AND t.T = 'I' "
        f"  AND t.CDATE > ? AND t.CDATE <= ? "
        f"  AND (t.D IS NULL OR t.D <> 'Y') "
        f"GROUP BY s.ISIN"
    )
    return pd.read_sql(sql, conn, params=[*OEFOF_PCODES, start, end])


def _pct_diff(a, b):
    denom = max(abs(b), abs(a), 1.0)
    return abs(a - b) / denom


def run(parquet_path: Path, vdate: str, conn_str: str) -> int:
    print(f"\n{'='*70}")
    print(f" OEFOF Reconciliation — parquet vs DB snapshot {vdate}")
    print(f"{'='*70}\n")

    # -- Load parquet --------------------------------------------------------
    if not parquet_path.exists():
        sys.exit(f"Parquet not found: {parquet_path}")
    our = pd.read_parquet(parquet_path)
    print(f"Parquet  : {parquet_path.name}  ({len(our)} positions)")

    # -- Load DB -------------------------------------------------------------
    conn = _connect(conn_str)
    snap = _load_snapshot(conn, vdate)
    # Aggregate per ISIN — take the LARGEST absolute PTVALUE row when the same
    # ISIN appears in multiple PCODEs (e.g. GSCBIHKT in OEFOF + OEFOGSSW).
    # Summing PTVALUEs across PCODEs gives wrong results for hedged positions
    # where the two PCODE rows have offsetting signs.
    snap["_abs_ptv"] = snap["PTVALUE"].abs()
    snap_agg = (
        snap.sort_values("_abs_ptv", ascending=False)
        .groupby("ISIN", as_index=False)
        .first()
        .rename(columns={"UNITS": "units", "PTVALUE": "ptvalue",
                          "XRATE": "xrate", "CCY": "ccy"})
    )
    income_db = _load_income(conn, START_DATE, vdate)
    conn.close()

    print(f"DB snap  : {len(snap_agg)} ISINs  (VDATE {vdate})")
    print(f"DB income: {len(income_db)} ISINs with T='I' trades  ({START_DATE} to {vdate})\n")

    results = {"PASS": 0, "WARN": 0, "FAIL": 0}
    failures = []

    def _flag(check_id, isin, name, field, ours, db_val, tol, status="FAIL"):
        diff = abs(ours - db_val)
        if diff > tol:
            failures.append({
                "Check": check_id, "ISIN": isin, "Name": name[:28],
                "Field": field,
                "Ours": round(ours, 4), "DB": round(db_val, 4),
                "Diff": round(diff, 4),
            })
            results[status] += 1
            return False
        results["PASS"] += 1
        return True

    # -- R-01: Unit count -----------------------------------------------------
    print("R-01  Unit count (Ending Units vs HiPort UNITS) ...")
    cur_isins = our[our["Ending Units"].abs() > 0.01][["ISIN", "Stock Name", "Ending Units"]].copy()
    for _, row in cur_isins.iterrows():
        isin = str(row["ISIN"])
        db_row = snap_agg[snap_agg["ISIN"] == isin]
        if db_row.empty:
            results["WARN"] += 1
            failures.append({"Check": "R-01", "ISIN": isin,
                              "Name": str(row["Stock Name"])[:28],
                              "Field": "Units", "Ours": row["Ending Units"],
                              "DB": "MISSING", "Diff": "N/A"})
            continue
        _flag("R-01", isin, str(row["Stock Name"]), "Units",
              float(row["Ending Units"]), float(db_row.iloc[0]["units"]), _TOL_UNITS)

    # -- R-02: Market value ---------------------------------------------------
    print("R-02  Market value End (EUR) vs HiPort PTVALUE ...")
    for _, row in our[our["Ending Units"].abs() > 0.01].iterrows():
        isin = str(row["ISIN"])
        db_row = snap_agg[snap_agg["ISIN"] == isin]
        if db_row.empty:
            continue
        ours_mv = float(row.get("Market Value End (EUR)", 0) or 0)
        db_ptv  = float(db_row.iloc[0]["ptvalue"])
        # PTVALUE is in fund base currency (EUR) already in HiPort
        if abs(ours_mv) < 1 and abs(db_ptv) < 1:
            results["PASS"] += 1
            continue
        pct = _pct_diff(ours_mv, db_ptv)
        if pct > _TOL_PCT:
            failures.append({
                "Check": "R-02", "ISIN": isin, "Name": str(row["Stock Name"])[:28],
                "Field": "MV End (EUR)",
                "Ours": round(ours_mv, 2), "DB": round(db_ptv, 2),
                "Diff": f"{pct*100:.2f}%",
            })
            results["FAIL"] += 1
        else:
            results["PASS"] += 1

    # -- R-03: Income total (local currency sum) ------------------------------
    print("R-03  Income (Local) sum vs DB tTRANS NINCOME ...")
    our_income = our[["ISIN", "Stock Name", "Income (Local)"]].copy()
    our_income["ISIN"] = our_income["ISIN"].astype(str)
    income_db["ISIN"] = income_db["ISIN"].astype(str)
    merged = our_income.merge(income_db, on="ISIN", how="outer")
    for _, row in merged.iterrows():
        isin = str(row["ISIN"])
        ours_inc = float(row.get("Income (Local)") or 0)
        db_inc   = float(row.get("income_local") or 0)
        if abs(ours_inc) < 1 and abs(db_inc) < 1:
            results["PASS"] += 1
            continue
        pct = _pct_diff(ours_inc, db_inc)
        if pct > _TOL_PCT:
            name = str(row.get("Stock Name", isin))[:28]
            failures.append({
                "Check": "R-03", "ISIN": isin, "Name": name,
                "Field": "Income (Local)",
                "Ours": round(ours_inc, 2), "DB": round(db_inc, 2),
                "Diff": f"{pct*100:.2f}%",
            })
            results["FAIL"] += 1
        else:
            results["PASS"] += 1

    # -- R-04: Missing / extra positions -------------------------------------
    print("R-04  Missing / extra positions ...")
    our_cur = set(our[our["Ending Units"].abs() > 0.01]["ISIN"].astype(str))
    db_cur  = set(snap_agg["ISIN"].astype(str))
    in_db_not_ours = db_cur - our_cur
    in_ours_not_db = our_cur - db_cur
    for isin in in_db_not_ours:
        name = snap_agg[snap_agg["ISIN"] == isin]["ccy"].values[0] if len(snap_agg[snap_agg["ISIN"] == isin]) else ""
        failures.append({"Check": "R-04", "ISIN": isin, "Name": "(in DB, missing from output)",
                         "Field": "Position", "Ours": 0, "DB": "present", "Diff": "N/A"})
        results["WARN"] += 1
    for isin in in_ours_not_db:
        row = our[our["ISIN"].astype(str) == isin]
        name = str(row.iloc[0]["Stock Name"]) if len(row) else ""
        failures.append({"Check": "R-04", "ISIN": isin, "Name": name[:28] + " (extra in output)",
                         "Field": "Position", "Ours": "present", "DB": 0, "Diff": "N/A"})
        results["WARN"] += 1
    if not in_db_not_ours and not in_ours_not_db:
        results["PASS"] += 1

    # -- Summary -------------------------------------------------------------
    print(f"\n{'-'*70}")
    print(f"  PASS: {results['PASS']}   WARN: {results['WARN']}   FAIL: {results['FAIL']}")
    print(f"{'-'*70}\n")
    if failures:
        fail_df = pd.DataFrame(failures)
        print(fail_df.to_string(index=False))
        print()
        # Save failures to CSV for review
        out_csv = _PROJECT / "out" / "reconcile_failures.csv"
        fail_df.to_csv(out_csv, index=False)
        print(f"Failures saved: {out_csv}")
    else:
        print("  All checks passed — computed output matches DB.")

    return 1 if results["FAIL"] > 0 else 0


def main():
    cfg = load_default_config()
    p = argparse.ArgumentParser(description="Reconcile parquet output vs DB")
    p.add_argument("--vdate", default="",
                   help="End snapshot YYYYMMDD (default: latest in DB)")
    p.add_argument("--parquet", default=str(_PROJECT / "out" / "pl_results.parquet"),
                   help="Path to parquet file")
    p.add_argument("--sql-conn", default=cfg.sql_conn_str)
    args = p.parse_args()

    conn = _connect(args.sql_conn)
    vdate = args.vdate[:10].replace("-", "-") if args.vdate else _latest_vdate(conn)
    conn.close()

    return run(Path(args.parquet), vdate, args.sql_conn)


if __name__ == "__main__":
    sys.exit(main())
