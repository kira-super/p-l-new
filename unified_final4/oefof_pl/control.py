"""Strict-run control gates: approvals, reconciliation, anomalies, audit."""

from __future__ import annotations

import csv
import hashlib
import json
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from .exceptions import OefofError
from .validate.checks import CheckResult, ValidationResult

_UNITS_TOL = 0.01
_EUR_TOL = 1.0
_PCT_TOL = 0.005
_STRICT_WARN_TO_FAIL = frozenset({"VAL-10", "VAL-12", "VAL-13"})


@dataclass(frozen=True)
class KnownExceptions:
    allowed: set[tuple[str, str, str]]

    def allows(self, date_yyyymmdd: str, scope: str, key: str) -> bool:
        d = date_yyyymmdd.strip()
        s = scope.strip().upper()
        k = key.strip().upper()
        return (
            (d, s, k) in self.allowed
            or (d, s, "*") in self.allowed
            or ("*", s, k) in self.allowed
            or ("*", s, "*") in self.allowed
        )


def load_known_exceptions(path: Path | None) -> KnownExceptions:
    if path is None or not path.exists():
        return KnownExceptions(allowed=set())
    allowed: set[tuple[str, str, str]] = set()
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            date = str(row.get("date", "*")).strip() or "*"
            scope = str(row.get("scope", "")).strip().upper()
            key = str(row.get("key", "*")).strip().upper() or "*"
            if not scope:
                continue
            allowed.add((date, scope, key))
    return KnownExceptions(allowed=allowed)


def load_json(path: Path | None) -> dict[str, Any]:
    if path is None or not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def write_json(path: Path, payload: dict[str, Any]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    tmp.replace(path)


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def collect_input_hashes(paths: list[Path]) -> dict[str, str]:
    out: dict[str, str] = {}
    for p in paths:
        if p.exists() and p.is_file():
            out[str(p)] = file_sha256(p)
    return out


def enforce_release_discipline(*, strict_run: bool, cwd: Path) -> None:
    if not strict_run:
        return
    try:
        dirty = subprocess.run(
            ["git", "status", "--porcelain"],
            cwd=str(cwd),
            check=False,
            capture_output=True,
            text=True,
        )
        if dirty.returncode != 0:
            raise OefofError("strict run: unable to query git status.")
        if dirty.stdout.strip():
            raise OefofError("strict run: git working tree is dirty.")
        tagged = subprocess.run(
            ["git", "describe", "--tags", "--exact-match"],
            cwd=str(cwd),
            check=False,
            capture_output=True,
            text=True,
        )
        if tagged.returncode != 0 or not tagged.stdout.strip():
            raise OefofError("strict run: HEAD is not an exact git tag.")
    except FileNotFoundError as e:
        raise OefofError("strict run: git executable not available.") from e


def enforce_override_approvals(
    *,
    strict_run: bool,
    date_yyyymmdd: str,
    approvals_path: Path | None,
    bonus_patched_rows: int,
    manual_override_count: int,
) -> None:
    if not strict_run:
        return
    if bonus_patched_rows <= 0 and manual_override_count <= 0:
        return
    if approvals_path is None or not approvals_path.exists():
        raise OefofError("strict run: override approvals file missing.")
    approvals: set[tuple[str, str]] = set()
    with approvals_path.open("r", encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            date = str(row.get("date", "")).strip()
            kind = str(row.get("kind", "")).strip().lower()
            approved = str(row.get("approved", "")).strip().lower()
            if date and kind and approved in {"y", "yes", "true", "1"}:
                approvals.add((date, kind))
    if bonus_patched_rows > 0 and (date_yyyymmdd, "bonus_price") not in approvals:
        raise OefofError(f"strict run: bonus_price override usage is not approved for {date_yyyymmdd}.")
    if manual_override_count > 0 and (date_yyyymmdd, "manual_ca") not in approvals:
        raise OefofError(f"strict run: manual_ca override usage is not approved for {date_yyyymmdd}.")


def escalate_validation_for_strict(
    validation: ValidationResult,
    *,
    strict_run: bool,
) -> ValidationResult:
    if not strict_run:
        return validation
    checks: list[CheckResult] = []
    for c in validation.checks:
        if c.id in _STRICT_WARN_TO_FAIL and c.status == "WARN":
            checks.append(CheckResult(
                id=c.id,
                name=c.name,
                status="FAIL",
                detail=f"[STRICT] promoted from WARN: {c.detail}",
                failures=list(c.failures),
            ))
        else:
            checks.append(c)
    return ValidationResult(checks=checks)


def _pct_diff(a: float, b: float) -> float:
    denom = max(abs(a), abs(b), 1.0)
    return abs(a - b) / denom


def reconcile_against_db_views(
    *,
    pl_df: pd.DataFrame,
    end_df: pd.DataFrame,
    trades_df: pd.DataFrame,
    date_yyyymmdd: str,
    known_exceptions: KnownExceptions,
) -> dict[str, Any]:
    issues: list[dict[str, Any]] = []

    our = pl_df.copy() if not pl_df.empty else pd.DataFrame()
    if our.empty:
        return {"ok": True, "issues": [], "summary": {"checked": 0}}
    our["ISIN"] = our["ISIN"].astype(str).str.strip().str.upper()
    our_open = our.loc[our["Ending Units"].astype(float).abs() > _UNITS_TOL].copy()
    our_agg = our_open.groupby("ISIN", as_index=False).agg({
        "Ending Units": "sum",
        "Market Value End (EUR)": "sum",
        "Income (Local)": "sum",
    })

    snap = end_df.copy() if not end_df.empty else pd.DataFrame(columns=["ISIN", "UNITS", "PTVALUE_EUR"])
    if not snap.empty:
        snap["ISIN"] = snap["ISIN"].astype(str).str.strip().str.upper()
        snap = snap.loc[snap["ISIN"] != ""]
    snap_agg = snap.groupby("ISIN", as_index=False).agg({
        "UNITS": "sum",
        "PTVALUE_EUR": "sum",
    }) if not snap.empty else pd.DataFrame(columns=["ISIN", "UNITS", "PTVALUE_EUR"])

    inc = trades_df.copy() if not trades_df.empty else pd.DataFrame(columns=["ISIN", "T", "INCOME_LOCAL"])
    if not inc.empty:
        inc["ISIN"] = inc["ISIN"].astype(str).str.strip().str.upper()
        inc = inc.loc[inc["T"].astype(str).str.upper() == "I"]
    inc_agg = inc.groupby("ISIN", as_index=False).agg({"INCOME_LOCAL": "sum"}) if not inc.empty else pd.DataFrame(columns=["ISIN", "INCOME_LOCAL"])

    merged = our_agg.merge(snap_agg, on="ISIN", how="outer").merge(inc_agg, on="ISIN", how="left")
    merged = merged.fillna(0.0)

    for _, r in merged.iterrows():
        isin = str(r["ISIN"])
        if not isin:
            continue
        if known_exceptions.allows(date_yyyymmdd, "ISIN", isin):
            continue
        u_ours = float(r.get("Ending Units", 0.0) or 0.0)
        u_db = float(r.get("UNITS", 0.0) or 0.0)
        if abs(u_ours - u_db) > _UNITS_TOL:
            issues.append({"type": "units", "ISIN": isin, "ours": u_ours, "db": u_db})
        mv_ours = float(r.get("Market Value End (EUR)", 0.0) or 0.0)
        mv_db = float(r.get("PTVALUE_EUR", 0.0) or 0.0)
        if _pct_diff(mv_ours, mv_db) > _PCT_TOL and abs(mv_ours - mv_db) > _EUR_TOL:
            issues.append({"type": "market_value", "ISIN": isin, "ours": mv_ours, "db": mv_db})
        inc_ours = float(r.get("Income (Local)", 0.0) or 0.0)
        inc_db = float(r.get("INCOME_LOCAL", 0.0) or 0.0)
        if _pct_diff(inc_ours, inc_db) > _PCT_TOL and abs(inc_ours - inc_db) > _EUR_TOL:
            issues.append({"type": "income_local", "ISIN": isin, "ours": inc_ours, "db": inc_db})

    return {
        "ok": len(issues) == 0,
        "issues": issues,
        "summary": {"checked": int(len(merged)), "issues": int(len(issues))},
    }


def detect_day_over_day_anomalies(
    *,
    pl_df: pd.DataFrame,
    previous_audit: dict[str, Any],
    date_yyyymmdd: str,
    known_exceptions: KnownExceptions,
    analyst_threshold_eur: float,
    isin_threshold_eur: float,
    total_threshold_eur: float,
) -> dict[str, Any]:
    if not previous_audit:
        return {"ok": True, "anomalies": [], "summary": {"baseline": "missing"}}

    anomalies: list[dict[str, Any]] = []
    cur_total = float(pl_df.get("Total P&L (EUR)", pd.Series(dtype=float)).sum()) if not pl_df.empty else 0.0
    prev_total = float(previous_audit.get("totals", {}).get("total_pl_eur", 0.0))
    if abs(cur_total - prev_total) > total_threshold_eur and not known_exceptions.allows(date_yyyymmdd, "TOTAL_PL", "*"):
        anomalies.append({"scope": "TOTAL_PL", "key": "*", "delta_eur": cur_total - prev_total})

    cur_analyst = (
        pl_df.groupby("Analyst")["Total P&L (EUR)"].sum().to_dict() if not pl_df.empty and "Analyst" in pl_df.columns else {}
    )
    prev_analyst = previous_audit.get("analyst_total_pl_eur", {}) or {}
    for k in sorted(set(cur_analyst) | set(prev_analyst)):
        d = float(cur_analyst.get(k, 0.0)) - float(prev_analyst.get(k, 0.0))
        if abs(d) > analyst_threshold_eur and not known_exceptions.allows(date_yyyymmdd, "ANALYST", str(k)):
            anomalies.append({"scope": "ANALYST", "key": str(k), "delta_eur": d})

    cur_isin = (
        pl_df.groupby("ISIN")["Total P&L (EUR)"].sum().to_dict() if not pl_df.empty and "ISIN" in pl_df.columns else {}
    )
    prev_isin = previous_audit.get("isin_total_pl_eur", {}) or {}
    for k in sorted(set(cur_isin) | set(prev_isin)):
        d = float(cur_isin.get(k, 0.0)) - float(prev_isin.get(k, 0.0))
        if abs(d) > isin_threshold_eur and not known_exceptions.allows(date_yyyymmdd, "ISIN", str(k)):
            anomalies.append({"scope": "ISIN", "key": str(k), "delta_eur": d})

    return {"ok": len(anomalies) == 0, "anomalies": anomalies, "summary": {"count": len(anomalies)}}


def build_audit_payload(
    *,
    cfg_name: str,
    date_yyyymmdd: str,
    strict_run: bool,
    pl_df: pd.DataFrame,
    input_hashes: dict[str, str],
    reconciliation: dict[str, Any],
    anomalies: dict[str, Any],
) -> dict[str, Any]:
    analyst_totals = (
        pl_df.groupby("Analyst")["Total P&L (EUR)"].sum().to_dict()
        if not pl_df.empty and "Analyst" in pl_df.columns
        else {}
    )
    isin_totals = (
        pl_df.groupby("ISIN")["Total P&L (EUR)"].sum().to_dict()
        if not pl_df.empty and "ISIN" in pl_df.columns
        else {}
    )
    return {
        "run_date": date_yyyymmdd,
        "strict_run": strict_run,
        "golden_source_contract": {
            "snapshot": "ccl.dbo.vw_RPT_VAL",
            "trades": "ccl.dbo.tTRANS + ccl.dbo.tSecurity",
            "authoritative": True,
            "exception_inputs": ["ca_overrides.csv", "ca_bonus_prices.csv"],
            "fund": cfg_name,
        },
        "totals": {
            "total_pl_eur": float(pl_df.get("Total P&L (EUR)", pd.Series(dtype=float)).sum()) if not pl_df.empty else 0.0,
            "rows": int(len(pl_df)),
        },
        "analyst_total_pl_eur": {str(k): float(v) for k, v in analyst_totals.items()},
        "isin_total_pl_eur": {str(k): float(v) for k, v in isin_totals.items()},
        "input_hashes": input_hashes,
        "reconciliation": reconciliation,
        "anomalies": anomalies,
    }


def build_inflation_diagnostics(
    *,
    pl_df: pd.DataFrame,
    trades_df: pd.DataFrame,
    out_path: Path,
    manual_ca_count: int,
    auto_ca_count: int,
    zzzz_ps_dropped: int,
) -> None:
    rows: list[dict[str, Any]] = []
    rows.append({
        "record_type": "summary",
        "metric": "manual_ca_count",
        "value": manual_ca_count,
    })
    rows.append({
        "record_type": "summary",
        "metric": "auto_ca_count",
        "value": auto_ca_count,
    })
    rows.append({
        "record_type": "summary",
        "metric": "zzzz_ps_dropped",
        "value": zzzz_ps_dropped,
    })
    if not trades_df.empty and "TRADE_ID" in trades_df.columns:
        dup = trades_df[trades_df.duplicated(subset=["TRADE_ID"], keep=False)]
        for _, r in dup.head(100).iterrows():
            rows.append({
                "record_type": "duplicate_trade",
                "metric": "TRADE_ID",
                "value": str(r.get("TRADE_ID", "")),
                "ISIN": str(r.get("ISIN", "")),
            })
    if not pl_df.empty and "Total P&L (EUR)" in pl_df.columns:
        movers = pl_df.assign(_abs=pl_df["Total P&L (EUR)"].abs()).sort_values("_abs", ascending=False).head(50)
        for _, r in movers.iterrows():
            rows.append({
                "record_type": "top_mover",
                "metric": "total_pl_eur",
                "value": float(r.get("Total P&L (EUR)", 0.0) or 0.0),
                "ISIN": str(r.get("ISIN", "")),
                "Analyst": str(r.get("Analyst", "")),
            })
    pd.DataFrame(rows).to_csv(out_path, index=False)
