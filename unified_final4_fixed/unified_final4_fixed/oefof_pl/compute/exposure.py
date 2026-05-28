"""Exposure data-types and pure-compute helpers.

This module owns three things:

1. The :class:`ExposureMetrics` / :class:`ExposureBlock` / :class:`PeriodMeta`
   value objects — shared by the pipeline and the workbook writer.
2. Pure-compute helpers that build exposure blocks from snapshot DataFrames
   (no I/O, no COM, no globals).
3. NAV helper utilities used by the Summary sheet.

Design note: these types live in ``compute/`` (not ``output/``) because they
are produced by the pipeline *before* workbook construction.  The workbook
writer consumes them as read-only inputs.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd


# ─── Value types ─────────────────────────────────────────────────────────────

@dataclass
class ExposureMetrics:
    """Exposure figures for one entity (fund or analyst).

    ``short_eur`` is always a positive absolute magnitude (not signed).
    net  = long − short
    gross = long + short
    """
    long_eur: float = 0.0
    short_eur: float = 0.0
    nav_eur: float = 0.0

    @property
    def net_eur(self) -> float:
        return self.long_eur - self.short_eur

    @property
    def gross_eur(self) -> float:
        return self.long_eur + self.short_eur

    def pct_of_nav(self, value: float) -> float | None:
        return value / self.nav_eur if self.nav_eur > 1e-9 else None


@dataclass
class ExposureBlock:
    """Snapshot or AUM-weighted exposure for the fund and per analyst."""
    label: str = ""
    as_of: pd.Timestamp | None = None
    fund: ExposureMetrics = field(default_factory=ExposureMetrics)
    by_analyst: dict[str, ExposureMetrics] = field(default_factory=dict)


@dataclass(frozen=True)
class PeriodMeta:
    """All non-numeric context the workbook needs."""
    fund_name: str
    start_date: pd.Timestamp
    end_date: pd.Timestamp
    fund_currency: str = "EUR"
    snapshot_path: str = ""
    bottler_path: str = ""
    run_timestamp: datetime = field(default_factory=datetime.now)
    # Optional NAV reconciliation components (EUR). Ordered tuple of
    # (label, value, is_total) rows rendered on the Summary sheet. Bridges
    # the equity-only stock MV to the HP_VAL fund portfolio total.
    nav_components: tuple[tuple[str, float, bool], ...] | None = None
    # Optional total NAV at start of period (EUR). When present, the
    # Summary headline ``Total P&L %`` is reported as Total P&L / NAV start
    # — the apples-to-apples denominator vs. the firm's TWR dashboard,
    # which divides the period P&L by starting fund NAV (stocks + cash
    # − accruals at 31-Dec).
    nav_total_start: float | None = None
    # Current-snapshot exposure (from end-period HiPort snapshot).
    snapshot_exposure: ExposureBlock | None = None
    # AUM-weighted daily average exposure over the full period.
    ytd_weighted_exposure: ExposureBlock | None = None


# ─── NAV helpers ─────────────────────────────────────────────────────────────

def compute_nav_components(
    end_df_fund: pd.DataFrame, pl_df: pd.DataFrame,
) -> tuple[tuple[str, float, bool], ...] | None:
    """Build the NAV reconciliation rows for the Summary sheet.

    Bridges the equity-only stock MV (what our P&L engine reports) to the
    HP_VAL fund portfolio total (sum of every PTVALUE_EUR row HiPort
    carries for the fund PCODE family). The gap is non-equity items —
    cash, FX equivalents, accrued fees, swap P&L cash buckets, dividend
    receivables — which are intentionally outside the stock-level model.

    Returns ``None`` if the snapshot is missing the required columns
    (lets unit tests using stub frames skip rendering).
    """
    if end_df_fund.empty:
        return None
    needed = {"PTVALUE_EUR", "ISIN", "CAT", "SNAME"}
    if not needed.issubset(end_df_fund.columns):
        return None

    df = end_df_fund.copy()
    df["__isin__"] = df["ISIN"].astype(str).str.strip()
    df["__val__"] = pd.to_numeric(df["PTVALUE_EUR"], errors="coerce").fillna(0.0)

    # Stock MV: take from pl_df so it matches what's printed elsewhere on
    # the Summary (uses our positions engine, not raw snapshot).
    stock_mv = float(pd.to_numeric(
        pl_df.get("Market Value End (EUR)", pd.Series(dtype=float)),
        errors="coerce",
    ).fillna(0.0).sum())

    # Cash bucket: rows with no real ISIN. Split positives from negatives
    # so the user sees what's an asset vs an accrual/expense.
    no_isin = df[df["__isin__"].isin(["", "nan", "NAN", "None", "NaN"])]
    cash_pos = float(no_isin.loc[no_isin["__val__"] > 0, "__val__"].sum())
    cash_neg = float(no_isin.loc[no_isin["__val__"] < 0, "__val__"].sum())

    # HP_VAL fund total = sum of every PTVALUE_EUR (equity + cash + accruals).
    hp_val_total = float(df["__val__"].sum())

    return (
        ("Stock Market Value (Equities, our P&L)", stock_mv, False),
        ("+ Cash & FX Equivalents", cash_pos, False),
        ("− Accrued Fees, Expenses & Swap P&L", cash_neg, False),
        ("Total — HP_VAL OEFOF Portfolio (EUR)", hp_val_total, True),
    )


def compute_nav_total(snap_df: pd.DataFrame) -> float | None:
    """Sum every PTVALUE_EUR row in an HP_VAL snapshot → total fund NAV.

    Used to obtain the **starting NAV** denominator for the headline
    Total P&L %: matches what the firm's per-share TWR dashboard divides
    by (stocks + cash + accruals at the period start date). Returns
    ``None`` when the snapshot lacks ``PTVALUE_EUR`` (e.g. unit-test
    stubs) so the headline gracefully falls back to cost-basis %.
    """
    if snap_df is None or snap_df.empty or "PTVALUE_EUR" not in snap_df.columns:
        return None
    return float(pd.to_numeric(snap_df["PTVALUE_EUR"], errors="coerce").fillna(0.0).sum())


# ─── Exposure compute helpers ─────────────────────────────────────────────────

def build_analyst_lookup(pl_df: pd.DataFrame) -> dict[str, dict[str, str]]:
    """Build ISIN→analyst and name→analyst lookup dicts from a P&L DataFrame."""
    by_isin: dict[str, str] = {}
    by_name: dict[str, str] = {}
    if pl_df is None or pl_df.empty:
        return {"isin": by_isin, "name": by_name}
    for _, row in pl_df.iterrows():
        analyst = str(row.get("Analyst", "")).strip().upper() or "UNASSIGNED"
        isin = str(row.get("ISIN", "")).strip().upper()
        sname = str(row.get("Stock Name", "")).strip().upper()
        if isin and isin not in by_isin:
            by_isin[isin] = analyst
        if sname and sname not in by_name:
            by_name[sname] = analyst
    return {"isin": by_isin, "name": by_name}


def resolve_snapshot_analyst(row: pd.Series, analyst_keys: dict[str, dict[str, str]]) -> str:
    """Resolve analyst code for a single snapshot row."""
    isin = str(row.get("ISIN", "")).strip().upper()
    sname = str(row.get("SNAME", "")).strip().upper()
    analyst = analyst_keys["isin"].get(isin) or analyst_keys["name"].get(sname)
    return analyst or "UNASSIGNED"


def snapshot_exposure_magnitude(df: pd.DataFrame) -> pd.Series:
    """Return the exposure magnitude series for each row of a snapshot DataFrame.

    For futures, swaps and FT-swaps the ``EXPOSURE`` column is used when
    available (notional exposure > market value). For all other instruments
    the absolute ``PTVALUE_EUR`` is used.
    """
    ptvalue = pd.to_numeric(df.get("PTVALUE_EUR", 0.0), errors="coerce").fillna(0.0).abs()
    if "EXPOSURE" in df.columns:
        exposure = pd.to_numeric(df["EXPOSURE"], errors="coerce").fillna(0.0).abs()
    else:
        exposure = pd.Series(0.0, index=df.index, dtype=float)
    cat = df.get("CAT", "").astype(str).str.upper()
    use_exposure = cat.isin(["FUT", "SWAP", "FTSWAP"])
    return exposure.where(use_exposure & exposure.gt(0.0), ptvalue)


def build_snapshot_exposure(
    snap_df: pd.DataFrame,
    pl_df: pd.DataFrame,
    as_of: pd.Timestamp,
) -> ExposureBlock:
    """Build point-in-time long/short/net/gross exposure from the end snapshot."""
    if snap_df is None or snap_df.empty:
        return ExposureBlock(label="Current Snapshot", as_of=as_of)

    analyst_keys = build_analyst_lookup(pl_df)
    long_mask = snap_df["LS"].astype(str).str.upper().ne("S")
    short_mask = snap_df["LS"].astype(str).str.upper().eq("S")
    ptvalue = pd.to_numeric(snap_df["PTVALUE_EUR"], errors="coerce").fillna(0.0)
    magnitude = snapshot_exposure_magnitude(snap_df)
    nav_eur = float(ptvalue.sum())

    by_analyst: dict[str, ExposureMetrics] = {}
    work = snap_df.copy()
    work["__MAGNITUDE__"] = magnitude
    work["__ANALYST__"] = work.apply(
        lambda row: resolve_snapshot_analyst(row, analyst_keys), axis=1,
    )
    for analyst, sub in work.groupby("__ANALYST__", dropna=False):
        code = str(analyst).strip().upper() or "UNASSIGNED"
        analyst_nav = nav_eur
        analyst_long = float(sub.loc[sub["LS"].astype(str).str.upper().ne("S"), "__MAGNITUDE__"].sum())
        analyst_short = float(sub.loc[sub["LS"].astype(str).str.upper().eq("S"), "__MAGNITUDE__"].sum())
        by_analyst[code] = ExposureMetrics(
            long_eur=max(analyst_long, 0.0),
            short_eur=analyst_short,
            nav_eur=analyst_nav,
        )

    return ExposureBlock(
        label="Current Snapshot",
        as_of=as_of,
        fund=ExposureMetrics(
            long_eur=float(magnitude[long_mask].sum()),
            short_eur=float(magnitude[short_mask].sum()),
            nav_eur=nav_eur,
        ),
        by_analyst=by_analyst,
    )


def _aum_weight_metrics(daily_df: pd.DataFrame) -> ExposureMetrics:
    if daily_df.empty:
        return ExposureMetrics()
    nav = pd.to_numeric(daily_df["NAV_EUR"], errors="coerce").fillna(0.0)
    nav_sum = float(nav.sum())
    if nav_sum <= 1e-9:
        return ExposureMetrics()
    long_eur = float(pd.to_numeric(daily_df["long_eur"], errors="coerce").fillna(0.0).mul(nav).sum() / nav_sum)
    short_eur = float(pd.to_numeric(daily_df["short_eur"], errors="coerce").fillna(0.0).mul(nav).sum() / nav_sum)
    nav_avg = float(nav.mean())
    return ExposureMetrics(long_eur=long_eur, short_eur=short_eur, nav_eur=nav_avg)


def build_weighted_exposure(
    history_df: pd.DataFrame,
    nav_history_df: pd.DataFrame,
    pl_df: pd.DataFrame,
) -> ExposureBlock:
    """AUM-weighted daily average exposure over the pipeline period."""
    if history_df.empty or nav_history_df.empty:
        return ExposureBlock(label="YTD AUM-Weighted Average Exposure")

    analyst_keys = build_analyst_lookup(pl_df)
    work = history_df.copy()
    work["VDATE"] = pd.to_datetime(work["VDATE"], errors="coerce").dt.normalize()
    work["__MAGNITUDE__"] = snapshot_exposure_magnitude(work)
    work["__SHORT__"] = work["LS"].astype(str).str.upper().eq("S")
    work["__ANALYST__"] = work.apply(
        lambda row: resolve_snapshot_analyst(row, analyst_keys), axis=1,
    )
    nav = nav_history_df.copy()
    nav["VDATE"] = pd.to_datetime(nav["VDATE"], errors="coerce").dt.normalize()
    nav = nav.groupby("VDATE", as_index=False)["NAV_EUR"].last()

    daily_fund = work.groupby("VDATE").apply(
        lambda sub: pd.Series({
            "long_eur": float(sub.loc[~sub["__SHORT__"], "__MAGNITUDE__"].sum()),
            "short_eur": float(sub.loc[sub["__SHORT__"], "__MAGNITUDE__"].sum()),
        })
    ).reset_index()
    merged_fund = daily_fund.merge(nav, on="VDATE", how="inner")
    fund_metrics = _aum_weight_metrics(merged_fund)

    by_analyst: dict[str, ExposureMetrics] = {}
    for analyst, sub in work.groupby("__ANALYST__", dropna=False):
        code = str(analyst).strip().upper() or "UNASSIGNED"
        daily = sub.groupby("VDATE").apply(
            lambda day: pd.Series({
                "long_eur": float(day.loc[~day["__SHORT__"], "__MAGNITUDE__"].sum()),
                "short_eur": float(day.loc[day["__SHORT__"], "__MAGNITUDE__"].sum()),
            })
        ).reset_index()
        merged = daily.merge(nav, on="VDATE", how="inner")
        by_analyst[code] = _aum_weight_metrics(merged)

    as_of = None if merged_fund.empty else pd.Timestamp(merged_fund["VDATE"].max())
    return ExposureBlock(
        label="YTD AUM-Weighted Average Exposure",
        as_of=as_of,
        fund=fund_metrics,
        by_analyst=by_analyst,
    )
