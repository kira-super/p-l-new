"""End-to-end orchestration for the OEFOF P&L pipeline.

The pipeline is a flat list of stages. Each stage either succeeds and
returns its artifacts in a :class:`StageResult`, or fails and the run
aborts immediately. There is no try/except wrapping that swallows errors
silently — every failure mode emits a structured ``StageResult`` so the
CLI / UI can render exactly what went wrong.

Stages (executed in order)
--------------------------
1.  ``load-end-snapshot`` — read today's HiPort snapshot from
    ``ccl.dbo.vw_RPT_VAL``; **VDATE is the only authoritative end date**
    (BUG-8). Aborts if missing.
2.  ``load-start-snapshot`` — find the archive snapshot for the start date,
    load it, ``assert_vdate_matches`` with the requested string.
3.  ``load-trades`` — pull the OEFOF trade blotter from ``ccl.dbo.tTRANS``
    joined to ``ccl.dbo.tSecurity`` (replaces the legacy ``StockTrList.xlsb``).
4.  ``filter-trade-window`` — keep trades with ``start < CDATE ≤ end``.
6.  ``aggregate-portfolios`` — per-ISIN aggregation for both snapshots.
7.  ``build-fx-tables`` — start FX, end FX, EUR/USD start, EUR/USD end.
8.  ``load-analyst-map`` — CSV-backed mapping; new ISINs recorded for review.
9.  ``compute-positions`` — pure compute, returns ``(pl_df, income_df)``.
10. ``validate`` — VAL-01..VAL-11 in order.
11. ``write-results-parquet`` — always emits ``pl_results.parquet``.
12. **Critical branch**: on any FAIL check, write
    ``OAKS_EM_VALIDATION_FAILURE_{ts}.xlsx`` + failure parquet, **no P&L
    workbook**. Otherwise write ``OAKS_EM_Stock_PL_Report.xlsx``.
13. ``build-valuation`` — ValuationA workbook (skipped if critical).
14. ``send-email`` *(optional)* — Outlook attachments for both workbooks.
15. ``save-analyst-map`` — persist newly-recorded UNASSIGNED rows.

The pipeline never deletes ``pl_results.parquet`` — every run replaces it
atomically. That single file is the audit trail.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

import pandas as pd

from .analyst.map import UNASSIGNED, AnalystMap
from .compute.aggregate import AggLine, aggregate_portfolio
from .compute.fx import build_fx_supplement, build_fx_table, resolve_eur_usd
from .compute.pl import compute_positions, split_ca_rows
from .config import Config
from .data.ca_overrides import (
    apply_bonus_price_overrides,
    append_history,
    build_suggestions,
    derive_ca_auto,
    load_bonus_prices,
    load_ca_overrides,
    load_history,
    load_overrides_merged,
    overrides_to_trades,
    write_suggestions_csv,
)
from .data.loader import assert_vdate_matches
from .data.sql_loader import (
    fetch_eur_usd_rate,
    fetch_missing_fx_rates,
    load_nav_history_sql,
    load_snapshot_history_sql,
    load_snapshot_sql,
    load_th_val_sql,
    load_trades_sql,
)
from .email_sender.outlook import (
    EmailMessage,
    EmailSendResult,
    build_email_html_body,
    build_summary_html_table,
    normalize_recipients,
    send_via_outlook,
)
from .exceptions import OefofError
from .output.valuation import (
    DATA_FIRST_ROW as _VAL_DATA_FIRST_ROW,
    HPV_FMCID_IDX as _VAL_HPV_FMCID_IDX,
    HPV_FUTVAL_IDX as _VAL_HPV_FUTVAL_IDX,
    HPV_INITIALS_IDX as _VAL_HPV_INITIALS_IDX,
    HPV_SECURITY_NAME_IDX as _VAL_HPV_SECURITY_NAME_IDX,
    HPV_SORT1_IDX as _VAL_HPV_SORT1_IDX,
    close_open_excel_workbook,
    extract_valuation_a_via_sql,
)
from .output.workbook import ExposureBlock, ExposureMetrics, PeriodMeta, build_workbook
from .validate import ValidationResult, run_all_checks


# ─── Output filenames ────────────────────────────────────────────────────────

RESULTS_PARQUET = "pl_results.parquet"
FAILURE_PARQUET_PREFIX = "pl_results_failed_"
FAILURE_XLSX_PREFIX = "VALIDATION_FAILURE_"
DIAGNOSTICS_PREFIX = "pnl_diagnostics_"


# ─── Result records ──────────────────────────────────────────────────────────

@dataclass(frozen=True)
class StageResult:
    name: str
    ok: bool
    message: str = ""
    artifacts: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class PipelineResult:
    ok: bool
    stages: tuple[StageResult, ...]
    aborted_at: str = ""
    pl_df: pd.DataFrame | None = None
    income_df: pd.DataFrame | None = None
    validation: ValidationResult | None = None
    pl_report_path: Path | None = None
    valuation_path: Path | None = None
    archive_path: Path | None = None
    failure_parquet_path: Path | None = None
    failure_xlsx_path: Path | None = None
    results_parquet_path: Path | None = None
    email_result: EmailSendResult | None = None
    ca_suggestions_path: Path | None = None
    unassigned_isins_path: Path | None = None
    end_date: pd.Timestamp | None = None  # period end VDATE (for CA persistence)

    def by_stage(self, name: str) -> StageResult:
        for s in self.stages:
            if s.name == name:
                return s
        raise KeyError(name)


# ─── Options ─────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class PipelineOptions:
    """All runtime knobs. Distinct from :class:`Config` (paths/codes)."""
    start_yyyymmdd: str = ""              # "" means use cfg.start_snapshot_default
    end_yyyymmdd: str = ""                # "" means take VDATE from HP_VAL
    send_email: bool = False
    email_to: str = ""
    email_cc: str = ""
    email_subject: str = ""
    email_body: str = "P&L pipeline output attached."
    email_from_smtp: str = ""
    fund_name: str = "OAKS Emerging and Frontier Fund"
    extra_ca_override_paths: tuple[Path, ...] = ()  # e.g. --apply-suggestions out/ca_suggestions_*.csv


# ─── Public entry ────────────────────────────────────────────────────────────

def run_pipeline(
    cfg: Config,
    options: PipelineOptions,
    *,
    extractor=None,
    outlook_dispatcher: Callable[[str], object] | None = None,
    closer: Callable[[Any], bool] | None = None,
    progress: Callable[[StageResult], None] | None = None,
) -> PipelineResult:
    """Execute the pipeline. Returns :class:`PipelineResult`.

    Injectable hooks:
        ``extractor`` — replaces ValuationA COM extraction (tests).
        ``outlook_dispatcher`` — replaces Outlook dispatch (tests).
        ``closer`` — replaces the open-workbook close-helper (tests).
        ``progress`` — invoked with each StageResult as it is appended,
            allowing the CLI to print live progress instead of waiting
            for the whole pipeline to finish.
    """
    class _StagesList(list):
        def append(self, item):  # type: ignore[override]
            super().append(item)
            if progress is not None:
                try:
                    progress(item)
                except Exception:  # never let UI break compute
                    pass

    stages: list[StageResult] = _StagesList()
    pipeline_start_time = datetime.now()
    output_dir = Path(cfg.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    diag_rows: list[dict[str, object]] = []
    diag_path = output_dir / f"{DIAGNOSTICS_PREFIX}{pipeline_start_time.strftime('%Y%m%d_%H%M%S')}.csv"

    def _diag_summary(stage: str, metric: str, value, detail: str = "") -> None:
        diag_rows.append({
            "record_type": "summary",
            "stage": stage,
            "metric": metric,
            "value": value,
            "isin": "",
            "stock_name": "",
            "analyst": "",
            "detail": detail,
        })

    def _diag_position(stage: str, row: pd.Series) -> None:
        def _num(key: str) -> float:
            return float(pd.to_numeric(row.get(key), errors="coerce") or 0.0)

        diag_rows.append({
            "record_type": "position",
            "stage": stage,
            "metric": "total_pl_eur",
            "value": _num("Total P&L (EUR)"),
            "isin": str(row.get("ISIN", "")),
            "stock_name": str(row.get("Stock Name", "")),
            "analyst": str(row.get("Analyst", "")),
            "detail": (
                f"realised={_num('Realised P&L (EUR)'):.2f}; "
                f"unrealised={_num('Unrealised P&L (EUR)'):.2f}; "
                f"income={_num('Income (EUR)'):.2f}; "
                f"bonus_units={_num('Bonus Units'):.0f}; "
                f"ending_units={_num('Ending Units'):.0f}"
            ),
        })

    def _write_diagnostics() -> None:
        if not diag_rows:
            return
        pd.DataFrame(diag_rows).to_csv(diag_path, index=False)

    def _abort(stage_name: str) -> PipelineResult:
        _write_diagnostics()
        return PipelineResult(
            ok=False, stages=tuple(stages), aborted_at=stage_name,
        )

    # ── 1. Load end snapshot (SQL: ccl.dbo.vw_RPT_VAL) ──────────────────────
    # End VDATE: explicit --end-yyyymmdd if given, else the latest VDATE
    # the server has for the OEFOF family. We resolve "latest" with a
    # tiny pre-query so the main loader still uses an exact VDATE filter
    # and the BUG-8 ``assert_vdate_matches`` guard stays meaningful.
    end_yyyymmdd_req = options.end_yyyymmdd
    try:
        if not end_yyyymmdd_req:
            end_yyyymmdd_req = _resolve_latest_sql_vdate(
                conn_str=cfg.sql_conn_str, fund_pcodes=cfg.fund_pcodes,
            )
        end_df, end_vdate = load_snapshot_sql(
            end_yyyymmdd_req,
            conn_str=cfg.sql_conn_str,
            fund_pcodes=cfg.fund_pcodes,
        )
    except OefofError as exc:
        stages.append(StageResult("load-end-snapshot", False, str(exc)))
        return _abort("load-end-snapshot")

    end_yyyymmdd = end_yyyymmdd_req or _to_yyyymmdd(end_vdate)
    try:
        assert_vdate_matches(end_vdate, end_yyyymmdd, where="vw_RPT_VAL[end]")
    except OefofError as exc:
        stages.append(StageResult("load-end-snapshot", False, str(exc)))
        return _abort("load-end-snapshot")
    stages.append(StageResult(
        "load-end-snapshot", True,
        f"end_vdate={end_vdate.date()} rows={len(end_df)}",
        {"vdate": end_vdate, "rows": len(end_df), "yyyymmdd": end_yyyymmdd},
    ))

    # ── 3. Load start snapshot (SQL) ─────────────────────────────────────────
    start_yyyymmdd = options.start_yyyymmdd or cfg.start_snapshot_default
    try:
        start_df, start_vdate = load_snapshot_sql(
            start_yyyymmdd,
            conn_str=cfg.sql_conn_str,
            fund_pcodes=cfg.fund_pcodes,
        )
        assert_vdate_matches(start_vdate, start_yyyymmdd, where="vw_RPT_VAL[start]")
    except (OefofError, ValueError) as exc:
        stages.append(StageResult("load-start-snapshot", False, str(exc)))
        return _abort("load-start-snapshot")
    stages.append(StageResult(
        "load-start-snapshot", True,
        f"start_vdate={start_vdate.date()} rows={len(start_df)}",
        {"vdate": start_vdate, "rows": len(start_df), "yyyymmdd": start_yyyymmdd},
    ))

    # ── 3b. Load tH_VAL end snapshot (sector taxonomy + ISIN_VALID gate) ────
    # Best-effort: a tH_VAL outage must NOT block the run. The loader has a
    # superset of vw_RPT_VAL columns; we use it for VAL-12 (ISIN_VALID) and
    # the sector breakdown sheet only.
    th_val_df: pd.DataFrame | None = None
    try:
        th_val_df, _ = load_th_val_sql(
            end_yyyymmdd,
            conn_str=cfg.sql_conn_str,
            fund_pcodes=cfg.fund_pcodes,
        )
        stages.append(StageResult(
            "load-th-val", True,
            f"rows={len(th_val_df)}",
            {"rows": len(th_val_df)},
        ))
    except Exception as exc:
        stages.append(StageResult(
            "load-th-val", True,  # non-fatal
            f"skipped ({exc})",
            {"error": str(exc)},
        ))
        th_val_df = None

    # ── 3c. Load tH_VAL start snapshot — FX supplement only ─────────────────
    # Best-effort. Used ONLY to recover XRATE for currencies that have
    # NULL XRATE in the start vw_RPT_VAL snapshot (see stage 7 below).
    # A closed position (e.g. NOK stock sold YTD) may have valid XRATE in
    # tH_VAL at Dec-31 even when vw_RPT_VAL carries NULL for that row.
    th_val_start_df: pd.DataFrame | None = None
    try:
        th_val_start_df, _ = load_th_val_sql(
            start_yyyymmdd,
            conn_str=cfg.sql_conn_str,
            fund_pcodes=cfg.fund_pcodes,
        )
    except Exception:
        th_val_start_df = None  # non-fatal; fx_start will be used as-is

    # ── 4. Load trades (SQL: ccl.dbo.tTRANS + tSecurity) ────────────────────
    try:
        trades_df = load_trades_sql(
            conn_str=cfg.sql_conn_str, fund_pcodes=cfg.fund_pcodes,
        )
    except OefofError as exc:
        stages.append(StageResult("load-trades", False, str(exc)))
        return _abort("load-trades")

    # Back-fill missing trade ISINs from the snapshot's SCODE→ISIN map.
    # Some securities (e.g. CFD/FT swaps such as 'GSCBIHKT CFD') have NULL
    # ``tSecurity.ISIN`` but get a synthetic ``FD<short>`` ISIN inside
    # ``vw_RPT_VAL``. ``load_trades_sql`` joins ``tSecurity`` directly so it
    # ends up with empty ISIN, which causes ``compute.pl`` (line ~184) to
    # silently drop every trade for these positions — zeroing out income,
    # realised P&L and trade dates on swap CFDs. Fix by remapping using the
    # snapshot view's synthetic ISIN.
    if "SCODE" in trades_df.columns:
        scode_isin: dict[str, str] = {}
        for snap in (end_df, start_df):
            if snap is None or snap.empty or "SCODE" not in snap.columns:
                continue
            sub = snap.loc[:, ["SCODE", "ISIN"]].dropna(subset=["SCODE"])
            for sc, isin in zip(sub["SCODE"], sub["ISIN"]):
                sc_s = str(sc).strip()
                isin_s = str(isin).strip().upper() if isin is not None else ""
                if sc_s and isin_s and sc_s not in scode_isin:
                    scode_isin[sc_s] = isin_s
        if scode_isin:
            mask = trades_df["ISIN"].fillna("").astype(str).str.strip() == ""
            if mask.any():
                trades_df.loc[mask, "ISIN"] = (
                    trades_df.loc[mask, "SCODE"].astype(str).str.strip().map(scode_isin).fillna("")
                )

    stages.append(StageResult(
        "load-trades", True, f"rows={len(trades_df)}",
        {"rows": len(trades_df)},
    ))
    _diag_summary("load-trades", "rows", len(trades_df))

    # ── 5. Filter trade window (start_vdate < CDATE <= end_vdate) ────────────
    try:
        filtered = _filter_trade_window(trades_df, start_vdate, end_vdate)
    except Exception as exc:
        stages.append(StageResult("filter-trade-window", False, str(exc)))
        return _abort("filter-trade-window")
    stages.append(StageResult(
        "filter-trade-window", True,
        f"kept={len(filtered)}/{len(trades_df)}",
        {"rows": len(filtered)},
    ))
    _diag_summary("filter-trade-window", "rows_kept", len(filtered), f"from={len(trades_df)}")
    _diag_summary("filter-trade-window", "rows_dropped", len(trades_df) - len(filtered))
    trades_df = filtered

    # ── 5a. Optional bonus-price overrides (price-only, no unit change) ─────
    try:
        bonus_prices = load_bonus_prices(cfg.ca_bonus_prices_path)
        trades_df, patched_bonus = apply_bonus_price_overrides(trades_df, bonus_prices)
        stages.append(StageResult(
            "apply-bonus-price-overrides", True,
            f"overrides={len(bonus_prices)} patched_rows={patched_bonus}",
            {
                "overrides": len(bonus_prices),
                "patched_rows": patched_bonus,
                "source": str(cfg.ca_bonus_prices_path),
            },
        ))
        _diag_summary("apply-bonus-price-overrides", "overrides", len(bonus_prices))
        _diag_summary("apply-bonus-price-overrides", "patched_rows", patched_bonus)
    except Exception as exc:
        stages.append(StageResult("apply-bonus-price-overrides", False, str(exc)))
        return _abort("apply-bonus-price-overrides")

    # ── 5b. Load + INJECT manual corporate-action overrides ────────────────
    # Manual overrides exist for the rare cases the snapshot itself is
    # suspect or units must be steered to a non-snapshot value (e.g.
    # mid-period correction). We inject them FIRST so the auto-derive
    # stage (6b) sees them as already-applied trades and only fills the
    # residual gap, never duplicating manual amounts.
    manual_ca_overrides: list = []
    try:
        all_paths = (cfg.ca_overrides_path, *options.extra_ca_override_paths)
        manual_ca_overrides = load_overrides_merged(*all_paths)
        if manual_ca_overrides:
            manual_df = overrides_to_trades(
                manual_ca_overrides,
                end_vdate=pd.Timestamp(end_vdate),
                template_columns=list(trades_df.columns),
            )
            trades_df = pd.concat([trades_df, manual_df], ignore_index=True)
        sources = ", ".join(str(p) for p in all_paths if Path(p).exists()) or "(none)"
        stages.append(StageResult(
            "load-manual-ca-overrides", True,
            f"manual={len(manual_ca_overrides)} from={sources}",
            {"applied": len(manual_ca_overrides), "sources": tuple(str(p) for p in all_paths)},
        ))
        _diag_summary("load-manual-ca-overrides", "manual_rows", len(manual_ca_overrides), sources)
    except Exception as exc:
        stages.append(StageResult("load-manual-ca-overrides", False, str(exc)))
        return _abort("load-manual-ca-overrides")

    # ── 6. Aggregate portfolios ──────────────────────────────────────────────
    # HP_VAL is a *firm-wide* snapshot (~30 PCODEs). The bottler trades are
    # already filtered to OEFOF_PCODES, so the snapshot must be too —
    # otherwise unit reconciliation (VAL-01) is meaningless because end-units
    # carry positions for funds we have no trades for.
    # FX tables are built from the *unfiltered* snapshot (more CCY coverage,
    # rates are firm-wide).
    try:
        start_df_oefof = _filter_snapshot_to_pcodes(start_df, cfg.fund_pcodes)
        end_df_oefof = _filter_snapshot_to_pcodes(end_df, cfg.fund_pcodes)
        start_agg = aggregate_portfolio(start_df_oefof)
        end_agg = aggregate_portfolio(end_df_oefof)
    except OefofError as exc:
        stages.append(StageResult("aggregate-portfolios", False, str(exc)))
        return _abort("aggregate-portfolios")
    stages.append(StageResult(
        "aggregate-portfolios", True,
        f"start_isins={len(start_agg)} end_isins={len(end_agg)} "
        f"(snapshot rows {len(start_df_oefof)}/{len(start_df)} start, "
        f"{len(end_df_oefof)}/{len(end_df)} end)",
        {"start_isins": len(start_agg), "end_isins": len(end_agg)},
    ))
    _diag_summary("aggregate-portfolios", "start_rows_kept", len(start_df_oefof), f"from={len(start_df)}")
    _diag_summary("aggregate-portfolios", "end_rows_kept", len(end_df_oefof), f"from={len(end_df)}")

    # ── 6b. Auto-derive CA adjustments from snapshot delta + inject ─────────
    # For each ISIN, computes residual = end_units − (start + Σbuys − Σsells)
    # using the trade frame that ALREADY contains manual CA rows from 5b.
    # So if a manual override fully closes the gap, auto produces nothing
    # for that ISIN; if manual under-closes, auto fills the residual.
    # VAL-01 then reconciles by construction.
    try:
        start_units = {isin: float(line.units) for isin, line in start_agg.items()}
        end_units = {isin: float(line.units) for isin, line in end_agg.items()}
        auto_ca_overrides = derive_ca_auto(
            start_units=start_units,
            end_units=end_units,
            trades_df=trades_df,
        )
        ca_overrides = list(manual_ca_overrides) + list(auto_ca_overrides)
        if auto_ca_overrides:
            auto_df = overrides_to_trades(
                auto_ca_overrides,
                end_vdate=pd.Timestamp(end_vdate),
                template_columns=list(trades_df.columns),
                base_id=-2_000_000,   # disjoint from manual range (-1_000_000…)
            )
            trades_df = pd.concat([trades_df, auto_df], ignore_index=True)
        stages.append(StageResult(
            "derive-ca-auto", True,
            f"manual={len(manual_ca_overrides)} auto={len(auto_ca_overrides)} "
            f"total={len(ca_overrides)}",
            {"manual": len(manual_ca_overrides), "auto": len(auto_ca_overrides)},
        ))
        _diag_summary("derive-ca-auto", "manual_rows", len(manual_ca_overrides))
        _diag_summary("derive-ca-auto", "auto_rows", len(auto_ca_overrides))
    except Exception as exc:
        stages.append(StageResult("derive-ca-auto", False, str(exc)))
        return _abort("derive-ca-auto")

    # ── 7. Build FX tables and EUR/USD ───────────────────────────────────────
    try:
        fx_start = build_fx_table(start_df)
        fx_end = build_fx_table(end_df)
        # Tier-2: tH_VAL start fills CCYs with NULL XRATE in the primary
        # start snapshot (e.g. a position where HiPort stored NULL).
        fx_start = build_fx_supplement(th_val_start_df, fx_start)
        # Tier-3: for positions opened AND closed entirely within the YTD
        # period (never in either snapshot — e.g. Capital Tankers MHY1096C1093
        # and SED Energy CY0101162119, both NOK, bought and sold in 2026),
        # fetch the most recent valid XRATE from vw_RPT_VAL across all PCODEs.
        # Include CCYs from trades as well as snapshots: positions opened
        # AND closed within the YTD period (e.g. Capital Tankers MHY1096C1093,
        # SED Energy CY0101162119 — both NOK, bought and sold in 2026) never
        # appear in start_agg or end_agg, so their CCY is only visible in trades.
        all_agg_ccys = {
            agg.ccy for agg in list(start_agg.values()) + list(end_agg.values())
            if agg.ccy
        }
        if not trades_df.empty and "CCY" in trades_df.columns:
            all_agg_ccys |= {
                str(c).strip().upper()
                for c in trades_df["CCY"].dropna().unique()
                if str(c).strip().upper() not in ("EUR", "USD", "")
            }
        missing_ccys = sorted(all_agg_ccys - set(fx_start) - set(fx_end))
        fallback_fx: dict[str, float] = {}
        if missing_ccys:
            fallback_fx = fetch_missing_fx_rates(
                missing_ccys,
                conn_str=cfg.sql_conn_str,
                as_of_yyyymmdd=end_yyyymmdd,
            )
            # Lowest priority — any real snapshot rate beats this.
            fx_start = {**fallback_fx, **fx_start}
        # For USD-base funds the snapshot has no EUR positions, so
        # resolve_eur_usd will fail.  In that case eur_usd = 1.0 (all
        # values are already in the fund currency) with a SQL fallback.
        if cfg.fund_currency.upper() == "USD":
            fb_start = fetch_eur_usd_rate(conn_str=cfg.sql_conn_str,
                                          as_of_yyyymmdd=start_yyyymmdd)
            fb_end   = fetch_eur_usd_rate(conn_str=cfg.sql_conn_str,
                                          as_of_yyyymmdd=end_yyyymmdd)
            eur_usd_start = fb_start if fb_start else 1.0
            eur_usd_end   = fb_end   if fb_end   else 1.0
        else:
            eur_usd_start = resolve_eur_usd(start_df)
            eur_usd_end = resolve_eur_usd(end_df)
    except OefofError as exc:
        stages.append(StageResult("build-fx-tables", False, str(exc)))
        return _abort("build-fx-tables")
    _fx_detail = (
        f"eur_usd start={eur_usd_start:.4f} end={eur_usd_end:.4f} "
        f"ccy_start={len(fx_start)} ccy_end={len(fx_end)}"
    )
    if fallback_fx:
        _fx_detail += f" fx_fallback={sorted(fallback_fx)}"
    if missing_ccys and set(missing_ccys) - set(fallback_fx):
        _fx_detail += f" fx_still_missing={sorted(set(missing_ccys)-set(fallback_fx))}"
    stages.append(StageResult(
        "build-fx-tables", True, _fx_detail,
        {"eur_usd_start": eur_usd_start, "eur_usd_end": eur_usd_end},
    ))

    # ── 8. Load analyst map ──────────────────────────────────────────────────
    try:
        analyst_map = AnalystMap.load(cfg.analyst_map_path, overrides=cfg.person_overrides)
    except OefofError as exc:
        stages.append(StageResult("load-analyst-map", False, str(exc)))
        return _abort("load-analyst-map")
    stages.append(StageResult(
        "load-analyst-map", True, f"loaded={len(analyst_map.by_isin)}",
        {"loaded": len(analyst_map.by_isin)},
    ))

    # ── 9. Compute positions ─────────────────────────────────────────────────
    try:
        compute_diag: dict[str, object] = {}
        pl_df, income_df = compute_positions(
            trades_df, start_agg, end_agg,
            fx_start=fx_start, fx_end=fx_end,
            eur_usd_start=eur_usd_start, eur_usd_end=eur_usd_end,
            end_date_ts=pd.Timestamp(end_vdate),
            analyst_resolver=lambda isin, sname: analyst_map.resolve(
                isin, sname, record_unknown=False,
            ),
            ftswap_isins=cfg.ftswap_isins,
            drop_zzzz_ps_trades=cfg.drop_zzzz_ps_trades,
            diagnostics=compute_diag,
        )
    except OefofError as exc:
        stages.append(StageResult("compute-positions", False, str(exc)))
        return _abort("compute-positions")
    for key, value in compute_diag.items():
        _diag_summary("compute-positions", key, value)
    split_before = len(pl_df)
    if cfg.split_ca_rows:
        pl_df = split_ca_rows(pl_df)
    _diag_summary("compute-positions", "rows_before_split", split_before)
    _diag_summary("compute-positions", "rows_after_split", len(pl_df))
    _diag_summary("compute-positions", "ca_split_rows_added", len(pl_df) - split_before)
    if not pl_df.empty and "Total P&L (EUR)" in pl_df.columns:
        worst = pl_df.sort_values("Total P&L (EUR)", ascending=True, na_position="last").head(10)
        for _, row in worst.iterrows():
            _diag_position("top-losses", row)
    stages.append(StageResult(
        "compute-positions", True,
        f"positions={len(pl_df)} income_records={len(income_df)}",
        {"positions": len(pl_df), "income_records": len(income_df)},
    ))
    # ── 9b. Extract ValuationA rows once (used by enrichment + workbook embed) ─
    # Done BEFORE validate so the parquet/failure artifacts also benefit from
    # the enriched analyst column.
    pl_report_path = output_dir / cfg.pl_report_filename
    val_rows: list | None = None
    used_extractor = extractor if extractor is not None else extract_valuation_a_via_sql
    used_closer = closer if closer is not None else close_open_excel_workbook
    if progress is not None:
        try:
            progress(StageResult(
                "extract-valuation", True,
                "starting (SQL: ccl.dbo.vw_RPT_VAL @ FCEDATA01)…",
            ))
        except Exception:
            pass
    try:
        used_closer(pl_report_path)
        val_rows = list(used_extractor(cfg.sql_conn_str, target_pname=cfg.target_pname))
        stages.append(StageResult(
            "extract-valuation", True,
            f"rows={max(0, len(val_rows) - 1)}",
            {"row_count": len(val_rows)},
        ))
    except Exception as exc:
        stages.append(StageResult("extract-valuation", False, str(exc)))
        val_rows = None

    # ── 9c. Enrich pl_df Analyst column from HiPort ValuationA initials ─────
    # Runs whenever we have either val_rows (pivot) or end_df_oefof (val_data).
    try:
        reassigned, recorded = _enrich_analysts_from_valuation(
            val_rows or [], pl_df, analyst_map,
            end_port_df=end_df_oefof,
        )
        stages.append(StageResult(
            "enrich-analysts", True,
            f"reassigned={reassigned} recorded_isins={recorded}",
            {"reassigned": reassigned, "recorded": recorded},
        ))
    except Exception as exc:
        stages.append(StageResult("enrich-analysts", False, str(exc)))
    # ── 10. Validate ─────────────────────────────────────────────────────────
    validation = run_all_checks(
        pl_df=pl_df, trades_df=trades_df,
        start_agg=start_agg, end_agg=end_agg,
        eur_usd_end=eur_usd_end,
        th_val_df=th_val_df,
    )
    fail_ids = [c.id for c in validation.checks if c.status == "FAIL"]
    warn_ids = [c.id for c in validation.checks if c.status == "WARN"]
    stages.append(StageResult(
        "validate", True,
        f"FAIL={len(fail_ids)} WARN={len(warn_ids)}",
        {"fail_ids": tuple(fail_ids), "warn_ids": tuple(warn_ids)},
    ))

    # ── 11. Always drop pl_results.parquet ───────────────────────────────────

    results_parquet = output_dir / RESULTS_PARQUET
    try:
        _write_parquet_atomic(pl_df, results_parquet)
        stages.append(StageResult(
            "write-results-parquet", True, str(results_parquet),
            {"path": str(results_parquet)},
        ))
    except Exception as exc:
        stages.append(StageResult("write-results-parquet", False, str(exc)))
        return _abort("write-results-parquet")

    # ── 12. Branch: critical → failure artifacts, no P&L workbook ────────────
    # NAV reconciliation must use the full SQL-filtered OEFOF-family snapshot
    # (already constrained in load_snapshot_sql by PCODE), not the stricter
    # PCODE_ORIG filter used for trade-aligned stock reconciliation.
    # Otherwise cash/accrual rows under sibling ORIG codes are dropped and the
    # Summary NAV block under-reports HP_VAL by several million EUR.
    nav_components = _compute_nav_components(end_df, pl_df)
    nav_total_start = _compute_nav_total(start_df)
    snapshot_exposure = _build_snapshot_exposure(end_df_oefof, pl_df, pd.Timestamp(end_vdate))
    ytd_weighted_exposure: ExposureBlock | None = None
    try:
        history_df = load_snapshot_history_sql(
            start_yyyymmdd=start_yyyymmdd,
            end_yyyymmdd=end_yyyymmdd,
            conn_str=cfg.sql_conn_str,
            fund_pcodes=cfg.fund_pcodes,
        )
        nav_history_df = load_nav_history_sql(
            start_yyyymmdd=start_yyyymmdd,
            end_yyyymmdd=end_yyyymmdd,
            conn_str=cfg.sql_conn_str,
            pcode_pattern=cfg.nav_pcode_pattern,
            nav_table=cfg.nav_sql_table,
            pcode_column=cfg.nav_pcode_column,
            vdate_column=cfg.nav_vdate_column,
            nav_value_column=cfg.nav_value_column,
        )
        ytd_weighted_exposure = _build_weighted_exposure(history_df, nav_history_df, pl_df)
        stages.append(StageResult(
            "load-exposure-history", True,
            f"days={len(nav_history_df)} snapshot_rows={len(history_df)}",
            {"days": len(nav_history_df), "rows": len(history_df)},
        ))
    except Exception as exc:
        stages.append(StageResult(
            "load-exposure-history", True,
            f"skipped ({exc})",
            {"error": str(exc)},
        ))
    period_meta = PeriodMeta(
        fund_name=cfg.fund_name,
        start_date=pd.Timestamp(start_vdate),
        end_date=pd.Timestamp(end_vdate),
        fund_currency=cfg.fund_currency,
        snapshot_path="ccl.dbo.vw_RPT_VAL @ FCEDATA01",
        bottler_path="ccl.dbo.tTRANS @ FCEDATA01",
        run_timestamp=pipeline_start_time,
        nav_components=nav_components,
        nav_total_start=nav_total_start,
        snapshot_exposure=snapshot_exposure,
        ytd_weighted_exposure=ytd_weighted_exposure,
    )
    ts_str = pipeline_start_time.strftime("%Y%m%d_%H%M%S")

    if validation.has_critical:
        failure_parquet = output_dir / f"{FAILURE_PARQUET_PREFIX}{ts_str}.parquet"
        failure_xlsx = output_dir / f"{FAILURE_XLSX_PREFIX}{ts_str}.xlsx"
        # Auto-detect VAL-01 → suggestions CSV BEFORE the workbook so the
        # workbook's "How to unblock" block can point straight at the file.
        suggestions_path: Path | None = None
        try:
            try:
                val01 = validation.by_id("VAL-01")
            except KeyError:
                val01 = None
            if val01 is not None and val01.failures:
                history = load_history(cfg.ca_history_path)
                suggestions = build_suggestions(val01.failures, history)
                if suggestions:
                    suggestions_path = output_dir / f"ca_suggestions_{ts_str}.csv"
                    write_suggestions_csv(suggestions_path, suggestions)
                    stages.append(StageResult(
                        "write-ca-suggestions", True,
                        f"file={suggestions_path.name} rows={len(suggestions)}",
                        {"path": str(suggestions_path), "rows": len(suggestions)},
                    ))
        except Exception as exc:
            stages.append(StageResult("write-ca-suggestions", False, str(exc)))
            # Non-fatal: continue with failure report.
        try:
            _write_parquet_atomic(pl_df, failure_parquet)
            build_workbook(
                out_path=failure_xlsx,
                pl_df=pl_df.iloc[0:0],          # empty holdings sheet
                income_df=income_df.iloc[0:0],
                trades_df=trades_df.iloc[0:0],
                validation=validation,
                period_meta=period_meta,
                ca_suggestions_path=suggestions_path,
                applied_ca_overrides=tuple(ca_overrides),
                ca_override_sources=tuple(
                    Path(p) for p in (cfg.ca_overrides_path, *options.extra_ca_override_paths)
                ),
                th_val_df=th_val_df,
            )
        except Exception as exc:
            stages.append(StageResult("write-failure-artifacts", False, str(exc)))
            return _abort("write-failure-artifacts")
        stages.append(StageResult(
            "write-failure-artifacts", True,
            f"FAIL={fail_ids}",
            {"parquet": str(failure_parquet), "xlsx": str(failure_xlsx)},
        ))
        _write_diagnostics()
        # Persist analyst map even on failure (new UNASSIGNED rows are useful).
        _save_analyst_map_safely(analyst_map, stages)
        unassigned_path = _write_unassigned_isins(pl_df, output_dir, ts_str, stages)
        return PipelineResult(
            ok=False,
            stages=tuple(stages),
            aborted_at="validate",
            pl_df=pl_df,
            income_df=income_df,
            validation=validation,
            failure_parquet_path=failure_parquet,
            failure_xlsx_path=failure_xlsx,
            results_parquet_path=results_parquet,
            ca_suggestions_path=suggestions_path,
            unassigned_isins_path=unassigned_path,
        )

    # ── 12b. Build P&L workbook (with optional ValuationA embed) ─────────────
    pl_lookup = _build_pl_lookup(pl_df)
    if progress is not None:
        try:
            progress(StageResult(
                "build-pl-workbook", True,
                f"starting (writing {len(pl_df)} positions, {len(income_df)} income rows, {len(trades_df)} trades)…",
            ))
        except Exception:
            pass
    try:
        build_workbook(
            out_path=pl_report_path,
            pl_df=pl_df, income_df=income_df, trades_df=trades_df,
            validation=validation, period_meta=period_meta,
            valuation_rows=val_rows,
            applied_ca_overrides=tuple(ca_overrides),
            ca_override_sources=tuple(
                Path(p) for p in (cfg.ca_overrides_path, *options.extra_ca_override_paths)
            ),
            th_val_df=th_val_df,
            analyst_sname_overrides=dict(analyst_map.sname_overrides),
        )
        stages.append(StageResult(
            "build-pl-workbook", True, str(pl_report_path),
            {"path": str(pl_report_path)},
        ))
    except Exception as exc:
        stages.append(StageResult("build-pl-workbook", False, str(exc)))
        return _abort("build-pl-workbook")

    # ── 12c. Archive workbook to per-month OneDrive folder ───────────────────
    # Disabled per user request 2026-05-05 — re-enable by uncommenting.
    # archive_path = _archive_pl_report(
    #     cfg.report_archive_root, pl_report_path, pd.Timestamp(end_vdate), stages,
    # )
    archive_path = None

    # ── 13. (removed) Standalone ValuationA workbook ─────────────────────────
    # Previously emitted ``current portfolio.xlsx``; the same data is now
    # embedded directly in the main P&L workbook (Summary + ValuationA +
    # per-analyst sheets), so a duplicate file would only confuse readers.

    # ── 14. Send email (optional) ────────────────────────────────────────────
    email_result: EmailSendResult | None = None
    if options.send_email:
        email_result = _send_email(
            options=options, cfg=cfg, period_meta=period_meta,
            attachments=(pl_report_path.resolve(),),
            valuation_path=pl_report_path,
            outlook_dispatcher=outlook_dispatcher,
        )
        stages.append(StageResult(
            "send-email", email_result.sent,
            email_result.error or f"to={email_result.to}",
            {"sent": email_result.sent, "to": email_result.to,
             "attachment_count": email_result.attachment_count},
        ))

    # ── 15. Save analyst map ─────────────────────────────────────────────────
    _save_analyst_map_safely(analyst_map, stages)
    unassigned_path = _write_unassigned_isins(pl_df, output_dir, ts_str, stages)

    # ── 16. Append applied CA overrides to history (audit + recurrence) ─────
    if ca_overrides:
        try:
            append_history(
                cfg.ca_history_path, ca_overrides,
                period_end=pd.Timestamp(end_vdate),
            )
            stages.append(StageResult(
                "append-ca-history", True,
                f"appended={len(ca_overrides)} -> {cfg.ca_history_path.name}",
                {"appended": len(ca_overrides)},
            ))
        except Exception as exc:
            stages.append(StageResult("append-ca-history", False, str(exc)))

    _write_diagnostics()
    return PipelineResult(
        ok=True, stages=tuple(stages),
        pl_df=pl_df, income_df=income_df, validation=validation,
        pl_report_path=pl_report_path, valuation_path=None,
        archive_path=archive_path,
        results_parquet_path=results_parquet,
        email_result=email_result,
        unassigned_isins_path=unassigned_path,
        end_date=pd.Timestamp(end_vdate),
    )


# ─── Helpers ─────────────────────────────────────────────────────────────────

def _to_yyyymmdd(ts: pd.Timestamp) -> str:
    return pd.Timestamp(ts).strftime("%Y%m%d")


def _resolve_latest_sql_vdate(
    *, conn_str: str, fund_pcodes: tuple[str, ...],
) -> str:
    """Return the latest VDATE (YYYYMMDD) the SQL view holds for OEFOF.

    Used when the caller didn't pass ``--end-yyyymmdd``: equivalent to the
    legacy "open HP_VAL.xlsm and read the VDATE column" behaviour.
    Tiny query — single date — so it adds milliseconds, not seconds.
    """
    try:
        import pyodbc  # type: ignore
    except ImportError as e:  # pragma: no cover
        raise OefofError(
            "pyodbc is required for SQL snapshot loading."
        ) from e
    pcodes = tuple(p.strip().upper() for p in fund_pcodes if str(p).strip())
    if not pcodes:
        raise OefofError("fund_pcodes is empty; cannot resolve latest VDATE")
    placeholders = ",".join("?" for _ in pcodes)
    sql = (
        "SELECT MAX(VDATE) FROM dbo.vw_RPT_VAL "
        f"WHERE UPPER(PCODE) IN ({placeholders})"
    )
    try:
        with pyodbc.connect(conn_str, timeout=15) as cn:
            row = cn.cursor().execute(sql, pcodes).fetchone()
    except pyodbc.Error as e:
        raise OefofError(f"vw_RPT_VAL: latest-VDATE probe failed ({e})") from e
    if not row or row[0] is None:
        raise OefofError(
            f"vw_RPT_VAL: no rows for OEFOF PCODEs {pcodes}; "
            "server has no snapshot to use as 'latest'."
        )
    return pd.Timestamp(row[0]).strftime("%Y%m%d")


def _archive_pl_report(
    archive_root: Path,
    src_path: Path,
    end_vdate: pd.Timestamp,
    stages: list,
) -> Path | None:
    """Copy the published P&L workbook into the per-month OneDrive archive.

    Layout: ``<archive_root>/<YYYY-MM Month>/P&L-OEFOF-<YYYYMMDD>.xlsx``.
    The end-of-period date drives both the month folder and the filename so
    re-runs overwrite the same file rather than spawning duplicates.
    Failures are recorded as a non-fatal stage and never abort the pipeline.
    """
    import shutil

    end = pd.Timestamp(end_vdate)
    month_dir = archive_root / end.strftime("%Y-%m %B")
    dest = month_dir / f"P&L-OEFOF-{end.strftime('%Y%m%d')}.xlsx"
    try:
        month_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_path, dest)
        stages.append(StageResult(
            "archive-pl-workbook", True, str(dest),
            {"path": str(dest), "month_dir": str(month_dir)},
        ))
        return dest
    except Exception as exc:
        stages.append(StageResult(
            "archive-pl-workbook", False, f"{type(exc).__name__}: {exc}",
        ))
        return None


def _filter_trade_window(
    trades_df: pd.DataFrame, start_vdate: pd.Timestamp, end_vdate: pd.Timestamp,
) -> pd.DataFrame:
    """Keep trades with start_vdate < CDATE ≤ end_vdate.

    Trades on the start date itself are *excluded* (they are part of the
    start-of-day position). Trades on the end date are included.
    """
    if trades_df.empty or "CDATE" not in trades_df.columns:
        return trades_df
    cdate = pd.to_datetime(trades_df["CDATE"], errors="coerce")
    start = pd.Timestamp(start_vdate).normalize()
    end = pd.Timestamp(end_vdate).normalize()
    mask = (cdate.dt.normalize() > start) & (cdate.dt.normalize() <= end)
    return trades_df.loc[mask].reset_index(drop=True)


def _filter_snapshot_to_pcodes(
    port_df: pd.DataFrame, fund_pcodes: tuple[str, ...],
) -> pd.DataFrame:
    """Keep only rows whose ``PCODE_ORIG`` is in the OEFOF set.

    HP_VAL is firm-wide. Without this filter, ``aggregate_portfolio``
    would sum units across funds we have no trades for, breaking
    every per-ISIN reconciliation downstream.
    """
    if port_df.empty or "PCODE_ORIG" not in port_df.columns:
        return port_df
    pcodes_upper = {str(p).strip().upper() for p in fund_pcodes}
    mask = port_df["PCODE_ORIG"].astype(str).str.strip().str.upper().isin(pcodes_upper)
    return port_df.loc[mask].reset_index(drop=True)


def _compute_nav_components(
    end_df_oefof: pd.DataFrame, pl_df: pd.DataFrame,
) -> tuple[tuple[str, float, bool], ...] | None:
    """Build the NAV reconciliation rows for the Summary sheet.

    Bridges the equity-only stock MV (what our P&L engine reports) to the
    HP_VAL fund portfolio total (sum of every PTVALUE_EUR row HiPort
    carries for the OEFOF PCODE family). The gap is non-equity items —
    cash, FX equivalents, accrued fees, swap P&L cash buckets, dividend
    receivables — which are intentionally outside the stock-level model.

    Returns ``None`` if the snapshot is missing the required columns
    (lets unit tests using stub frames skip rendering).
    """
    if end_df_oefof.empty:
        return None
    needed = {"PTVALUE_EUR", "ISIN", "CAT", "SNAME"}
    if not needed.issubset(end_df_oefof.columns):
        return None

    df = end_df_oefof.copy()
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


def _compute_nav_total(snap_df: pd.DataFrame) -> float | None:
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


def _build_snapshot_exposure(
    snap_df: pd.DataFrame,
    pl_df: pd.DataFrame,
    as_of: pd.Timestamp,
) -> ExposureBlock:
    """Build point-in-time long/short/net/gross exposure from the end snapshot."""
    if snap_df is None or snap_df.empty:
        return ExposureBlock(label="Current Snapshot", as_of=as_of)

    analyst_keys = _build_analyst_lookup(pl_df)
    long_mask = snap_df["LS"].astype(str).str.upper().ne("S")
    short_mask = snap_df["LS"].astype(str).str.upper().eq("S")
    ptvalue = pd.to_numeric(snap_df["PTVALUE_EUR"], errors="coerce").fillna(0.0)
    magnitude = _snapshot_exposure_magnitude(snap_df)
    nav_eur = float(ptvalue.sum())

    by_analyst: dict[str, ExposureMetrics] = {}
    work = snap_df.copy()
    work["__MAGNITUDE__"] = magnitude
    work["__ANALYST__"] = work.apply(
        lambda row: _resolve_snapshot_analyst(row, analyst_keys), axis=1,
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


def _build_weighted_exposure(
    history_df: pd.DataFrame,
    nav_history_df: pd.DataFrame,
    pl_df: pd.DataFrame,
) -> ExposureBlock:
    """AUM-weighted daily average exposure over the pipeline period."""
    if history_df.empty or nav_history_df.empty:
        return ExposureBlock(label="YTD AUM-Weighted Average Exposure")

    analyst_keys = _build_analyst_lookup(pl_df)
    work = history_df.copy()
    work["VDATE"] = pd.to_datetime(work["VDATE"], errors="coerce").dt.normalize()
    work["__MAGNITUDE__"] = _snapshot_exposure_magnitude(work)
    work["__SHORT__"] = work["LS"].astype(str).str.upper().eq("S")
    work["__ANALYST__"] = work.apply(
        lambda row: _resolve_snapshot_analyst(row, analyst_keys), axis=1,
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


def _build_analyst_lookup(pl_df: pd.DataFrame) -> dict[str, dict[str, str]]:
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


def _resolve_snapshot_analyst(row: pd.Series, analyst_keys: dict[str, dict[str, str]]) -> str:
    isin = str(row.get("ISIN", "")).strip().upper()
    sname = str(row.get("SNAME", "")).strip().upper()
    analyst = analyst_keys["isin"].get(isin) or analyst_keys["name"].get(sname)
    return analyst or "UNASSIGNED"


def _snapshot_exposure_magnitude(df: pd.DataFrame) -> pd.Series:
    ptvalue = pd.to_numeric(df.get("PTVALUE_EUR", 0.0), errors="coerce").fillna(0.0).abs()
    if "EXPOSURE" in df.columns:
        exposure = pd.to_numeric(df["EXPOSURE"], errors="coerce").fillna(0.0).abs()
    else:
        exposure = pd.Series(0.0, index=df.index, dtype=float)
    cat = df.get("CAT", "").astype(str).str.upper()
    use_exposure = cat.isin(["FUT", "SWAP", "FTSWAP"])
    return exposure.where(use_exposure & exposure.gt(0.0), ptvalue)


def _enrich_analysts_from_valuation(
    val_rows: Sequence[Sequence[object]],
    pl_df: pd.DataFrame,
    analyst_map: AnalystMap,
    *,
    end_port_df: pd.DataFrame | None = None,
) -> tuple[int, int]:
    """Use HiPort ValuationA's carried-down analyst initials to fill in
    pl_df rows still tagged ``UNASSIGNED`` and persist new ISIN→analyst
    rows back to the analyst map.

    Two sources, in order of priority:

      1. ``end_port_df`` — the raw ``val_data`` snapshot. Column
         ``Initials`` (col AS) is HiPort's per-row analyst code keyed by
         ISIN (col AL). This is the authoritative ISIN→analyst map for
         every position currently in the snapshot.
      2. ValuationA pivot (``val_rows``) — initials in col B carried down
         per group. Keyed by SNAME since the pivot has no ISIN. Used as a
         fallback for positions that exited (no longer in the snapshot)
         but still appear in ``pl_df``.

    Every currently-held ISIN (``Ending Units != 0``) is also persisted
    into the analyst map so the CSV reflects the live portfolio next run,
    even if the analyst is still unknown.

    Returns ``(reassigned, recorded)`` for the stage message.
    """
    if pl_df.empty:
        return (0, 0)

    # Source 1 — direct ISIN → Initials from val_data.
    isin_to_analyst: dict[str, str] = {}
    if end_port_df is not None and not end_port_df.empty \
            and "ISIN" in end_port_df.columns and "Initials" in end_port_df.columns:
        for isin_v, init_v in zip(
            end_port_df["ISIN"].tolist(),
            end_port_df["Initials"].tolist(),
        ):
            try:
                isin_k = str(isin_v).strip().upper()
                init_k = str(init_v).strip()
            except Exception:
                continue
            if not isin_k or isin_k in ("NAN", "NONE"):
                continue
            if not init_k or init_k.lower() in ("nan", "(blank)", "none"):
                continue
            isin_to_analyst.setdefault(isin_k, init_k)

    sname_to_analyst: dict[str, str] = {}
    current_sort1: str | None = None
    current_initials: str | None = None
    for idx, row in enumerate(val_rows):
        if idx + 1 < _VAL_DATA_FIRST_ROW:
            continue
        row_padded = list(row) + [None] * (
            max(0, _VAL_HPV_FUTVAL_IDX + 1 - len(row))
        )
        sort1 = row_padded[_VAL_HPV_SORT1_IDX]
        if sort1 not in (None, ""):
            current_sort1 = str(sort1).strip()
        if sort1 is not None and "Total" in str(sort1):
            current_initials = None
        initials = row_padded[_VAL_HPV_INITIALS_IDX]
        if initials not in (None, ""):
            txt = str(initials).strip()
            # HiPort uses "(blank)" as a group label for un-tagged rows.
            # Treat it like no initials so we don't persist it as an
            # analyst code or override good map entries with it.
            if txt and txt.lower() != "(blank)":
                current_initials = txt
            else:
                current_initials = None
        if current_sort1 != "A":
            continue
        sname = row_padded[_VAL_HPV_SECURITY_NAME_IDX]
        if not sname or not current_initials:
            continue
        key = str(sname).strip().upper()
        if key:
            sname_to_analyst.setdefault(key, current_initials)

    reassigned = 0
    recorded = 0
    for i in pl_df.index:
        cur = str(pl_df.at[i, "Analyst"]).strip().upper()
        isin_val = str(pl_df.at[i, "ISIN"]).strip().upper()
        sname_val = str(pl_df.at[i, "Stock Name"]).strip().upper()
        # Prefer val_data ISIN match; fall back to pivot SNAME match.
        new_a = isin_to_analyst.get(isin_val) or sname_to_analyst.get(sname_val)
        if cur == UNASSIGNED and new_a:
            pl_df.at[i, "Analyst"] = new_a
            reassigned += 1
            cur = new_a
        # Persist any currently-held ISIN into the map so the CSV mirrors
        # the live portfolio. UNASSIGNED entries surface unmapped names.
        try:
            held = float(pl_df.at[i, "Ending Units"]) != 0.0
        except (TypeError, ValueError):
            held = False
        if not held:
            continue
        isin = str(pl_df.at[i, "ISIN"]).strip().upper()
        if not isin:
            continue
        existing = analyst_map.by_isin.get(isin)
        if existing != cur:
            analyst_map.by_isin[isin] = cur
            analyst_map._dirty = True
            recorded += 1
    return (reassigned, recorded)


def _build_pl_lookup(pl_df: pd.DataFrame) -> dict[str, list[float]]:
    """Map ISIN → 8 P&L values (M:T) for the ValuationA workbook.

    Order matches ``oefof_pl.output.valuation.VAL_HEADERS_M_TO_T``:
    Realised EUR, Realised %, Unrealised EUR, Unrealised %,
    Income EUR, Income Yield %, Total EUR, Total %.
    """
    if pl_df.empty:
        return {}
    cols = (
        "Realised P&L (EUR)", "Realised P&L (%)",
        "Unrealised P&L (EUR)", "Unrealised P&L (%)",
        "Income (EUR)", "Income Yield (%)",
        "Total P&L (EUR)", "Total P&L (%)",
    )
    missing = [c for c in cols if c not in pl_df.columns]
    if missing:
        raise KeyError(f"_build_pl_lookup: pl_df missing columns {missing}")
    out: dict[str, list[float]] = {}
    for _, row in pl_df.iterrows():
        isin = str(row.get("ISIN", "")).strip()
        if not isin:
            continue
        out[isin] = [float(row[c]) if pd.notna(row[c]) else 0.0 for c in cols]
    return out


def _write_parquet_atomic(df: pd.DataFrame, path: Path) -> None:
    """Write `df` to parquet via a sibling tempfile + os.replace."""
    import os
    import tempfile
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_str = tempfile.mkstemp(
        prefix=path.stem + ".", suffix=".parquet.tmp", dir=str(path.parent),
    )
    os.close(fd)
    tmp = Path(tmp_str)
    try:
        df.to_parquet(tmp, index=False)
        os.replace(tmp, path)
    finally:
        try:
            if tmp.exists():
                tmp.unlink()
        except OSError:
            pass


def _save_analyst_map_safely(analyst_map: AnalystMap, stages: list[StageResult]) -> None:
    try:
        analyst_map.save()
        stages.append(StageResult(
            "save-analyst-map", True,
            f"entries={len(analyst_map.by_isin)} dirty={analyst_map.dirty}",
            {"entries": len(analyst_map.by_isin)},
        ))
    except Exception as exc:
        stages.append(StageResult("save-analyst-map", False, str(exc)))


def _write_unassigned_isins(
    pl_df: pd.DataFrame, output_dir: Path, ts_str: str,
    stages: list[StageResult],
) -> Path | None:
    """Emit out/unassigned_isins_<ts>.csv listing positions still UNASSIGNED.

    Format: ISIN, SNAME, Status (Current/Exited), Ending Units, ANALYST.
    The ANALYST column is left blank for the user to fill in; the file is
    designed to be merged back into ``isin_analyst_map.csv`` row-by-row.
    Returns the path written, or None if nothing to write.
    """
    import csv as _csv
    if pl_df is None or pl_df.empty or "Analyst" not in pl_df.columns:
        return None
    try:
        sub = pl_df[pl_df["Analyst"].astype(str).str.upper() == UNASSIGNED]
        if sub.empty:
            return None
        path = output_dir / f"unassigned_isins_{ts_str}.csv"
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8") as fh:
            w = _csv.writer(fh)
            w.writerow(("ISIN", "SNAME", "Status", "Ending Units", "ANALYST"))
            for _, r in sub.sort_values("Stock Name").iterrows():
                ending = r.get("Ending Units", 0)
                status = "Current" if ending not in (0, 0.0) else "Exited"
                w.writerow((
                    str(r.get("ISIN", "")), str(r.get("Stock Name", "")),
                    status, ending, "",
                ))
        stages.append(StageResult(
            "write-unassigned-isins", True,
            f"file={path.name} rows={len(sub)}",
            {"path": str(path), "rows": int(len(sub))},
        ))
        return path
    except Exception as exc:
        stages.append(StageResult("write-unassigned-isins", False, str(exc)))
        return None


def _send_email(
    *,
    options: PipelineOptions,
    cfg: Config,
    period_meta: PeriodMeta,
    attachments: tuple[Path, ...],
    valuation_path: Path,
    outlook_dispatcher,
) -> EmailSendResult:
    to = normalize_recipients(options.email_to or cfg.email_to_default)
    cc = normalize_recipients(options.email_cc)
    subject = options.email_subject or cfg.email_subject_default
    summary_html = build_summary_html_table(valuation_path)
    footer = (
        f"Period: {period_meta.start_date.date()} → {period_meta.end_date.date()}",
        f"Source: {period_meta.snapshot_path}",
        f"Computed at: {period_meta.run_timestamp:%Y-%m-%d %H:%M}",
    )
    html_body = build_email_html_body(options.email_body, summary_html, footer_lines=footer)
    msg = EmailMessage(
        to=to, cc=cc, subject=subject,
        body_text=options.email_body, body_html=html_body,
        attachments=attachments, from_smtp=options.email_from_smtp,
    )
    kwargs = {}
    if outlook_dispatcher is not None:
        kwargs["dispatcher"] = outlook_dispatcher
    return send_via_outlook(msg, **kwargs)
