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

from dataclasses import replace
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence

import pandas as pd

from .analyst.map import UNASSIGNED, AnalystMap
from .compute.aggregate import AggLine, aggregate_portfolio
from .compute.ca_split import split_ca_rows
from .compute.exposure import (
    ExposureBlock,
    ExposureMetrics,
    PeriodMeta,
    build_snapshot_exposure as _build_snapshot_exposure,
    build_weighted_exposure as _build_weighted_exposure,
    compute_nav_total as _compute_nav_total,
)
from .compute.fx import (
    build_fx_from_trades,
    build_fx_supplement,
    build_fx_table,
    local_to_eur_at,
    resolve_eur_usd,
)
from .compute.pl import compute_positions
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
from .data.arb_pairs_loader import load_arb_pairs
from .data.live_prices import fetch_live_prices, load_yahoo_ticker_map
from .data.openfigi import lookup_tickers as _openfigi_lookup
from .data.loader import assert_vdate_matches
from .data.sql_loader import (
    fetch_eur_usd_rate,
    fetch_missing_fx_rates,
    load_snapshot_history_sql,
    load_snapshot_sql,
    load_th_val_sql,
    load_trades_sql,
)
from .email_sender.outlook import EmailSendResult
from .exceptions import OefofError
from .output.valuation import (
    HPV_FMCID_IDX as _VAL_HPV_FMCID_IDX,
    close_open_excel_workbook,
    extract_valuation_a_via_sql,
)
from .output.workbook import build_workbook
from .pipeline_helpers import (
    archive_pl_report as _archive_pl_report,
    build_pl_lookup as _build_pl_lookup,
    enrich_analysts_from_valuation as _enrich_analysts_from_valuation,
    filter_snapshot_to_pcodes as _filter_snapshot_to_pcodes,
    filter_trade_window as _filter_trade_window,
    resolve_latest_sql_vdate as _resolve_latest_sql_vdate,
    save_analyst_map_safely as _save_analyst_map_safely,
    send_email as _send_email,
    send_unassigned_alert_email as _send_unassigned_alert_email,
    to_yyyymmdd as _to_yyyymmdd,
    write_parquet_atomic as _write_parquet_atomic,
    write_unassigned_isins as _write_unassigned_isins,
)
from .pipeline_types import (  # noqa: F401  (re-exported for test imports)
    PipelineOptions,
    PipelineResult,
    StageResult,
)
from .validate import ValidationResult, run_all_checks


# ─── Output filenames ────────────────────────────────────────────────────────

PL_REPORT_FILENAME = "OAKS_EM_Stock_PL_Report.xlsx"
RESULTS_PARQUET = "pl_results.parquet"
FAILURE_PARQUET_PREFIX = "pl_results_failed_"
FAILURE_XLSX_PREFIX = "VALIDATION_FAILURE_"
DIAGNOSTICS_PREFIX = "pnl_diagnostics_"


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

    def _live_prices_to_eur(
        live_prices_local: Mapping[str, float],
    ) -> dict[str, float]:
        """Convert live prices in local CCY into EUR for workbook display."""
        if not live_prices_local or pl_df.empty:
            return {}

        isin_series = pl_df.get("ISIN", pd.Series(dtype=object)).astype(str).str.strip().str.upper()
        ccy_series = pl_df.get("CCY", pd.Series(dtype=object)).astype(str).str.strip().str.upper()
        isin_to_ccy = {
            isin: ccy
            for isin, ccy in zip(isin_series, ccy_series)
            if isin and ccy
        }
        fx_effective = {**build_fx_from_trades(trades_df), **dict(fx_start), **dict(fx_end)}

        out: dict[str, float] = {}
        for isin_raw, px_local_raw in live_prices_local.items():
            isin = str(isin_raw).strip().upper()
            ccy = isin_to_ccy.get(isin)
            if not isin or not ccy:
                continue
            try:
                px_local = float(px_local_raw)
            except (TypeError, ValueError):
                continue
            xrate = 0.0 if ccy in {"EUR", "USD"} else float(fx_effective.get(ccy, float("nan")))
            try:
                out[isin] = float(local_to_eur_at(px_local, ccy, xrate, eur_usd_end))
            except Exception:
                continue
        return out

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

    # ── 8b. Resolve tickers + fetch live prices (open positions) ────────────
    # Done BEFORE compute_positions so live prices feed into the P&L math,
    # not just the display layer.
    #
    # Step 1: collect open-position ISINs from end_agg (units != 0)
    open_isins = [
        isin for isin, agg in end_agg.items()
        if agg.units != 0
    ]
    # Step 2: build ISIN->CCY map for OpenFIGI exchange selection
    isin_to_ccy = {isin: end_agg[isin].ccy for isin in open_isins}

    # Step 3: load manual ticker overrides
    try:
        yahoo_tickers_manual = load_yahoo_ticker_map(cfg.yahoo_tickers_path)
        stages.append(StageResult(
            "load-yahoo-tickers", True,
            f"rows={len(yahoo_tickers_manual)}",
            {"rows": len(yahoo_tickers_manual)},
        ))
    except Exception as exc:
        yahoo_tickers_manual = {}
        stages.append(StageResult(
            "load-yahoo-tickers", True,
            f"skipped ({exc})",
            {"error": str(exc)},
        ))

    # Step 4: OpenFIGI auto-lookup for ISINs not covered by manual overrides
    openfigi_tickers: dict[str, str] = {}
    try:
        openfigi_cache = cfg.yahoo_tickers_path.parent / "openfigi_cache.json"
        openfigi_tickers = _openfigi_lookup(
            open_isins,
            isin_to_ccy=isin_to_ccy,
            cache_path=openfigi_cache,
            manual_overrides=yahoo_tickers_manual,
        )
        stages.append(StageResult(
            "openfigi-lookup", True,
            f"resolved={len(openfigi_tickers)}/{len(open_isins)} open ISINs",
            {"resolved": len(openfigi_tickers), "open_isins": len(open_isins)},
        ))
    except Exception as exc:
        stages.append(StageResult(
            "openfigi-lookup", True,
            f"skipped ({exc})",
            {"error": str(exc)},
        ))

    # Step 5: fetch live prices for all resolvable open-position ISINs
    live_prices_eur: dict[str, float] = {}
    try:
        arb_pairs = load_arb_pairs(cfg.arb_pairs_path)
    except Exception as exc:
        stages.append(StageResult("load-arb-pairs", False, str(exc)))
        arb_pairs = ()

    try:
        live_prices_local = fetch_live_prices(
            arb_pairs,
            yahoo_tickers=yahoo_tickers_manual,
            open_isins=open_isins,
            openfigi_tickers=openfigi_tickers,
        )
        # Convert local CCY prices to EUR inline — can't use _live_prices_to_eur
        # here because pl_df doesn't exist yet. Use isin_to_ccy from end_agg instead.
        fx_effective = {**build_fx_from_trades(trades_df), **dict(fx_start), **dict(fx_end)}
        live_prices_eur = {}
        for isin_raw, px_local in live_prices_local.items():
            isin = str(isin_raw).strip().upper()
            ccy = isin_to_ccy.get(isin, "")
            if not isin or not ccy:
                continue
            try:
                xrate = 0.0 if ccy in {"EUR", "USD"} else float(fx_effective.get(ccy, float("nan")))
                live_prices_eur[isin] = float(local_to_eur_at(float(px_local), ccy, xrate, eur_usd_end))
            except Exception:
                continue
        stages.append(StageResult(
            "fetch-live-prices", True,
            f"fetched={len(live_prices_local)} converted_to_eur={len(live_prices_eur)}",
            {"fetched": len(live_prices_local), "converted": len(live_prices_eur)},
        ))
    except Exception as exc:
        live_prices_eur = {}
        stages.append(StageResult(
            "fetch-live-prices", True,
            f"skipped ({exc})",
            {"error": str(exc)},
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
            live_prices_eur=live_prices_eur or None,
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
    snapshot_exposure = _build_snapshot_exposure(end_df_oefof, pl_df, pd.Timestamp(end_vdate))
    ytd_weighted_exposure: ExposureBlock | None = None
    try:
        history_df = load_snapshot_history_sql(
            start_yyyymmdd=start_yyyymmdd,
            end_yyyymmdd=end_yyyymmdd,
            conn_str=cfg.sql_conn_str,
            fund_pcodes=cfg.fund_pcodes,
        )
        ytd_weighted_exposure = _build_weighted_exposure(history_df, pl_df)
        stages.append(StageResult(
            "load-exposure-history", True,
            f"days={history_df['VDATE'].nunique()} snapshot_rows={len(history_df)}",
            {"days": int(history_df['VDATE'].nunique()), "rows": len(history_df)},
        ))
    except Exception as exc:
        stages.append(StageResult(
            "load-exposure-history", True,
            f"skipped ({exc})",
            {"error": str(exc)},
        ))
    start_nav_eur = _compute_nav_total(start_df)
    period_meta = PeriodMeta(
        fund_name=cfg.fund_name,
        start_date=pd.Timestamp(start_vdate),
        end_date=pd.Timestamp(end_vdate),
        fund_currency=cfg.fund_currency,
        snapshot_path="ccl.dbo.vw_RPT_VAL @ FCEDATA01",
        bottler_path="ccl.dbo.tTRANS @ FCEDATA01",
        run_timestamp=pipeline_start_time,
        snapshot_exposure=snapshot_exposure,
        ytd_weighted_exposure=ytd_weighted_exposure,
        start_nav_eur=start_nav_eur,
        analyst_codes=dict(cfg.analyst_codes),
    )
    ts_str = pipeline_start_time.strftime("%Y%m%d_%H%M%S")

    # live_prices_eur already computed in stage 8b above and passed into
    # compute_positions. Re-use it here for the workbook display layer too
    # (live prices now affect both P&L math AND the Last Price column).
    live_prices = live_prices_eur  # alias for workbook call sites below

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
                arb_pairs=arb_pairs,
                live_prices=live_prices,
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
            arb_pairs=arb_pairs,
            live_prices=live_prices,
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

    # ── 14a. Write UNASSIGNED CSV (if any) before email sending ─────────────
    unassigned_path = _write_unassigned_isins(pl_df, output_dir, ts_str, stages)

    # ── 14b. Send email (optional) ───────────────────────────────────────────
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
        if unassigned_path is not None:
            alert_result = _send_unassigned_alert_email(
                unassigned_csv_path=unassigned_path,
                period_meta=period_meta,
                outlook_dispatcher=outlook_dispatcher,
                from_smtp=options.email_from_smtp,
            )
            stages.append(StageResult(
                "send-unassigned-alert", alert_result.sent,
                alert_result.error or f"to={alert_result.to}",
                {"sent": alert_result.sent, "to": alert_result.to,
                 "attachment_count": alert_result.attachment_count},
            ))

    # ── 15. Save analyst map ─────────────────────────────────────────────────
    _save_analyst_map_safely(analyst_map, stages)

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