"""Tests for ``oefof_pl.pipeline``.

These tests exercise the orchestrator end-to-end with monkey-patched
data loaders, COM dispatchers, and ValuationA extractor. The goal is to
verify ordering, branching (critical → no P&L workbook), artefact paths,
and that injectable hooks are invoked correctly.
"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pandas as pd
import pytest

from oefof_pl import pipeline as P
from oefof_pl.compute.aggregate import AggLine
from oefof_pl.config import load_default_config
from oefof_pl.email_sender.outlook import EmailSendResult
from oefof_pl.exceptions import DataLoadError
from oefof_pl.pipeline import (
    PL_REPORT_FILENAME,
    RESULTS_PARQUET,
    PipelineOptions,
    PipelineResult,
    StageResult,
    run_pipeline,
)


# ─── Fixture builders ───────────────────────────────────────────────────────

START_VDATE = pd.Timestamp("2025-12-31")
END_VDATE = pd.Timestamp("2026-04-29")


def _port_df(*, vdate: pd.Timestamp, units: float = 100.0) -> pd.DataFrame:
    """Build a minimal post-loader portfolio DataFrame (PTVALUE_EUR etc.)."""
    return pd.DataFrame([
        {
            "ISIN": "EUR0000000001", "SNAME": "ALPHA EUR", "CCY": "EUR",
            "XRATE": 1.0, "UNITS": units,
            "PTVALUE_EUR": 120.0 * units, "PTCOST_EUR": 100.0 * units,
            "NUCOST_LOCAL": 100.0, "LST_PRICE_LOCAL": 120.0,
            "CAT": "ORD", "LS": "L", "PCODE_ORIG": "OEFOF",
            "VDATE": vdate, "EXCODE1": "DE",
        },
        {
            "ISIN": "EUR0000000001", "SNAME": "ALPHA EUR", "CCY": "EUR",
            "XRATE": 1.0, "UNITS": 0.0,
            "PTVALUE_EUR": 0.0, "PTCOST_EUR": 0.0,
            "NUCOST_LOCAL": 0.0, "LST_PRICE_LOCAL": 120.0,
            "CAT": "ORD", "LS": "L", "PCODE_ORIG": "OEFOGSSC",
            "VDATE": vdate, "EXCODE1": "DE",
        },
    ])


def _trades_df() -> pd.DataFrame:
    """One mid-period buy of 50 units at €110 — keeps things simple."""
    return pd.DataFrame([{
        "TRADE_ID": 1,
        "PCODE_ORIG": "OEFOF", "ISIN": "EUR0000000001", "SNAME": "ALPHA EUR",
        "CCY": "EUR", "T": "P",
        "CDATE": pd.Timestamp("2026-02-15"),
        "UNITS": 50.0,
        "GROSSPRICE_LOCAL": 110.0,
        "BUY_NET_LOCAL": 5_500.0, "SELL_NET_LOCAL": 0.0,
        "INCOME_LOCAL": 0.0,
    }])


def _patch_loaders(monkeypatch, *, start_df=None, end_df=None, trades_df=None,
                   start_vdate=START_VDATE, end_vdate=END_VDATE):
    """Install in-memory replacements for the three loaders."""
    # Buy of 50 units in trades → end units must be start + 50 to satisfy VAL-01.
    s = start_df if start_df is not None else _port_df(vdate=start_vdate, units=100.0)
    e = end_df if end_df is not None else _port_df(vdate=end_vdate, units=150.0)
    t = trades_df if trades_df is not None else _trades_df()

    def fake_load_snapshot_sql(yyyymmdd, *, conn_str, oefof_pcodes):
        # Disambiguate by requested date: anything matching the start vdate
        # returns the start frame, anything else the end frame.
        if pd.Timestamp(yyyymmdd) == pd.Timestamp(start_vdate).normalize():
            return s, start_vdate
        return e, end_vdate

    def fake_resolve_latest(*, conn_str, oefof_pcodes):
        return pd.Timestamp(end_vdate).strftime("%Y%m%d")

    def fake_load_trades_sql(*, conn_str, oefof_pcodes):
        return t.copy()

    monkeypatch.setattr(P, "load_snapshot_sql", fake_load_snapshot_sql)
    monkeypatch.setattr(P, "_resolve_latest_sql_vdate", fake_resolve_latest)
    monkeypatch.setattr(P, "load_trades_sql", fake_load_trades_sql)


def _make_cfg(tmp_path: Path):
    cfg = load_default_config()
    return replace(
        cfg,
        sql_conn_str="DRIVER={fake};",  # never reached due to monkeypatch
        bottler_path=tmp_path / "bottler.xlsb",
        analyst_map_path=tmp_path / "analyst.csv",
        output_dir=tmp_path / "out",
        ca_overrides_path=tmp_path / "ca_overrides.csv",
        ca_history_path=tmp_path / "ca_overrides_history.csv",
    )


# ─── Helpers for assertions ─────────────────────────────────────────────────

def _stage_names(result: PipelineResult) -> list[str]:
    return [s.name for s in result.stages]


# ─── Happy path ─────────────────────────────────────────────────────────────

def test_pipeline_happy_path_writes_all_artifacts(monkeypatch, tmp_path: Path):
    _patch_loaders(monkeypatch)
    cfg = _make_cfg(tmp_path)
    options = PipelineOptions(start_yyyymmdd="20251231")

    result = run_pipeline(
        cfg, options,
        extractor=lambda hp_val_path, *, target_pname: [[None]*12 for _ in range(8)],
        closer=lambda p: False,
    )

    assert result.ok is True, _stage_names(result) + [s.message for s in result.stages]
    assert result.pl_report_path == cfg.output_dir / PL_REPORT_FILENAME
    assert result.valuation_path is None
    assert result.results_parquet_path == cfg.output_dir / RESULTS_PARQUET
    assert result.pl_report_path.exists()
    assert result.results_parquet_path.exists()

    # Stage ordering — refresh-bottler skipped, end loaded before start.
    names = _stage_names(result)
    assert names[:2] == ["load-end-snapshot", "load-start-snapshot"]
    assert "compute-positions" in names
    assert "validate" in names
    assert "build-pl-workbook" in names
    assert "save-analyst-map" in names

    diag_files = list(cfg.output_dir.glob("pnl_diagnostics_*.csv"))
    assert len(diag_files) == 1
    diag = pd.read_csv(diag_files[0])
    assert "load-trades" in set(diag["stage"])
    assert "compute-positions" in set(diag["stage"])


def test_pipeline_uses_vdate_when_end_yyyymmdd_blank(monkeypatch, tmp_path: Path):
    _patch_loaders(monkeypatch)
    cfg = _make_cfg(tmp_path)
    result = run_pipeline(
        cfg, PipelineOptions(start_yyyymmdd="20251231"),
        extractor=lambda hp_val_path, *, target_pname: [[None]*12 for _ in range(8)],
        closer=lambda p: False,
    )
    assert result.ok
    end_stage = result.by_stage("load-end-snapshot")
    assert end_stage.artifacts["yyyymmdd"] == "20260429"


def test_pipeline_propagates_explicit_end_yyyymmdd(monkeypatch, tmp_path: Path):
    _patch_loaders(monkeypatch)
    cfg = _make_cfg(tmp_path)
    # Provide the same date as VDATE — must succeed (assert_vdate_matches passes).
    result = run_pipeline(
        cfg, PipelineOptions(start_yyyymmdd="20251231", end_yyyymmdd="20260429"),
        extractor=lambda hp_val_path, *, target_pname: [[None]*12 for _ in range(8)],
        closer=lambda p: False,
    )
    assert result.ok


def test_pipeline_rejects_end_yyyymmdd_when_vdate_disagrees(monkeypatch, tmp_path: Path):
    _patch_loaders(monkeypatch)
    cfg = _make_cfg(tmp_path)
    result = run_pipeline(
        cfg, PipelineOptions(start_yyyymmdd="20251231", end_yyyymmdd="20260101"),
    )
    assert result.ok is False
    assert result.aborted_at == "load-end-snapshot"
    assert not result.by_stage("load-end-snapshot").ok


# ─── Failure paths ──────────────────────────────────────────────────────────

def test_pipeline_aborts_when_hp_val_load_fails(monkeypatch, tmp_path: Path):
    cfg = _make_cfg(tmp_path)

    def fake_load(yyyymmdd, *, conn_str, oefof_pcodes):
        raise DataLoadError("vw_RPT_VAL unreachable")

    monkeypatch.setattr(P, "load_snapshot_sql", fake_load)
    monkeypatch.setattr(P, "_resolve_latest_sql_vdate",
                        lambda *, conn_str, oefof_pcodes: "20260504")
    monkeypatch.setattr(P, "load_trades_sql", lambda **k: pd.DataFrame())

    result = run_pipeline(cfg, PipelineOptions())
    assert result.ok is False
    assert result.aborted_at == "load-end-snapshot"
    assert result.by_stage("load-end-snapshot").ok is False
    # Subsequent stages were never attempted.
    assert "load-start-snapshot" not in _stage_names(result)


def test_pipeline_aborts_when_bottler_load_fails(monkeypatch, tmp_path: Path):
    _patch_loaders(monkeypatch)

    def boom(**k):
        raise DataLoadError("trades source missing")

    monkeypatch.setattr(P, "load_trades_sql", boom)

    result = run_pipeline(_make_cfg(tmp_path), PipelineOptions(start_yyyymmdd="20251231"))
    assert result.ok is False
    assert result.aborted_at == "load-trades"


# ─── Critical-validation branch ─────────────────────────────────────────────

def test_pipeline_critical_validation_writes_failure_artifacts(monkeypatch, tmp_path: Path):
    _patch_loaders(monkeypatch)
    cfg = _make_cfg(tmp_path)

    # Build a ValidationResult that reports a critical FAIL.
    from oefof_pl.validate.checks import CheckResult, ValidationResult
    crit = ValidationResult(checks=(
        CheckResult(id="VAL-XX", name="Forced fail", status="FAIL",
                    detail="injected", failures=()),
    ))
    monkeypatch.setattr(P, "run_all_checks", lambda **kw: crit)

    result = run_pipeline(
        cfg, PipelineOptions(start_yyyymmdd="20251231"),
        extractor=lambda hp_val_path, *, target_pname: [[None]*12],
        closer=lambda p: False,
    )

    assert result.ok is False
    assert result.aborted_at == "validate"
    assert result.failure_parquet_path is not None
    assert result.failure_xlsx_path is not None
    assert result.failure_parquet_path.exists()
    assert result.failure_xlsx_path.exists()
    # No regular P&L report in the critical branch.
    assert not (cfg.output_dir / PL_REPORT_FILENAME).exists()
    # But the always-on parquet still got written.
    assert result.results_parquet_path.exists()
    # Analyst map still saved on failure.
    assert "save-analyst-map" in _stage_names(result)


# ─── Email branch ────────────────────────────────────────────────────────────

class _FakeMail:
    def __init__(self):
        self.To = self.CC = self.Subject = self.Body = self.HTMLBody = ""
        self.SendUsingAccount = None
        self.attached: list[str] = []
        self.sent = False

        class _A:
            def Add(s_self, p): self.attached.append(p)
        self.Attachments = _A()

    def Send(self):
        self.sent = True


class _FakeOutlook:
    def __init__(self):
        self.last_mail = None

        class _S:
            class _Accounts:
                def __iter__(s_self): return iter([])
            Accounts = _Accounts()
        self.Session = _S()

    def CreateItem(self, kind):
        m = _FakeMail()
        self.last_mail = m
        return m


def test_pipeline_send_email_uses_injected_outlook(monkeypatch, tmp_path: Path):
    _patch_loaders(monkeypatch)
    cfg = _make_cfg(tmp_path)
    outlook = _FakeOutlook()
    options = PipelineOptions(
        start_yyyymmdd="20251231", send_email=True,
        email_to="dest@x.com", email_subject="Hi",
    )

    result = run_pipeline(
        cfg, options,
        extractor=lambda hp_val_path, *, target_pname: [[None]*12 for _ in range(8)],
        closer=lambda p: False,
        outlook_dispatcher=lambda _: outlook,
    )

    assert result.ok
    assert result.email_result is not None
    assert result.email_result.sent is True
    assert outlook.last_mail.To == "dest@x.com"
    assert outlook.last_mail.Subject == "Hi"
    # One attachment — the main P&L workbook (ValuationA is embedded inside).
    assert len(outlook.last_mail.attached) == 1


def test_pipeline_send_email_skipped_when_flag_off(monkeypatch, tmp_path: Path):
    _patch_loaders(monkeypatch)
    cfg = _make_cfg(tmp_path)
    result = run_pipeline(
        cfg, PipelineOptions(start_yyyymmdd="20251231"),
        extractor=lambda hp_val_path, *, target_pname: [[None]*12 for _ in range(8)],
        closer=lambda p: False,
    )
    assert result.ok
    assert result.email_result is None
    assert "send-email" not in _stage_names(result)


# ─── Trade window filter ────────────────────────────────────────────────────

def test_filter_trade_window_keeps_inside_excludes_start():
    df = pd.DataFrame({
        "CDATE": [
            pd.Timestamp("2025-12-31"),     # = start → excluded
            pd.Timestamp("2026-01-01"),     # > start → kept
            pd.Timestamp("2026-04-29"),     # = end → kept
            pd.Timestamp("2026-04-30"),     # > end → excluded
        ],
        "ISIN": ["A", "B", "C", "D"],
    })
    out = P._filter_trade_window(df, START_VDATE, END_VDATE)
    assert list(out["ISIN"]) == ["B", "C"]


def test_filter_trade_window_empty_input():
    out = P._filter_trade_window(pd.DataFrame(), START_VDATE, END_VDATE)
    assert out.empty


# ─── _build_pl_lookup ───────────────────────────────────────────────────────

def test_build_pl_lookup_extracts_eight_columns_per_isin():
    cols = (
        "ISIN", "Realised P&L (EUR)", "Realised P&L (%)",
        "Unrealised P&L (EUR)", "Unrealised P&L (%)",
        "Income (EUR)", "Income Yield (%)",
        "Total P&L (EUR)", "Total P&L (%)",
    )
    df = pd.DataFrame([
        ["X1",  1.0, 0.01, 2.0, 0.02, 3.0, 0.03, 6.0, 0.06],
        ["X2", 10.0, 0.10, 20.0, 0.20, 30.0, 0.30, 60.0, 0.60],
    ], columns=cols)
    out = P._build_pl_lookup(df)
    assert set(out) == {"X1", "X2"}
    assert out["X1"] == [1.0, 0.01, 2.0, 0.02, 3.0, 0.03, 6.0, 0.06]


def test_build_pl_lookup_empty_df_returns_empty():
    assert P._build_pl_lookup(pd.DataFrame()) == {}
