"""Tests for ``oefof_pl.__main__`` CLI argument parsing and exit codes."""

from __future__ import annotations

from pathlib import Path

import pytest

from oefof_pl import __main__ as cli
from oefof_pl.config import load_default_config
from oefof_pl.pipeline import PipelineOptions, PipelineResult, StageResult


# ─── _parse ─────────────────────────────────────────────────────────────────

def test_parse_defaults_are_blank_strings():
    args = cli._parse([])
    assert args.start == ""
    assert args.end == ""
    assert args.send_email is False
    assert args.strict_run is False


def test_parse_path_overrides(tmp_path: Path):
    args = cli._parse([
        "--sql-conn", "DRIVER={fake};Server=test;",
        "--analyst-map", str(tmp_path / "a.csv"),
        "--output-dir", str(tmp_path / "o"),
        "--send-email",
        "--email-to", "a@b.com",
        "--strict-run",
        "--override-approvals", str(tmp_path / "approvals.csv"),
        "--known-exceptions", str(tmp_path / "exceptions.csv"),
        "--prior-audit", str(tmp_path / "audit.json"),
    ])
    cfg = cli._apply_overrides(load_default_config(), args)
    assert cfg.sql_conn_str == "DRIVER={fake};Server=test;"
    assert cfg.analyst_map_path == tmp_path / "a.csv"
    assert cfg.output_dir == tmp_path / "o"
    opts = cli._options_from_args(args)
    assert opts.send_email is True
    assert opts.email_to == "a@b.com"
    assert opts.strict_run is True
    assert opts.override_approvals_path == tmp_path / "approvals.csv"
    assert opts.known_exceptions_path == tmp_path / "exceptions.csv"
    assert opts.prior_audit_path == tmp_path / "audit.json"


def test_apply_overrides_no_args_returns_same_cfg():
    cfg = load_default_config()
    out = cli._apply_overrides(cfg, cli._parse([]))
    assert out is cfg


# ─── main exit codes ───────────────────────────────────────────────────────

def _ok_result() -> PipelineResult:
    return PipelineResult(
        ok=True,
        stages=(StageResult("x", True, "done"),),
        pl_report_path=Path("pl.xlsx"),
        valuation_path=Path("v.xlsx"),
        results_parquet_path=Path("r.parquet"),
    )


def _fail_result() -> PipelineResult:
    return PipelineResult(
        ok=False,
        stages=(StageResult("validate", False, "bad"),),
        aborted_at="validate",
        failure_xlsx_path=Path("fail.xlsx"),
        failure_parquet_path=Path("fail.parquet"),
    )


def test_main_returns_0_on_success(monkeypatch, capsys):
    monkeypatch.setattr(cli, "run_pipeline", lambda cfg, opts, **kw: _ok_result())
    rc = cli.main([])
    out = capsys.readouterr().out
    assert rc == 0
    assert "OK" in out


def test_main_returns_1_on_failure(monkeypatch, capsys):
    monkeypatch.setattr(cli, "run_pipeline", lambda cfg, opts, **kw: _fail_result())
    rc = cli.main([])
    out = capsys.readouterr().out
    assert rc == 1
    assert "ABORTED" in out
    assert "fail.xlsx" in out


def test_main_passes_start_and_end_into_options(monkeypatch):
    captured: dict = {}

    def fake_run(cfg, opts, **kw):
        captured["opts"] = opts
        return _ok_result()

    monkeypatch.setattr(cli, "run_pipeline", fake_run)
    cli.main(["--start", "20251231", "--end", "20260429"])
    assert captured["opts"].start_yyyymmdd == "20251231"
    assert captured["opts"].end_yyyymmdd == "20260429"
