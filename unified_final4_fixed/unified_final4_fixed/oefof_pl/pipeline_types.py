"""Pipeline result types.

Defines the frozen dataclasses that carry pipeline state across modules
(:class:`StageResult`, :class:`PipelineResult`, :class:`PipelineOptions`).

These live in a separate module so that :mod:`oefof_pl.pipeline` (the
orchestrator) and :mod:`oefof_pl.pipeline_helpers` (utility functions) can
both import them without creating a circular dependency.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Mapping

import pandas as pd

from .email_sender.outlook import EmailSendResult
from .validate import ValidationResult


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
