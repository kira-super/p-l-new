"""Command-line entry: ``python -m oefof_pl``."""
from __future__ import annotations

import argparse
import os
import sys
from dataclasses import replace
from pathlib import Path

# Ensure relative paths in config (inputs/, out/, isin_analyst_map.csv) resolve
# against the project root regardless of where ``python -m`` was invoked from.
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
os.chdir(_PROJECT_ROOT)

from .config import Config, load_default_config
from .pipeline import PipelineOptions, run_pipeline


def _parse(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="oefof-pl",
        description="OEFOF Stock-Level P&L pipeline.",
    )
    p.add_argument("--start", default="", help="Start snapshot YYYYMMDD (default: cfg.start_snapshot_default).")
    p.add_argument("--end", default="", help="End snapshot YYYYMMDD (default: latest VDATE in vw_RPT_VAL).")
    p.add_argument("--sql-conn", default=None, help="pyodbc connection string for the HiPort SQL view (overrides cfg.sql_conn_str).")
    p.add_argument("--analyst-map", default=None, help="Path to analyst CSV (overrides cfg.analyst_map_path).")
    p.add_argument("--output-dir", default=None, help="Output directory (overrides cfg.output_dir).")
    p.add_argument("--send-email", action="store_true", help="Send Outlook email with both workbooks attached.")
    p.add_argument("--email-to", default="", help="Recipient(s); comma/semicolon separated.")
    p.add_argument("--email-cc", default="", help="CC recipient(s); comma/semicolon separated.")
    p.add_argument("--email-subject", default="", help="Override email subject (default from cfg).")
    p.add_argument("--email-body", default="OEFOF P&L pipeline output attached.", help="Plain-text email body.")
    p.add_argument("--email-from-smtp", default="", help="Outlook account SMTP to send from.")
    p.add_argument("--apply-suggestions", action="append", default=[],
                   metavar="PATH",
                   help="Additional CA-overrides CSV to apply (typically out/ca_suggestions_*.csv). "
                        "Repeat the flag to apply multiple files.")
    return p.parse_args(argv)


def _apply_overrides(cfg: Config, args: argparse.Namespace) -> Config:
    overrides: dict = {}
    if args.sql_conn:
        overrides["sql_conn_str"] = args.sql_conn
    if args.analyst_map:
        overrides["analyst_map_path"] = Path(args.analyst_map)
    if args.output_dir:
        overrides["output_dir"] = Path(args.output_dir)
    return replace(cfg, **overrides) if overrides else cfg


def _options_from_args(args: argparse.Namespace) -> PipelineOptions:
    return PipelineOptions(
        start_yyyymmdd=args.start,
        end_yyyymmdd=args.end,
        send_email=args.send_email,
        email_to=args.email_to,
        email_cc=args.email_cc,
        email_subject=args.email_subject,
        email_body=args.email_body,
        email_from_smtp=args.email_from_smtp,
        extra_ca_override_paths=tuple(Path(p) for p in args.apply_suggestions),
    )


def _print_stages(result) -> None:
    for s in result.stages:
        marker = "OK  " if s.ok else "FAIL"
        print(f"  [{marker}] {s.name:<24} {s.message}")


def main(argv=None) -> int:
    args = _parse(argv)
    cfg = _apply_overrides(load_default_config(), args)
    options = _options_from_args(args)

    print(f"Running OEFOF P&L pipeline (start={options.start_yyyymmdd or '<default>'}, "
          f"end={options.end_yyyymmdd or '<from VDATE>'})")
    print()
    print("Stages (live):")

    def _emit(stage):
        marker = "OK  " if stage.ok else "FAIL"
        print(f"  [{marker}] {stage.name:<24} {stage.message}", flush=True)

    result = run_pipeline(cfg, options, progress=_emit)
    print()
    if result.ok:
        print(f"OK — P&L report: {result.pl_report_path}")
        if result.archive_path:
            print(f"     Archived  : {result.archive_path}")
        print(f"     Results   : {result.results_parquet_path}")
        if result.unassigned_isins_path:
            print(f"     Unassigned: {result.unassigned_isins_path}")
        return 0
    print(f"ABORTED at {result.aborted_at!r}")
    if result.failure_xlsx_path:
        print(f"Failure workbook: {result.failure_xlsx_path}")
    if result.failure_parquet_path:
        print(f"Failure parquet : {result.failure_parquet_path}")
    if result.ca_suggestions_path:
        print(f"CA suggestions  : {result.ca_suggestions_path}")
        print(f"  Re-run with: python -m oefof_pl --apply-suggestions \"{result.ca_suggestions_path}\"")
    if result.unassigned_isins_path:
        print(f"Unassigned ISINs: {result.unassigned_isins_path}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
