"""Unified P&L pipeline launcher.

Select the fund with --oefof or --absa (default: --oefof).

Examples
--------
Run OEFOF today and email it::

    python app.py --oefof

Run ABSA for a specific period, no email::

    python app.py --absa --start 20251231 --end 20260513 --no-email

Run OEFOF without opening the workbook::

    python app.py --oefof --no-email --no-open
"""
from __future__ import annotations

import argparse
import os
import sys
from dataclasses import replace
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent
os.chdir(_PROJECT_ROOT)

from oefof_pl.config import Config
from oefof_pl.data.ca_overrides import persist_auto_ca
from oefof_pl.pipeline import PipelineOptions, run_pipeline

# Registry: flag name → (display label, loader import path)
_FUNDS = {
    "oefof": ("OEFOF (OAKS Emerging & Frontier)", "oefof_pl.funds.oefof"),
    "absa":  ("ABSA Fund",                         "oefof_pl.funds.absa"),
}
_DEFAULT_FUND = "oefof"


def _load_fund_config(fund_key: str) -> Config:
    import importlib
    mod = importlib.import_module(_FUNDS[fund_key][1])
    return mod.load_config()


def _parse(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="app",
        description="Multi-fund P&L pipeline launcher.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="\n".join(
            f"  --{k:<8}  {label}" for k, (label, _) in _FUNDS.items()
        ),
    )

    # ── Fund selector (mutually exclusive flags) ──────────────────────────
    fund_grp = p.add_mutually_exclusive_group()
    for key, (label, _) in _FUNDS.items():
        fund_grp.add_argument(
            f"--{key}",
            dest="fund",
            action="store_const",
            const=key,
            help=f"Run pipeline for {label}.",
        )
    p.set_defaults(fund=_DEFAULT_FUND)

    # ── Period / data sources ─────────────────────────────────────────────
    p.add_argument("--start", default="",
                   help="Start snapshot YYYYMMDD (default: cfg.start_snapshot_default).")
    p.add_argument("--end", default="",
                   help="End snapshot YYYYMMDD (default: latest VDATE in vw_RPT_VAL).")
    p.add_argument("--sql-conn", default=None,
                   help="pyodbc connection string override.")
    p.add_argument("--analyst-map", default=None,
                   help="Path to analyst CSV override.")
    p.add_argument("--output-dir", default=None,
                   help="Output directory override.")
    p.add_argument("--apply-suggestions", action="append", default=[],
                   metavar="PATH",
                   help="Additional CA-overrides CSV (repeatable).")
    p.add_argument("--strict-run", action="store_true",
                   help="Enable production hard gates (override approvals, reconciliation, anomaly/release checks).")
    p.add_argument("--override-approvals", default="",
                   help="CSV of approved override usage (default from config).")
    p.add_argument("--known-exceptions", default="",
                   help="CSV register of approved anomaly/reconciliation exceptions (default from config).")
    p.add_argument("--prior-audit", default="",
                   help="Prior run audit JSON for day-over-day anomaly gate.")

    zzzz = p.add_mutually_exclusive_group()
    zzzz.add_argument("--drop-zzzz-ps-trades", dest="drop_zzzz_ps_trades",
                      action="store_true", default=None)
    zzzz.add_argument("--keep-zzzz-ps-trades", dest="drop_zzzz_ps_trades",
                      action="store_false")
    split = p.add_mutually_exclusive_group()
    split.add_argument("--split-ca-rows", dest="split_ca_rows",
                       action="store_true", default=None)
    split.add_argument("--no-split-ca-rows", dest="split_ca_rows",
                       action="store_false")
    auto = p.add_mutually_exclusive_group()
    auto.add_argument("--auto-apply-suggestions", dest="auto_apply_suggestions",
                      action="store_true", default=True)
    auto.add_argument("--no-auto-apply-suggestions", dest="auto_apply_suggestions",
                      action="store_false")

    # ── Email (on by default) ─────────────────────────────────────────────
    email = p.add_mutually_exclusive_group()
    email.add_argument("--send-email", dest="send_email",
                       action="store_true", default=True)
    email.add_argument("--no-email", dest="send_email",
                       action="store_false",
                       help="Build workbooks but skip email.")
    p.add_argument("--email-to", default="")
    p.add_argument("--email-cc", default="")
    p.add_argument("--email-subject", default="")
    p.add_argument("--email-body", default="P&L pipeline output attached.")
    p.add_argument("--email-from-smtp", default="")

    # ── Auto-open ─────────────────────────────────────────────────────────
    open_grp = p.add_mutually_exclusive_group()
    open_grp.add_argument("--open", dest="open_output",
                          action="store_true", default=True)
    open_grp.add_argument("--no-open", dest="open_output",
                          action="store_false")

    return p.parse_args(argv)


def _apply_overrides(cfg: Config, args: argparse.Namespace) -> Config:
    overrides: dict = {}
    if args.sql_conn:
        overrides["sql_conn_str"] = args.sql_conn
    if args.analyst_map:
        overrides["analyst_map_path"] = Path(args.analyst_map)
    if args.output_dir:
        overrides["output_dir"] = Path(args.output_dir)
    if args.drop_zzzz_ps_trades is not None:
        overrides["drop_zzzz_ps_trades"] = args.drop_zzzz_ps_trades
    if args.split_ca_rows is not None:
        overrides["split_ca_rows"] = args.split_ca_rows
    return replace(cfg, **overrides) if overrides else cfg


def _options(args: argparse.Namespace, cfg: Config) -> PipelineOptions:
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
        strict_run=bool(args.strict_run),
        override_approvals_path=(Path(args.override_approvals) if args.override_approvals else cfg.override_approvals_path),
        known_exceptions_path=(Path(args.known_exceptions) if args.known_exceptions else cfg.known_exceptions_path),
        prior_audit_path=(Path(args.prior_audit) if args.prior_audit else None),
    )


def main(argv: list[str] | None = None) -> int:
    args = _parse(argv)
    fund_label, _ = _FUNDS[args.fund]
    cfg = _apply_overrides(_load_fund_config(args.fund), args)
    options = _options(args, cfg)

    print("=" * 72)
    print(f" {fund_label} — P&L pipeline")
    print("=" * 72)
    print(f" Period       : start={options.start_yyyymmdd or '<default>'} "
          f"end={options.end_yyyymmdd or '<latest VDATE>'}")
    print(f" SQL source   : ccl.dbo.vw_RPT_VAL @ FCEDATA01")
    print(f" Trades       : ccl.dbo.tTRANS @ FCEDATA01")
    print(f" Output dir   : {cfg.output_dir}")
    print(f" Send email   : {options.send_email}")
    if options.send_email:
        to = options.email_to or cfg.email_to_default
        print(f"   to         : {to}")
        if options.email_cc:
            print(f"   cc         : {options.email_cc}")
    print()
    print("Stages (live):")

    def _emit(stage):
        marker = "OK  " if stage.ok else "FAIL"
        print(f"  [{marker}] {stage.name:<24} {stage.message}", flush=True)

    result = run_pipeline(cfg, options, progress=_emit)

    # ── Auto-retry with CA suggestions ────────────────────────────────────
    if (
        not result.ok
        and args.auto_apply_suggestions
        and result.ca_suggestions_path is not None
        and Path(result.ca_suggestions_path).exists()
    ):
        sugg_path = Path(result.ca_suggestions_path)
        print()
        print("=" * 72)
        print(f" Auto-applying CA suggestions: {sugg_path}")
        print(" (disable with --no-auto-apply-suggestions)")
        print("=" * 72)
        print("Stages (live):")
        retry_options = replace(
            options,
            extra_ca_override_paths=tuple([*options.extra_ca_override_paths, sugg_path]),
        )
        result = run_pipeline(cfg, retry_options, progress=_emit)

        if result.ok and result.end_date is not None:
            try:
                persisted = persist_auto_ca(
                    suggestions_path=sugg_path,
                    overrides_path=cfg.ca_overrides_path,
                    history_path=cfg.ca_history_path,
                    period_end=result.end_date,
                )
                if persisted:
                    print()
                    print(f"  [AUTO] Persisted {len(persisted)} CA override(s) to "
                          f"{cfg.ca_overrides_path}")
                    for ov in persisted:
                        print(f"         {ov.isin:20s}  {ov.units:+.4g} units")
            except Exception as exc:
                print(f"  [WARN] Could not persist CA overrides: {exc}")

    print()
    if result.ok:
        print(f"OK  P&L report : {result.pl_report_path}")
        if result.archive_path:
            print(f"    Archived   : {result.archive_path}")
        print(f"    Results    : {result.results_parquet_path}")
        if result.unassigned_isins_path:
            print(f"    Unassigned : {result.unassigned_isins_path}")
        if result.email_result is not None:
            er = result.email_result
            if er.sent:
                print(f"    Email sent : to={er.to} cc={er.cc} "
                      f"attachments={er.attachment_count}")
            else:
                print(f"    Email FAIL : {er.error}")
        if args.open_output and result.pl_report_path:
            try:
                os.startfile(str(result.pl_report_path))  # type: ignore[attr-defined]
            except OSError as exc:
                print(f"    Open FAIL  : {exc}")
        return 0

    print(f"ABORTED at {result.aborted_at!r}")
    if result.failure_xlsx_path:
        print(f"  Failure workbook: {result.failure_xlsx_path}")
    if result.failure_parquet_path:
        print(f"  Failure parquet : {result.failure_parquet_path}")
    if result.ca_suggestions_path:
        print(f"  CA suggestions  : {result.ca_suggestions_path}")
        print(f"  Re-run with: python app.py --{args.fund} "
              f'--apply-suggestions "{result.ca_suggestions_path}"')
    if result.unassigned_isins_path:
        print(f"  Unassigned ISINs: {result.unassigned_isins_path}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
