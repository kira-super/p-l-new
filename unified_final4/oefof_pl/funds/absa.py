"""ABSA fund configuration.

Call ``load_config()`` to get the production Config for the ABSA Fund.
All fund-specific constants are defined here; the shared Config dataclass
and pipeline live in the parent package.
"""
from __future__ import annotations

import os
from dataclasses import replace
from pathlib import Path
from typing import Mapping

from oefof_pl.config import Config, load_default_config


# ─── Production paths ────────────────────────────────────────────────────────

# Uses the shared isin_analyst_map.csv (same as OEFOF and all future funds)
CA_OVERRIDES_PATH = Path("inputs/ca_overrides_absa.csv")
CA_HISTORY_PATH = Path("inputs/ca_overrides_history_absa.csv")
CA_BONUS_PRICES_PATH = Path("inputs/ca_bonus_prices_absa.csv")
OUTPUT_DIR = Path("out/absa")
REPORT_ARCHIVE_ROOT = Path(
    r"C:\Users\kgontar\Fiera Capital Corporation\Oaks - Documents\OAKS P&L Reports\ABSA"
)


# ─── SQL Server ───────────────────────────────────────────────────────────────
# Shares the same SQL Server as OEFOF; override via ABSA_SQL_CONN if needed.
SQL_CONN_STR = os.environ.get("ABSA_SQL_CONN", "")  # empty = use OEFOF default


# ─── Fund universe ────────────────────────────────────────────────────────────
def _parse_pcodes(raw: str) -> tuple[str, ...]:
    tokens = [p.strip().upper() for p in raw.replace(";", ",").split(",")]
    cleaned = [p for p in tokens if p]
    return tuple(dict.fromkeys(cleaned)) if cleaned else ("ABSA",)


ABSA_PCODES: tuple[str, ...] = _parse_pcodes(
    os.environ.get("ABSA_PCODES", "ABSA,L118")
)


# ─── Instrument classification ────────────────────────────────────────────────
# No index-level CFDs for ABSA.
FTSWAP_ISINS: frozenset[str] = frozenset()


# ─── Analyst codes ────────────────────────────────────────────────────────────
ANALYST_CODES: dict[str, str] = {
    "IS": "Ian Simmons",
    "HK": "Hayden Kwan",
    "JB": "Julius Bottcher",
    "SB": "Stefan Bottcher",
    "KX": "Karen Xiao",
    "VS": "Vijay Singh",
    "AS": "Alexander Short",
}

PERSON_OVERRIDES_BY_SECURITY: dict[str, str] = {}


# ─── NAV source ───────────────────────────────────────────────────────────────
NAV_PCODE_PATTERN = os.environ.get("ABSA_NAV_PCODE_PATTERN", "ABSA%")


# ─── Fund identity ────────────────────────────────────────────────────────────
FUND_NAME = "ABSA Fund"
FUND_CURRENCY = os.environ.get("ABSA_FUND_CURRENCY", "USD").strip().upper() or "USD"
PL_REPORT_FILENAME = "ABSA_Stock_PL_Report.xlsx"
START_SNAPSHOT_DEFAULT = "20251231"


# ─── Email defaults ───────────────────────────────────────────────────────────
EMAIL_TO_DEFAULT = "kgontar@fieracapital.com"
EMAIL_SUBJECT_DEFAULT = "ABSA ValuationA and P&L (inc per analyst)"


def load_config() -> Config:
    """Return the ABSA production Config."""
    # Start from the OEFOF defaults (shared SQL server, paths layout, etc.)
    base = load_default_config()
    overrides: dict = dict(
        # analyst_map_path: inherits shared isin_analyst_map.csv from base config
        ca_overrides_path=CA_OVERRIDES_PATH,
        ca_history_path=CA_HISTORY_PATH,
        ca_bonus_prices_path=CA_BONUS_PRICES_PATH,
        output_dir=OUTPUT_DIR,
        report_archive_root=REPORT_ARCHIVE_ROOT,
        fund_pcodes=ABSA_PCODES,
        ftswap_isins=FTSWAP_ISINS,
        analyst_codes=dict(ANALYST_CODES),
        person_overrides=dict(PERSON_OVERRIDES_BY_SECURITY),
        nav_pcode_pattern=NAV_PCODE_PATTERN,
        fund_name=FUND_NAME,
        fund_currency=FUND_CURRENCY,
        pl_report_filename=PL_REPORT_FILENAME,
        start_snapshot_default=START_SNAPSHOT_DEFAULT,
        email_to_default=EMAIL_TO_DEFAULT,
        email_subject_default=EMAIL_SUBJECT_DEFAULT,
    )
    if SQL_CONN_STR:
        overrides["sql_conn_str"] = SQL_CONN_STR
    return replace(base, **overrides)
