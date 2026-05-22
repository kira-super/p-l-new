"""Central configuration for the oefof_pl package.

This is the only module that owns paths, fund codes, analyst codes, and
default tunables. Every other module receives a `Config` instance (or the
specific fields it needs) as a parameter — no other module reads from this
module's globals directly. This keeps the package testable: tests build a
`Config` pointing at temp directories and never touch production data.

All constants below reflect decisions recorded in DESIGN.md §18 (resolved
2026-04-29).
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Mapping


def _env_bool(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    val = str(raw).strip().lower()
    if val in {"1", "true", "t", "yes", "y", "on"}:
        return True
    if val in {"0", "false", "f", "no", "n", "off"}:
        return False
    return default


# ─── Production paths ────────────────────────────────────────────────────────

# HP_VAL.xlsm and StockTrList.xlsb have both been fully retired. The
# portfolio snapshot, ValuationA pivot rows, and trade blotter are now
# read directly from SQL Server (see ``SQL_CONN_STR`` below). The
# ``BOTTLER_PATH`` constant is kept only as a vestigial label echoed in
# the workbook header; nothing on disk is read from it.
BOTTLER_PATH = Path("inputs/StockTrList.xlsb")
ANALYST_MAP_PATH = Path("isin_analyst_map.csv")
CA_OVERRIDES_PATH = Path("inputs/ca_overrides.csv")
CA_HISTORY_PATH = Path("inputs/ca_overrides_history.csv")
CA_BONUS_PRICES_PATH = Path("inputs/ca_bonus_prices.csv")
OUTPUT_DIR = Path("out")
# Per-month OneDrive archive of the published P&L workbook. The pipeline
# copies every successful build into <REPORT_ARCHIVE_ROOT>/<YYYY-MM Month>/
# P&L-OEFOF-<YYYYMMDD>.xlsx (folder created on demand).
REPORT_ARCHIVE_ROOT = Path(
    r"C:\Users\kgontar\Fiera Capital Corporation\Oaks - Documents\OAKS P&L Reports"
)


# ─── SQL Server (HiPort valuation source) ────────────────────────────────────
# Both end- and start-of-period snapshots are read from the
# ``ccl.dbo.vw_RPT_VAL`` view on FCEDATA01. Override via the
# ``OEFOF_SQL_CONN`` env var when running on a host that needs a
# different driver / auth / server name.
SQL_CONN_STR_DEFAULT = (
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=fcedata01.fieracapital.ca;"
    "Database=ccl;"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)
SQL_CONN_STR = os.environ.get("OEFOF_SQL_CONN", SQL_CONN_STR_DEFAULT)
NAV_SQL_TABLE = os.environ.get("OEFOF_NAV_SQL_TABLE", "NAV.dbo.tNAV")
NAV_PCODE_COLUMN = os.environ.get("OEFOF_NAV_PCODE_COLUMN", "PCODE")
NAV_PCODE_PATTERN = os.environ.get("OEFOF_NAV_PCODE_PATTERN", "OEFOF%")
NAV_VDATE_COLUMN = os.environ.get("OEFOF_NAV_VDATE_COLUMN", "vDATE")
NAV_VALUE_COLUMN = os.environ.get("OEFOF_NAV_VALUE_COLUMN", "MKTCAP")


# ─── Year-to-date anchor ─────────────────────────────────────────────────────
# Manual annual roll. Bump this on Jan 1 each year (audit trail in git).
START_SNAPSHOT_DEFAULT = "20251231"
START_SNAPSHOT_DEFAULT = os.environ.get(
    "OEFOF_START_SNAPSHOT_DEFAULT", START_SNAPSHOT_DEFAULT
)


# ─── Fund universe ───────────────────────────────────────────────────────────
# Decision §18 Q3: legacy 5-code list. The 15-code list in the original prompt
# would double-count OEFOF via OEFOFMR (mirror) and the various swap-collateral
# accounts that are already represented in OEFOGSSC/OEFOGSSW/OEFOHFSC/OEFOHFSW.
OEFOF_PCODES: tuple[str, ...] = (
    "OEFOF",      # OAKS Emerging and Frontier — main equity account
    "OEFOGSSC",   # OEFOF GS Swap Collateral
    "OEFOGSSW",   # OEFOF GS Swap
    "OEFOHFSW",   # OEFOF HSBC Swap
    "OEFOHFSC",   # OEFOF HSBC Swap Collateral
)
_OEFOF_PCODES_ENV = os.environ.get("OEFOF_PCODES")
if _OEFOF_PCODES_ENV:
    OEFOF_PCODES = tuple(
        p.strip().upper()
        for p in _OEFOF_PCODES_ENV.split(",")
        if p.strip()
    )


# ─── Compute toggles ─────────────────────────────────────────────────────────
# Keep defaults matching current behaviour; allow explicit override for
# diagnostics and controlled what-if runs.
DROP_ZZZZ_PS_TRADES = _env_bool("OEFOF_DROP_ZZZZ_PS", True)
SPLIT_CA_ROWS = _env_bool("OEFOF_SPLIT_CA_ROWS", True)


# ─── Instrument classification ───────────────────────────────────────────────
# Index-level CFDs are reported separately as FTSWAP. Extend as new ones are
# discovered.
FTSWAP_ISINS: frozenset[str] = frozenset({
    "SX7E INDEX",   # Euro Stoxx Banks CFD
    "FDGSCBIHKT",   # GS basket CFD (actual ISIN in HiPort tTRANS / vw_RPT_VAL)
    "GSCBIHKT",     # Legacy alias — tSecurity.ISIN is NULL so sname fallback fires
})


# ─── Analyst codes ───────────────────────────────────────────────────────────
# Decision §18 Q8: 7 analysts. Full names filled in over time; missing names
# default to the code itself (the pipeline never blocks on an unknown name).
ANALYST_CODES: dict[str, str] = {
    "IS": "Ian Simmons",
    "HK": "Hayden Kwan",
    "JB": "Julius Bottcher",
    "SB": "Stefan Bottcher",
    "KX": "Karen Xiao",
    "VS": "Vijay Singh",
    "AS": "Alexander Short",         
}


# Manual SNAME-based analyst overrides. The single source of truth for both
# ISIN- and SNAME-based mappings is ``isin_analyst_map.csv``: rows with a
# blank ISIN but a populated SNAME are SNAME overrides. This dict is kept
# only as a runtime fallback / extension point and is empty by default.
PERSON_OVERRIDES_BY_SECURITY: dict[str, str] = {}


# ─── Email defaults ──────────────────────────────────────────────────────────
# Decision §18 Q6: production recipient is oaks@fieracapital.com.
EMAIL_TO_DEFAULT = "kgontar@fieracapital.com"
EMAIL_SUBJECT_DEFAULT = "OEFOF ValuationA and P&L (inc per analyst)"


# ─── Valuation workbook ──────────────────────────────────────────────────────
TARGET_PNAME = "OAKS Emerging and Frontier Fund"


# ─── Config dataclass (the thing that flows through the pipeline) ────────────

@dataclass(frozen=True)
class Config:
    """Immutable runtime configuration.

    Built once at process start by `load_default_config()`, then passed to
    `Pipeline(cfg)`. Tests construct one with `dataclasses.replace(cfg, ...)`
    to redirect paths into a temp directory.
    """

    bottler_path: Path
    analyst_map_path: Path
    ca_overrides_path: Path
    ca_history_path: Path
    ca_bonus_prices_path: Path
    output_dir: Path
    report_archive_root: Path
    sql_conn_str: str
    nav_sql_table: str
    nav_pcode_column: str
    nav_pcode_pattern: str
    nav_vdate_column: str
    nav_value_column: str

    fund_pcodes: tuple[str, ...]        # was oefof_pcodes
    ftswap_isins: frozenset[str]
    analyst_codes: Mapping[str, str]
    person_overrides: Mapping[str, str]
    drop_zzzz_ps_trades: bool
    split_ca_rows: bool

    start_snapshot_default: str

    # Fund identity — defaults work for OEFOF; override per fund in funds/
    fund_currency: str = "EUR"
    fund_name: str = TARGET_PNAME
    pl_report_filename: str = "OAKS_EM_Stock_PL_Report.xlsx"

    target_pname: str = TARGET_PNAME

    email_to_default: str = EMAIL_TO_DEFAULT
    email_subject_default: str = EMAIL_SUBJECT_DEFAULT


def load_default_config() -> Config:
    """Build the production `Config` from the module-level constants above."""
    return Config(
        bottler_path=BOTTLER_PATH,
        analyst_map_path=ANALYST_MAP_PATH,
        ca_overrides_path=CA_OVERRIDES_PATH,
        ca_history_path=CA_HISTORY_PATH,
        ca_bonus_prices_path=CA_BONUS_PRICES_PATH,
        output_dir=OUTPUT_DIR,
        report_archive_root=REPORT_ARCHIVE_ROOT,
        sql_conn_str=SQL_CONN_STR,
        nav_sql_table=NAV_SQL_TABLE,
        nav_pcode_column=NAV_PCODE_COLUMN,
        nav_pcode_pattern=NAV_PCODE_PATTERN,
        nav_vdate_column=NAV_VDATE_COLUMN,
        nav_value_column=NAV_VALUE_COLUMN,
        fund_pcodes=OEFOF_PCODES,
        fund_currency="EUR",
        fund_name=TARGET_PNAME,
        pl_report_filename="OAKS_EM_Stock_PL_Report.xlsx",
        ftswap_isins=FTSWAP_ISINS,
        analyst_codes=dict(ANALYST_CODES),
        person_overrides=dict(PERSON_OVERRIDES_BY_SECURITY),
        drop_zzzz_ps_trades=DROP_ZZZZ_PS_TRADES,
        split_ca_rows=SPLIT_CA_ROWS,
        start_snapshot_default=START_SNAPSHOT_DEFAULT,
    )


