# oefof_pl — OAKS Emerging & Frontier Fund · Stock-Level P&L

Greenfield Python pipeline that produces a single, audit-grade Excel
report covering every position the fund touched in a chosen window:

- Fund-level dashboard (KPIs, instrument breakdown, by-analyst exposure).
- HiPort `ValuationA` reconciliation rebuilt as a hierarchical pivot
  (Analyst → Country → Stocks → Country Total → Analyst Total → Grand Total).
- One tab per analyst with focused 14-column layout, a header card, and a
  grand-total row.
- Current Holdings, Exited Positions, Income & Dividends, Trade Detail.
- Validation Report (11 gates, complete failure listings) + Methodology.
- Optional Outlook email with both workbooks attached and a rich HTML
  summary in the body.

The compute layer is pure Python (no Excel, no COM) and fully unit
tested. Excel/Outlook are only touched at the edges — `data/loader.py`
for HP_VAL extraction, `refresh/bottler.py` for the StockTrList macro,
`output/valuation.py` for the live ValuationA dump, and
`email_sender/outlook.py` for sending.

---

## Quick start

### 1. Install (once)

```powershell
cd oefof_pl
pip install -e .            # core dependencies
pip install -e .[com,dev]   # + pywin32 (Outlook/Excel COM) + pytest
```

`pywin32` is only required for the optional COM-driven steps
(ValuationA extraction, bottler refresh, Outlook send). The compute
layer, validation, and the workbook builder run on plain CPython.

### 2. Run

The simplest way — **double-click or invoke the launcher**:

```powershell
cd oefof_pl
python app.py
```

`app.py` runs the full pipeline **and emails the report** by default
(equivalent to `python -m oefof_pl --send-email`). Pass `--no-email` to
build the workbooks without sending.

Same launcher, common overrides:

```powershell
# Specific period, no email
python app.py --start 20251231 --end 20260429 --no-email

# Send to a different recipient
python app.py --email-to "trader@fieracapital.com" --email-cc "team@fieracapital.com"
```

### 3. Module entry (no auto-email)

The package is also runnable as a module — same code path, but
`--send-email` is **off** by default:

```powershell
python -u -m oefof_pl
python -u -m oefof_pl --send-email --email-to "oaks@fieracapital.com"
```

Use `-u` so live stage progress is flushed line-by-line. Avoid piping
through `Tee-Object` / `Select-Object` — PowerShell pipes buffer the
full stream until the process exits.

After `pip install -e .` the console-script alias is also available:

```powershell
oefof-pl --send-email
```

---

## What the workbook contains

| # | Sheet                | Purpose                                                                 |
|---|----------------------|-------------------------------------------------------------------------|
| 1 | **Summary**          | Title block, KPI cards, Instrument Breakdown, By-Analyst (incl. Gross). |
| 2 | **ValuationA**       | HiPort reconciliation, hierarchical pivot with subtotals & Grand Total. |
| 3 | One tab per analyst  | KPI strip + Current / Exited / ANALYST TOTAL — 14-col focused layout.   |
| 4 | **Current Holdings** | Full 31-column position detail.                                         |
| 5 | **Exited Positions** | Full 31-column position detail; Units & Last Px hidden.                 |
| 6 | **Income & Dividends** + **Income Detail** | By position and by trade.                            |
| 7 | **Trade Detail**     | Every trade in the analysis window.                                     |
| 8 | **Validation Report**| Every check ID, status, and **complete** failure listing (no truncation).|
| 9 | **Methodology**      | The five rules the numbers depend on.                                   |

Every sheet has frozen panes, autofilter on the data range, banded
rows, and tab colour-coding by purpose (teal = dashboard, mid-teal =
analyst, grey = raw data, gold = audit). Negative P&L renders red.

### Per-analyst sheet layout

Columns (in order): ISIN, Stock Name, Total P&L (EUR), Total P&L %,
Realised (EUR), Unrealised (EUR), Cost Basis (EUR), MV End (EUR),
Income & Financing (EUR), CCY, Type, L/S, Units, Last Px.

- Subtotal — Current and Subtotal — Exited rows close each section.
- The Exited section blanks the Units & Last Px columns (data and
  header) since both are stale once the position is closed.
- ANALYST TOTAL row sums Total P&L (EUR), Realised, Unrealised, Cost
  Basis, MV End, Income & Financing; the Total P&L % cell is a
  Σ(P&L) / |Σ(Cost Basis)| ratio. CCY / Type / L/S / Units / Last Px are
  blanked since they don't aggregate meaningfully.

### ValuationA layout

The sheet is rebuilt from the live `HP_VAL.xlsm › ValuationA` pivot:

- Pivot subtotal rows (`<Country> Total`, `<Analyst> Total`,
  `Grand Total`) are dropped on read and recomputed on write so that
  blank-Initials rows can be reassigned via the SNAME → analyst lookup
  built from the P&L frame.
- Output emits one Initials value at the top of each analyst block, one
  Country value at the top of each (Analyst, Country) sub-block, then
  the stock rows, then a `<Country> Total` row, then the next country.
  After the last country of an analyst a bold `<Analyst> Total` row
  closes the block. The sheet ends with a banded `Grand Total` row.

### Email summary

When `--send-email` is on, the body is generated by
`email_sender/outlook.py › build_email_html_body()` — light theme,
"Portfolio Summary" h2 with the project teal accent (`#2e6168`), KPI
grid, then the Instrument Breakdown and By-Analyst tables scraped from
the Summary sheet. Both `OAKS_EM_Stock_PL_Report.xlsx` and the
`current portfolio.xlsx` standalone ValuationA copy are attached using
absolute paths (Outlook COM rejects relative paths).

---

## Inputs & analyst mapping

| Path (default)                              | Purpose                                |
|---------------------------------------------|----------------------------------------|
| `T:\CHARLEMAGNE CAPITAL\Hiport\HP_VAL.xlsm` | End-of-period snapshot (`val_data`).   |
| `T:\CHARLEMAGNE CAPITAL\Hiport\<YYYY>\<Mon>\<YYYYMMDD>.xls(m)` | Start snapshot. |
| `inputs/StockTrList.xlsb`                   | *(retired — trades now from `ccl.dbo.tTRANS`; file archived)* |
| `isin_analyst_map.csv`                      | ISIN/SNAME → analyst code.             |
| `inputs/ca_overrides.csv`                   | *(optional safety hatch — routine CAs auto-derived from snapshot delta; manual rows still win on contested ISINs)* |

`isin_analyst_map.csv` is the single source of truth for analyst
attribution. Resolution order:

1. SNAME override (`config.PERSON_OVERRIDES_BY_SECURITY`, plus rows in
   the CSV with a blank ISIN but populated SNAME).
2. ISIN row in the CSV.
3. `UNASSIGNED`.

Every run appends new ISINs (currently held or recently traded) as
`UNASSIGNED` so they show up on the workbook's `Unassigned` tab. Fill
those rows in the CSV with the correct analyst code (`KX`, `SB`, …) and
they will route correctly on the next run — and stay correct after the
position is sold (the bottler trades sheet has no analyst column).

---

## Outputs

In `cfg.output_dir` (default `out/` next to the package):

| File                                       | When           |
|--------------------------------------------|----------------|
| `OAKS_EM_Stock_PL_Report.xlsx`             | OK runs        |
| `current portfolio.xlsx` (ValuationA copy) | OK runs        |
| `pl_results.parquet`                       | every run      |
| `ca_suggestions_<ts>.csv`                  | when CA gaps detected |
| `OAKS_EM_VALIDATION_FAILURE_<ts>.xlsx`     | critical FAIL  |
| `pl_results_failed_<ts>.parquet`           | critical FAIL  |

Workbook writes are atomic (`tempfile.mkstemp` + `os.replace`); same
for the analyst CSV.

---

## Configuration

All defaults live in [`oefof_pl/config.py`](oefof_pl/config.py) and are
plain module-level constants frozen into a `Config` dataclass at start.
There are **no environment variables** — every override is either an
edit to `config.py` or a CLI flag.

Highlights:

| Constant                  | Default                                              |
|---------------------------|------------------------------------------------------|
| `HP_VAL_PATH`             | `T:\CHARLEMAGNE CAPITAL\Hiport\HP_VAL.xlsm`          |
| `HP_ARCHIVE_ROOT`         | `T:\CHARLEMAGNE CAPITAL\Hiport`                      |
| `BOTTLER_PATH`            | *(deprecated — vestigial label only; trades come from SQL)* |
| `ANALYST_MAP_PATH`        | `isin_analyst_map.csv`                               |
| `OUTPUT_DIR`              | `out/`                                               |
| `START_SNAPSHOT_DEFAULT`  | `"20251231"` *(bump on Jan 1 each year)*             |
| `OEFOF_PCODES`            | OEFOF, OEFOGSSC, OEFOGSSW, OEFOHFSW, OEFOHFSC        |
| `FTSWAP_ISINS`            | `SX7E INDEX`, `GSCBIHKT`                             |
| `ANALYST_CODES`           | IS, HK, JB, SB, KX, VS, AS                           |
| `EMAIL_TO_DEFAULT`        | `oaks@fieracapital.com`                              |
| `EMAIL_SUBJECT_DEFAULT`   | `OEFOF ValuationA and P&L (inc per analyst)`         |

---

## CLI reference

`app.py` and `python -m oefof_pl` accept the same flags. The only
behavioural difference is that `app.py` defaults `--send-email` to **on**.

| Flag                          | Description                                                          |
|-------------------------------|----------------------------------------------------------------------|
| `--start YYYYMMDD`            | Start snapshot. Default: `cfg.start_snapshot_default`.               |
| `--end YYYYMMDD`              | End snapshot. Default: VDATE from `HP_VAL.xlsm`.                     |
| `--hp-val PATH`               | Override `cfg.hp_val_path`.                                          |
| `--analyst-map PATH`          | Override `cfg.analyst_map_path`.                                     |
| `--output-dir PATH`           | Override `cfg.output_dir`.                                           |
| `--apply-suggestions PATH`    | Apply additional CA-overrides CSV. Repeatable.                       |
| `--strict-run`                | Enable production hard gates (release, approvals, reconciliation, anomaly). |
| `--override-approvals PATH`   | CSV approvals register for manual CA/bonus override usage.           |
| `--known-exceptions PATH`     | CSV exception register (`date,scope,key,reason,approver`) for controlled bypasses. |
| `--prior-audit PATH`          | Previous `run_audit_latest.json` used for day-over-day anomaly checks. |
| `--send-email` / `--no-email` | Send (or skip) the Outlook email. `app.py` defaults to **on**.       |
| `--email-to ADDR`             | Recipient(s); comma/semicolon separated.                             |
| `--email-cc ADDR`             | CC recipient(s).                                                     |
| `--email-subject TEXT`        | Override default subject.                                            |
| `--email-body TEXT`           | Plain-text body (HTML summary is appended).                          |
| `--email-from-smtp ADDR`      | Outlook account to send from (matches Account.SmtpAddress).          |

---

## Tests

```powershell
python -m pytest -q
```

257 tests, all green. Coverage is heaviest in the compute, validate,
analyst-map, and email-render modules — anything pure-Python. The COM
modules (`refresh/bottler.py`, `output/valuation.py`,
`email_sender/outlook.py`) have dispatcher-injection seams so the unit
tests run without Outlook or Excel.

A guard test (`tests/test_no_com_in_compute.py`) AST-scans the repo and
fails the build if `win32com` ever leaks into a non-edge module.

---

## Architecture

Bottom-up, strictly layered. Compute is pure; I/O is at the edges.

```
app.py             — top-level launcher (auto-email by default)
oefof_pl/
  __main__.py      — argparse CLI (no auto-email)
  config.py        — paths, fund codes, analyst codes, email defaults
  exceptions.py    — OefofError hierarchy
  pipeline.py      — 15-stage orchestrator
  data/            — HP_VAL.xlsm (val_data) + Bottler .xlsb readers
  compute/         — fx, wac, classify, aggregate, pl   (NO I/O, NO COM)
  analyst/         — ISIN/SNAME → analyst CSV map
  validate/        — VAL-01..VAL-11 critical/warning gates
  output/          — workbook builder (openpyxl), valuation extractor (COM)
  refresh/         — Bottler refresh via Excel COM
  email_sender/    — Outlook COM send + HTML body builder
tests/             — 15 modules, 257 tests
```

The pipeline runs 15 named stages; each one returns a `StageResult` and
critical failures abort cleanly via `_abort()`. `PipelineResult`
(frozen dataclass) carries every output path, the validation report,
and the email result back to the caller.

### Invariants enforced by tests

- VDATE is the only date authority — filenames are never parsed.
- `win32com` only imported in `refresh/bottler.py`,
  `output/valuation.py`, and `email_sender/outlook.py` (AST scan).
- Compute layer takes FX rates as explicit args, never from a global.
- Snapshot `PTVALUE` / `ORD` columns are dropped at the loader.
- Percentages are stored as ratios; Excel formats them at display time.
- Critical validation failure → no P&L workbook (only failure artifacts).
- `pl_results.parquet` is always written via atomic
  `tempfile.mkstemp` + `os.replace`.
- Validation failures listed completely (no truncation).

---

## Project review notes

Recent code review highlights (kept here so the next maintainer doesn't
have to rediscover them):

- `output/workbook.py` is large (~1.4k lines) and deliberately so —
  every sheet's layout, formatting, and totals logic lives in one place.
  The hierarchical `_write_valuation_a()` (≈400 lines) is the most
  complex section; if you touch it, run the full smoke test
  (`python app.py --no-email`) and eyeball the ValuationA tab.
- Email HTML is unbounded — a 500-position fund will produce a long
  table in the body. Acceptable today; if Outlook starts truncating,
  switch the by-analyst block to a CSV attachment.
- Outlook COM has no send timeout. If a hung Outlook ever blocks the
  pipeline, kill the process and re-run.
- `Config` is fully overridable from the CLI; tests build their own
  `Config` pointing at temp dirs.
