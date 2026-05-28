# OEFOF P&L System — Design Document (v2.0)

> Status: **IMPLEMENTED.** This document describes the shipped v2 design;
> the codebase under `oefof_pl/` reflects every decision below and the
> 257-test suite enforces the invariants. Update this header when the
> design changes, not before.
> Audience: maintainer, reviewer, fund operations.
> Goal: make every one of the 11 known bugs **structurally impossible**, not just patched.

---

## 0. Why this document exists

The previous system was four scripts (`master_pl_pipeline.py` → `daily_pl_pipeline.py` → `build_workbook.py` → `build_valuationA.py`) glued together by `run_all.py` plus a `_pipeline_meta.json` side-channel. It accumulated 11 production bugs, the worst of which (BUG-3, BUG-4, BUG-7, BUG-11) caused **multi-million EUR misstatements** that were caught only by the fund manager noticing.

Reading the legacy code, the *root* failure modes were not arithmetic — they were **architectural**:

| Legacy failure mode | Concrete evidence in legacy code | Structural fix in v2 |
|---|---|---|
| Globals leak FX into compute | `daily_pl_pipeline.py` uses `EUR_USD_END`, `all_fx` as module globals inside `local_to_eur` | `local_to_eur_at(amount, ccy, xrate, eur_usd)` — **explicit args only**, no module state |
| Stage handoff via files / globals | `_pipeline_meta.json`, four CSVs, `START_PORT_SHEET` constant | Single in-process `Pipeline` object; DataFrames flow as return values |
| Silent fallback FX | `EUR_USD_START = fx_start.get('EUR', 0.851462)` — magic number, no warning | Missing FX → **WARN logged + recorded in ValidationResult**; never silent |
| Filename-derived dates | `re.match(r'PORT_(\d{2})(\d{2})(\d{4})', name)` — drove BUG-8 | `VDATE` is the **only** date authority; filename regex is *never* trusted, only used as a tiebreaker for archive lookup |
| Compute mixed with COM | `master_pl_pipeline.py` runs the Bottler macro **and** computes FX adjustments in the same function | `refresh/bottler.py` and `email_sender/outlook.py` are the **only** modules allowed to import `win32com`; enforced by an import-lint test |
| Classification reads wrong column | `_instrument()` originally branched on `ORD` column ('A' for everything) | `classify_instrument()` takes the **`CAT` column from aggregated portfolio lines** as primary input; ORD column is *not in the function signature at all* |
| PTVALUE used as EUR | `start_mv = s.get('PTVALUE', 0)` — comment said "EUR (from portfolio)", was actually local CCY | P&L row builder accepts only **EUR-converted** MV fields; the raw `PTVALUE` field is renamed `ptvalue_local` everywhere it appears so any direct use is an obvious type error |
| No tests | Zero tests in legacy | Pure-function compute layer; full pytest suite runs on a machine **without Excel, T:, or win32com** |
| Empty DataFrames returned silently | `df_dec_filtered = pd.DataFrame()` with only a `print` | Loader functions that return 0 rows raise `EmptyDataError` *or* return a `LoadReport(rows=0, reason=...)` that the pipeline must inspect |

Everything below follows from those decisions.

---

## 1. System context (one diagram)

```
                     ┌──────────────────────────────────────────────────────┐
                     │                T:\CHARLEMAGNE CAPITAL\Hiport          │
                     │   HP_VAL.xlsm   +   {YYYY}\{Mon}\{YYYYMMDD}.xls(m)   │
                     └─────────────────────┬────────────────────────────────┘
                                           │ read-only
                                           ▼
   ┌───────────────────┐    macro     ┌────────────┐    DataFrames    ┌─────────────────┐
   │ refresh/bottler.py│ ───────────▶ │ Bottler    │ ───────────────▶ │ data/loader.py  │
   │   (win32com)      │              │ .xlsb      │                  │ data/normalise  │
   └───────────────────┘              └────────────┘                  └────────┬────────┘
                                                                                │ pure pd.DataFrame
                                                                                ▼
   ┌──────────────────────────────────────────────────────────────────────────────────┐
   │                              compute/  (PURE, no I/O, no COM)                    │
   │   fx.local_to_eur_at()  classify.classify_instrument()  wac.WACEngine            │
   │   pl.compute_positions()  →  (pl_df, income_df)                                  │
   └──────────────────────────────────────────────────────────────────────────────────┘
                                           │
                                           ▼
              ┌────────────────────────────────────────────────────┐
              │ validate/checks.run_all_checks() → ValidationResult │
              │      VAL-01..VAL-11   (FAIL aborts; WARN continues) │
              └────────────────────────────────────────────────────┘
                                           │  if no FAIL
                       ┌───────────────────┴───────────────────┐
                       ▼                                       ▼
        ┌────────────────────────────┐        ┌────────────────────────────────┐
        │ output/workbook.py         │        │ output/valuation.py            │
        │ OAKS_EM_Stock_PL_Report    │        │ current portfolio.xlsx         │
        │ (9 sheets, openpyxl only)  │        │ (uses HP_VAL via refresh/COM)  │
        └────────────────────────────┘        └────────────────────────────────┘
                       │                                       │
                       └───────────────────┬───────────────────┘
                                           ▼
                                ┌─────────────────────────┐
                                │ email_sender/outlook.py │  (optional, never aborts)
                                └─────────────────────────┘
```

**Hard rule:** the only modules that may `import win32com` are `refresh/bottler.py`, `output/valuation.py` (for the HP_VAL pivot extraction step) and `email_sender/outlook.py`. A test (`tests/test_no_com_in_compute.py`) greps the package and fails the build if any other module imports it.

---

## 2. Layered architecture

| Layer | Modules | Allowed dependencies | Forbidden |
|---|---|---|---|
| **Config** | `config.py` | stdlib only | importing any other oefof_pl module |
| **Data** | `data/loader.py`, `data/normalise.py` | `config`, `pandas`, `openpyxl`, `pyxlsb`, `xlrd` | `win32com`, `compute/*`, network |
| **Compute (pure core)** | `compute/fx.py`, `compute/wac.py`, `compute/classify.py`, `compute/pl.py` | `config`, `pandas`, `numpy`, stdlib | **any** I/O, COM, file paths, logging side-effects beyond `logger.warning` |
| **Analyst** | `analyst/map.py` | `config`, `pandas`, csv | win32com |
| **Validate** | `validate/checks.py` | `config`, `pandas`, `numpy` | I/O |
| **Output** | `output/styles.py`, `output/workbook.py`, `output/valuation.py` | `openpyxl`; `valuation.py` may use `win32com` for pivot extraction only | direct re-computation of P&L |
| **Refresh** | `refresh/bottler.py` | `win32com`, `pythoncom` | `compute/*` |
| **Email** | `email_sender/outlook.py` | `win32com` | `compute/*` |
| **Pipeline** | `pipeline.py` | everything above | adding any computation logic of its own |

The `Pipeline` class is **pure orchestration**. If it ever contains an `if`/`for` that touches a P&L number, the design has been violated.

---

## 3. Data contracts (the canonical schemas)

All DataFrames passed across module boundaries have a fixed schema. They are validated at the boundary using a tiny `_assert_schema(df, required_cols)` helper. Out-of-schema columns are allowed; missing or wrongly-typed columns raise `SchemaError` immediately — *not* silently coerced.

### 3.1 `port_df` (output of `data/loader.load_snapshot`)

| Column | dtype | Notes |
|---|---|---|
| `ISIN` | str (stripped, upper) | Normalised at load time |
| `SNAME` | str | |
| `CCY` | str | |
| `XRATE` | float | local per 1 USD; 1.0 for USD; **NaN** if missing (not 0, not 1) |
| `UNITS` | float | signed |
| `PTVALUE_LOCAL` | float | **renamed from PTVALUE at load time** to make local-vs-EUR confusion impossible |
| `PTCOST_LOCAL` | float | renamed from PTCOST |
| `NUCOST_LOCAL` | float | renamed from NUCOST |
| `LST_PRICE_LOCAL` | float | renamed from LST_PRICE |
| `CAT` | str | `'ORD'`, `'FUT'`, etc. **Sole instrument-type input.** |
| `LS` | str | `'L'` / `'S'` |
| `PCODE_ORIG` | str | |
| `VDATE` | datetime64 | **Authoritative valuation date** |
| `EXCODE1` | str | → Country |

> The `_LOCAL` suffix is non-negotiable. `compute/pl.py` will not accept a column called `PTVALUE` — it will raise `SchemaError`. This is BUG-11 made impossible.

### 3.2 `trades_df` (output of `data/loader.load_bottler` filtered to `OEFOF_PCODES`)

| Column | dtype | Notes |
|---|---|---|
| `PCODE_ORIG` | str | |
| `ISIN` | str | normalised |
| `SNAME` | str | unified from `SNAME` / `SHORT_NAME` |
| `CCY` | str | per-trade CCY (matters for hybrids) |
| `T` | str | `'P'` / `'S'` / `'I'` |
| `CDATE` | datetime64 | parsed at load |
| `UNITS` | float | always positive in source |
| `GROSSPRICE_LOCAL` | float | renamed from GROSSPRICE |
| `BUY_NET_LOCAL` | float | renamed from `BUY NET` |
| `SELL_NET_LOCAL` | float | renamed from `SELL NET` |
| `INCOME_LOCAL` | float | renamed from `INCOME` |
| `BCOMM_LOCAL` | float | |
| `EXPENSES_LOCAL` | float | |
| `CCL_CAT` | str | classification fallback only |
| `DELETED` | str | rows where `DELETED=='Y'` are dropped at load |

> `TRADEGROSS_USD` is **deliberately not in this schema**. It is dropped at the loader layer. Compute code cannot accidentally use it because it does not exist in the DataFrame they receive. This is BUG-4 made structurally impossible.

### 3.3 `agg_dict` (output of `compute.classify.aggregate_portfolio`)

```python
agg: dict[str, AggLine] = {
    'AEE001901017': AggLine(
        sname='EMAAR DEVELOPMENT', ccy='AED', ls='L', country='AE',
        units=893072.0,
        ptvalue_local=13_530_000.0,        # local CCY — guaranteed
        ptcost_local=...,
        nucost_local=9.1042,                # weighted average if multi-line
        lst_price_local=15.15,
        cat='ORD',                          # escalated to 'FUT' if any line was FUT
        xrate=3.6729,
        n_source_rows=1,
    ),
    ...
}
```

`AggLine` is a `@dataclass(frozen=True)`. Frozen dataclasses prevent mutation bugs and give clear KeyError messages — better than `dict.get(..., 0)` which masked BUG-11.

### 3.4 `pl_df` (output of `compute.pl.compute_positions`)

The 30 columns from §9 of the prompt, in a deterministic order, all `float64` for numeric columns. Percentage columns are stored as **ratios** (0.44 = 44%) per BUG-1. A schema-validation step at the end of `compute_positions` asserts this.

### 3.5 `income_df`

Long-format, one row per `I`-trade, with the per-trade `CCY` preserved (BUG-7).

---

## 4. The bug → structural fix matrix

| Bug | Root cause | Structural prevention in v2 | Test that proves it |
|---|---|---|---|
| **BUG-1** Pct ×100 wrong | Stored as 44 not 0.44 | `pl_df` percentage columns produced **only** by `_to_ratio(numerator, denominator)` helper that returns `num/den`. Output layer applies `'0.00%'` format. A schema check asserts `\|max(pct_col)\| ≤ 5` at end of compute. | `test_pl.py::test_total_pct_is_ratio_not_percent` |
| **BUG-2** No Start Price | Column missing from output | Column is in the `pl_df` schema; output layer iterates the schema. A test verifies the column exists in the saved workbook. | `test_workbook.py::test_current_holdings_columns` |
| **BUG-3** SWAP double count | Code added Income + ΔPTVALUE | `compute_positions` has two **separate code paths** dispatched on `Instrument`: `_compute_ord_pl(...)` and `_compute_swap_pl(...)`. The SWAP path **does not receive** `start_mv_eur` or `end_mv_eur` as inputs to its total — they are passed only as informational outputs. Cannot accidentally add. | `test_pl.py::test_swap_total_equals_income_only` and `VAL-04` |
| **BUG-4** TRADEGROSS_USD per-trade FX | Code converted each trade at trade-date USD rate | `TRADEGROSS_USD` is dropped at the loader. `local_to_eur_at` requires explicit `(xrate, eur_usd)` pair. Trade conversion call sites pass `fx_end[ccy]`, `eur_usd_end` — uniform period FX. | `test_pl.py::test_corporate_action_pair_nets_to_zero` |
| **BUG-5** SWAP cost basis inflated | Used \|StartMV\|+\|BuyEUR\| | `_compute_swap_pl` returns `cost_basis = abs(nucost_local * units)` converted via `local_to_eur_at`. The ORD code path is the only one that uses `\|StartMV\|+\|BuyEUR\|`. | `test_pl.py::test_swap_cost_basis_is_notional` |
| **BUG-6** No First Trade Date | Column missing | Schema requires `First Trade Date`. Filled only when `ISIN not in start_isins`. | `test_pl.py::test_first_trade_date_only_for_new_positions` |
| **BUG-7** Vietnam Dairy hybrid | (a) used ORD col, (b) no escalation, (c) wrong income FX | (a) `ORD` column dropped from schema. (b) `aggregate_portfolio` escalates `cat='FUT'` if any source row is FUT. (c) income loop iterates per-trade and reads `trade['CCY']`, *not* the position CCY. | `test_classify.py::test_hybrid_escalates_to_swap` + `test_pl.py::test_income_uses_per_trade_ccy` |
| **BUG-8** Filename date typo | Sheet name parsed for date | `data/loader.load_snapshot` returns `(df, vdate)`; vdate is read from the `VDATE` column. Filename is **never parsed for date**; the only thing the filename gives is "this is the file at this archive path", and that mapping is `find_snapshot_path(date) → path`, not the inverse. The pipeline asserts `vdate == requested_date` (within 0 days) and **FAIL**s otherwise. | `test_loader.py::test_vdate_is_authoritative` and a fixture file with mismatched name vs VDATE |
| **BUG-9** ORD column for class | Used the always-'A' column | `classify_instrument` signature **does not accept** an `ord_flag` parameter. The `port_df` schema does not include the legacy `ORD` column. | `test_classify.py::test_classifier_does_not_use_ord_column` (greps source) |
| **BUG-10** SWAP decomp wrong | Realised=Income, Unrealised=PTVALUE | `_compute_swap_pl` returns hard-coded `realised=0.0, unrealised=0.0`. There is no input parameter that could change that. | `test_pl.py::test_swap_realised_and_unrealised_are_zero` |
| **BUG-11** PTVALUE used as EUR | Raw PTVALUE used in EUR formula | Schema renames to `PTVALUE_LOCAL`. `compute_positions` builds `start_mv_eur = local_to_eur_at(agg.ptvalue_local, agg.ccy, fx_start[ccy], eur_usd_start)` for ORD; the SWAP path uses `end_mv_eur` only for informational display. The string `PTVALUE` does not appear anywhere in compute code. | `test_pl.py::test_ord_ptvalue_converted_with_period_fx` (KRW position with real-world numbers) + `VAL-11` |

The matrix is the contract for the test suite. Each cell with "Test" is a required pytest test.

---

## 5. FX module design (the smallest, most-tested module)

`compute/fx.py` is ~40 lines and pure. Three responsibilities:

```
build_fx_dict(port_df) -> dict[str, float]      # CCY -> XRATE; warns on duplicates with diffs
get_eur_usd(fx) -> float                         # raises FXMissingError if 'EUR' not in fx — no silent default
local_to_eur_at(amount, ccy, xrate, eur_usd) -> float
```

Design decisions:

1. **No fallback EUR/USD.** The legacy `0.868018` fallback was a silent BUG-11-class problem waiting to happen. If EUR is not in the FX dict, **the pipeline aborts** (the snapshot is corrupt — better to fail loudly than emit a wrong report).
2. **`local_to_eur_at` raises on `xrate=0` for non-USD/EUR.** Returning 0.0 like the legacy code did was a silent zeroing of real positions. Caller is responsible for checking before calling, or wrapping in `try/except` if it really wants to zero.
3. **Two FX dicts flow through compute:** `fx_start` and `fx_end`. They are passed as separate arguments to `compute_positions` and are *never merged* into a single `all_fx` dict (the legacy `{**fx_start, **fx_end}` merge silently lost start-period rates for any CCY also present at end).

---

## 6. WAC engine

`WACEngine` is a small state machine:

```
state: held_units (float), total_cost_local (float), realised_local (float)

methods:
  process_buy(units, net_amount_local, gross_price_local)
  process_sell(units, net_amount_local, gross_price_local)

properties:
  wac, held, total_cost, realised
```

Decisions:
- Trades **must be fed in `CDATE` order**; the engine asserts this (`prev_cdate <= cdate`) and raises if violated. The legacy `pd.concat([buys,sells]).sort_values('CDATE')` was easy to forget; making it an assertion catches future regressions.
- Bonus shares (`gross_price == 0 and net == 0` on a buy) → dilute WAC, no cost added — exactly per spec.
- Zero-cost transfers (same on a sell) → remove cost at WAC.
- Sells **never** modify WAC. Asserted in tests.
- Engine is initialised with `(initial_units, initial_nucost_local)` from the start-period `AggLine`. This means the engine knows nothing about the portfolio snapshot directly — it is testable without any DataFrames.

---

## 7. Classification

```
classify_instrument(
    isin: str,
    cat_values: list[str],          # CAT values from all aggregated portfolio lines for this ISIN
    sname: str,
    trade_ccl_cats: list[str],      # CCL_CAT from trades for this ISIN
    ftswap_isins: frozenset[str],
) -> Literal['ORD', 'SWAP', 'FTSWAP']
```

The function is pure and takes no DataFrames — only the few primitive lists/strings it needs. This makes it trivially testable and impossible to accidentally read the wrong column. Decision order is exactly the 6-rule priority from §6 of the prompt.

`aggregate_portfolio` is a separate function that:
- groups `port_df` by `ISIN`,
- sums `units`, `ptvalue_local`, `ptcost_local`,
- weighted-averages `nucost_local` by `abs(units)` (matching legacy `_aggregate_port`),
- escalates `cat` to `'FUT'` if any row is FUT,
- takes `ccy`, `xrate`, `sname`, `country`, `lst_price_local`, `ls` from the **first row** (these should be invariant across rows for a given ISIN; if not, it logs a WARN and continues).

The output is `dict[str, AggLine]` (frozen dataclass) — not a dict of dicts. `AggLine.ptvalue_eur` does not exist; it is computed in `compute/pl.py` only.

---

## 8. P&L compute (`compute/pl.py`)

Single entry point:

```python
def compute_positions(
    trades_df: pd.DataFrame,
    start_agg: dict[str, AggLine],
    end_agg: dict[str, AggLine],
    fx_start: dict[str, float],
    fx_end: dict[str, float],
    eur_usd_start: float,
    eur_usd_end: float,
    end_date_ts: pd.Timestamp,
    analyst_resolver: Callable[[str, str], tuple[str, bool]],   # (isin,sname) -> (code, assigned)
    ftswap_isins: frozenset[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
```

Internally it splits per-ISIN, calls one of two private helpers:

- `_compute_ord_position(isin, start, end, trades, fx_start, fx_end, eur_usd_start, eur_usd_end) -> PositionRow`
- `_compute_swap_position(isin, start, end, trades, fx_end, eur_usd_end) -> PositionRow`

Critical design choices:

1. **The SWAP function does not take `fx_start` or `eur_usd_start`** — it cannot accidentally produce a start-MV-based total.
2. **The ORD function builds `start_mv_eur` and `end_mv_eur` exactly once each**, at the top, using `local_to_eur_at`. They are computed values; the raw `ptvalue_local` is never used downstream.
3. **Income is always computed by iterating `I`-trades and using each trade's own CCY** — the position CCY is not even passed to that loop.
4. **The grand-total identity** `Σ Total = Σ(EndMV − StartMV) + Σ Sell − Σ Buy + Σ Income` is checked inside `compute_positions` for ORD-only positions before returning. (SWAP positions satisfy `Total = Income` separately.) If it fails by more than 1 EUR, a `ComputeIntegrityError` is raised — this is *before* `validate/` even runs, because it indicates a code bug, not a data issue.

`PositionRow` is a `@dataclass(frozen=True)` with all 30 fields. The DataFrame is built by `pd.DataFrame([asdict(p) for p in rows])` — guarantees column order and types.

---

## 9. Analyst resolution

```
load_map(csv_path) -> dict[str, AnalystRecord]
apply_overrides(sname, default, overrides) -> str
lookup_analyst(isin, sname, map_dict, overrides) -> tuple[code, is_assigned]
```

The resolver passed into `compute_positions` is a `functools.partial(lookup_analyst, map_dict=..., overrides=...)` — keeps `compute/pl.py` decoupled from any I/O.

`update_map` writes atomically (temp file + rename) so a Ctrl-C mid-write cannot corrupt the CSV.

---

## 10. Validation gates

`validate/checks.py` exposes one function and three dataclasses:

```python
@dataclass(frozen=True)
class CheckResult:
    id: str; name: str; status: Literal['PASS','WARN','FAIL']
    detail: str; failures: list[dict]

@dataclass(frozen=True)
class ValidationResult:
    checks: list[CheckResult]
    @property def has_critical(self) -> bool: ...
    @property def has_warnings(self) -> bool: ...
```

Run order:

1. VAL-01 unit reconciliation (FAIL)
2. VAL-02 duplicate trades (FAIL)
3. VAL-11 PTVALUE conversion sanity (FAIL) — runs **before** VAL-03 because if VAL-11 fails, VAL-03 will also fail uselessly
4. VAL-03 grand total identity (FAIL)
5. VAL-04 SWAP identity (FAIL)
6. VAL-05..10 (WARN)

`Pipeline.run()` writes Excel **only if `not result.has_critical`**. If critical, it writes a **single-sheet error report** `OAKS_EM_VALIDATION_FAILURE_{timestamp}.xlsx` listing every failing row with ISIN/name/expected/got, and returns False. The fund manager never sees a wrong number.

`failures` lists are **complete** — never truncated to the first 3 like the legacy code did.

---

## 11. Output layer

`output/workbook.py` consumes only `(pl_df, income_df, validation_result, period_meta)` and produces the 9-sheet workbook. Layout constants live in `output/styles.py`:

```python
COLOR_HEADER = '1B474D'
COLOR_POS    = '437A22'
COLOR_NEG    = 'A13544'
COLOR_TOTAL  = 'E8F4F5'
FONT_BODY    = Font(name='Calibri', size=10)
...
COLUMN_LAYOUT_HOLDINGS = [
    ColSpec('ISIN',           'ISIN',                 width=14, fmt='text'),
    ColSpec('Stock Name',     'Stock Name',           width=32, fmt='text'),
    ...
]
```

`ColSpec` is a tiny dataclass — `(header_label, df_column, width, format_kind)`. The same list drives header writing, data writing, and the total row. Adding a column is a one-line change; the legacy code required edits in 4 places per column.

The Validation Report sheet is built **last** so it can include a "writing took N seconds" diagnostic.

`output/valuation.py` is the only output module that uses `win32com` — it opens HP_VAL.xlsm, sets the pivot page filter to `OAKS Emerging and Frontier Fund`, refreshes, extracts cols A:L of `ValuationA`, then closes Excel. The post-processing (FMCID lookup, Summary, per-analyst tabs) is pure openpyxl.

The legacy file-lock workaround (ROT enumeration to close stale Excel handles) is preserved because it's a real-world necessity — but isolated in `_close_open_excel_workbook(path)` inside `output/valuation.py` and unit-tested with a mock.

---

## 12. Pipeline orchestration

```python
class Pipeline:
    def __init__(self, cfg: Config): ...
    def refresh_bottler(self) -> StageResult: ...
    def load_data(self, start: str, end: str | None) -> StageResult: ...
    def compute(self) -> StageResult: ...
    def validate(self) -> ValidationResult: ...
    def write_pl_report(self, out_dir: Path) -> StageResult: ...
    def write_valuation(self, out_dir: Path) -> StageResult: ...
    def send_email(self, **kw) -> StageResult: ...
    def run(self, ...) -> int:    # returns process exit code
```

`StageResult` is `(ok: bool, message: str, artifacts: dict)`. `Pipeline.run` is a flat list of `if not stage.ok: return abort(...)` lines — no nested logic.

The **end-date defaulting rule**:
- If `--end` not passed, look for today's `HP_VAL.xlsm`, read its `VDATE`. That's the end date.
- If today's snapshot doesn't exist, abort with a clear message — do NOT silently use yesterday's.

Start-date defaulting: `START_SNAPSHOT_DEFAULT = "20251231"` from config (i.e., last year-end). Operator can override.

---

## 13. Configuration

`config.py` is the **only** place with paths, codes, and tunable constants. All other modules receive a `Config` instance (or specific fields) as a parameter. No module reads `config.OEFOF_PCODES` directly — it's passed in.

```python
@dataclass(frozen=True)
class Config:
    hp_val_path: Path
    hp_archive_root: Path
    bottler_path: Path
    analyst_map_path: Path
    output_dir: Path
    oefof_pcodes: tuple[str, ...]
    ftswap_isins: frozenset[str]
    analyst_codes: dict[str, str]
    person_overrides: dict[str, str]
    start_snapshot_default: str
    target_pname: str = "OAKS Emerging and Frontier Fund"
    email_to_default: str = "oaks@fieracapital.com"
    email_subject_default: str = "OEFOF ValuationA and P&L (inc per analyst)"

def load_default_config() -> Config: ...     # for prod
```

Config is **immutable**. `dataclasses.replace(cfg, output_dir=tmp)` for testing.

---

## 14. Logging

One logger `oefof_pl`, one handler installed by the script entry points (not the library). Format:

```
2026-04-29 09:12:01 INFO  load_data    Loaded start snapshot: 245 rows, VDATE=2025-12-31
2026-04-29 09:12:02 INFO  load_data    Loaded end snapshot:   238 rows, VDATE=2026-04-29
2026-04-29 09:12:02 WARN  load_bottler PCODE filter excluded 14_822 of 20_350 rows (5_528 retained)
2026-04-29 09:12:05 INFO  compute      127 positions: 98 ORD, 27 SWAP, 2 FTSWAP
2026-04-29 09:12:05 INFO  fx           EUR/USD start=0.851462 end=0.868018
2026-04-29 09:12:06 INFO  validate     PASS:9  WARN:2  FAIL:0
2026-04-29 09:12:11 INFO  pipeline     Wrote OAKS_EM_Stock_PL_Report.xlsx
```

WARN-level entries always include counts so silent-empty bugs can never recur unnoticed.

---

## 15. Test strategy

Three tiers:

| Tier | What it tests | Runs without |
|---|---|---|
| **Unit** (`test_fx`, `test_wac`, `test_classify`, `test_analyst`) | Pure functions, edge cases | Excel, T:, COM, win32com |
| **Component** (`test_pl`, `test_validate`) | `compute_positions` end-to-end with synthetic DataFrames; every bug-regression case from §4 | Excel, T:, COM |
| **Smoke / integration** (`test_workbook_smoke`) | `output/workbook.build` runs to completion, file opens, expected sheets exist, Current Holdings has 28 columns | win32com (uses openpyxl read-back) |

Tests for COM modules (`refresh/bottler.py`, `output/valuation.py` COM extraction, `email_sender/outlook.py`) use `unittest.mock.patch('win32com.client.DispatchEx')` — they verify the dispatch sequence and argument shapes, not real Excel behaviour. A separate manual smoke checklist (in README) covers the COM-touching path on a real Windows box.

A **lint test** (`test_no_com_in_compute.py`) walks `oefof_pl/` source, parses each file with `ast`, and fails if any module outside the allow-list imports `win32com` or `pythoncom`.

A **bug-regression test** (`test_bug_regressions.py`) has one test per bug from §4, named `test_bug_{n}_{slug}`. Each carries a docstring quoting the bug from `BUG_FIXES_LOG.docx` and an explicit assertion. These are non-negotiable and never deleted.

Coverage target: **100% on `compute/`, `validate/`, `analyst/`**. ≥80% overall.

---

## 16. Directory layout (final)

```
v2/oefof_pl/
├── DESIGN.md                       ← this file
├── pyproject.toml
├── README.md
├── isin_analyst_map.csv            ← template, headers only on first run
├── oefof_pl/
│   ├── __init__.py
│   ├── config.py
│   ├── pipeline.py
│   ├── data/
│   │   ├── __init__.py
│   │   ├── loader.py
│   │   └── normalise.py
│   ├── compute/
│   │   ├── __init__.py
│   │   ├── fx.py
│   │   ├── wac.py
│   │   ├── classify.py
│   │   └── pl.py
│   ├── analyst/
│   │   ├── __init__.py
│   │   └── map.py
│   ├── validate/
│   │   ├── __init__.py
│   │   └── checks.py
│   ├── output/
│   │   ├── __init__.py
│   │   ├── styles.py
│   │   ├── workbook.py
│   │   └── valuation.py
│   ├── refresh/
│   │   ├── __init__.py
│   │   └── bottler.py
│   └── email_sender/
│       ├── __init__.py
│       └── outlook.py
├── scripts/
│   ├── run_daily.py
│   └── update_analysts.py
└── tests/
    ├── conftest.py
    ├── test_fx.py
    ├── test_wac.py
    ├── test_classify.py
    ├── test_analyst.py
    ├── test_pl.py
    ├── test_validate.py
    ├── test_loader.py
    ├── test_workbook_smoke.py
    ├── test_no_com_in_compute.py
    └── test_bug_regressions.py
```

---

## 17. Build order (proposed)

The package will be built bottom-up so that every layer is fully tested before anything depends on it.

| Phase | Deliverable | Acceptance |
|---|---|---|
| 1 | `config.py`, `data/normalise.py` schemas, exception classes | importable, pytest collects 0 tests |
| 2 | `compute/fx.py` + `test_fx.py` | all FX tests green |
| 3 | `compute/wac.py` + `test_wac.py` | all WAC tests green |
| 4 | `compute/classify.py` + `test_classify.py` | all classification tests green |
| 5 | `analyst/map.py` + `test_analyst.py` | green |
| 6 | `compute/pl.py` + `test_pl.py` + `test_bug_regressions.py` | **all 11 bug regressions green** |
| 7 | `validate/checks.py` + `test_validate.py` | green |
| 8 | `data/loader.py` + `test_loader.py` (with synthetic .xlsx fixtures) | green |
| 9 | `output/styles.py`, `output/workbook.py` + `test_workbook_smoke.py` | green |
| 10 | `output/valuation.py` (mocked COM in tests) | green |
| 11 | `refresh/bottler.py`, `email_sender/outlook.py` (mocked COM) | green |
| 12 | `pipeline.py`, `scripts/run_daily.py`, `scripts/update_analysts.py` | end-to-end dry run passes on synthetic inputs |
| 13 | `pyproject.toml`, `README.md`, `test_no_com_in_compute.py` | lint test green; `pip install -e .[dev]` works |

Each phase ends with `pytest -q`. No phase begins until the previous phase's tests are 100% green.

---

## 18. Open design questions — RESOLVED

The reviewer answers from 2026-04-29 are recorded below. These are now fixed inputs to implementation.

| # | Question | Decision |
|---|---|---|
| 1 | Pre-2008 `.xls` schema support | **Out of scope.** First release supports YTD only (start = `20251231`, end = today's `HP_VAL.xlsm`). The `data/loader.py` will only handle the modern `.xlsm` schema. Pre-2008 `.xls` and the `hiportfolio_valuationDDMMYYYY.xls` filename pattern are deliberately not implemented. The compute and output layers are built so that **all P&L rows include enough raw fields (Analyst, CCY, Country, Instrument, L/S, Start MV, End MV, Realised, Unrealised, Income)** to feed downstream charts (per-analyst, per-instrument, per-country, per-CCY) without recomputation. A `pl_results.parquet` snapshot is dropped alongside the workbook on every run for any future BI tooling. |
| 2 | End-date defaulting | **Option (a): end = `VDATE` of live `HP_VAL.xlsm`.** Operator can override via `--end YYYYMMDD`. Today's calendar date is never assumed. |
| 3 | PCODE filter list | **Use the legacy 5-code list:** `['OEFOF', 'OEFOGSSC', 'OEFOGSSW', 'OEFOHFSW', 'OEFOHFSC']`. `OEFOFMR` and the other 9 codes from the prompt are excluded — including them would double-count vs `OEFOF`. Documented in `config.py` with the rationale. |
| 4 | `isin_analyst_map.csv` location | **Lives in the package root only.** No OneDrive, no env-var. Operator maintains it manually via `scripts/update_analysts.py`. |
| 5 | Valuation Summary layout | **Keep the legacy 8-column layout** (Person, Long Exposure, Long %, Short Exposure, Short %, Net Exposure, Long P&L, Short P&L). Can trim later. |
| 6 | Email body format | **Plain-text only.** No HTML. Body = the same KPI summary that prints to stdout (totals, per-analyst totals, per-instrument totals, validation status). **For testing, recipient hard-coded to `kgontar@fieracapital.com`** instead of `oaks@fieracapital.com`. A `--email-to` flag still works for overrides; production switch-over is a single-line config change. |
| 7 | `START_SNAPSHOT_DEFAULT` rollover | **Manual annual roll.** `config.py` ships with `START_SNAPSHOT_DEFAULT = "20251231"` (the YTD anchor for the 2026 reporting year). Annual update is documented in README under "Year-end maintenance." |
| 8 | Canonical analyst list | **7 analysts:** `KX, SB, VS, HK, AS, IS, JB`. Full names TBD per analyst — `AS` is new (not in original prompt list). The map is a `dict[str, str]` in `config.py`; full names default to the code itself when unknown and can be filled in over time. The pipeline never blocks on a missing full name. |

### Other notes captured from the review

- **Schema language must be finance-correct and client-presentable.** Output column headers will use industry-standard wording: "Market Value (Start, EUR)", "Market Value (End, EUR)", "Realised P&L (EUR)", "Unrealised P&L (EUR)", "Income / Dividends (EUR)", "Total Return (EUR)", "Total Return (%)", "Notional Exposure (EUR)" for swap cost basis, etc. Internal column names in `pl_df` track the same wording one-to-one. The `_LOCAL` suffix is kept on **internal** schema fields (e.g. `ptvalue_local`) but never appears in any user-facing sheet.
- **VAL-11 before VAL-03 ordering.** Confirmed acceptable; the rationale (VAL-11 fixes the inputs that VAL-03 sums) stands.
- **Abort-on-CRITICAL writes no Excel.** Confirmed acceptable. The behaviour explicitly **does not** interfere with computed data — `pl_df` and `income_df` are always produced and persisted to a side-channel `pl_results_failed_{timestamp}.parquet` alongside the validation report so the developer can debug, but the formatted Excel report — the thing the fund manager sees — is never written when a critical check fails.
- **Build proceeds in phases (§17).** Each phase ends with `pytest -q` green before the next begins.

---

## 19. What this design explicitly is NOT

- It is not a Spark/streaming pipeline. The data volume (≤ 5,000 trades, ≤ 250 positions per period) makes pandas in-memory the right tool.
- It is not a database-backed system. Source-of-truth stays in HP_VAL and the Bottler. We only persist `isin_analyst_map.csv`.
- It is not a web service. CLI-only; the operator's one command is `python run_daily.py`.
- It does not refactor the Bottler `.xlsb` macro. We continue to call it via COM exactly as today.
- It does not rewrite HP_VAL extraction logic in `output/valuation.py` to avoid COM. The pivot table is the only reliable filter mechanism for that file. COM stays — but isolated.

---

## 20. Sign-off checklist — RESOLVED

| Item | Decision |
|---|---|
| §3 schemas | **Approved** — but all user-facing column names use finance-standard, client-presentable wording (Market Value, Total Return, Notional Exposure, etc.). The `_LOCAL` suffix is internal only and never appears in any sheet. |
| §4 bug → fix mapping | **Approved as written.** No additional bugs identified by the reviewer. |
| §10 VAL-11 before VAL-03 | **Approved.** |
| §12 abort-on-CRITICAL | **Approved**, with the explicit guarantee that abort never destroys computed data — `pl_df` and `income_df` are still persisted to a `pl_results_failed_{timestamp}.parquet` for diagnosis. |
| §17 build phases | **Approved.** Phase-by-phase, `pytest -q` green between phases. |
| §18 open questions | **All 8 resolved** above. |

**This design is approved for implementation. Phase 1 begins on next instruction.**
