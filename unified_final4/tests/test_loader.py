"""Tests for data.loader — HP_VAL snapshot reader and Bottler trade reader.

These tests build small real .xlsx fixtures (openpyxl) so the pandas
read path is exercised end-to-end. The Bottler reader (.xlsb) is tested
by monkey-patching ``pandas.read_excel`` because writing .xlsb files in
Python is impractical; the loader's post-read normalisation is what
matters.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from oefof_pl.data import loader
from oefof_pl.data.loader import (
    assert_vdate_matches,
    find_snapshot_path,
    load_bottler,
    load_snapshot,
)
from oefof_pl.exceptions import (
    DataLoadError,
    EmptyDataError,
    SchemaError,
    VDateMismatchError,
)


# ─── Fixtures ──────────────────────────────────────────────────────────────

def _write_snapshot_xlsx(path: Path, rows: list[dict], sheet: str = "val_data"):
    """Write a minimal HP_VAL-style workbook to ``path``."""
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=sheet, index=False)


def _row(**kw) -> dict:
    """Build a portfolio row in legacy column names (PTVALUE not _LOCAL)."""
    base = {
        "ISIN": " kr7005930003 ", "SNAME": "SAMSUNG ELEC", "CCY": "krw ",
        "XRATE": 1300.0, "UNITS": 100.0,
        "PTVALUE": 7_000_000.0, "PTCOST": 5_000_000.0,
        "NUCOST": 50_000.0, "LST_PRICE": 70_000.0,
        "CAT": "ORD", "LS": "L", "PCODE_ORIG": "OEFOF",
        "VDATE": "2026-01-31", "EXCODE1": "KR",
        "ORD": "A",  # legacy column that must be dropped
    }
    base.update(kw)
    return base


# ─── load_snapshot — happy path ────────────────────────────────────────────

def test_load_snapshot_renames_ptvalue_and_drops_ord(tmp_path: Path):
    p = tmp_path / "HP_VAL.xlsm"
    _write_snapshot_xlsx(p, [_row()])
    df, vdate = load_snapshot(p)

    assert "PTVALUE_EUR" in df.columns
    assert "PTVALUE" not in df.columns
    assert "NUCOST_LOCAL" in df.columns
    assert "LST_PRICE_LOCAL" in df.columns
    assert "ORD" not in df.columns  # BUG-9: dropped at loader
    assert df["ISIN"].iloc[0] == "KR7005930003"
    assert df["CCY"].iloc[0] == "KRW"
    assert vdate == pd.Timestamp("2026-01-31")


def test_load_snapshot_returns_vdate_from_column():
    """VDATE source of truth: the column, not the filename. BUG-8."""
    # Filename says one date; VDATE column says another. We MUST get the column.
    import tempfile
    with tempfile.TemporaryDirectory() as td:
        # Filename intentionally misleading.
        p = Path(td) / "PORT_31122025.xlsm"
        _write_snapshot_xlsx(p, [_row(VDATE="2026-01-31")])
        df, vdate = load_snapshot(p)
        assert vdate == pd.Timestamp("2026-01-31")


# ─── load_snapshot — error paths ──────────────────────────────────────────

def test_load_snapshot_missing_file(tmp_path: Path):
    with pytest.raises(DataLoadError):
        load_snapshot(tmp_path / "nope.xlsm")


def test_load_snapshot_missing_sheet(tmp_path: Path):
    p = tmp_path / "HP_VAL.xlsm"
    _write_snapshot_xlsx(p, [_row()], sheet="WrongName")
    with pytest.raises(DataLoadError):
        load_snapshot(p)


def test_load_snapshot_empty(tmp_path: Path):
    p = tmp_path / "HP_VAL.xlsm"
    # An empty sheet.
    with pd.ExcelWriter(p, engine="openpyxl") as w:
        pd.DataFrame().to_excel(w, sheet_name="val_data", index=False)
    with pytest.raises(EmptyDataError):
        load_snapshot(p)


def test_load_snapshot_schema_error_missing_column(tmp_path: Path):
    p = tmp_path / "HP_VAL.xlsm"
    bad = _row()
    del bad["UNITS"]
    _write_snapshot_xlsx(p, [bad])
    with pytest.raises(SchemaError):
        load_snapshot(p)


def test_load_snapshot_mixed_vdates_raises(tmp_path: Path):
    """A snapshot file with two distinct VDATE values is structurally wrong."""
    p = tmp_path / "HP_VAL.xlsm"
    _write_snapshot_xlsx(p, [
        _row(ISIN="X", VDATE="2026-01-30"),
        _row(ISIN="Y", VDATE="2026-01-31"),
    ])
    with pytest.raises(VDateMismatchError):
        load_snapshot(p)


# ─── BUG-8: assert_vdate_matches ──────────────────────────────────────────

def test_assert_vdate_matches_ok():
    assert_vdate_matches(pd.Timestamp("2026-01-31"), "20260131")


def test_assert_vdate_matches_mismatch_raises():
    with pytest.raises(VDateMismatchError):
        assert_vdate_matches(pd.Timestamp("2026-01-31"), "20251231")


def test_loader_does_not_parse_filename_for_date():
    """Structural lock: the loader source contains no regex that extracts
    a date from a filename. BUG-8 prevention."""
    import re
    src = Path(loader.__file__).read_text(encoding="utf-8")
    # Disallow patterns that pull digits out of file names.
    forbidden = [
        r"PORT_\(\\d",       # legacy `re.match(r'PORT_(\d{2})...')`
        r"\.name.*\\d\{",    # `path.name` followed by digit-class regex
        r"\\d\{4\}.*\.xls",  # year regex against filename
    ]
    for pat in forbidden:
        assert not re.search(pat, src), (
            f"loader.py contains forbidden filename-date pattern: {pat}"
        )


# ─── find_snapshot_path ────────────────────────────────────────────────────

def test_find_snapshot_path_layout(tmp_path: Path):
    root = tmp_path / "Hiport"
    folder = root / "2025" / "Dec"
    folder.mkdir(parents=True)
    f = folder / "20251231.xlsm"
    f.write_bytes(b"")
    out = find_snapshot_path("20251231", archive_root=root)
    assert out == f


def test_find_snapshot_path_falls_back_to_xls(tmp_path: Path):
    root = tmp_path / "Hiport"
    folder = root / "2024" / "Jun"
    folder.mkdir(parents=True)
    f = folder / "20240630.xls"
    f.write_bytes(b"")
    out = find_snapshot_path("20240630", archive_root=root)
    assert out == f


def test_find_snapshot_path_returns_xlsm_default_when_missing(tmp_path: Path):
    root = tmp_path / "Hiport"
    out = find_snapshot_path("20260229", archive_root=root)
    # Note: 2026-02-29 is invalid but find_snapshot_path doesn't validate calendar.
    # Just check the layout.
    assert out == root / "2026" / "Feb" / "20260229.xlsm"


def test_find_snapshot_path_rejects_bad_input(tmp_path: Path):
    with pytest.raises(ValueError):
        find_snapshot_path("2026", archive_root=tmp_path)
    with pytest.raises(ValueError):
        find_snapshot_path("20261331", archive_root=tmp_path)  # month 13


# ─── load_bottler — happy & error paths via monkey-patch ──────────────────

def _trades_df_legacy(rows):
    """Build a Bottler-style DataFrame in **legacy** column names so the
    loader's renames are exercised. Each row is auto-assigned a unique ``ID``
    (raw bottler PK) which the loader renames to ``TRADE_ID``."""
    cols = ["PCODE_ORIG", "ISIN", "SHORT_NAME", "CCY", "T", "CDATE", "UNITS",
            "GROSSPRICE", "BUY NET", "SELL NET", "INCOME",
            "TRADEGROSS_USD", "DELETED"]
    if not rows:
        df = pd.DataFrame(columns=cols)
        df["ID"] = pd.Series(dtype="int64")
        return df
    df = pd.DataFrame(rows, columns=cols)
    df.insert(0, "ID", range(1, len(df) + 1))
    return df


def _patch_read_excel(monkeypatch, df: pd.DataFrame):
    def fake(path, *, sheet_name=None, engine=None, skiprows=0, **_):
        return df.copy()
    monkeypatch.setattr(loader.pd, "read_excel", fake)


def test_load_bottler_renames_and_drops_tradegross_usd(tmp_path: Path, monkeypatch):
    df = _trades_df_legacy([
        ("OEFOF", "kr7005930003 ", "SAMSUNG", "krw", "P", "2026-01-10",
         100, 70_000, 7_000_000, 0, 0, 99_999, ""),
    ])
    _patch_read_excel(monkeypatch, df)
    p = tmp_path / "bottler.xlsb"
    p.write_bytes(b"")  # need file to exist for path check
    out = load_bottler(p, fund_pcodes=("OEFOF", "OEFOGSSC"))

    assert "TRADEGROSS_USD" not in out.columns  # BUG-4 structural drop
    assert "GROSSPRICE_LOCAL" in out.columns
    assert "BUY_NET_LOCAL" in out.columns
    assert "SELL_NET_LOCAL" in out.columns
    assert "INCOME_LOCAL" in out.columns
    assert "SNAME" in out.columns  # SHORT_NAME → SNAME
    assert out["ISIN"].iloc[0] == "KR7005930003"
    assert out["CCY"].iloc[0] == "KRW"
    assert pd.api.types.is_datetime64_any_dtype(out["CDATE"])


def test_load_bottler_preserves_derivative_isin_suffix(tmp_path: Path, monkeypatch):
    df = _trades_df_legacy([
        ("OEFOF", "kyg8879r1048 : fut", "TIME INTERCONNECT", "eur", "P", "2026-01-10",
         1, 1, 1, 0, 0, 0, ""),
    ])
    _patch_read_excel(monkeypatch, df)
    p = tmp_path / "bottler.xlsb"
    p.write_bytes(b"")
    out = load_bottler(p, fund_pcodes=("OEFOF",))

    assert out["ISIN"].iloc[0] == "KYG8879R1048:FUT"


def test_load_bottler_filters_pcode(tmp_path: Path, monkeypatch):
    df = _trades_df_legacy([
        ("OEFOF",   "X", "S", "EUR", "P", "2026-01-10", 1, 1, 1, 0, 0, 0, ""),
        ("OTHER",   "Y", "S", "EUR", "P", "2026-01-10", 1, 1, 1, 0, 0, 0, ""),
        ("OEFOHFSC", "Z", "S", "EUR", "P", "2026-01-10", 1, 1, 1, 0, 0, 0, ""),
    ])
    _patch_read_excel(monkeypatch, df)
    p = tmp_path / "bottler.xlsb"
    p.write_bytes(b"")
    out = load_bottler(p, fund_pcodes=("OEFOF", "OEFOHFSC"))
    assert set(out["PCODE_ORIG"]) == {"OEFOF", "OEFOHFSC"}
    assert "OTHER" not in set(out["PCODE_ORIG"])


def test_load_bottler_drops_deleted_rows(tmp_path: Path, monkeypatch):
    df = _trades_df_legacy([
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-10", 1, 1, 1, 0, 0, 0, ""),
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-11", 2, 2, 2, 0, 0, 0, "Y"),
    ])
    _patch_read_excel(monkeypatch, df)
    p = tmp_path / "bottler.xlsb"
    p.write_bytes(b"")
    out = load_bottler(p, fund_pcodes=("OEFOF",))
    assert len(out) == 1
    assert float(out["UNITS"].iloc[0]) == 1.0


def test_load_bottler_empty_after_filter_raises(tmp_path: Path, monkeypatch):
    df = _trades_df_legacy([
        ("OTHER", "X", "S", "EUR", "P", "2026-01-10", 1, 1, 1, 0, 0, 0, ""),
    ])
    _patch_read_excel(monkeypatch, df)
    p = tmp_path / "bottler.xlsb"
    p.write_bytes(b"")
    with pytest.raises(EmptyDataError):
        load_bottler(p, fund_pcodes=("OEFOF",))


def test_load_bottler_missing_file(tmp_path: Path):
    with pytest.raises(DataLoadError):
        load_bottler(tmp_path / "nope.xlsb", fund_pcodes=("OEFOF",))


def test_load_bottler_empty_pcodes_rejected(tmp_path: Path, monkeypatch):
    df = _trades_df_legacy([
        ("OEFOF", "X", "S", "EUR", "P", "2026-01-10", 1, 1, 1, 0, 0, 0, ""),
    ])
    _patch_read_excel(monkeypatch, df)
    p = tmp_path / "bottler.xlsb"
    p.write_bytes(b"")
    with pytest.raises(ValueError):
        load_bottler(p, fund_pcodes=())


def test_load_bottler_missing_pcode_column_raises(tmp_path: Path, monkeypatch):
    df = pd.DataFrame({"ISIN": ["X"], "T": ["P"]})  # no PCODE_ORIG
    _patch_read_excel(monkeypatch, df)
    p = tmp_path / "bottler.xlsb"
    p.write_bytes(b"")
    with pytest.raises(SchemaError):
        load_bottler(p, fund_pcodes=("OEFOF",))
