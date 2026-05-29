"""Tests for analyst.map — ISIN → analyst-code resolution and persistence."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from oefof_pl.analyst.map import UNASSIGNED, AnalystMap
from oefof_pl.exceptions import AnalystMapError


# ─── load ────────────────────────────────────────────────────────────────────

def test_load_missing_file_returns_empty_map(tmp_path: Path):
    m = AnalystMap.load(tmp_path / "missing.csv")
    assert m.by_isin == {}
    assert not m.path.exists()  # not created until save


def test_load_existing_csv(tmp_path: Path):
    p = tmp_path / "map.csv"
    pd.DataFrame({"ISIN": ["KR7005930003", "DE000HLAG475"],
                  "ANALYST": ["HK", "SB"]}).to_csv(p, index=False)
    m = AnalystMap.load(p)
    assert m.by_isin == {"KR7005930003": "HK", "DE000HLAG475": "SB"}


def test_load_normalises_isin(tmp_path: Path):
    p = tmp_path / "map.csv"
    pd.DataFrame({"ISIN": [" kr7005930003 "], "ANALYST": ["HK"]}).to_csv(p, index=False)
    m = AnalystMap.load(p)
    assert m.resolve("KR7005930003") == "HK"


def test_load_bad_schema_raises(tmp_path: Path):
    p = tmp_path / "bad.csv"
    pd.DataFrame({"isin": ["x"], "code": ["y"]}).to_csv(p, index=False)
    with pytest.raises(AnalystMapError):
        AnalystMap.load(p)


# ─── resolve ─────────────────────────────────────────────────────────────────

def test_resolve_csv_hit(tmp_path: Path):
    p = tmp_path / "map.csv"
    pd.DataFrame({"ISIN": ["KR7005930003"], "ANALYST": ["HK"]}).to_csv(p, index=False)
    m = AnalystMap.load(p)
    assert m.resolve("KR7005930003") == "HK"


def test_resolve_unknown_returns_unassigned_without_recording(tmp_path: Path):
    m = AnalystMap.load(tmp_path / "map.csv")
    assert m.resolve("NEW-ISIN") == UNASSIGNED
    assert "NEW-ISIN" not in m.by_isin
    assert not m.dirty


def test_resolve_unknown_records_when_requested(tmp_path: Path):
    m = AnalystMap.load(tmp_path / "map.csv")
    assert m.resolve("NEW-ISIN", record_unknown=True) == UNASSIGNED
    assert m.by_isin["NEW-ISIN"] == UNASSIGNED
    assert m.dirty


def test_repo_map_contains_known_isin_mapping():
    repo_map = Path(__file__).resolve().parents[1] / "isin_analyst_map.csv"
    m = AnalystMap.load(repo_map)
    assert m.resolve("KYG070341048") == "HK"


def test_sname_override_takes_priority_over_csv(tmp_path: Path):
    p = tmp_path / "map.csv"
    pd.DataFrame({"ISIN": ["X"], "ANALYST": ["IS"]}).to_csv(p, index=False)
    m = AnalystMap.load(p, overrides={"VICTORY GIANT TECH": "KX"})
    assert m.resolve("X", sname="VICTORY GIANT TECH") == "KX"
    # Without the SNAME, falls through to CSV
    assert m.resolve("X") == "IS"


def test_sname_override_case_insensitive(tmp_path: Path):
    m = AnalystMap.load(
        tmp_path / "map.csv",
        overrides={"VICTORY GIANT TECH": "KX"},
    )
    assert m.resolve("UNKNOWN", sname=" victory giant tech ") == "KX"


def test_sname_override_does_not_record(tmp_path: Path):
    """SNAME-override hits must NOT pollute the CSV."""
    m = AnalystMap.load(
        tmp_path / "map.csv",
        overrides={"VICTORY GIANT TECH": "KX"},
    )
    m.resolve("UNKNOWN-ISIN", sname="VICTORY GIANT TECH", record_unknown=True)
    assert "UNKNOWN-ISIN" not in m.by_isin
    assert not m.dirty


def test_resolve_empty_isin_returns_unassigned(tmp_path: Path):
    m = AnalystMap.load(tmp_path / "map.csv")
    assert m.resolve("") == UNASSIGNED
    assert m.resolve(None) == UNASSIGNED
    assert m.resolve(float("nan")) == UNASSIGNED


# ─── save (atomicity) ────────────────────────────────────────────────────────

def test_save_writes_sorted_csv(tmp_path: Path):
    m = AnalystMap.load(tmp_path / "map.csv")
    m.resolve("ZZZ", record_unknown=True)
    m.resolve("AAA", record_unknown=True)
    m.by_isin["AAA"] = "HK"
    m._dirty = True
    m.save()

    df = pd.read_csv(tmp_path / "map.csv")
    assert list(df["ISIN"]) == ["AAA", "ZZZ"]
    assert list(df["ANALYST"]) == ["HK", "UNASSIGNED"]


def test_save_is_atomic_no_tmp_left_behind(tmp_path: Path):
    m = AnalystMap.load(tmp_path / "map.csv")
    m.resolve("AAA", record_unknown=True)
    m.save()
    leftovers = list(tmp_path.glob(".analyst_map_*.tmp"))
    assert leftovers == []


def test_save_idempotent_when_clean(tmp_path: Path):
    p = tmp_path / "map.csv"
    pd.DataFrame({"ISIN": ["X"], "ANALYST": ["IS"]}).to_csv(p, index=False)
    m = AnalystMap.load(p)
    mtime_before = p.stat().st_mtime_ns
    m.save()  # no-op, file unchanged
    assert p.stat().st_mtime_ns == mtime_before


def test_save_overwrites_existing(tmp_path: Path):
    p = tmp_path / "map.csv"
    pd.DataFrame({"ISIN": ["X"], "ANALYST": ["IS"]}).to_csv(p, index=False)
    m = AnalystMap.load(p)
    m.by_isin["X"] = "HK"
    m._dirty = True
    m.save()
    df = pd.read_csv(p)
    assert df.loc[df["ISIN"] == "X", "ANALYST"].iloc[0] == "HK"


def test_round_trip(tmp_path: Path):
    p = tmp_path / "map.csv"
    m1 = AnalystMap.load(p)
    m1.by_isin.update({"A": "HK", "B": "IS", "C": UNASSIGNED})
    m1._dirty = True
    m1.save()

    m2 = AnalystMap.load(p)
    assert m2.by_isin == m1.by_isin
