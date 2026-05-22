"""Shared dataclasses for corporate-action modules."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd


@dataclass(frozen=True)
class CaOverride:
    isin: str
    units: float       # signed
    reason: str
    price: float = 0.0


@dataclass(frozen=True)
class CaSuggestion:
    """One auto-detected unit drift, presented to the user for sign-off."""
    isin: str
    sname: str
    units: float                     # signed (= VAL-01 Diff)
    start_units: float
    expected_end_units: float
    actual_end_units: float
    recurrence_count: int            # times same-ISIN/same-direction appeared in history
    last_seen_period: str            # "" if never; YYYY-MM-DD otherwise
    suggested_reason: str            # heuristic, includes recurrence hint


@dataclass(frozen=True)
class HistoryEntry:
    isin: str
    units: float
    price: float
    reason: str
    period_end: str                  # YYYY-MM-DD


@dataclass(frozen=True)
class BonusPriceOverride:
    isin: str
    cdate: pd.Timestamp
    price: float
    reason: str
