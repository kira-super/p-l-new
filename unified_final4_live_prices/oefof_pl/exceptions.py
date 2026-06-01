"""Exception hierarchy for the oefof_pl package.

All exceptions raised by this package inherit from `OefofError`. Every
exception is intentionally specific so callers can catch only what they want
to handle (e.g. the pipeline catches `EmptyDataError` to log a clear message
but lets `ComputeIntegrityError` propagate because it indicates a code bug).
"""

from __future__ import annotations


class OefofError(Exception):
    """Base class for all package-specific exceptions."""


# ── data layer ────────────────────────────────────────────────────────────────

class DataLoadError(OefofError):
    """Raised when a source file cannot be read (missing file, missing sheet,
    missing engine dependency such as pyxlsb, etc.)."""


class SchemaError(OefofError):
    """Raised when a DataFrame crossing a module boundary does not match the
    declared schema (missing required columns, wrong dtypes)."""


class EmptyDataError(OefofError):
    """Raised when a load or filter step produces zero rows where ≥1 was
    expected. Prevents the legacy silent-empty-DataFrame failure mode."""


class VDateMismatchError(OefofError):
    """Raised when the `VDATE` read from a snapshot does not match the date
    that was requested. Prevents BUG-8 (filename / VDATE divergence)."""


# ── compute layer ─────────────────────────────────────────────────────────────

class FXMissingError(OefofError):
    """Raised when a required FX rate (typically EUR/USD) is absent from the
    portfolio snapshot. The legacy code silently used a hard-coded fallback
    (0.868018); v2 aborts so a wrong report can never be emitted."""


class FXInvalidError(OefofError):
    """Raised when an FX rate is present but unusable (zero, negative, NaN)
    for a non-USD/EUR currency."""


class ComputeIntegrityError(OefofError):
    """Raised when an internal P&L identity check fails inside
    `compute.pl.compute_positions` (e.g. grand total ≠ sum of components by
    more than 1 EUR). Indicates a code bug, not a data issue."""


# ── analyst layer ─────────────────────────────────────────────────────────────

class AnalystMapError(OefofError):
    """Raised when the analyst map CSV is malformed or cannot be written
    atomically."""
