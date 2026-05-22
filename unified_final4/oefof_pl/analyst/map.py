"""ISIN/SNAME → analyst-code mapping.

Resolution chain (highest priority first):

1. **SNAME override** — entries with a non-blank ``SNAME`` in
   ``isin_analyst_map.csv`` (or supplied via ``overrides=``).
   Used for securities whose ISIN coverage is unreliable across snapshots
   (CFDs, swaps) or where the issuer name is more stable than the ISIN.
2. **CSV map** — ``isin_analyst_map.csv`` keyed by ISIN.
3. **Default** — ``UNASSIGNED``. Newly-seen ISINs are appended to the CSV
   with this code so the user can fill them in for the next run.

Design constraints:

* The CSV is the single durable source of truth for both kinds of override.
  A row with both ISIN and SNAME blank is ignored; a row with only SNAME
  populated is a pure SNAME override.
* CSV writes are **atomic** (write to ``.tmp`` then ``os.replace``). A
  crash mid-write cannot leave a half-written file.
* The "auto-append unknown ISINs" behaviour is opt-in via
  ``record_unknown=True`` on ``resolve``. Compute code never auto-appends
  by accident; only the pipeline's analyst step does.
"""

from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Mapping

import pandas as pd

from ..exceptions import AnalystMapError


CSV_COLUMNS = ("ISIN", "ANALYST")
CSV_COLUMNS_FULL = ("ISIN", "SNAME", "ANALYST")
UNASSIGNED = "UNASSIGNED"


@dataclass
class AnalystMap:
    """Mutable in-memory mapping that persists to a CSV.

    Use ``AnalystMap.load(path, overrides=...)`` to construct. Use
    ``resolve(isin, sname)`` for lookups, optionally with
    ``record_unknown=True`` to append missing ISINs to the file.
    """

    path: Path
    by_isin: dict[str, str]
    sname_overrides: dict[str, str]
    # SNAME entries that came from the CSV (and so should be persisted on
    # save). Entries supplied only via the ``overrides=`` argument are
    # treated as runtime supplements and are NOT written back, to keep the
    # CSV's contents stable across runs that pass different ``overrides``.
    _persisted_snames: dict[str, str] = field(default_factory=dict, repr=False)
    _dirty: bool = field(default=False, init=False, repr=False)

    # ─── construction ──────────────────────────────────────────────────────

    @classmethod
    def load(
        cls,
        path: Path,
        *,
        overrides: Mapping[str, str] | None = None,
    ) -> "AnalystMap":
        """Load the CSV at ``path``. Creates an empty file if missing.

        Args:
            path: Location of the CSV. Created if absent.
            overrides: Optional SNAME → analyst-code map (typically
                ``Config.person_overrides``). Keys are normalised
                (uppercased, stripped) on load.
        """
        path = Path(path)
        persisted_snames: dict[str, str] = {}
        if path.exists():
            try:
                df = pd.read_csv(path, dtype=str)
            except Exception as exc:
                raise AnalystMapError(
                    f"Could not read analyst map at {path}: {exc}"
                ) from exc
            missing = [c for c in CSV_COLUMNS if c not in df.columns]
            if missing:
                raise AnalystMapError(
                    f"Analyst map {path} missing columns {missing}; "
                    f"got {list(df.columns)}"
                )
            has_sname_col = "SNAME" in df.columns
            by_isin: dict[str, str] = {}
            for i in range(len(df)):
                isin_raw = df["ISIN"].iloc[i]
                sname_raw = df["SNAME"].iloc[i] if has_sname_col else ""
                analyst_raw = df["ANALYST"].iloc[i]
                analyst = (str(analyst_raw).strip() or UNASSIGNED) if pd.notna(analyst_raw) else UNASSIGNED
                isin_key = _norm(isin_raw)
                sname_key = _norm(sname_raw)
                if sname_key:
                    persisted_snames[sname_key] = analyst
                if isin_key:
                    by_isin[isin_key] = analyst
        else:
            by_isin = {}

        # Merge: CSV-loaded SNAMEs take precedence; ``overrides=`` arg fills
        # in any SNAMEs not present in the CSV (typically empty post-migration).
        merged_snames: dict[str, str] = {}
        for k, v in (overrides or {}).items():
            nk = _norm(k)
            if nk and v:
                merged_snames[nk] = str(v).strip()
        merged_snames.update(persisted_snames)

        return cls(
            path=path,
            by_isin=by_isin,
            sname_overrides=merged_snames,
            _persisted_snames=persisted_snames,
        )

    # ─── lookup ────────────────────────────────────────────────────────────

    def resolve(
        self,
        isin: object,
        sname: object = "",
        *,
        record_unknown: bool = False,
    ) -> str:
        """Return the analyst code for ``isin``.

        Resolution order:

        1. SNAME override (case/whitespace-insensitive exact match).
        2. CSV map by ISIN.
        3. ``UNASSIGNED``.

        If ``record_unknown=True`` and the ISIN was not in the CSV, the new
        entry is added to the in-memory dict (and persisted on the next
        ``save()``). SNAME-override hits are NOT recorded, because the
        override always wins regardless of CSV contents.
        """
        sname_key = _norm(sname)
        if sname_key and sname_key in self.sname_overrides:
            return self.sname_overrides[sname_key]

        isin_key = _norm(isin)
        if not isin_key:
            return UNASSIGNED

        if isin_key in self.by_isin:
            return self.by_isin[isin_key]

        if record_unknown:
            self.by_isin[isin_key] = UNASSIGNED
            self._dirty = True
        return UNASSIGNED

    # ─── persistence ───────────────────────────────────────────────────────

    @property
    def dirty(self) -> bool:
        """True if there are unsaved changes."""
        return self._dirty

    def save(self) -> None:
        """Atomically write the in-memory map to ``self.path``.

        Writes to a sibling ``.tmp`` file then ``os.replace``s it over the
        target. A crash mid-write cannot corrupt the existing map.
        """
        if not self._dirty and self.path.exists():
            return

        self.path.parent.mkdir(parents=True, exist_ok=True)

        # Sort for stable diffs in version control. SNAME-only rows are
        # written first (alphabetical by SNAME), then ISIN rows alphabetical
        # by ISIN. Two-column files (no SNAME entries) keep the legacy
        # ``ISIN,ANALYST`` schema for backward compatibility with consumers.
        sname_rows = sorted(self._persisted_snames.items())
        isin_rows = sorted(self.by_isin.items())
        if sname_rows:
            rows = (
                [("", sname, analyst) for sname, analyst in sname_rows]
                + [(isin, "", analyst) for isin, analyst in isin_rows]
            )
            df = pd.DataFrame(rows, columns=list(CSV_COLUMNS_FULL))
        else:
            df = pd.DataFrame(isin_rows, columns=list(CSV_COLUMNS))

        # Atomic write via NamedTemporaryFile in the target directory so
        # os.replace is on the same filesystem.
        fd, tmp_path = tempfile.mkstemp(
            prefix=".analyst_map_", suffix=".csv.tmp",
            dir=str(self.path.parent),
        )
        try:
            with os.fdopen(fd, "w", newline="", encoding="utf-8") as fh:
                df.to_csv(fh, index=False)
            os.replace(tmp_path, self.path)
        except Exception as exc:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            raise AnalystMapError(f"Could not write analyst map: {exc}") from exc

        self._dirty = False


# ─── helpers ────────────────────────────────────────────────────────────────

def _norm(x: object) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and x != x:  # NaN
        return ""
    return str(x).strip().upper()
