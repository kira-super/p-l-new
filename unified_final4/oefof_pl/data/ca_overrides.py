"""Corporate-action unit overrides — re-export shim.

All implementation has been split into focused modules:
  - :mod:`.ca_types`   — shared dataclasses
  - :mod:`.ca_manual`  — load/apply/write manual (curated) overrides
  - :mod:`.ca_auto`    — derive auto overrides and build suggestions
  - :mod:`.ca_history` — append-only audit log

This module re-exports every public name for backwards compatibility.
"""

from __future__ import annotations

# noqa: F401 — all re-exports are intentional
from .ca_auto import (  # noqa: F401
    build_suggestions,
    derive_ca_auto,
    write_suggestions_csv,
)
from .ca_history import (  # noqa: F401
    _HISTORY_HEADER,
    append_history,
    deduplicate_history,
    load_history,
)
from .ca_manual import (  # noqa: F401
    apply_bonus_price_overrides,
    load_bonus_prices,
    load_ca_overrides,
    load_overrides_merged,
    overrides_to_trades,
    persist_auto_ca,
    write_ca_overrides,
)
from .ca_types import (  # noqa: F401
    BonusPriceOverride,
    CaOverride,
    CaSuggestion,
    HistoryEntry,
)

__all__ = [
    "CaOverride",
    "CaSuggestion",
    "HistoryEntry",
    "BonusPriceOverride",
    "load_ca_overrides",
    "load_overrides_merged",
    "derive_ca_auto",
    "overrides_to_trades",
    "load_history",
    "append_history",
    "deduplicate_history",
    "load_bonus_prices",
    "apply_bonus_price_overrides",
    "build_suggestions",
    "write_ca_overrides",
    "persist_auto_ca",
    "write_suggestions_csv",
]
