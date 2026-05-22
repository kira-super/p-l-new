"""Fund-specific configuration modules.

Each module exposes a single ``load_config() -> Config`` function that
returns a fully-populated :class:`~oefof_pl.config.Config` for that fund.

Usage::

    from oefof_pl.funds.oefof import load_config
    cfg = load_config()

    from oefof_pl.funds.absa import load_config
    cfg = load_config()

    # Or use the dispatcher:
    from oefof_pl.funds import load_config
    cfg = load_config("oefof")
    cfg = load_config("absa")
"""
from __future__ import annotations

from oefof_pl.config import Config

_REGISTRY: dict[str, str] = {
    "oefof": "oefof_pl.funds.oefof",
    "absa": "oefof_pl.funds.absa",
}


def load_config(fund: str) -> Config:
    """Return the production :class:`~oefof_pl.config.Config` for *fund*.

    Parameters
    ----------
    fund:
        Case-insensitive fund identifier, e.g. ``"oefof"`` or ``"absa"``.

    Raises
    ------
    ValueError
        If *fund* is not a known fund identifier.
    """
    key = fund.strip().lower()
    if key not in _REGISTRY:
        known = ", ".join(sorted(_REGISTRY))
        raise ValueError(f"Unknown fund {fund!r}. Known funds: {known}")

    import importlib
    mod = importlib.import_module(_REGISTRY[key])
    return mod.load_config()
