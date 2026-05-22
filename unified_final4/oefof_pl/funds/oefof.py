"""OEFOF fund configuration.

Call ``load_config()`` to get the production Config for the
OAKS Emerging and Frontier Fund.

All OEFOF-specific constants live in the parent ``config.py``.
This module exists so fund launchers can do a symmetric import::

    from oefof_pl.funds.oefof import load_config
    from oefof_pl.funds.absa  import load_config
"""
from __future__ import annotations

from oefof_pl.config import Config, load_default_config


def load_config() -> Config:
    """Return the OEFOF production Config."""
    return load_default_config()
