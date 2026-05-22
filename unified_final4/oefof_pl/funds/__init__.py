"""Fund-specific configuration modules.

Each module exposes a single ``load_config() -> Config`` function that
returns a fully-populated :class:`~oefof_pl.config.Config` for that fund.

Usage::

    from oefof_pl.funds.oefof import load_config
    cfg = load_config()

    from oefof_pl.funds.absa import load_config
    cfg = load_config()
"""
