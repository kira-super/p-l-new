"""Convenience script. Equivalent to ``python -m oefof_pl``."""
from __future__ import annotations

import sys

from oefof_pl.__main__ import main


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
