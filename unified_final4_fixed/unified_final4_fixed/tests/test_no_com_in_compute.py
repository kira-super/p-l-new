"""AST lint: ``win32com`` may only be imported by the three designated
COM-touching modules. The compute core, loaders, validators, workbook
builder, and pipeline must remain COM-free so the package can be
installed and tested on non-Windows environments and so unit tests never
spawn Excel or Outlook.

Allowed importers:
    * ``oefof_pl/output/valuation.py``    — ValuationA workbook
    * ``oefof_pl/email_sender/outlook.py`` — Outlook send

Any other ``import win32com`` (or ``from win32com...``) anywhere under
``oefof_pl/`` is a structural regression and fails this test.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest


PKG_ROOT = Path(__file__).resolve().parents[1] / "oefof_pl"

ALLOWED = {
    PKG_ROOT / "output" / "valuation.py",
    PKG_ROOT / "email_sender" / "outlook.py",
}


def _python_files() -> list[Path]:
    return [p for p in PKG_ROOT.rglob("*.py") if p.is_file()]


def _imports_win32com(path: Path) -> bool:
    """Return True if *path* contains any module-level or function-level
    ``import win32com`` or ``from win32com...`` statement."""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "win32com" or alias.name.startswith("win32com."):
                    return True
        elif isinstance(node, ast.ImportFrom):
            mod = node.module or ""
            if mod == "win32com" or mod.startswith("win32com."):
                return True
    return False


def test_pkg_root_exists():
    assert PKG_ROOT.is_dir(), f"missing package root: {PKG_ROOT}"


def test_allowed_modules_actually_exist():
    for p in ALLOWED:
        assert p.is_file(), f"allowed COM module not found: {p}"


def test_no_unexpected_win32com_imports():
    """Fail if any non-allowed module imports win32com."""
    offenders: list[Path] = []
    for path in _python_files():
        if path in ALLOWED:
            continue
        if _imports_win32com(path):
            offenders.append(path.relative_to(PKG_ROOT))

    assert not offenders, (
        "win32com may only be imported by the three designated COM modules. "
        f"Offenders: {[str(p) for p in offenders]}"
    )


@pytest.mark.parametrize("module", sorted(ALLOWED))
def test_allowed_module_imports_win32com_lazily(module: Path):
    """The three allowed modules MUST import win32com inside a function,
    not at module top level. This keeps the package importable when
    pywin32 is not installed (e.g. on CI / Linux dev boxes).
    """
    tree = ast.parse(module.read_text(encoding="utf-8"))
    top_level = [n for n in tree.body if isinstance(n, (ast.Import, ast.ImportFrom))]
    for node in top_level:
        if isinstance(node, ast.Import):
            for alias in node.names:
                assert not alias.name.startswith("win32com"), (
                    f"{module.name}: top-level 'import win32com' forbidden — "
                    "must be function-local."
                )
        elif isinstance(node, ast.ImportFrom):
            mod = node.module or ""
            assert not mod.startswith("win32com"), (
                f"{module.name}: top-level 'from win32com...' forbidden — "
                "must be function-local."
            )
