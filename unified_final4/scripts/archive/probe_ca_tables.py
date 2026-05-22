"""Probe v2: search both table names AND column names across accessible DBs."""
import pyodbc
from oefof_pl.config import load_default_config

cfg = load_default_config()
cn = pyodbc.connect(cfg.sql_conn_str, timeout=30)
cur = cn.cursor()

ACCESSIBLE = ["CCL", "GICS", "MKT_LIVE", "NAV", "SR", "SRSQL", "SunSystemsDomain"]
NAME_PATS = [
    "%CORP%", "%CORPORATE%", "%CACTION%", "%CA[_]%", "%[_]CA",
    "%DIVIDEND%", "%SPLIT%", "%BONUS%", "%RIGHTS%", "%MERGER%",
    "%ENTITLEMENT%", "%STKDIV%", "%CASHDIV%", "%REORG%",
    "%EVENT%", "%NOTICE%", "%ANNOUNC%", "%TENDER%", "%SPINOFF%",
]

def section(title: str) -> None:
    print(f"\n{'=' * 60}\n {title}\n{'=' * 60}")

# 1) Table/view names
section("TABLES / VIEWS with CA-related names")
for db in ACCESSIBLE:
    try:
        seen: set[tuple[str, str, str]] = set()
        for pat in NAME_PATS:
            sql = (
                f"SELECT TABLE_SCHEMA, TABLE_NAME, TABLE_TYPE "
                f"FROM [{db}].INFORMATION_SCHEMA.TABLES "
                f"WHERE TABLE_NAME LIKE ?"
            )
            cur.execute(sql, pat)
            for r in cur.fetchall():
                seen.add((r[0], r[1], r[2]))
        for sch, name, ttype in sorted(seen):
            print(f"  [{db}].[{sch}].[{name}]  ({ttype})")
    except Exception as e:
        print(f"  [{db}] ERR: {e.__class__.__name__}")

# 2) Column names
section("COLUMNS with CA-related names")
COL_PATS = [
    "%CORP[_]ACTION%", "%CA[_]TYPE%", "%EX[_]DATE%",
    "%RECORD[_]DATE%", "%ANNOUNCE%", "%PAY[_]DATE%",
    "%DIVIDEND%", "%SPLIT[_]RATIO%", "%BONUS[_]RATIO%",
]
for db in ACCESSIBLE:
    try:
        seen2: set[tuple[str, str, str]] = set()
        for pat in COL_PATS:
            sql = (
                f"SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME "
                f"FROM [{db}].INFORMATION_SCHEMA.COLUMNS "
                f"WHERE COLUMN_NAME LIKE ?"
            )
            cur.execute(sql, pat)
            for r in cur.fetchall():
                seen2.add((r[0], r[1], r[2]))
        if seen2:
            print(f"\n[{db}]")
            by_tbl: dict[tuple[str, str], list[str]] = {}
            for sch, tbl, col in seen2:
                by_tbl.setdefault((sch, tbl), []).append(col)
            for (sch, tbl), cols in sorted(by_tbl.items()):
                print(f"  {sch}.{tbl} -> {sorted(set(cols))}")
    except Exception as e:
        print(f"\n[{db}] ERR: {e.__class__.__name__}")
