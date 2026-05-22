import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from oefof_pl.config import load_default_config, TARGET_PNAME
from oefof_pl.output.valuation import extract_valuation_a_via_sql

rows = extract_valuation_a_via_sql(load_default_config().sql_conn_str)
print(f"Rows total (incl headers): {len(rows)}")
print(f"Header row 7 (index 6): {rows[6] if len(rows) > 6 else '(missing)'}")
data = rows[7:]
print(f"Data rows: {len(data)}")
for r in data[:5]:
    print(" ", r)
