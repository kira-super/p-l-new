"""Comprehensive scan of accessible DBs for objects useful to OEFOF P&L.

Focus areas:
  1. Anything with SCODE / ISIN + units / shares / price / date  (potential CA log)
  2. tSecurity columns we don't already use
  3. tTRANS columns we don't already use (e.g. INCOME / FX / COMMISSION buckets)
  4. Anything mentioning corporate actions, bonus, split, rights, reorg, merger
  5. Anything with WAC, COST_BASIS, PERFORMANCE we could cross-check against
  6. PCODE-keyed tables (other than what we use) that could give portfolio-level
     reconciliation values (NAV, TWR, etc.)
"""
import pyodbc
from collections import defaultdict
from oefof_pl.config import load_default_config

cn = pyodbc.connect(load_default_config().sql_conn_str, timeout=30)
cur = cn.cursor()

DBS = ["CCL", "GICS", "MKT_LIVE", "NAV", "SR", "SRSQL", "SunSystemsDomain"]

# ---------- 1. Hunt by interesting column names ----------
KEYWORDS = {
    "CA event log":     ["CA_TYPE", "CA_EVENT", "CACTION", "EX_DATE",
                         "RECORD_DATE", "PAY_DATE", "ANNOUNCE", "BONUS_RATIO",
                         "SPLIT_RATIO", "TERMS", "ENTITLEMENT", "ELECTION"],
    "CA inline trades": ["CADJ", "ADJ_FACTOR", "REORG", "TENDER", "SPINOFF",
                         "MERGER", "STKDIV", "STK_DIV"],
    "Cost / WAC ref":   ["WAC", "AVG_COST", "AVERAGE_COST", "COST_BASIS",
                         "ACQ_COST", "BOOK_COST"],
    "FX / market":      ["XRATE", "EX_RATE", "SPOT_RATE", "FX_RATE", "RATE_DATE"],
    "Brokerage":        ["BROKERAGE", "COMMISSION", "BCOMM", "STAMP_DUTY",
                         "EXCHANGE_FEE", "REG_FEE", "TAX_FEE"],
    "Performance":      ["TWR", "NAV_PER_SHARE", "RETURN", "PERF"],
}

print("=" * 70)
print(" 1. COLUMN-NAME HUNT (anything we might be missing)")
print("=" * 70)
for db in DBS:
    found_db = defaultdict(list)
    for category, kws in KEYWORDS.items():
        for kw in kws:
            try:
                cur.execute(
                    f"SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, DATA_TYPE "
                    f"FROM {db}.INFORMATION_SCHEMA.COLUMNS "
                    f"WHERE COLUMN_NAME LIKE ?", f"%{kw}%")
                for r in cur.fetchall():
                    found_db[category].append((r[0], r[1], r[2], r[3]))
            except Exception:
                pass
    if found_db:
        print(f"\n[{db}]")
        for cat, rows in found_db.items():
            uniq = sorted(set(rows))
            print(f"  {cat}:")
            for sch, tbl, col, dt in uniq:
                print(f"    {sch}.{tbl}.{col} ({dt})")

# ---------- 2. tSecurity full column list (we use few of 63) ----------
print("\n" + "=" * 70)
print(" 2. tSecurity FULL column list (currently we read SCODE/ISIN/SHORT_NAME/CCY)")
print("=" * 70)
cur.execute("""
    SELECT COLUMN_NAME, DATA_TYPE
    FROM CCL.INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME='tSecurity' ORDER BY ORDINAL_POSITION
""")
for r in cur.fetchall():
    print(f"  {r[0]:<32} {r[1]}")

# ---------- 3. tTRANS full column list ----------
print("\n" + "=" * 70)
print(" 3. tTRANS FULL column list")
print("=" * 70)
cur.execute("""
    SELECT COLUMN_NAME, DATA_TYPE
    FROM CCL.INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME='tTRANS' ORDER BY ORDINAL_POSITION
""")
for r in cur.fetchall():
    print(f"  {r[0]:<32} {r[1]}")

# ---------- 4. All distinct BCODE values for OEFOF (to reveal CA-style codes) ----------
print("\n" + "=" * 70)
print(" 4. tTRANS distinct (T, BCODE) combinations for OEFOF, top 50 by count")
print("=" * 70)
cur.execute("""
    SELECT TOP 50 T, BCODE, COUNT(*) AS n
    FROM CCL.dbo.tTRANS
    WHERE UPPER(PCODE)='OEFOF' AND (D IS NULL OR D <> 'Y')
    GROUP BY T, BCODE
    ORDER BY n DESC
""")
for r in cur.fetchall():
    print(f"  T='{r[0]}'  BCODE='{r[1]}'  count={r[2]}")

# ---------- 5. All ZZZZ (CA inline) trades for OEFOF this period ----------
print("\n" + "=" * 70)
print(" 5. All BCODE='ZZZZ' trades for OEFOF since 2025-12-31")
print("=" * 70)
cur.execute("""
    SELECT s.ISIN, s.SHORT_NAME, t.T, t.CDATE, t.UNITS, t.NUPRICEG,
           t.NSETTLE_AMOUNT
    FROM CCL.dbo.tTRANS t
    LEFT JOIN CCL.dbo.tSecurity s ON s.SCODE = t.SCODE
    WHERE UPPER(t.PCODE)='OEFOF'
      AND (t.D IS NULL OR t.D <> 'Y')
      AND UPPER(t.BCODE)='ZZZZ'
      AND t.CDATE >= '2025-12-31'
    ORDER BY t.CDATE, s.ISIN
""")
n = 0
for r in cur.fetchall():
    print(f"  {r[3]:%Y-%m-%d} {r[2]} {r[0]:<14} {r[1]:<25} {float(r[4]):>15,.0f} px={float(r[5]):>10.4f} settle={float(r[6]):>15,.0f}")
    n += 1
print(f"  ({n} ZZZZ rows in period)")

# ---------- 6. tH_VAL (history) - might give intraperiod snapshots ----------
print("\n" + "=" * 70)
print(" 6. tH_VAL columns (historical snapshots)")
print("=" * 70)
cur.execute("""
    SELECT COLUMN_NAME, DATA_TYPE
    FROM CCL.INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME='tH_VAL' ORDER BY ORDINAL_POSITION
""")
for r in cur.fetchall():
    print(f"  {r[0]:<32} {r[1]}")

# ---------- 7. NAV.tNAV (we have it but don't use it - daily NAV?) ----------
print("\n" + "=" * 70)
print(" 7. NAV.dbo.tNAV columns + sample rows for OEFOF")
print("=" * 70)
cur.execute("""
    SELECT COLUMN_NAME, DATA_TYPE
    FROM NAV.INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME='tNAV' ORDER BY ORDINAL_POSITION
""")
for r in cur.fetchall():
    print(f"  {r[0]:<32} {r[1]}")
try:
    cur.execute("SELECT TOP 5 * FROM NAV.dbo.tNAV WHERE PCode LIKE '%OEFOF%' ORDER BY vDate DESC")
    cols = [c[0] for c in cur.description]
    print("  sample:")
    print("   ", " | ".join(cols))
    for r in cur.fetchall():
        print("   ", " | ".join(str(v)[:25] for v in r))
except Exception as e:
    print(f"  sample ERR: {e}")

# ---------- 8. CCL.tCGT_Work and tH_CGT_DETAILS - capital gains tax details (book cost!) ----
print("\n" + "=" * 70)
print(" 8. CCL.tH_CGT_DETAILS columns (might give us BOOK COST per ISIN)")
print("=" * 70)
cur.execute("""
    SELECT COLUMN_NAME, DATA_TYPE
    FROM CCL.INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME='tH_CGT_DETAILS' ORDER BY ORDINAL_POSITION
""")
for r in cur.fetchall():
    print(f"  {r[0]:<32} {r[1]}")

# ---------- 9. tStockAvailability - shorting borrow rate? ----------
print("\n" + "=" * 70)
print(" 9. CCL.tStockAvailability columns")
print("=" * 70)
cur.execute("""
    SELECT COLUMN_NAME, DATA_TYPE
    FROM CCL.INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_NAME='tStockAvailability' ORDER BY ORDINAL_POSITION
""")
for r in cur.fetchall():
    print(f"  {r[0]:<32} {r[1]}")

# ---------- 10. List ALL views in CCL for OEFOF/RPT (might find a ready P&L view) ----------
print("\n" + "=" * 70)
print(" 10. CCL views matching '%RPT%' or '%PL%' or '%PERF%' or '%CGT%'")
print("=" * 70)
cur.execute("""
    SELECT TABLE_SCHEMA, TABLE_NAME
    FROM CCL.INFORMATION_SCHEMA.VIEWS
    WHERE TABLE_NAME LIKE '%RPT%' OR TABLE_NAME LIKE '%PL[_]%'
       OR TABLE_NAME LIKE '%PERF%' OR TABLE_NAME LIKE '%CGT%'
       OR TABLE_NAME LIKE '%PNL%' OR TABLE_NAME LIKE '%REALIS%'
    ORDER BY TABLE_NAME
""")
for r in cur.fetchall():
    print(f"  {r[0]}.{r[1]}")
