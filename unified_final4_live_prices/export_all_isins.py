import pyodbc
import pandas as pd

conn = pyodbc.connect(
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=fcedata01.fieracapital.ca;"
    "Database=ccl;"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)

# Last known price per ISIN across all history
df = pd.read_sql("""
    SELECT v.ISIN, v.SNAME, v.CCY, v.LST_PRICE, v.UNITS, v.VDATE
    FROM ccl.dbo.vw_RPT_VAL v
    INNER JOIN (
        SELECT ISIN, MAX(VDATE) AS MAX_VDATE
        FROM ccl.dbo.vw_RPT_VAL
        WHERE PCODE IN ('OEFOF','OEFOGSSC','OEFOGSSW')
        GROUP BY ISIN
    ) m ON v.ISIN = m.ISIN AND v.VDATE = m.MAX_VDATE
    WHERE v.PCODE IN ('OEFOF','OEFOGSSC','OEFOGSSW')
""", conn)

df.to_csv("hiport_all_isins.csv", index=False)
print(f"Saved {len(df)} rows to hiport_all_isins.csv")
