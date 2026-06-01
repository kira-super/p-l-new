import pyodbc
import pandas as pd

conn = pyodbc.connect(
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=fcedata01.fieracapital.ca;"
    "Database=ccl;"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)

df = pd.read_sql("""
    SELECT ISIN, SNAME, CCY, LST_PRICE, UNITS, VDATE
    FROM ccl.dbo.vw_RPT_VAL
    WHERE PCODE IN ('OEFOF','OEFOGSSC','OEFOGSSW')
    AND VDATE = (
        SELECT MAX(VDATE) FROM ccl.dbo.vw_RPT_VAL
        WHERE PCODE IN ('OEFOF','OEFOGSSC','OEFOGSSW')
    )
""", conn)

df.to_csv("hiport_snapshot.csv", index=False)
print(f"Saved {len(df)} rows to hiport_snapshot.csv")
