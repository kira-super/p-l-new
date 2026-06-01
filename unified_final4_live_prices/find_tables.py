import pyodbc
import csv

CONN_STR = (
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=fcedata01.fieracapital.ca;"
    "Database=CCL;"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)

TABLES = [
    "dbo.tFactSet_AnalystsAssignments_V6",
    "dbo.tFactSet_Analysts_V6",
]

conn = pyodbc.connect(CONN_STR)
cursor = conn.cursor()

for table in TABLES:
    filename = table.replace("dbo.", "") + ".csv"
    cursor.execute(f"SELECT * FROM {table}")
    cols = [d[0] for d in cursor.description]
    rows = cursor.fetchall()
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(cols)
        writer.writerows(rows)
    print(f"✓ {table} → {filename}  ({len(rows)} rows)")

conn.close()