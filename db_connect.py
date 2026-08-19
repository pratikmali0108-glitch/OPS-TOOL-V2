import pyodbc

SERVER = 'SJSERVER'  # Or your server name / IP (e.g., '.\SQLEXPRESS')
DATABASE = 'TSHISJEP'

# --- Connection String for ODBC Driver 18 ---
conn_str = (
    "DRIVER={ODBC Driver 18 for SQL Server};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)

try:
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    print("Successfully connected to SQL Server!")

    # Test query
    cursor.execute("SELECT @@VERSION;")
    print("SQL Server Version:", cursor.fetchone()[0])

    cursor.close()
    conn.close()

except pyodbc.Error as e:
    print(f"Error: {e}")