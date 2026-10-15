import os
from mssql_python import connect

CONN_STR = os.environ.get("FABRIC_CONN_STR", (
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=2syk64xvbs6etmnq4gvqaftbom-lxq4dihtl5nuzpponrs3hni6cm.database.fabric.microsoft.com,1433;"
    "Database={Marketplace-482d29d2-1576-45e2-9bd5-4cd16827fea9};"
    "Encrypt=yes;TrustServerCertificate=no;"
    "Authentication=ActiveDirectoryInteractive"
))


def get_connection():
    return connect(CONN_STR)


if __name__ == "__main__":
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT 1")
    print(cursor.fetchone())
    conn.close()
