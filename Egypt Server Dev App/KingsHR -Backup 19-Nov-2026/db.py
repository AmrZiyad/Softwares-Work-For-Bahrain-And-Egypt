"""Database helpers — swap CONN_STR for Egypt server credentials."""
import pyodbc

CONN_STR = (
    "DRIVER={ODBC Driver 18 for SQL Server};"
    "SERVER=localhost;DATABASE=KingsSchoolEgypt;"
    "UID=businessapp;PWD=Kings@123;TrustServerCertificate=yes;"
    "Connection Timeout=5;"
)

def get_conn():
    return pyodbc.connect(CONN_STR)

def qy(sql, p=()):
    try:
        with get_conn() as c:
            cur = c.cursor(); cur.execute(sql, p)
            return [tuple(r) for r in cur.fetchall()]
    except pyodbc.Error as e:
        return []

def ex(sql, p=()):
    try:
        with get_conn() as c:
            cur = c.cursor(); cur.execute(sql, p); c.commit()
    except pyodbc.Error as e:
        raise Exception(f"Database error: {str(e)[:200]}")

def has_col(table, col):
    try:
        r = qy("SELECT COL_LENGTH(?, ?)", (table, col))
        return bool(r and r[0][0] is not None)
    except Exception:
        return False
