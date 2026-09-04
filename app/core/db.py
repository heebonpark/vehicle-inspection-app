import sqlite3
from contextlib import contextmanager

DB_NAME = "vehicle_safe_system.db"

@contextmanager
def get_db():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS integrated_inspections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT,
                inspect_date TEXT,
                inspector TEXT,
                hq_name TEXT,
                branch_name TEXT,
                car_no TEXT,
                check_data TEXT,
                accumulated_km TEXT,
                signature_name TEXT,
                img_front BLOB,
                img_rear BLOB,
                img_right BLOB,
                img_left BLOB
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS vehicles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT,
                hq_name TEXT,
                branch_name TEXT,
                car_no TEXT UNIQUE
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id TEXT PRIMARY KEY,
                name TEXT,
                role TEXT,
                branch TEXT,
                salt TEXT,
                pw_hash TEXT
            )
        """)
