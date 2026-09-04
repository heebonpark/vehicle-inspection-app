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
                signature_image BLOB,
                img_front BLOB,
                img_rear BLOB,
                img_right BLOB,
                img_left BLOB
            )
        """)
        # 기존에 만들어진 DB 파일에는 signature_image 컬럼이 없을 수 있으므로
        # 없는 경우에만 추가한다(가벼운 마이그레이션).
        try:
            conn.execute("ALTER TABLE integrated_inspections ADD COLUMN signature_image BLOB")
        except sqlite3.OperationalError:
            pass
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
