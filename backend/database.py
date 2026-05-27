"""
database.py v4
[v4 변경]
 - depletion_plans: created_by_name, updated_by_name 컬럼 추가
 - 마이그레이션: 기존 DB 유지하며 컬럼 추가 (ALTER TABLE IF NOT EXISTS)
 - 권한 관리: role 기반 (admin / user)
 - 트랜잭션 롤백 처리 적용
"""
import os, sqlite3, hashlib, logging
from pathlib import Path

logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent.parent
DB_PATH  = BASE_DIR / "data" / "inventory.db"

ROLES = {
    "admin": "관리자",
    "user":  "일반사용자",
}

def get_conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA busy_timeout=30000")
    return conn

def _add_col_if_not_exists(conn, table, col, dtype):
    """기존 데이터 유지하며 컬럼 추가 (마이그레이션)"""
    try:
        existing = [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
        if col not in existing:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {dtype}")
            logger.info(f"마이그레이션: {table}.{col} 추가")
    except Exception as e:
        logger.warning(f"컬럼 추가 실패 {table}.{col}: {e}")

def init_db():
    conn = get_conn()
    try:
        conn.execute("PRAGMA journal_mode=WAL")
    except Exception:
        pass

    conn.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        username      TEXT NOT NULL UNIQUE,
        password_hash TEXT NOT NULL,
        display_name  TEXT,
        role          TEXT NOT NULL DEFAULT 'user',
        department    TEXT,
        is_active     INTEGER DEFAULT 1,
        last_login    TEXT,
        created_at    TEXT DEFAULT (datetime('now','localtime')),
        created_by    TEXT
    );
    CREATE TABLE IF NOT EXISTS inventory_items (
        id               INTEGER PRIMARY KEY AUTOINCREMENT,
        upload_id        TEXT NOT NULL,
        ref_date         TEXT NOT NULL,
        factory          TEXT,
        item_type        TEXT,
        item_group       TEXT,
        item_code        TEXT,
        item_name        TEXT,
        cost_center      TEXT,
        cost_center_name TEXT,
        lot_no           TEXT NOT NULL,
        wo_no            TEXT,
        qty              REAL DEFAULT 0,
        weight_kg        REAL DEFAULT 0,
        weight_ton       REAL DEFAULT 0,
        amount           REAL DEFAULT 0,
        qty_consumed     REAL DEFAULT 0,
        amount_consumed  REAL DEFAULT 0,
        base_date        TEXT,
        months_label     TEXT,
        is_new           INTEGER DEFAULT 0,
        source_sheet     TEXT,
        created_at       TEXT DEFAULT (datetime('now','localtime'))
    );
    CREATE INDEX IF NOT EXISTS idx_inv_ref  ON inventory_items(ref_date);
    CREATE INDEX IF NOT EXISTS idx_inv_lot  ON inventory_items(lot_no);
    CREATE INDEX IF NOT EXISTS idx_inv_fac  ON inventory_items(factory);
    CREATE INDEX IF NOT EXISTS idx_inv_cc   ON inventory_items(cost_center);
    CREATE TABLE IF NOT EXISTS depletion_plans (
        id               INTEGER PRIMARY KEY AUTOINCREMENT,
        lot_no           TEXT NOT NULL UNIQUE,
        item_code        TEXT,
        item_name        TEXT,
        factory          TEXT,
        cost_center      TEXT,
        cost_center_name TEXT,
        item_type        TEXT,
        dept             TEXT,
        reason           TEXT,
        plan_type        TEXT,
        plan_date        TEXT,
        detail_plan      TEXT,
        is_complete      INTEGER DEFAULT 0,
        created_by       TEXT,
        created_by_name  TEXT,
        updated_by       TEXT,
        updated_by_name  TEXT,
        created_at       TEXT DEFAULT (datetime('now','localtime')),
        updated_at       TEXT DEFAULT (datetime('now','localtime'))
    );
    CREATE INDEX IF NOT EXISTS idx_dp_lot ON depletion_plans(lot_no);
    CREATE TABLE IF NOT EXISTS depletion_actuals (
        id                 INTEGER PRIMARY KEY AUTOINCREMENT,
        upload_id          TEXT NOT NULL,
        ref_date           TEXT NOT NULL,
        factory            TEXT,
        item_type          TEXT,
        item_group         TEXT,
        item_code          TEXT,
        item_name          TEXT,
        cost_center        TEXT,
        lot_no             TEXT NOT NULL,
        wo_no              TEXT,
        qty                REAL DEFAULT 0,
        weight_kg          REAL DEFAULT 0,
        weight_ton         REAL DEFAULT 0,
        qty_consumed       REAL DEFAULT 0,
        amount_consumed    REAL DEFAULT 0,
        actual_type_raw    TEXT,
        actual_type        TEXT,
        actual_type_manual TEXT,
        process_date       TEXT,
        processor          TEXT,
        source_sheet       TEXT,
        created_at         TEXT DEFAULT (datetime('now','localtime'))
    );
    CREATE INDEX IF NOT EXISTS idx_da_lot      ON depletion_actuals(lot_no);
    CREATE INDEX IF NOT EXISTS idx_da_ref_date ON depletion_actuals(ref_date);
    CREATE TABLE IF NOT EXISTS upload_history (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        upload_id    TEXT NOT NULL UNIQUE,
        filename     TEXT,
        ref_date     TEXT,
        inv_count    INTEGER DEFAULT 0,
        wip_count    INTEGER DEFAULT 0,
        act_count    INTEGER DEFAULT 0,
        total_amount REAL DEFAULT 0,
        uploaded_by  TEXT,
        created_at   TEXT DEFAULT (datetime('now','localtime'))
    );
    CREATE TABLE IF NOT EXISTS settings (
        key        TEXT PRIMARY KEY,
        value      TEXT,
        updated_by TEXT,
        updated_at TEXT DEFAULT (datetime('now','localtime'))
    );
    """)

    # 마이그레이션: 기존 DB에 컬럼 추가 (기존 데이터 유지)
    _add_col_if_not_exists(conn, "depletion_plans", "created_by_name", "TEXT")
    _add_col_if_not_exists(conn, "depletion_plans", "updated_by_name",  "TEXT")
    _add_col_if_not_exists(conn, "users",           "created_by",       "TEXT")

    _ensure_admin(conn)
    for k, v in [("company_name","귀사"), ("current_upload_id","")]:
        conn.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (k, v))

    conn.commit()
    conn.close()

def _hash_pw(plain: str) -> str:
    return hashlib.sha256(plain.encode()).hexdigest()

def verify_pw(plain: str, hashed: str) -> bool:
    return _hash_pw(plain) == hashed

def _ensure_admin(conn):
    if not conn.execute("SELECT id FROM users WHERE username='admin'").fetchone():
        conn.execute(
            "INSERT INTO users(username,password_hash,display_name,role) VALUES(?,?,?,?)",
            ("admin", _hash_pw("admin1234"), "시스템 관리자", "admin")
        )

def authenticate(username: str, password: str):
    conn = get_conn()
    try:
        row = conn.execute(
            "SELECT * FROM users WHERE username=? AND is_active=1", (username,)
        ).fetchone()
        if not row or not verify_pw(password, row["password_hash"]):
            return None
        conn.execute(
            "UPDATE users SET last_login=datetime('now','localtime') WHERE username=?",
            (username,)
        )
        conn.commit()
        return dict(row)
    except Exception as e:
        logger.error(f"인증 오류: {e}")
        return None
    finally:
        conn.close()

def get_setting(key: str, default: str = "") -> str:
    conn = get_conn()
    try:
        row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row and row["value"] is not None else default
    finally:
        conn.close()
