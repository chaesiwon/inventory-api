"""
migrate.py - DB 마이그레이션 (기존 데이터 유지)
Render/Railway 시작 시 자동 실행
"""
import sys
import sqlite3
import os
from pathlib import Path

DB_PATH = Path(__file__).parent / "data" / "inventory.db"

def migrate():
    # data 디렉토리 생성
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)

    if not DB_PATH.exists():
        print("[migrate] DB 없음 - 신규 생성 예정 (init_db에서 처리)")
        return True

    try:
        conn = sqlite3.connect(str(DB_PATH))
        migrations = [
            ("depletion_plans", "created_by_name", "TEXT"),
            ("depletion_plans", "updated_by_name",  "TEXT"),
            ("users",           "created_by",        "TEXT"),
            ("inventory_items", "qty_consumed",      "REAL DEFAULT 0"),
            ("inventory_items", "amount_consumed",   "REAL DEFAULT 0"),
        ]
        applied = 0
        for table, col, dtype in migrations:
            # 테이블 존재 여부 확인
            tbl_exists = conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,)
            ).fetchone()
            if not tbl_exists:
                continue
            existing = [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
            if col not in existing:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {dtype}")
                print(f"[migrate] 컬럼 추가: {table}.{col}")
                applied += 1
        conn.commit()
        conn.close()
        print(f"[migrate] 완료 ({applied}개 변경)")
        return True
    except Exception as e:
        print(f"[migrate] 오류 (무시하고 계속): {e}")
        return True  # 실패해도 서버 시작은 계속

if __name__ == "__main__":
    success = migrate()
    sys.exit(0)  # 항상 0 반환 (빌드/시작 중단 방지)
