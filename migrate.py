"""
migrate.py - DB 마이그레이션 스크립트
기존 데이터를 유지하면서 v4 스키마로 업그레이드
실행: py migrate.py
"""
import sqlite3, os
from pathlib import Path

DB_PATH = Path(__file__).parent / "data" / "inventory.db"

def migrate():
    if not DB_PATH.exists():
        print("DB 없음 - main.py 실행 시 자동 생성됩니다.")
        return

    conn = sqlite3.connect(str(DB_PATH))
    print(f"DB 경로: {DB_PATH}")

    migrations = [
        # (테이블, 컬럼, 타입, 기본값)
        ("depletion_plans", "created_by_name", "TEXT",    None),
        ("depletion_plans", "updated_by_name",  "TEXT",    None),
        ("users",           "created_by",        "TEXT",    None),
        ("inventory_items", "qty_consumed",      "REAL",    "0"),
        ("inventory_items", "amount_consumed",   "REAL",    "0"),
    ]

    applied = 0
    for table, col, dtype, default in migrations:
        existing = [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
        if col not in existing:
            sql = f"ALTER TABLE {table} ADD COLUMN {col} {dtype}"
            if default is not None:
                sql += f" DEFAULT {default}"
            conn.execute(sql)
            print(f"  추가: {table}.{col} ({dtype})")
            applied += 1
        else:
            print(f"  이미 존재: {table}.{col}")

    conn.commit()
    conn.close()
    print(f"\n마이그레이션 완료: {applied}개 컬럼 추가")

if __name__ == "__main__":
    migrate()
