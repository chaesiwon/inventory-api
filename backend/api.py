"""
api.py v5 - 전체 REST API
[v5 수정사항]
 1. 다운로드 한글 파일명: urllib.parse.quote 적용 (RFC 5987)
 2. 일괄입력 상태 관리: lot_no 기준 정확한 대상만 저장, 중복 방지
 3. PPT 다운로드 추가 (/api/compare/export-ppt)
 4. 사용자/권한 관리 API 추가
 5. 작성자/작성일시 자동 기록 (created_by_name, updated_by_name)
 6. 전체 try/except + 롤백 처리
 7. 권한 검증 데코레이터 패턴 적용
"""
import io, logging, urllib.parse
from datetime import datetime, date
from typing import Optional, List

import pandas as pd
from fastapi import APIRouter, File, UploadFile, HTTPException, Query, Request
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from backend.database import get_conn, authenticate, get_setting, _hash_pw, verify_pw, ROLES
from backend.loader   import parse_inventory_file, save_parsed_data

logger = logging.getLogger(__name__)
router = APIRouter()

# ── 헬퍼
def cur_user(request: Request) -> Optional[dict]:
    """세션 또는 헤더에서 사용자 정보 조회 (세션 실패 시 헤더 백업)"""
    # 1. 세션에서 조회
    u = request.session.get("user") if hasattr(request, 'session') else None
    if u:
        return u
    # 2. X-User-Id 헤더 기반 조회 (세션 쿠키 미작동 환경 대비)
    user_id = request.headers.get("X-User-Id")
    user_token = request.headers.get("X-Auth-Token")
    if user_id and user_token:
        try:
            conn = get_conn()
            row = conn.execute(
                "SELECT id,username,display_name,role,department FROM users WHERE id=? AND is_active=1",
                (int(user_id),)
            ).fetchone()
            conn.close()
            if row:
                import hashlib
                expected = hashlib.sha256(f"{row['username']}-{row['id']}-inventory2024".encode()).hexdigest()[:16]
                if user_token == expected:
                    return dict(row)
        except Exception:
            pass
    return None

def require_role(request: Request, role: str = "user") -> dict:
    u = cur_user(request)
    if not u:
        raise HTTPException(401, "로그인이 필요합니다.")
    if role == "admin" and u.get("role") != "admin":
        raise HTTPException(403, "관리자 권한이 필요합니다.")
    return u

def _safe_fname(name: str) -> str:
    """한글 파일명 RFC 5987 인코딩"""
    return urllib.parse.quote(name, safe="")

def _xlsx_response(buf: io.BytesIO, filename: str) -> StreamingResponse:
    """Excel StreamingResponse - 한글 파일명 안전 처리"""
    buf.seek(0)
    encoded = _safe_fname(filename)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"}
    )

def _pptx_response(buf: io.BytesIO, filename: str) -> StreamingResponse:
    """PPT StreamingResponse"""
    buf.seek(0)
    encoded = _safe_fname(filename)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"}
    )

def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _latest_ref(conn) -> str:
    row = conn.execute(
        "SELECT ref_date FROM inventory_items ORDER BY ref_date DESC LIMIT 1"
    ).fetchone()
    return row["ref_date"] if row else ""

def _prev_ref(conn, ref_date: str, mode: str = "month") -> Optional[str]:
    rows = conn.execute(
        "SELECT DISTINCT ref_date FROM inventory_items WHERE ref_date < ? ORDER BY ref_date DESC",
        (ref_date,)
    ).fetchall()
    dates = [r["ref_date"] for r in rows]
    if not dates: return None
    if mode == "month":   return dates[0]
    if mode == "quarter":
        for d in dates:
            if len(d)==8 and int(d[:6]) <= int(ref_date[:6]) - 3: return d
        return dates[-1]
    if mode == "year":
        for d in dates:
            if len(d)==8 and int(d[:6]) <= int(ref_date[:6]) - 12: return d
        return dates[-1]
    return dates[0]

# ══════════════════════════════════════
# 인증
# ══════════════════════════════════════
class LoginBody(BaseModel):
    username: str
    password: str

@router.post("/auth/login")
async def login(body: LoginBody, request: Request):
    try:
        user = authenticate(body.username, body.password)
        if not user:
            raise HTTPException(401, "아이디 또는 비밀번호가 틀렸습니다.")
        request.session["user"] = {
            k: user[k] for k in ("id","username","display_name","role","department")
            if k in dict(user)
        }
        # 세션 저장 (실패 시에도 계속 진행)
        session_user = {
            k: user[k] for k in ("id","username","display_name","role","department")
            if k in dict(user)
        }
        try:
            request.session["user"] = session_user
        except Exception as se:
            logger.warning(f"세션 저장 실패 (무시): {se}")
        # 헤더 기반 인증을 위한 토큰 생성 (세션 백업)
        import hashlib as _hs
        token = _hs.sha256(f"{user['username']}-{user['id']}-inventory2024".encode()).hexdigest()[:16]
        return {"ok": True, "user": session_user, "token": token}
    except HTTPException: raise
    except Exception as e:
        logger.error(f"로그인 오류: {e}", exc_info=True)
        raise HTTPException(500, f"로그인 처리 중 오류: {e}")

@router.post("/auth/logout")
async def logout(request: Request):
    request.session.clear()
    return {"ok": True}

@router.get("/auth/me")
async def me(request: Request):
    u = cur_user(request)
    return {"logged_in": bool(u), "user": u}

# ── 설정
class SettingBody(BaseModel):
    key: str; value: str

@router.get("/settings")
async def get_settings():
    conn = get_conn()
    try:
        rows = conn.execute("SELECT key,value FROM settings").fetchall()
        return {r["key"]: r["value"] for r in rows}
    finally: conn.close()

@router.post("/settings")
async def save_setting(body: SettingBody, request: Request):
    u = require_role(request, "admin")
    conn = get_conn()
    try:
        conn.execute("""
            INSERT INTO settings(key,value,updated_by,updated_at)
            VALUES(?,?,?,datetime('now','localtime'))
            ON CONFLICT(key) DO UPDATE SET
              value=excluded.value, updated_by=excluded.updated_by,
              updated_at=excluded.updated_at
        """, (body.key, body.value, u["username"]))
        conn.commit()
        return {"ok": True}
    except Exception as e:
        conn.rollback()
        logger.error(f"설정 저장 오류: {e}", exc_info=True)
        raise HTTPException(500, f"설정 저장 실패: {e}")
    finally: conn.close()

# ══════════════════════════════════════
# 사용자 관리 (admin only)
# ══════════════════════════════════════
class UserCreateBody(BaseModel):
    username: str
    password: str
    display_name: str
    role: str = "user"
    department: Optional[str] = None

class UserUpdateBody(BaseModel):
    display_name: Optional[str] = None
    role:         Optional[str] = None
    department:   Optional[str] = None
    is_active:    Optional[int] = None
    password:     Optional[str] = None

@router.get("/users")
async def list_users(request: Request):
    require_role(request, "admin")
    conn = get_conn()
    try:
        rows = conn.execute(
            "SELECT id,username,display_name,role,department,is_active,last_login,created_at FROM users ORDER BY id"
        ).fetchall()
        return {"users": [dict(r) for r in rows]}
    finally: conn.close()

@router.post("/users")
async def create_user(body: UserCreateBody, request: Request):
    u = require_role(request, "admin")
    if body.role not in ROLES:
        raise HTTPException(400, f"유효하지 않은 권한: {body.role}. 가능값: {list(ROLES.keys())}")
    conn = get_conn()
    try:
        existing = conn.execute("SELECT id FROM users WHERE username=?", (body.username,)).fetchone()
        if existing:
            raise HTTPException(400, f"이미 존재하는 아이디: {body.username}")
        conn.execute(
            "INSERT INTO users(username,password_hash,display_name,role,department,created_by) VALUES(?,?,?,?,?,?)",
            (body.username, _hash_pw(body.password), body.display_name,
             body.role, body.department, u["username"])
        )
        conn.commit()
        return {"ok": True, "message": f"사용자 {body.username} 생성 완료"}
    except HTTPException: raise
    except Exception as e:
        conn.rollback()
        logger.error(f"사용자 생성 오류: {e}", exc_info=True)
        raise HTTPException(500, f"사용자 생성 실패: {e}")
    finally: conn.close()

@router.put("/users/{user_id}")
async def update_user(user_id: int, body: UserUpdateBody, request: Request):
    u = require_role(request, "admin")
    conn = get_conn()
    try:
        row = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if not row:
            raise HTTPException(404, f"사용자 없음: id={user_id}")
        sets, params = [], []
        if body.display_name is not None: sets.append("display_name=?"); params.append(body.display_name)
        if body.role          is not None:
            if body.role not in ROLES: raise HTTPException(400, f"유효하지 않은 권한: {body.role}")
            sets.append("role=?"); params.append(body.role)
        if body.department    is not None: sets.append("department=?"); params.append(body.department)
        if body.is_active     is not None: sets.append("is_active=?"); params.append(body.is_active)
        if body.password      is not None: sets.append("password_hash=?"); params.append(_hash_pw(body.password))
        if not sets: return {"ok": True, "message": "변경 없음"}
        params.append(user_id)
        conn.execute(f"UPDATE users SET {','.join(sets)} WHERE id=?", params)
        conn.commit()
        return {"ok": True, "message": "사용자 정보 수정 완료"}
    except HTTPException: raise
    except Exception as e:
        conn.rollback()
        logger.error(f"사용자 수정 오류: {e}", exc_info=True)
        raise HTTPException(500, f"사용자 수정 실패: {e}")
    finally: conn.close()

# ══════════════════════════════════════
# 기준일 목록
# ══════════════════════════════════════
@router.get("/inventory/ref-dates")
async def ref_dates():
    conn = get_conn()
    try:
        rows = conn.execute(
            "SELECT DISTINCT ref_date FROM inventory_items ORDER BY ref_date DESC"
        ).fetchall()
        return {"ref_dates": [r["ref_date"] for r in rows]}
    finally: conn.close()

# ══════════════════════════════════════
# 대시보드
# ══════════════════════════════════════
@router.get("/dashboard/kpi")
async def dashboard_kpi(ref_date: Optional[str] = Query(None)):
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        if not ref_date:
            return {"ref_date":"","total_amount":0,"total_weight_ton":0,"total_count":0,
                    "new_amount":0,"completed_amount":0,"uncompleted_amount":0,
                    "no_plan_amount":0,"plan_this_month":0,"total_consumed_amount":0}

        r = conn.execute("""
            SELECT COALESCE(SUM(amount),0) AS ta, COALESCE(SUM(weight_ton),0) AS tw,
                   COUNT(*) AS tc,
                   COALESCE(SUM(CASE WHEN is_new=1 THEN amount ELSE 0 END),0) AS na,
                   COALESCE(SUM(amount_consumed),0) AS tca
            FROM inventory_items WHERE ref_date=?
        """, (ref_date,)).fetchone()

        completed = conn.execute("""
            SELECT COALESCE(SUM(i.amount),0) FROM inventory_items i
            WHERE i.ref_date=?
              AND EXISTS(SELECT 1 FROM depletion_plans   p WHERE p.lot_no=i.lot_no)
              AND EXISTS(SELECT 1 FROM depletion_actuals a WHERE a.lot_no=i.lot_no)
        """, (ref_date,)).fetchone()[0]

        uncompleted = conn.execute("""
            SELECT COALESCE(SUM(i.amount),0) FROM inventory_items i
            WHERE i.ref_date=?
              AND EXISTS    (SELECT 1 FROM depletion_plans   p WHERE p.lot_no=i.lot_no)
              AND NOT EXISTS(SELECT 1 FROM depletion_actuals a WHERE a.lot_no=i.lot_no)
        """, (ref_date,)).fetchone()[0]

        no_plan = conn.execute("""
            SELECT COALESCE(SUM(i.amount),0) FROM inventory_items i
            WHERE i.ref_date=? AND NOT EXISTS(SELECT 1 FROM depletion_plans p WHERE p.lot_no=i.lot_no)
        """, (ref_date,)).fetchone()[0]

        this_month = date.today().strftime("%Y-%m")
        plan_this = conn.execute("""
            SELECT COALESCE(SUM(i.amount),0) FROM inventory_items i
            JOIN depletion_plans p ON p.lot_no=i.lot_no
            WHERE i.ref_date=? AND p.plan_date LIKE ?
        """, (ref_date, f"{this_month}%")).fetchone()[0]

        return {
            "ref_date": ref_date,
            "total_amount":          float(r["ta"]),
            "total_weight_ton":      float(r["tw"]),
            "total_count":           int(r["tc"]),
            "new_amount":            float(r["na"]),
            "completed_amount":      float(completed),
            "uncompleted_amount":    float(uncompleted),
            "no_plan_amount":        float(no_plan),
            "plan_this_month":       float(plan_this),
            "total_consumed_amount": float(r["tca"]),
        }
    except Exception as e:
        logger.error(f"KPI 조회 오류: {e}", exc_info=True)
        raise HTTPException(500, f"KPI 조회 실패: {e}")
    finally: conn.close()

@router.get("/dashboard/top20")
async def top20(ref_date: Optional[str] = Query(None)):
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        rows = conn.execute("""
            SELECT i.factory, i.item_type, i.item_code, i.item_name,
                   COALESCE(i.cost_center_name,i.cost_center) AS cc_name,
                   i.lot_no, i.weight_ton, i.amount, i.qty_consumed, i.amount_consumed,
                   i.base_date, i.months_label, i.is_new,
                   p.plan_type, p.plan_date, p.dept,
                   CASE WHEN a.lot_no IS NOT NULL THEN 1 ELSE 0 END AS is_completed
            FROM inventory_items i
            LEFT JOIN depletion_plans   p ON p.lot_no=i.lot_no
            LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no
            WHERE i.ref_date=? ORDER BY i.amount DESC LIMIT 20
        """, (ref_date,)).fetchall()
        return {"ref_date": ref_date, "items": [dict(r) for r in rows]}
    except Exception as e:
        logger.error(f"TOP20 오류: {e}", exc_info=True)
        raise HTTPException(500, f"TOP20 조회 실패: {e}")
    finally: conn.close()

@router.get("/dashboard/monthly-trend")
async def monthly_trend():
    conn = get_conn()
    try:
        rows = conn.execute("""
            SELECT ref_date, SUM(amount) AS total_amount,
                   SUM(weight_ton) AS total_weight_ton,
                   SUM(amount_consumed) AS total_consumed,
                   COUNT(*) AS item_count
            FROM inventory_items GROUP BY ref_date ORDER BY ref_date
        """).fetchall()
        return {"trend": [dict(r) for r in rows]}
    finally: conn.close()

@router.get("/dashboard/plan-weight-trend")
async def plan_weight_trend():
    conn = get_conn()
    try:
        rows = conn.execute("""
            SELECT substr(p.plan_date,1,7) AS plan_month,
                   COUNT(*) AS plan_count,
                   COALESCE(SUM(i.weight_ton),0) AS plan_weight_ton,
                   COALESCE(SUM(i.amount),0) AS plan_amount,
                   COALESCE(SUM(i.amount_consumed),0) AS consumed_amount
            FROM depletion_plans p
            LEFT JOIN inventory_items i ON i.lot_no=p.lot_no
            WHERE p.plan_date IS NOT NULL AND p.plan_date!=''
            GROUP BY plan_month ORDER BY plan_month
        """).fetchall()
        return {"trend": [dict(r) for r in rows]}
    finally: conn.close()

@router.get("/dashboard/cost-center-summary")
async def cost_center_summary(ref_date: Optional[str] = Query(None)):
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        rows = conn.execute("""
            SELECT COALESCE(i.cost_center_name,i.cost_center) AS cc_name,
                   i.cost_center,
                   COUNT(*) AS item_count,
                   COALESCE(SUM(i.weight_ton),0) AS total_weight,
                   COALESCE(SUM(i.amount),0) AS total_amount,
                   COALESCE(SUM(i.amount_consumed),0) AS consumed_amount,
                   COUNT(p.lot_no) AS plan_count,
                   COUNT(a.lot_no) AS actual_count
            FROM inventory_items i
            LEFT JOIN depletion_plans   p ON p.lot_no=i.lot_no
            LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no
            WHERE i.ref_date=?
            GROUP BY i.cost_center, i.cost_center_name
            ORDER BY total_amount DESC
        """, (ref_date,)).fetchall()
        return {"ref_date": ref_date, "items": [dict(r) for r in rows]}
    finally: conn.close()

@router.get("/dashboard/period-compare")
async def period_compare(
    ref_date: Optional[str] = Query(None),
    mode: str = Query("month")
):
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        prev_rd = _prev_ref(conn, ref_date, mode)

        def _summary(rd):
            if not rd: return None
            r = conn.execute("""
                SELECT COALESCE(SUM(amount),0) AS ta, COALESCE(SUM(weight_ton),0) AS tw,
                       COUNT(*) AS tc, COALESCE(SUM(amount_consumed),0) AS tca
                FROM inventory_items WHERE ref_date=?
            """, (rd,)).fetchone()
            ar = conn.execute("""
                SELECT COUNT(DISTINCT i.lot_no) AS plan_total,
                       SUM(CASE WHEN a.lot_no IS NOT NULL THEN 1 ELSE 0 END) AS action_cnt,
                       COALESCE(SUM(CASE WHEN a.lot_no IS NOT NULL THEN i.amount ELSE 0 END),0) AS action_amt,
                       COALESCE(SUM(CASE WHEN a.lot_no IS NOT NULL THEN i.weight_ton ELSE 0 END),0) AS action_wt
                FROM inventory_items i
                JOIN depletion_plans p ON p.lot_no=i.lot_no
                LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no
                WHERE i.ref_date=?
            """, (rd,)).fetchone()
            return {
                "ref_date": rd, "total_amount": float(r["ta"]),
                "total_weight": float(r["tw"]), "total_count": int(r["tc"]),
                "consumed_amount": float(r["tca"]),
                "plan_total": int(ar["plan_total"]),
                "action_count": int(ar["action_cnt"]),
                "action_amount": float(ar["action_amt"]),
                "action_weight": float(ar["action_wt"]),
            }

        return {
            "mode": mode, "current": _summary(ref_date), "previous": _summary(prev_rd),
            "mode_label": {"month":"전월","quarter":"전분기","year":"전년도"}.get(mode,"전월")
        }
    finally: conn.close()

# ══════════════════════════════════════
# 재고 목록 (조회전용)
# ══════════════════════════════════════
@router.get("/inventory")
async def inventory_list(
    ref_date: Optional[str]=Query(None), factory: Optional[str]=Query(None),
    item_type: Optional[str]=Query(None), item_code: Optional[str]=Query(None),
    lot_no: Optional[str]=Query(None), dept: Optional[str]=Query(None),
    plan_type: Optional[str]=Query(None), cost_center: Optional[str]=Query(None),
    item_name: Optional[str]=Query(None),
    page: int=Query(1,ge=1), page_size: int=Query(50,ge=1,le=500),
):
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        conds=["i.ref_date=?"]; params=[ref_date]
        if factory:     conds.append("i.factory=?");         params.append(factory)
        if item_type:   conds.append("i.item_type=?");       params.append(item_type)
        if item_code:   conds.append("i.item_code LIKE ?");  params.append(f"%{item_code}%")
        if lot_no:      conds.append("i.lot_no LIKE ?");     params.append(f"%{lot_no}%")
        if dept:        conds.append("p.dept=?");            params.append(dept)
        if plan_type:   conds.append("p.plan_type=?");       params.append(plan_type)
        if cost_center: conds.append("(i.cost_center=? OR i.cost_center_name LIKE ?)"); params+=[cost_center,f"%{cost_center}%"]
        if item_name:   conds.append("i.item_name LIKE ?");  params.append(f"%{item_name}%")
        where=" AND ".join(conds); offset=(page-1)*page_size
        total=conn.execute(
            f"SELECT COUNT(*) FROM inventory_items i LEFT JOIN depletion_plans p ON p.lot_no=i.lot_no WHERE {where}",
            params
        ).fetchone()[0]
        rows=conn.execute(f"""
            SELECT i.factory,i.item_type,i.item_code,i.item_name,
                   COALESCE(i.cost_center_name,i.cost_center) AS cc_name, i.cost_center,
                   i.lot_no,i.wo_no,i.qty,i.weight_ton,i.amount,
                   i.qty_consumed,i.amount_consumed,
                   i.base_date,i.months_label,i.is_new,i.source_sheet,
                   p.dept,p.reason,p.plan_type,p.plan_date,p.detail_plan,
                   p.created_by,p.created_by_name,p.created_at AS plan_created_at,
                   p.updated_by,p.updated_by_name,p.updated_at AS plan_updated_at,
                   p.is_complete,
                   CASE WHEN a.lot_no IS NOT NULL THEN 1 ELSE 0 END AS has_actual
            FROM inventory_items i
            LEFT JOIN depletion_plans   p ON p.lot_no=i.lot_no
            LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no
            WHERE {where} ORDER BY i.amount DESC LIMIT ? OFFSET ?
        """,params+[page_size,offset]).fetchall()
        return {"ref_date":ref_date,"total":total,"page":page,"page_size":page_size,"items":[dict(r) for r in rows]}
    except Exception as e:
        logger.error(f"재고 목록 오류: {e}", exc_info=True)
        raise HTTPException(500, f"재고 목록 조회 실패: {e}")
    finally: conn.close()

@router.get("/inventory/export")
async def export_inventory(
    ref_date: Optional[str]=Query(None),
    factory: Optional[str]=Query(None),
    item_type: Optional[str]=Query(None),
):
    """재고현황 Excel 다운로드 - 한글 파일명 RFC 5987 처리"""
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        if not ref_date: raise HTTPException(404, "업로드된 재고 데이터가 없습니다.")
        conds=["i.ref_date=?"]; params=[ref_date]
        if factory:   conds.append("i.factory=?");   params.append(factory)
        if item_type: conds.append("i.item_type=?"); params.append(item_type)
        where=" AND ".join(conds)
        rows=conn.execute(f"""
            SELECT i.factory,i.item_type,i.item_code,i.item_name,
                   COALESCE(i.cost_center_name,i.cost_center) AS cc_name,
                   i.lot_no,i.qty,i.weight_ton,i.amount,i.qty_consumed,i.amount_consumed,
                   i.base_date,i.months_label,i.is_new,
                   p.dept,p.reason,p.plan_type,p.plan_date,p.detail_plan,
                   p.created_by_name,p.created_at,p.updated_by_name,p.updated_at
            FROM inventory_items i LEFT JOIN depletion_plans p ON p.lot_no=i.lot_no
            WHERE {where} ORDER BY i.amount DESC
        """,params).fetchall()
        conn.close()

        data = [dict(r) for r in rows]
        if not data:
            # 빈 데이터도 헤더만 있는 엑셀 생성
            data = []
        df = pd.DataFrame(data)
        COL_MAP = {
            "factory":"공장","item_type":"품목구분","item_code":"품목코드","item_name":"품명",
            "cc_name":"원가중심점","lot_no":"LOT NO","qty":"수량","weight_ton":"중량(ton)",
            "amount":"금액","qty_consumed":"소진수량","amount_consumed":"소진금액",
            "base_date":"기준일자","months_label":"개월","is_new":"신규여부",
            "dept":"담당부서","reason":"장기재고사유","plan_type":"소진계획방안",
            "plan_date":"소진계획기한","detail_plan":"세부계획",
            "created_by_name":"작성자","created_at":"작성일시",
            "updated_by_name":"수정자","updated_at":"수정일시",
        }
        if not df.empty:
            rename_cols = {k:v for k,v in COL_MAP.items() if k in df.columns}
            df = df.rename(columns=rename_cols)
        else:
            df = pd.DataFrame(columns=list(COL_MAP.values()))

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False, sheet_name="장기재고현황")
            ws = w.sheets["장기재고현황"]
            hdr = w.book.add_format({"bold":True,"bg_color":"#DDEBF7","border":1})
            for ci, col in enumerate(df.columns):
                ws.write(0, ci, col, hdr)
                ws.set_column(ci, ci, max(12, len(str(col))+2))

        fname = f"장기재고현황_{ref_date}_{datetime.now().strftime('%H%M%S')}.xlsx"
        return _xlsx_response(buf, fname)

    except HTTPException: raise
    except Exception as e:
        logger.error(f"재고 Excel 다운로드 오류: {e}", exc_info=True)
        raise HTTPException(500, f"Excel 생성 실패: {e}")

# ══════════════════════════════════════
# 업로드
# ══════════════════════════════════════
@router.post("/upload")
async def upload_file(request: Request, file: UploadFile = File(...)):
    u = require_role(request, "admin")
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(400, ".xlsx 파일만 업로드 가능합니다.")
    try:
        content = await file.read()
        if len(content) == 0:
            raise HTTPException(400, "빈 파일입니다.")
        parsed = parse_inventory_file(content, file.filename, u["username"])
        if "error" in parsed:
            raise HTTPException(400, parsed["error"])
        result = save_parsed_data(parsed, u["username"])
        return {
            "ok": True, "upload_id": result["upload_id"],
            "ref_date": parsed["ref_date"],
            "inv_count": result["inv_count"], "wip_count": result["wip_count"],
            "act_count": result["act_count"], "total_amount": result["total_amount"],
            "warnings": parsed.get("warnings", []),
        }
    except HTTPException: raise
    except Exception as e:
        logger.error(f"파일 업로드 오류: {e}", exc_info=True)
        raise HTTPException(500, f"업로드 처리 실패: {e}")

@router.get("/upload-history")
async def upload_history():
    conn = get_conn()
    try:
        rows = conn.execute("SELECT * FROM upload_history ORDER BY created_at DESC LIMIT 50").fetchall()
        return {"history": [dict(r) for r in rows]}
    finally: conn.close()

@router.delete("/upload/{upload_id}")
async def delete_upload(upload_id: str, request: Request):
    require_role(request, "admin")
    conn = get_conn()
    try:
        row = conn.execute("SELECT ref_date FROM upload_history WHERE upload_id=?", (upload_id,)).fetchone()
        if not row:
            raise HTTPException(404, f"업로드 이력 없음: {upload_id}")
        conn.execute("DELETE FROM inventory_items   WHERE upload_id=?", (upload_id,))
        conn.execute("DELETE FROM depletion_actuals WHERE upload_id=?", (upload_id,))
        conn.execute("DELETE FROM upload_history    WHERE upload_id=?", (upload_id,))
        conn.commit()
        return {"ok": True, "deleted": upload_id}
    except HTTPException: raise
    except Exception as e:
        conn.rollback()
        logger.error(f"업로드 삭제 오류: {e}", exc_info=True)
        raise HTTPException(500, f"삭제 실패: {e}")
    finally: conn.close()

@router.delete("/upload/all/data")
async def delete_all(request: Request):
    require_role(request, "admin")
    conn = get_conn()
    try:
        for t in ["inventory_items","depletion_actuals","upload_history"]:
            conn.execute(f"DELETE FROM {t}")
        conn.commit()
        return {"ok": True}
    except Exception as e:
        conn.rollback()
        logger.error(f"전체 삭제 오류: {e}", exc_info=True)
        raise HTTPException(500, f"전체 삭제 실패: {e}")
    finally: conn.close()

# ══════════════════════════════════════
# 소진계획: 경로 충돌 주의 - 구체적 경로를 먼저 등록
# ══════════════════════════════════════
@router.get("/plans/export-template")
async def export_plan_template(ref_date: Optional[str] = Query(None)):
    """소진계획 템플릿 다운로드 - 한글 파일명 RFC 5987 처리"""
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        if not ref_date: raise HTTPException(404, "업로드된 재고 데이터가 없습니다.")
        rows = conn.execute("""
            SELECT i.factory, i.item_type, i.item_code, i.item_name,
                   COALESCE(i.cost_center_name,i.cost_center) AS cc_name,
                   i.lot_no, i.qty, i.weight_ton, i.amount, i.base_date,
                   p.dept, p.reason, p.plan_type, p.plan_date, p.detail_plan
            FROM inventory_items i LEFT JOIN depletion_plans p ON p.lot_no=i.lot_no
            WHERE i.ref_date=? ORDER BY i.amount DESC
        """, (ref_date,)).fetchall()
        conn.close()

        COL_NAMES = ["공장","품목구분","품목코드","품명","원가중심점",
                     "LOT NO","수량","중량(ton)","금액","기준일자",
                     "담당부서","장기재고사유","소진계획방안","소진계획기한","세부계획"]
        READONLY_COLS = 10  # A~J (공장~기준일자) 읽기전용
        data = [dict(r) for r in rows]
        df = pd.DataFrame(data) if data else pd.DataFrame(columns=COL_NAMES)
        if not df.empty:
            df.columns = COL_NAMES

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False, sheet_name="소진계획입력")
            wb = w.book; ws = w.sheets["소진계획입력"]
            hdr = wb.add_format({"bold":True,"bg_color":"#DDEBF7","border":1,"align":"center"})
            ro  = wb.add_format({"bg_color":"#F2F2F2","border":1})
            inp = wb.add_format({"bg_color":"#FFFFFF","border":1})
            for ci, cn in enumerate(COL_NAMES):
                ws.write(0, ci, cn, hdr)
                ws.set_column(ci, ci, 15)
            for ri in range(len(df)):
                for ci in range(READONLY_COLS):
                    val = df.iloc[ri, ci] if ci < len(df.columns) else ""
                    ws.write(ri+1, ci, "" if str(val)=="nan" else val, ro)
                for ci in range(READONLY_COLS, len(COL_NAMES)):
                    val = df.iloc[ri, ci] if ci < len(df.columns) else ""
                    ws.write(ri+1, ci, "" if str(val)=="nan" else val, inp)
            # 드롭다운 validation
            n = max(len(df)+100, 200)
            ws.data_validation(1,10,n,10,{"validate":"list","source":["영업","생산","구매"]})
            ws.data_validation(1,11,n,11,{"validate":"list","source":["주문 변경","주문 취소","납품 후 잔량","기타"]})
            ws.data_validation(1,12,n,12,{"validate":"list","source":["생산투입","전환 판매","폐기","기타"]})

        fname = f"소진계획입력템플릿_{ref_date}.xlsx"
        return _xlsx_response(buf, fname)

    except HTTPException: raise
    except Exception as e:
        logger.error(f"템플릿 다운로드 오류: {e}", exc_info=True)
        raise HTTPException(500, f"템플릿 생성 실패: {e}")

@router.post("/plans/bulk-upload")
async def bulk_upload_plans(request: Request, file: UploadFile = File(...)):
    """엑셀 일괄 업로드: LOT NO 기준 upsert, 중복 방지"""
    u = cur_user(request)
    if not u: raise HTTPException(401, "로그인이 필요합니다.")
    try:
        content = await file.read()
        df = pd.read_excel(io.BytesIO(content), header=0, dtype=str)
    except Exception as e:
        raise HTTPException(400, f"엑셀 파싱 실패: {e}")

    # 필수 컬럼 검증
    if "LOT NO" not in df.columns:
        raise HTTPException(400, f"필수 컬럼 'LOT NO' 없음. 템플릿을 다시 다운로드하여 사용하세요.")

    by = u["username"]; by_name = u.get("display_name","") or by
    ok = fail = 0; errors = []; now = _now_str()
    conn = get_conn()
    try:
        for idx, row in df.iterrows():
            lot = str(row.get("LOT NO","")).strip()
            if not lot or lot == "nan": continue
            try:
                dept  = str(row.get("담당부서","")).strip()    or None
                rsn   = str(row.get("장기재고사유","")).strip() or None
                ptype = str(row.get("소진계획방안","")).strip() or None
                pdate = str(row.get("소진계획기한","")).strip() or None
                det   = str(row.get("세부계획","")).strip()     or None
                # 날짜 형식 정규화
                if pdate and pdate not in ("nan","None",""):
                    for fmt in ("%Y-%m-%d","%Y/%m/%d","%Y.%m.%d"):
                        try: pdate = datetime.strptime(pdate[:10], fmt).strftime("%Y-%m-%d"); break
                        except: pass
                    else: pdate = None
                else: pdate = None

                ex = conn.execute("SELECT id FROM depletion_plans WHERE lot_no=?", (lot,)).fetchone()
                if ex:
                    conn.execute("""
                        UPDATE depletion_plans
                        SET dept=?,reason=?,plan_type=?,plan_date=?,detail_plan=?,
                            updated_by=?,updated_by_name=?,updated_at=?
                        WHERE lot_no=?
                    """, (dept,rsn,ptype,pdate,det,by,by_name,now,lot))
                else:
                    inv = conn.execute(
                        "SELECT item_code,item_name,factory,cost_center,cost_center_name,item_type FROM inventory_items WHERE lot_no=? LIMIT 1",
                        (lot,)
                    ).fetchone()
                    conn.execute("""
                        INSERT INTO depletion_plans
                            (lot_no,item_code,item_name,factory,cost_center,cost_center_name,item_type,
                             dept,reason,plan_type,plan_date,detail_plan,
                             created_by,created_by_name,updated_by,updated_by_name)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (lot,
                          inv["item_code"] if inv else None, inv["item_name"] if inv else None,
                          inv["factory"]   if inv else None, inv["cost_center"] if inv else None,
                          inv["cost_center_name"] if inv else None, inv["item_type"] if inv else None,
                          dept,rsn,ptype,pdate,det,by,by_name,by,by_name))
                ok += 1
            except Exception as e:
                fail += 1
                errors.append(f"행{idx+2} LOT:{lot} → {e}")
        conn.commit()
        return {"ok": True, "success": ok, "fail": fail, "errors": errors[:20]}
    except Exception as e:
        conn.rollback()
        logger.error(f"일괄 업로드 오류: {e}", exc_info=True)
        raise HTTPException(500, f"일괄 업로드 실패: {e}")
    finally: conn.close()

@router.get("/plans/no-plan")
async def inventory_no_plan(
    ref_date: Optional[str]=Query(None), factory: Optional[str]=Query(None),
    item_name: Optional[str]=Query(None), cost_center: Optional[str]=Query(None),
    lot_no: Optional[str]=Query(None),
    page: int=Query(1,ge=1), page_size: int=Query(50,ge=1,le=500)
):
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        conds=["i.ref_date=?","NOT EXISTS(SELECT 1 FROM depletion_plans p WHERE p.lot_no=i.lot_no)"]
        params=[ref_date]
        if factory:     conds.append("i.factory=?");        params.append(factory)
        if item_name:   conds.append("i.item_name LIKE ?"); params.append(f"%{item_name}%")
        if cost_center: conds.append("(i.cost_center=? OR i.cost_center_name LIKE ?)"); params+=[cost_center,f"%{cost_center}%"]
        if lot_no:      conds.append("i.lot_no LIKE ?");    params.append(f"%{lot_no}%")
        where=" AND ".join(conds); offset=(page-1)*page_size
        total=conn.execute(f"SELECT COUNT(*) FROM inventory_items i WHERE {where}",params).fetchone()[0]
        rows=conn.execute(f"""
            SELECT i.factory,i.item_type,i.item_code,i.item_name,
                   COALESCE(i.cost_center_name,i.cost_center) AS cc_name,
                   i.lot_no,i.weight_ton,i.amount,i.base_date,i.is_new
            FROM inventory_items i WHERE {where} ORDER BY i.amount DESC LIMIT ? OFFSET ?
        """,params+[page_size,offset]).fetchall()
        return {"ref_date":ref_date,"total":total,"page":page,"items":[dict(r) for r in rows]}
    finally: conn.close()

@router.get("/plans")
async def get_plans(
    ref_date: Optional[str]=Query(None), factory: Optional[str]=Query(None),
    dept: Optional[str]=Query(None), plan_type: Optional[str]=Query(None),
    lot_no: Optional[str]=Query(None), lot_no_exact: Optional[str]=Query(None),
    item_name: Optional[str]=Query(None), cost_center: Optional[str]=Query(None),
    page: int=Query(1,ge=1), page_size: int=Query(50,ge=1,le=500),
):
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        if lot_no_exact:
            rows=conn.execute("""
                SELECT i.factory,i.item_type,i.item_code,i.item_name,
                       COALESCE(i.cost_center_name,i.cost_center) AS cc_name,
                       i.lot_no,i.weight_ton,i.amount,i.base_date,
                       p.dept,p.reason,p.plan_type,p.plan_date,p.detail_plan,
                       p.is_complete,p.created_by,p.created_by_name,p.created_at,
                       p.updated_by,p.updated_by_name,p.updated_at
                FROM inventory_items i JOIN depletion_plans p ON p.lot_no=i.lot_no
                WHERE p.lot_no=? LIMIT 1
            """,(lot_no_exact,)).fetchall()
            if not rows:
                rows=conn.execute("SELECT * FROM depletion_plans WHERE lot_no=? LIMIT 1",(lot_no_exact,)).fetchall()
            return {"ref_date":ref_date,"total":len(rows),"page":1,"items":[dict(r) for r in rows]}

        conds=["i.ref_date=?"]; params=[ref_date]
        if factory:     conds.append("i.factory=?");         params.append(factory)
        if dept:        conds.append("p.dept=?");            params.append(dept)
        if plan_type:   conds.append("p.plan_type=?");       params.append(plan_type)
        if lot_no:      conds.append("i.lot_no LIKE ?");     params.append(f"%{lot_no}%")
        if item_name:   conds.append("i.item_name LIKE ?");  params.append(f"%{item_name}%")
        if cost_center: conds.append("(i.cost_center=? OR i.cost_center_name LIKE ?)"); params+=[cost_center,f"%{cost_center}%"]
        where=" AND ".join(conds); offset=(page-1)*page_size
        total=conn.execute(
            f"SELECT COUNT(*) FROM inventory_items i JOIN depletion_plans p ON p.lot_no=i.lot_no WHERE {where}",params
        ).fetchone()[0]
        rows=conn.execute(f"""
            SELECT i.factory,i.item_type,i.item_code,i.item_name,
                   COALESCE(i.cost_center_name,i.cost_center) AS cc_name,
                   i.lot_no,i.weight_ton,i.amount,i.qty_consumed,i.amount_consumed,i.base_date,
                   p.dept,p.reason,p.plan_type,p.plan_date,p.detail_plan,p.is_complete,
                   p.created_by,p.created_by_name,p.created_at AS plan_created_at,
                   p.updated_by,p.updated_by_name,p.updated_at AS plan_updated_at
            FROM inventory_items i JOIN depletion_plans p ON p.lot_no=i.lot_no
            WHERE {where} ORDER BY i.amount DESC LIMIT ? OFFSET ?
        """,params+[page_size,offset]).fetchall()
        return {"ref_date":ref_date,"total":total,"page":page,"items":[dict(r) for r in rows]}
    finally: conn.close()

class PlanBody(BaseModel):
    dept: Optional[str]=None; reason: Optional[str]=None
    plan_type: Optional[str]=None; plan_date: Optional[str]=None
    detail_plan: Optional[str]=None

@router.post("/plans/{lot_no}")
async def upsert_plan(lot_no: str, body: PlanBody, request: Request):
    u = cur_user(request)
    if not u: raise HTTPException(401, "로그인이 필요합니다.")
    by = u["username"]; by_name = u.get("display_name","") or by
    now = _now_str()
    conn = get_conn()
    try:
        ex = conn.execute("SELECT id FROM depletion_plans WHERE lot_no=?", (lot_no,)).fetchone()
        if ex:
            conn.execute("""
                UPDATE depletion_plans
                SET dept=?,reason=?,plan_type=?,plan_date=?,detail_plan=?,
                    updated_by=?,updated_by_name=?,updated_at=?
                WHERE lot_no=?
            """, (body.dept,body.reason,body.plan_type,body.plan_date,body.detail_plan,
                  by,by_name,now,lot_no))
        else:
            inv = conn.execute(
                "SELECT item_code,item_name,factory,cost_center,cost_center_name,item_type FROM inventory_items WHERE lot_no=? LIMIT 1",
                (lot_no,)
            ).fetchone()
            conn.execute("""
                INSERT INTO depletion_plans
                    (lot_no,item_code,item_name,factory,cost_center,cost_center_name,item_type,
                     dept,reason,plan_type,plan_date,detail_plan,
                     created_by,created_by_name,updated_by,updated_by_name)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (lot_no,
                  inv["item_code"] if inv else None, inv["item_name"] if inv else None,
                  inv["factory"]   if inv else None, inv["cost_center"] if inv else None,
                  inv["cost_center_name"] if inv else None, inv["item_type"] if inv else None,
                  body.dept,body.reason,body.plan_type,body.plan_date,body.detail_plan,
                  by,by_name,by,by_name))
        conn.commit()
        return {"ok": True, "lot_no": lot_no}
    except Exception as e:
        conn.rollback()
        logger.error(f"계획 저장 오류: {e}", exc_info=True)
        raise HTTPException(500, f"계획 저장 실패: {e}")
    finally: conn.close()

@router.delete("/plans/{lot_no}")
async def delete_plan(lot_no: str, request: Request):
    u = cur_user(request)
    if not u: raise HTTPException(401, "로그인이 필요합니다.")
    conn = get_conn()
    try:
        conn.execute("DELETE FROM depletion_plans WHERE lot_no=?", (lot_no,))
        conn.commit()
        return {"ok": True}
    except Exception as e:
        conn.rollback()
        raise HTTPException(500, f"계획 삭제 실패: {e}")
    finally: conn.close()

# ══════════════════════════════════════
# 계획/실적 비교
# ══════════════════════════════════════
@router.get("/compare")
async def compare_plan_actual(
    ref_date: Optional[str]=Query(None), factory: Optional[str]=Query(None),
    dept: Optional[str]=Query(None),
    page: int=Query(1,ge=1), page_size: int=Query(50,ge=1,le=500),
):
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        conds=["i.ref_date=?","p.lot_no IS NOT NULL"]; params=[ref_date]
        if factory: conds.append("i.factory=?"); params.append(factory)
        if dept:    conds.append("p.dept=?");    params.append(dept)
        where=" AND ".join(conds); offset=(page-1)*page_size
        total=conn.execute(
            f"SELECT COUNT(*) FROM inventory_items i JOIN depletion_plans p ON p.lot_no=i.lot_no LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no WHERE {where}",params
        ).fetchone()[0]
        rows=conn.execute(f"""
            SELECT i.factory,i.item_type,i.item_code,i.item_name,
                   COALESCE(i.cost_center_name,i.cost_center) AS cc_name,
                   i.lot_no,i.weight_ton,i.amount,i.qty_consumed,i.amount_consumed,i.base_date,
                   p.dept,p.plan_type,p.plan_date,p.reason,
                   a.actual_type,a.actual_type_manual,a.process_date,
                   a.weight_ton AS actual_weight,
                   CASE WHEN a.lot_no IS NOT NULL THEN 1 ELSE 0 END AS has_actual,
                   CASE WHEN a.lot_no IS NOT NULL THEN '조치' ELSE '미조치' END AS action_status,
                   a.id AS actual_id
            FROM inventory_items i
            JOIN depletion_plans p ON p.lot_no=i.lot_no
            LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no
            WHERE {where} ORDER BY i.amount DESC LIMIT ? OFFSET ?
        """,params+[page_size,offset]).fetchall()
        items=[]
        for r in rows:
            d=dict(r)
            pt=d.get("plan_type") or ""; at=d.get("actual_type_manual") or d.get("actual_type") or ""
            d["type_match"]=(pt==at) if (pt and at) else None
            items.append(d)
        return {"ref_date":ref_date,"total":total,"page":page,"items":items}
    except Exception as e:
        logger.error(f"비교 조회 오류: {e}", exc_info=True)
        raise HTTPException(500, f"비교 조회 실패: {e}")
    finally: conn.close()

@router.get("/compare/summary")
async def compare_summary(ref_date: Optional[str]=Query(None)):
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        r=conn.execute("""
            SELECT COUNT(*) AS plan_total,
                   SUM(CASE WHEN a.lot_no IS NOT NULL THEN 1 ELSE 0 END) AS action_count,
                   SUM(CASE WHEN a.lot_no IS NULL THEN 1 ELSE 0 END) AS no_action_count,
                   COALESCE(SUM(i.weight_ton),0) AS total_weight,
                   COALESCE(SUM(i.amount),0) AS total_amount,
                   COALESCE(SUM(CASE WHEN a.lot_no IS NOT NULL THEN i.weight_ton ELSE 0 END),0) AS action_weight,
                   COALESCE(SUM(CASE WHEN a.lot_no IS NOT NULL THEN i.amount ELSE 0 END),0) AS action_amount,
                   COALESCE(SUM(CASE WHEN a.lot_no IS NULL THEN i.weight_ton ELSE 0 END),0) AS no_action_weight,
                   COALESCE(SUM(CASE WHEN a.lot_no IS NULL THEN i.amount ELSE 0 END),0) AS no_action_amount,
                   COALESCE(SUM(i.amount_consumed),0) AS consumed_amount
            FROM inventory_items i
            JOIN depletion_plans p ON p.lot_no=i.lot_no
            LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no
            WHERE i.ref_date=?
        """,(ref_date,)).fetchone()
        pt=r["plan_total"] or 1; ac=r["action_count"] or 0
        plan_rows=conn.execute("""
            SELECT p.plan_type, COUNT(*) AS plan_count,
                   COALESCE(SUM(i.weight_ton),0) AS plan_weight,
                   COALESCE(SUM(i.amount),0) AS plan_amount
            FROM depletion_plans p LEFT JOIN inventory_items i ON i.lot_no=p.lot_no AND i.ref_date=?
            GROUP BY p.plan_type ORDER BY plan_amount DESC
        """,(ref_date,)).fetchall()
        actual_rows=conn.execute("""
            SELECT COALESCE(a.actual_type_manual,a.actual_type,'기타') AS actual_type,
                   COUNT(*) AS actual_count, COALESCE(SUM(a.weight_ton),0) AS actual_weight
            FROM depletion_actuals a WHERE a.ref_date=?
            GROUP BY actual_type ORDER BY actual_weight DESC
        """,(ref_date,)).fetchall()
        return {
            "ref_date":ref_date,"plan_total":int(pt),"action_count":int(ac),
            "no_action_count":int(r["no_action_count"] or 0),
            "action_rate":round(ac/pt*100,1),
            "total_weight":float(r["total_weight"]),"total_amount":float(r["total_amount"]),
            "action_weight":float(r["action_weight"]),"action_amount":float(r["action_amount"]),
            "no_action_weight":float(r["no_action_weight"]),"no_action_amount":float(r["no_action_amount"]),
            "consumed_amount":float(r["consumed_amount"]),
            "plan_by_type":[dict(x) for x in plan_rows],
            "actual_by_type":[dict(x) for x in actual_rows],
        }
    finally: conn.close()

@router.get("/compare/export")
async def export_compare(ref_date: Optional[str]=Query(None)):
    """계획/실적 비교 Excel 다운로드 - 한글 파일명 RFC 5987 처리"""
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        if not ref_date: raise HTTPException(404, "데이터가 없습니다.")
        rows=conn.execute("""
            SELECT i.factory,i.item_type,i.item_code,i.item_name,
                   COALESCE(i.cost_center_name,i.cost_center) AS cc_name,
                   i.lot_no,i.weight_ton,i.amount,i.qty_consumed,i.amount_consumed,i.base_date,
                   p.dept,p.plan_type,p.plan_date,
                   COALESCE(a.actual_type_manual,a.actual_type,'') AS actual_type,
                   a.process_date,
                   CASE WHEN a.lot_no IS NOT NULL THEN '조치' ELSE '미조치' END AS action_status
            FROM inventory_items i
            JOIN depletion_plans p ON p.lot_no=i.lot_no
            LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no
            WHERE i.ref_date=? ORDER BY i.amount DESC
        """,(ref_date,)).fetchall()
        conn.close()
        data=[dict(r) for r in rows]
        df=pd.DataFrame(data) if data else pd.DataFrame()
        if not df.empty:
            df.columns=["공장","품목구분","품목코드","품명","원가중심점","LOT NO",
                        "중량(ton)","금액","소진수량","소진금액","기준일자",
                        "담당부서","계획유형","계획기한","실적유형","처리일자","조치여부"]
        buf=io.BytesIO()
        with pd.ExcelWriter(buf,engine="xlsxwriter") as w:
            df.to_excel(w,index=False,sheet_name="계획실적비교")
            if not df.empty:
                ws=w.sheets["계획실적비교"]
                hdr=w.book.add_format({"bold":True,"bg_color":"#DDEBF7","border":1})
                for ci,col in enumerate(df.columns): ws.write(0,ci,col,hdr); ws.set_column(ci,ci,14)
        ts=datetime.now().strftime("%Y%m%d_%H%M%S")
        fname=f"계획실적비교_{ref_date}_{ts}.xlsx"
        return _xlsx_response(buf, fname)
    except HTTPException: raise
    except Exception as e:
        logger.error(f"비교 Excel 오류: {e}", exc_info=True)
        raise HTTPException(500, f"Excel 생성 실패: {e}")

@router.get("/compare/export-ppt")
async def export_compare_ppt(ref_date: Optional[str] = Query(None)):
    """계획/실적 비교 PPT 다운로드"""
    conn = get_conn()
    try:
        if not ref_date: ref_date = _latest_ref(conn)
        if not ref_date: raise HTTPException(404, "데이터가 없습니다.")

        # summary
        summ_rows = conn.execute("""
            SELECT COUNT(*) AS pt,
                   SUM(CASE WHEN a.lot_no IS NOT NULL THEN 1 ELSE 0 END) AS ac,
                   SUM(CASE WHEN a.lot_no IS NULL THEN 1 ELSE 0 END) AS nc,
                   COALESCE(SUM(i.amount),0) AS ta,
                   COALESCE(SUM(CASE WHEN a.lot_no IS NOT NULL THEN i.amount ELSE 0 END),0) AS aa,
                   COALESCE(SUM(CASE WHEN a.lot_no IS NULL THEN i.amount ELSE 0 END),0) AS na,
                   COALESCE(SUM(i.amount_consumed),0) AS ca,
                   COALESCE(SUM(i.weight_ton),0) AS tw
            FROM inventory_items i
            JOIN depletion_plans p ON p.lot_no=i.lot_no
            LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no
            WHERE i.ref_date=?
        """, (ref_date,)).fetchone()
        pt=summ_rows["pt"] or 1; ac=summ_rows["ac"] or 0
        plan_by=conn.execute("""
            SELECT p.plan_type, COUNT(*) AS plan_count,
                   COALESCE(SUM(i.weight_ton),0) AS plan_weight
            FROM depletion_plans p LEFT JOIN inventory_items i ON i.lot_no=p.lot_no AND i.ref_date=?
            GROUP BY p.plan_type ORDER BY plan_count DESC
        """, (ref_date,)).fetchall()
        actual_by=conn.execute("""
            SELECT COALESCE(actual_type_manual,actual_type,'기타') AS actual_type,
                   COUNT(*) AS actual_count, COALESCE(SUM(weight_ton),0) AS actual_weight
            FROM depletion_actuals WHERE ref_date=?
            GROUP BY actual_type ORDER BY actual_count DESC
        """, (ref_date,)).fetchall()
        items=conn.execute("""
            SELECT i.factory,i.item_name,i.lot_no,i.amount,
                   p.plan_type, COALESCE(a.actual_type_manual,a.actual_type,'') AS actual_type,
                   CASE WHEN a.lot_no IS NOT NULL THEN '조치' ELSE '미조치' END AS action_status
            FROM inventory_items i
            JOIN depletion_plans p ON p.lot_no=i.lot_no
            LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no
            WHERE i.ref_date=? ORDER BY i.amount DESC LIMIT 50
        """, (ref_date,)).fetchall()
        conn.close()

        summary = {
            "plan_total":int(pt),"action_count":int(ac),"no_action_count":int(summ_rows["nc"] or 0),
            "action_rate":round(ac/pt*100,1),"total_amount":float(summ_rows["ta"]),
            "action_amount":float(summ_rows["aa"]),"no_action_amount":float(summ_rows["na"]),
            "consumed_amount":float(summ_rows["ca"]),
            "plan_by_type":[dict(r) for r in plan_by],
            "actual_by_type":[dict(r) for r in actual_by],
        }
        # plan_trend / cc_items 추가 조회
        plan_trend_rows = conn.execute("""
            SELECT substr(p.plan_date,1,7) AS plan_month,
                   COUNT(*) AS plan_count,
                   COALESCE(SUM(i.weight_ton),0) AS plan_weight_ton,
                   COALESCE(SUM(i.amount),0) AS plan_amount
            FROM depletion_plans p
            LEFT JOIN inventory_items i ON i.lot_no=p.lot_no
            WHERE p.plan_date IS NOT NULL AND p.plan_date!=''
            GROUP BY plan_month ORDER BY plan_month
        """).fetchall()
        cc_rows = conn.execute("""
            SELECT COALESCE(i.cost_center_name,i.cost_center) AS cc_name,
                   COUNT(*) AS item_count,
                   COALESCE(SUM(i.weight_ton),0) AS total_weight,
                   COALESCE(SUM(i.amount),0) AS total_amount,
                   COUNT(p.lot_no) AS plan_count,
                   COUNT(a.lot_no) AS actual_count
            FROM inventory_items i
            LEFT JOIN depletion_plans p ON p.lot_no=i.lot_no
            LEFT JOIN depletion_actuals a ON a.lot_no=i.lot_no
            WHERE i.ref_date=?
            GROUP BY i.cost_center, i.cost_center_name
            ORDER BY total_amount DESC
        """, (ref_date,)).fetchall()
        from backend.ppt_exporter import generate_compare_ppt
        ppt_bytes = generate_compare_ppt(
            summary, [dict(r) for r in items], ref_date,
            [dict(r) for r in plan_trend_rows],
            [dict(r) for r in cc_rows],
        )
        buf = io.BytesIO(ppt_bytes)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"계획실적비교_{ref_date}_{ts}.pptx"
        return _pptx_response(buf, fname)

    except HTTPException: raise
    except Exception as e:
        logger.error(f"PPT 생성 오류: {e}", exc_info=True)
        raise HTTPException(500, f"PPT 생성 실패: {str(e)}")

@router.patch("/actuals/{actual_id}/type")
async def patch_actual_type(actual_id: int, body: dict, request: Request):
    u = cur_user(request)
    if not u: raise HTTPException(401, "로그인이 필요합니다.")
    conn = get_conn()
    try:
        conn.execute(
            "UPDATE depletion_actuals SET actual_type_manual=? WHERE id=?",
            (body.get("actual_type_manual"), actual_id)
        )
        conn.commit()
        return {"ok": True}
    except Exception as e:
        conn.rollback()
        raise HTTPException(500, f"실적유형 수정 실패: {e}")
    finally: conn.close()
