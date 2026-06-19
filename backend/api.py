"""
api.py v6 - 전체 REST API
[v6 핵심 수정사항 - 사용자 요구사항 반영]
 1. 금액 포맷: 원단위는 소수점 반올림 표시 안함, 백만원/억원은 소수점 둘째자리. 기본 표기는 억원.
 2. 당월소진예정금액: 소진계획기한(plan_date)이 "시스템 오늘 날짜" 기준 당월인 것만 계산 (그대로 유지, 명확화)
 3. 전월대비 소진금액: 재고+재공 시트 기준, 해당월 vs 전월 LOT 비교(전월 LOT가 당월에 사라지거나 감소한 만큼)
 4. 소진완료금액: LOT별 단가(=재고/재공 시트의 amount/weight_ton)를 구하고,
    상세시트(실적)의 weight_ton에 단가를 곱해 산출. 결과는 항상 양수로 표시.
 5. JOIN 중복버그 수정: depletion_plans/actuals와 JOIN하기 전에 inventory_items를 LOT 단위로
    먼저 GROUP BY 집계(서브쿼리)하여 1:1 매칭 보장.
 6. '미조치(계획미등록)' -> '당월계획분 미조치'로 명칭 변경.
    계산: 소진계획기한이 시스템 당월인 LOT 중, 아직 실적(소진)이 확인되지 않은 금액.
 7. 저장품(item_type='저장품')은 모든 계산/조회에서 제외 (WHERE i.item_type != '저장품' 강제)
 8. 대시보드 KPI에 금액/중량/건수/전월 비교를 모두 포함, 중복 카드(총금액/조치금액/소진금액 비교카드) 제거.

기준 원칙: 계획은 시스템에 입력된 depletion_plans, 실적은 업로드 파일의 상세시트(depletion_actuals)에서
조회기준일(ref_date) 기준으로 산출됨. 계획과 실적은 항상 LOT_NO로 매칭한다.
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

# 저장품 제외 조건 (모든 inventory_items 조회에 일관 적용)
EXCLUDE_STORAGE = "i.item_type != '저장품'"


# ══════════════════════════════════════════════
# 금액/중량 포맷 헬퍼
# ══════════════════════════════════════════════
def fmt_amount(value: float, unit: str = "HM") -> dict:
    """
    금액 포맷 통일 함수.
    unit: 'KRW'(원) | 'MN'(백만원) | 'HM'(억원)
    - 원단위: 소수점 반올림, 정수만 표시
    - 백만원/억원: 소수점 둘째자리까지 (일반 반올림 규칙, 0.5는 올림)
    반환: {value: 표시용 숫자(round 처리됨), raw: 원본 float, unit: unit}
    """
    import decimal
    def _round(n: float, ndigits: int) -> float:
        # Python 기본 round()는 banker's rounding(0.5를 짝수로) 이라 회계 표기와 다를 수 있음.
        # ROUND_HALF_UP으로 고정하여 "사사오입" 방식의 일반적인 반올림을 보장한다.
        q = decimal.Decimal(10) ** -ndigits
        return float(decimal.Decimal(str(n)).quantize(q, rounding=decimal.ROUND_HALF_UP))

    v = float(value or 0)
    if unit == "KRW":
        return {"value": int(_round(v, 0)), "raw": v, "unit": "KRW"}
    if unit == "MN":
        return {"value": _round(v / 1_000_000, 2), "raw": v, "unit": "MN"}
    # 기본: 억원
    return {"value": _round(v / 100_000_000, 2), "raw": v, "unit": "HM"}


def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _latest_ref(conn) -> str:
    row = conn.execute(
        f"SELECT ref_date FROM inventory_items i WHERE {EXCLUDE_STORAGE} ORDER BY ref_date DESC LIMIT 1"
    ).fetchone()
    return row["ref_date"] if row else ""


def _prev_ref(conn, ref_date: str, mode: str = "month") -> Optional[str]:
    """ref_date 이전의 가장 가까운 스냅샷 ref_date를 찾는다 (month/quarter/year 모드)."""
    rows = conn.execute(
        f"SELECT DISTINCT ref_date FROM inventory_items i WHERE {EXCLUDE_STORAGE} AND ref_date < ? ORDER BY ref_date DESC",
        (ref_date,)
    ).fetchall()
    dates = [r["ref_date"] for r in rows]
    if not dates:
        return None
    if mode == "month":
        return dates[0]
    if mode == "quarter":
        for d in dates:
            if len(d) == 8 and int(d[:6]) <= int(ref_date[:6]) - 3:
                return d
        return dates[-1]
    if mode == "year":
        for d in dates:
            if len(d) == 8 and int(d[:6]) <= int(ref_date[:6]) - 12:
                return d
        return dates[-1]
    return dates[0]


def _safe_fname(name: str) -> str:
    return urllib.parse.quote(name, safe="")


def _xlsx_response(buf: io.BytesIO, filename: str) -> StreamingResponse:
    buf.seek(0)
    encoded = _safe_fname(filename)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"}
    )


def _pptx_response(buf: io.BytesIO, filename: str) -> StreamingResponse:
    buf.seek(0)
    encoded = _safe_fname(filename)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded}"}
    )


# ══════════════════════════════════════════════
# 인증
# ══════════════════════════════════════════════
def cur_user(request: Request) -> Optional[dict]:
    return request.session.get("user")


def require_role(request: Request, role: str = "user") -> dict:
    u = cur_user(request)
    if not u:
        raise HTTPException(401, "로그인이 필요합니다.")
    if role == "admin" and u.get("role") != "admin":
        raise HTTPException(403, "관리자 권한이 필요합니다.")
    return u


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
            k: user[k] for k in ("id", "username", "display_name", "role", "department")
            if k in dict(user)
        }
        return {"ok": True, "user": request.session["user"]}
    except HTTPException:
        raise
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


class SettingBody(BaseModel):
    key: str
    value: str


@router.get("/settings")
async def get_settings():
    conn = get_conn()
    try:
        rows = conn.execute("SELECT key,value FROM settings").fetchall()
        return {r["key"]: r["value"] for r in rows}
    finally:
        conn.close()


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
    finally:
        conn.close()


# ══════════════════════════════════════════════
# 사용자 관리 (admin only)
# ══════════════════════════════════════════════
class UserCreateBody(BaseModel):
    username: str
    password: str
    display_name: str
    role: str = "user"
    department: Optional[str] = None


class UserUpdateBody(BaseModel):
    display_name: Optional[str] = None
    role: Optional[str] = None
    department: Optional[str] = None
    is_active: Optional[int] = None
    password: Optional[str] = None


@router.get("/users")
async def list_users(request: Request):
    require_role(request, "admin")
    conn = get_conn()
    try:
        rows = conn.execute(
            "SELECT id,username,display_name,role,department,is_active,last_login,created_at FROM users ORDER BY id"
        ).fetchall()
        return {"users": [dict(r) for r in rows]}
    finally:
        conn.close()


@router.post("/users")
async def create_user(body: UserCreateBody, request: Request):
    u = require_role(request, "admin")
    if body.role not in ROLES:
        raise HTTPException(400, f"유효하지 않은 권한: {body.role}")
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
    except HTTPException:
        raise
    except Exception as e:
        conn.rollback()
        logger.error(f"사용자 생성 오류: {e}", exc_info=True)
        raise HTTPException(500, f"사용자 생성 실패: {e}")
    finally:
        conn.close()


@router.put("/users/{user_id}")
async def update_user(user_id: int, body: UserUpdateBody, request: Request):
    u = require_role(request, "admin")
    conn = get_conn()
    try:
        row = conn.execute("SELECT * FROM users WHERE id=?", (user_id,)).fetchone()
        if not row:
            raise HTTPException(404, f"사용자 없음: id={user_id}")
        sets, params = [], []
        if body.display_name is not None:
            sets.append("display_name=?"); params.append(body.display_name)
        if body.role is not None:
            if body.role not in ROLES:
                raise HTTPException(400, f"유효하지 않은 권한: {body.role}")
            sets.append("role=?"); params.append(body.role)
        if body.department is not None:
            sets.append("department=?"); params.append(body.department)
        if body.is_active is not None:
            sets.append("is_active=?"); params.append(body.is_active)
        if body.password is not None:
            sets.append("password_hash=?"); params.append(_hash_pw(body.password))
        if not sets:
            return {"ok": True, "message": "변경 없음"}
        params.append(user_id)
        conn.execute(f"UPDATE users SET {','.join(sets)} WHERE id=?", params)
        conn.commit()
        return {"ok": True, "message": "사용자 정보 수정 완료"}
    except HTTPException:
        raise
    except Exception as e:
        conn.rollback()
        logger.error(f"사용자 수정 오류: {e}", exc_info=True)
        raise HTTPException(500, f"사용자 수정 실패: {e}")
    finally:
        conn.close()


# ══════════════════════════════════════════════
# 기준일 목록
# ══════════════════════════════════════════════
@router.get("/inventory/ref-dates")
async def ref_dates():
    conn = get_conn()
    try:
        rows = conn.execute(
            f"SELECT DISTINCT ref_date FROM inventory_items i WHERE {EXCLUDE_STORAGE} ORDER BY ref_date DESC"
        ).fetchall()
        return {"ref_dates": [r["ref_date"] for r in rows]}
    finally:
        conn.close()


@router.get("/inventory/filter-options")
async def inventory_filter_options(ref_date: Optional[str] = Query(None)):
    """대시보드/조회/비교 화면의 공장·원가중심점 드롭다운에 사용할 목록."""
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        factories = conn.execute(
            f"SELECT DISTINCT factory FROM inventory_items i WHERE {EXCLUDE_STORAGE} AND ref_date=? "
            f"AND factory IS NOT NULL AND factory != '' ORDER BY factory",
            (ref_date,)
        ).fetchall()
        cost_centers = conn.execute(
            f"SELECT DISTINCT cost_center, cost_center_name FROM inventory_items i WHERE {EXCLUDE_STORAGE} AND ref_date=? "
            f"AND cost_center_name IS NOT NULL AND cost_center_name != '' ORDER BY cost_center_name",
            (ref_date,)
        ).fetchall()
        return {
            "factories": [r["factory"] for r in factories],
            "cost_centers": [{"code": r["cost_center"], "name": r["cost_center_name"]} for r in cost_centers],
        }
    finally:
        conn.close()


# ══════════════════════════════════════════════
# 업로드
# ══════════════════════════════════════════════
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
            "all_ref_dates": parsed["all_ref_dates"],
            "inv_count": result["inv_count"], "wip_count": result["wip_count"],
            "act_count": result["act_count"], "total_amount": result["total_amount"],
            "excluded_count": parsed.get("excluded_count", 0),
            "warnings": parsed.get("warnings", []),
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"파일 업로드 오류: {e}", exc_info=True)
        raise HTTPException(500, f"업로드 처리 실패: {e}")


@router.get("/upload-history")
async def upload_history():
    conn = get_conn()
    try:
        rows = conn.execute("SELECT * FROM upload_history ORDER BY created_at DESC LIMIT 50").fetchall()
        return {"history": [dict(r) for r in rows]}
    finally:
        conn.close()


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
    except HTTPException:
        raise
    except Exception as e:
        conn.rollback()
        logger.error(f"업로드 삭제 오류: {e}", exc_info=True)
        raise HTTPException(500, f"삭제 실패: {e}")
    finally:
        conn.close()


@router.delete("/upload/all/data")
async def delete_all(request: Request):
    require_role(request, "admin")
    conn = get_conn()
    try:
        for t in ["inventory_items", "depletion_actuals", "upload_history"]:
            conn.execute(f"DELETE FROM {t}")
        conn.commit()
        return {"ok": True}
    except Exception as e:
        conn.rollback()
        logger.error(f"전체 삭제 오류: {e}", exc_info=True)
        raise HTTPException(500, f"전체 삭제 실패: {e}")
    finally:
        conn.close()


# ══════════════════════════════════════════════
# 대시보드
# ══════════════════════════════════════════════
#
# [계산 구조 핵심 설명]
#
# ① 총 장기재고 금액/중량/건수
#    -> 현재 ref_date 기준 inventory_items 합계 (저장품 제외)
#
# ② 당월소진예정금액 (요구사항 2)
#    -> depletion_plans.plan_date 의 연-월이 "시스템 오늘 날짜의 연-월"과 같은 LOT들의
#       현재 ref_date 기준 재고금액 합계.
#    -> "당월"은 재고 조회기준일이 아니라 캘린더상 오늘 날짜를 의미함 (사용자 확인사항)
#
# ③ 당월계획분 미조치 (요구사항 6, 구 "미조치(계획미등록)")
#    -> 당월(②와 동일 기준)이 계획기한인 LOT 중에서, 아직 실적(소진)이 확인되지 않은 금액.
#       "실적 확인됨"의 기준은 depletion_actuals에 해당 LOT가 존재하는지 여부.
#
# ④ 소진완료금액 (요구사항 4)
#    -> LOT 단위로 inventory_items에서 "단가 = 합산금액 / 합산중량(ton)"을 구하고,
#       depletion_actuals(상세시트, 이미 절대값으로 저장됨)의 LOT별 weight_ton 합계에
#       그 단가를 곱해서 산출. 중량이 0인 LOT(단가 계산 불가)는 0원으로 처리됨.
#       항상 양수로 표시 (요구사항: 마이너스=소진의미, 화면엔 양수로 표시)
#
# ⑤ 전월대비 소진금액 (요구사항 3)
#    -> inventory_items(재고+재공)만 사용. 전월 ref_date에 존재했던 LOT가
#       당월 ref_date에서 사라졌거나 금액/중량이 감소한 만큼을 합산.
#       계획/실적 테이블과 무관하게 순수 재고 스냅샷 비교로만 계산.
#
# [JOIN 중복버그 수정 - 요구사항 5]
#    inventory_items에는 동일 LOT_NO가 여러 자재 행에 걸쳐 나타날 수 있음(저장품 제외해도 발생).
#    depletion_plans/actuals와 JOIN하기 전에 반드시 LOT 단위로 먼저 GROUP BY 집계한
#    서브쿼리(lot_agg)를 만들어 1:1 매칭을 보장한다. 절대 inventory_items를 직접 JOIN하지 않는다.

def _lot_agg_subquery(ref_date_param_placeholder: str = "?", factory: Optional[str] = None,
                       cost_center: Optional[str] = None) -> str:
    """LOT 단위로 먼저 집계하는 서브쿼리. 저장품 제외 적용됨.
    SUM 결과가 NULL이 되는 경우(그룹 내 모든 값이 NULL)를 방지하기 위해 COALESCE 적용.

    [중요] factory/cost_center 필터를 켜면, 이 서브쿼리를 사용하는 SQL의 파라미터 바인딩 순서는
    반드시 (ref_date, [factory], [cost_center], ...나머지) 순서를 지켜야 한다.
    파라미터 개수가 동적으로 바뀌므로, 호출부는 _lot_agg_params() 헬퍼로 짝을 맞춰 사용한다.
    """
    extra = ""
    if factory:
        extra += " AND factory = ?"
    if cost_center:
        extra += " AND (cost_center = ? OR cost_center_name LIKE ?)"
    return f"""
        SELECT lot_no,
               MIN(factory) AS factory, MIN(item_type) AS item_type,
               MIN(item_code) AS item_code, MIN(item_name) AS item_name,
               MIN(cost_center) AS cost_center, MIN(cost_center_name) AS cost_center_name,
               MIN(base_date) AS base_date, MIN(months_label) AS months_label,
               MAX(is_new) AS is_new,
               COALESCE(SUM(amount),0) AS amount,
               COALESCE(SUM(weight_ton),0) AS weight_ton,
               COALESCE(SUM(qty),0) AS qty
        FROM inventory_items i
        WHERE {EXCLUDE_STORAGE} AND ref_date = {ref_date_param_placeholder}{extra}
        GROUP BY lot_no
    """


def _lot_agg_params(ref_date, factory: Optional[str] = None, cost_center: Optional[str] = None) -> list:
    """_lot_agg_subquery와 짝을 맞추는 파라미터 리스트. ref_date 뒤에 factory/cost_center를 순서대로 추가."""
    params = [ref_date]
    if factory:
        params.append(factory)
    if cost_center:
        params.append(cost_center)
        params.append(f"%{cost_center}%")
    return params


def _actual_type_to_plan_type(actual_type_raw: Optional[str]) -> Optional[str]:
    """실적유형(상세시트의 '유형' 컬럼) 문자열을 소진계획방안 카테고리로 매핑.
    [사용자 확정 규칙]
      - 'Sales'로 시작 -> '전환 판매' 실적
      - 'WIP'로 시작   -> '생산투입' 실적
      - 그 외(Account alias issue, Direct Org Transfer 등) -> 매핑 없음(None, 기타로 집계)
    대소문자 구분 없이 비교한다.
    """
    if not actual_type_raw:
        return None
    s = str(actual_type_raw).strip().lower()
    if s.startswith("sales"):
        return "전환 판매"
    if s.startswith("wip"):
        return "생산투입"
    return None


@router.get("/dashboard/kpi")
async def dashboard_kpi(
    ref_date: Optional[str] = Query(None),
    unit: str = Query("HM", description="KRW|MN|HM"),
    factory: Optional[str] = Query(None, description="공장 필터 (예: 임실공장). 비우면 전체"),
    cost_center: Optional[str] = Query(None, description="원가중심점 필터(부분일치). 비우면 전체"),
):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        if not ref_date:
            empty = {"amount": 0, "weight_ton": 0, "count": 0}
            return {
                "ref_date": "", "unit": unit,
                "total": {**empty, "prev_amount": 0},
                "plan_this_month": {**empty, "prev_amount": 0},
                "uncompleted_this_month": {**empty, "prev_amount": 0},
                "completed": {**empty, "prev_amount": 0},
                "consumed_mom": {**empty, "prev_amount": 0},
            }

        prev_rd = _prev_ref(conn, ref_date, "month")
        today_month = date.today().strftime("%Y-%m")  # 요구사항2,6: 시스템 오늘 날짜 기준 당월

        def _lot_sql(rd):
            return _lot_agg_subquery(factory=factory, cost_center=cost_center), _lot_agg_params(rd, factory, cost_center)

        # ── ① 총 장기재고 (LOT 집계 기준) ──
        def _total(rd):
            if not rd:
                return {"amount": 0.0, "weight_ton": 0.0, "count": 0}
            sql, params = _lot_sql(rd)
            r = conn.execute(f"""
                SELECT COALESCE(SUM(amount),0) AS ta, COALESCE(SUM(weight_ton),0) AS tw, COUNT(*) AS tc
                FROM ({sql}) lot_agg
            """, params).fetchone()
            return {"amount": float(r["ta"]), "weight_ton": float(r["tw"]), "count": int(r["tc"])}

        total_cur = _total(ref_date)
        total_prev = _total(prev_rd)

        # ── ② 당월소진예정금액: plan_date의 연월 = 오늘의 연월인 LOT의 현재 재고금액 ──
        def _plan_this_month(rd):
            if not rd:
                return {"amount": 0.0, "weight_ton": 0.0, "count": 0}
            sql, params = _lot_sql(rd)
            r = conn.execute(f"""
                SELECT COALESCE(SUM(la.amount),0) AS ta, COALESCE(SUM(la.weight_ton),0) AS tw, COUNT(*) AS tc
                FROM ({sql}) la
                JOIN depletion_plans p ON p.lot_no = la.lot_no
                WHERE substr(p.plan_date,1,7) = ?
            """, params + [today_month]).fetchone()
            return {"amount": float(r["ta"]), "weight_ton": float(r["tw"]), "count": int(r["tc"])}

        plan_this_cur = _plan_this_month(ref_date)
        plan_this_prev = _plan_this_month(prev_rd)

        # ── ③ 당월계획분 미조치: 당월계획 LOT 중 실적(소진) 미확인 금액 ──
        def _uncompleted_this_month(rd):
            if not rd:
                return {"amount": 0.0, "weight_ton": 0.0, "count": 0}
            sql, params = _lot_sql(rd)
            r = conn.execute(f"""
                SELECT COALESCE(SUM(la.amount),0) AS ta, COALESCE(SUM(la.weight_ton),0) AS tw, COUNT(*) AS tc
                FROM ({sql}) la
                JOIN depletion_plans p ON p.lot_no = la.lot_no
                WHERE substr(p.plan_date,1,7) = ?
                  AND NOT EXISTS (
                      SELECT 1 FROM depletion_actuals a
                      WHERE a.lot_no = la.lot_no AND a.ref_date = ?
                  )
            """, params + [today_month, rd]).fetchone()
            return {"amount": float(r["ta"]), "weight_ton": float(r["tw"]), "count": int(r["tc"])}

        uncompleted_cur = _uncompleted_this_month(ref_date)
        uncompleted_prev = _uncompleted_this_month(prev_rd)

        # ── ④ 소진완료금액: LOT단가 × 실적중량, 항상 양수 ──
        def _completed(rd):
            if not rd:
                return {"amount": 0.0, "weight_ton": 0.0, "count": 0}
            sql, params = _lot_sql(rd)
            # LOT별 단가 산출 (weight_ton > 0 인 경우만; amount/weight_ton 모두 COALESCE로 NULL 방지)
            unit_price_rows = conn.execute(f"""
                SELECT lot_no, COALESCE(amount,0) AS amount, COALESCE(weight_ton,0) AS weight_ton
                FROM ({sql}) la
            """, params).fetchall()
            unit_price = {
                row["lot_no"]: (row["amount"] / row["weight_ton"])
                for row in unit_price_rows
                if row["weight_ton"] is not None and row["weight_ton"] > 0
                and row["amount"] is not None
            }
            actual_rows = conn.execute("""
                SELECT lot_no, COALESCE(SUM(weight_ton),0) AS wt
                FROM depletion_actuals
                WHERE ref_date = ?
                GROUP BY lot_no
            """, (rd,)).fetchall()
            total_amt = 0.0
            total_wt = 0.0
            cnt = 0
            for a in actual_rows:
                up = unit_price.get(a["lot_no"])
                wt = a["wt"] or 0.0
                if up is not None:
                    amt = abs(wt) * up   # 항상 양수
                    total_amt += amt
                    total_wt += abs(wt)
                    cnt += 1
            return {"amount": total_amt, "weight_ton": total_wt, "count": cnt}

        completed_cur = _completed(ref_date)
        completed_prev = _completed(prev_rd)

        # ── ⑤ 전월대비 소진금액: 순수 재고 스냅샷 비교 (재고+재공만, 계획/실적 무관) ──
        def _consumed_mom(cur_rd, prv_rd):
            if not cur_rd or not prv_rd:
                return {"amount": 0.0, "weight_ton": 0.0, "count": 0}
            cur_sql, cur_params = _lot_sql(cur_rd)
            prv_sql, prv_params = _lot_sql(prv_rd)
            cur_lots = {
                row["lot_no"]: (row["amount"] or 0.0, row["weight_ton"] or 0.0)
                for row in conn.execute(
                    f"SELECT lot_no, COALESCE(amount,0) AS amount, COALESCE(weight_ton,0) AS weight_ton FROM ({cur_sql}) la",
                    cur_params
                ).fetchall()
            }
            prev_lots = {
                row["lot_no"]: (row["amount"] or 0.0, row["weight_ton"] or 0.0)
                for row in conn.execute(
                    f"SELECT lot_no, COALESCE(amount,0) AS amount, COALESCE(weight_ton,0) AS weight_ton FROM ({prv_sql}) la",
                    prv_params
                ).fetchall()
            }
            total_amt = 0.0
            total_wt = 0.0
            cnt = 0
            for lot_no, (p_amt, p_wt) in prev_lots.items():
                if lot_no not in cur_lots:
                    # 완전히 사라짐 = 전액 소진
                    total_amt += p_amt
                    total_wt += p_wt
                    cnt += 1
                else:
                    c_amt, c_wt = cur_lots[lot_no]
                    if p_amt > c_amt:
                        total_amt += (p_amt - c_amt)
                        total_wt += max(0.0, p_wt - c_wt)
                        cnt += 1
            return {"amount": total_amt, "weight_ton": total_wt, "count": cnt}

        consumed_cur = _consumed_mom(ref_date, prev_rd)
        prev_of_prev = _prev_ref(conn, prev_rd, "month") if prev_rd else None
        consumed_prev = _consumed_mom(prev_rd, prev_of_prev) if prev_rd else {"amount": 0.0, "weight_ton": 0.0, "count": 0}

        def _pack(cur: dict, prev: dict) -> dict:
            return {
                "amount": fmt_amount(cur["amount"], unit),
                "weight_ton": round(cur["weight_ton"], 3),
                "count": cur["count"],
                "prev_amount": fmt_amount(prev["amount"], unit),
                "prev_weight_ton": round(prev["weight_ton"], 3),
                "prev_count": prev["count"],
            }

        return {
            "ref_date": ref_date,
            "prev_ref_date": prev_rd,
            "unit": unit,
            "today_month": today_month,
            "factory": factory,
            "cost_center": cost_center,
            "total": _pack(total_cur, total_prev),
            "plan_this_month": _pack(plan_this_cur, plan_this_prev),
            "uncompleted_this_month": _pack(uncompleted_cur, uncompleted_prev),
            "completed": _pack(completed_cur, completed_prev),
            "consumed_mom": _pack(consumed_cur, consumed_prev),
        }
    except Exception as e:
        logger.error(f"KPI 조회 오류: {e}", exc_info=True)
        raise HTTPException(500, f"KPI 조회 실패: {e}")
    finally:
        conn.close()


@router.get("/dashboard/critical-stock")
async def critical_stock_summary(
    ref_date: Optional[str] = Query(None), unit: str = Query("HM"),
    factory: Optional[str] = Query(None), cost_center: Optional[str] = Query(None),
):
    """7개월이상 장기재고(months_label='7개월이상') 집중관리 요약."""
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        sql = _lot_agg_subquery(factory=factory, cost_center=cost_center)
        params = _lot_agg_params(ref_date, factory, cost_center)
        r = conn.execute(f"""
            SELECT COALESCE(SUM(amount),0) AS ta, COALESCE(SUM(weight_ton),0) AS tw, COUNT(*) AS tc
            FROM ({sql}) la
            WHERE months_label = '7개월이상'
        """, params).fetchone()
        return {
            "ref_date": ref_date,
            "amount": fmt_amount(r["ta"], unit)["value"],
            "weight_ton": round(r["tw"], 3),
            "count": r["tc"],
        }
    finally:
        conn.close()


@router.get("/dashboard/top20")
async def top20(
    ref_date: Optional[str] = Query(None),
    factory: Optional[str] = Query(None), cost_center: Optional[str] = Query(None),
):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        sql = _lot_agg_subquery(factory=factory, cost_center=cost_center)
        params = _lot_agg_params(ref_date, factory, cost_center)
        rows = conn.execute(f"""
            SELECT la.factory, la.item_type, la.item_code, la.item_name,
                   COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   la.lot_no, la.weight_ton, la.amount,
                   la.base_date, la.months_label, la.is_new,
                   p.plan_type, p.plan_date, p.dept,
                   CASE WHEN ax.lot_no IS NOT NULL THEN 1 ELSE 0 END AS is_completed
            FROM ({sql}) la
            LEFT JOIN depletion_plans p ON p.lot_no = la.lot_no
            LEFT JOIN (SELECT DISTINCT lot_no FROM depletion_actuals WHERE ref_date=?) ax
                   ON ax.lot_no = la.lot_no
            ORDER BY la.amount DESC LIMIT 20
        """, params + [ref_date]).fetchall()
        return {"ref_date": ref_date, "items": [dict(r) for r in rows]}
    except Exception as e:
        logger.error(f"TOP20 오류: {e}", exc_info=True)
        raise HTTPException(500, f"TOP20 조회 실패: {e}")
    finally:
        conn.close()


@router.get("/dashboard/monthly-trend")
async def monthly_trend(unit: str = Query("HM")):
    conn = get_conn()
    try:
        ref_dates = [r["ref_date"] for r in conn.execute(
            f"SELECT DISTINCT ref_date FROM inventory_items i WHERE {EXCLUDE_STORAGE} ORDER BY ref_date"
        ).fetchall()]
        trend = []
        for rd in ref_dates:
            r = conn.execute(f"""
                SELECT COALESCE(SUM(amount),0) AS ta, COALESCE(SUM(weight_ton),0) AS tw, COUNT(*) AS tc
                FROM ({_lot_agg_subquery()}) la
            """, (rd,)).fetchone()
            trend.append({
                "ref_date": rd,
                "total_amount": fmt_amount(r["ta"], unit)["value"],
                "total_weight_ton": round(r["tw"], 3),
                "item_count": int(r["tc"]),
            })
        return {"trend": trend, "unit": unit}
    finally:
        conn.close()


@router.get("/dashboard/plan-weight-trend")
async def plan_weight_trend(unit: str = Query("HM")):
    conn = get_conn()
    try:
        latest_rd = _latest_ref(conn)
        rows = conn.execute(f"""
            SELECT substr(p.plan_date,1,7) AS plan_month,
                   COUNT(*) AS plan_count,
                   COALESCE(SUM(la.weight_ton),0) AS plan_weight_ton,
                   COALESCE(SUM(la.amount),0) AS plan_amount
            FROM depletion_plans p
            LEFT JOIN ({_lot_agg_subquery()}) la ON la.lot_no = p.lot_no
            WHERE p.plan_date IS NOT NULL AND p.plan_date != ''
            GROUP BY plan_month ORDER BY plan_month
        """, (latest_rd,)).fetchall()
        trend = []
        for r in rows:
            trend.append({
                "plan_month": r["plan_month"],
                "plan_count": r["plan_count"],
                "plan_weight_ton": round(r["plan_weight_ton"], 3),
                "plan_amount": fmt_amount(r["plan_amount"], unit)["value"],
            })
        return {"trend": trend, "unit": unit}
    finally:
        conn.close()


@router.get("/dashboard/cost-center-summary")
async def cost_center_summary(
    ref_date: Optional[str] = Query(None), unit: str = Query("HM"),
    factory: Optional[str] = Query(None),
):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        sql = _lot_agg_subquery(factory=factory)
        params = _lot_agg_params(ref_date, factory)
        rows = conn.execute(f"""
            SELECT COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   la.cost_center,
                   COUNT(*) AS item_count,
                   COALESCE(SUM(la.weight_ton),0) AS total_weight,
                   COALESCE(SUM(la.amount),0) AS total_amount,
                   COUNT(p.lot_no) AS plan_count,
                   COUNT(ax.lot_no) AS actual_count
            FROM ({sql}) la
            LEFT JOIN depletion_plans p ON p.lot_no = la.lot_no
            LEFT JOIN (SELECT DISTINCT lot_no FROM depletion_actuals WHERE ref_date=?) ax
                   ON ax.lot_no = la.lot_no
            GROUP BY la.cost_center, la.cost_center_name
            ORDER BY total_amount DESC
        """, params + [ref_date]).fetchall()
        items = []
        for r in rows:
            d = dict(r)
            d["total_amount"] = fmt_amount(d["total_amount"], unit)["value"]
            d["total_weight"] = round(d["total_weight"], 3)
            items.append(d)
        return {"ref_date": ref_date, "unit": unit, "items": items}
    finally:
        conn.close()


@router.get("/dashboard/cost-center-plan-type-summary")
async def cost_center_plan_type_summary(
    ref_date: Optional[str] = Query(None), unit: str = Query("HM"),
    factory: Optional[str] = Query(None),
):
    """
    원가중심점 × 소진계획방안(plan_type) 별 교차 집계.
    각 조합에 대해 계획 기준(건수/중량/금액)과 실적 기준(건수/중량/금액)을 모두 산출.

    [요구사항 7] 실적 집계 기준 변경:
    실적유형(상세시트 '유형' 컬럼)이 'Sales'로 시작하면 '전환 판매' 실적로, 'WIP'로 시작하면
    '생산투입' 실적로 집계한다 (해당 LOT에 등록된 계획유형과는 무관하게, 실적유형 자체가
    어떤 방안에 대한 실적인지를 결정한다). 이 둘에 해당하지 않는 실적유형은 집계에서 제외한다
    (원가중심점x계획방안 표는 계획방안 4종 기준이므로, 매핑 불가능한 실적은 별도 노출하지 않음).
    """
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)

        sql = _lot_agg_subquery(factory=factory)
        params = _lot_agg_params(ref_date, factory)

        # 계획 기준: 원가중심점 x plan_type 별 LOT 건수/중량/금액
        plan_rows = conn.execute(f"""
            SELECT COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   COALESCE(p.plan_type, '미등록') AS plan_type,
                   COUNT(*) AS plan_count,
                   COALESCE(SUM(la.weight_ton),0) AS plan_weight,
                   COALESCE(SUM(la.amount),0) AS plan_amount
            FROM ({sql}) la
            LEFT JOIN depletion_plans p ON p.lot_no = la.lot_no
            GROUP BY cc_name, plan_type
            ORDER BY cc_name, plan_amount DESC
        """, params).fetchall()

        # LOT별 단가 및 원가중심점 매핑 (실적금액 산출용, factory 필터 동일 적용)
        unit_price_rows = conn.execute(f"""
            SELECT lot_no, COALESCE(cost_center_name, cost_center) AS cc_name,
                   COALESCE(amount,0) AS amount, COALESCE(weight_ton,0) AS weight_ton
            FROM ({sql}) la
        """, params).fetchall()
        unit_price = {}
        lot_cc = {}
        for r in unit_price_rows:
            lot_cc[r["lot_no"]] = r["cc_name"]
            if r["weight_ton"] and r["weight_ton"] > 0:
                unit_price[r["lot_no"]] = r["amount"] / r["weight_ton"]

        # 실적 기준: actual_type(원본 '유형' 문자열) -> plan_type 매핑(요구사항7 규칙)으로 분류
        actual_rows = conn.execute(
            "SELECT lot_no, actual_type_manual, actual_type_raw, weight_ton FROM depletion_actuals WHERE ref_date=?",
            (ref_date,)
        ).fetchall()

        actual_agg = {}  # (cc_name, plan_type) -> {count, weight, amount}
        for a in actual_rows:
            lot = a["lot_no"]
            if lot not in lot_cc:
                continue  # factory 필터에 의해 이번 집계 대상이 아닌 LOT
            # 사용자가 수동으로 실적유형을 보정했다면 그 값을 우선 사용
            raw = a["actual_type_manual"] or a["actual_type_raw"]
            mapped_pt = _actual_type_to_plan_type(raw)
            if mapped_pt is None:
                continue  # Sales/WIP 패턴에 해당하지 않는 실적유형은 이 표에서 제외
            cc = lot_cc.get(lot, "미배정")
            key = (cc, mapped_pt)
            wt = abs(a["weight_ton"] or 0.0)
            up = unit_price.get(lot)
            amt = wt * up if up is not None else 0.0
            if key not in actual_agg:
                actual_agg[key] = {"actual_count": 0, "actual_weight": 0.0, "actual_amount": 0.0}
            actual_agg[key]["actual_count"] += 1
            actual_agg[key]["actual_weight"] += wt
            actual_agg[key]["actual_amount"] += amt

        items = []
        for r in plan_rows:
            cc = r["cc_name"] or "미배정"
            pt = r["plan_type"]
            key = (cc, pt)
            a = actual_agg.get(key, {"actual_count": 0, "actual_weight": 0.0, "actual_amount": 0.0})
            items.append({
                "cc_name": cc,
                "plan_type": pt,
                "plan_count": r["plan_count"],
                "plan_weight": round(r["plan_weight"], 3),
                "plan_amount": fmt_amount(r["plan_amount"], unit)["value"],
                "actual_count": a["actual_count"],
                "actual_weight": round(a["actual_weight"], 3),
                "actual_amount": fmt_amount(a["actual_amount"], unit)["value"],
            })
        return {"ref_date": ref_date, "unit": unit, "items": items}
    except Exception as e:
        logger.error(f"cost_center_plan_type_summary 오류: {e}", exc_info=True)
        raise HTTPException(500, f"조회 실패: {e}")
    finally:
        conn.close()


@router.get("/dashboard/period-compare")
async def period_compare(
    ref_date: Optional[str] = Query(None),
    mode: str = Query("month"),
    unit: str = Query("HM"),
):
    """전월/전분기/전년 비교 - 총 장기재고 금액만 비교 (중복 카드 제거 후 단순화)"""
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        prev_rd = _prev_ref(conn, ref_date, mode)

        def _summary(rd):
            if not rd:
                return None
            r = conn.execute(f"""
                SELECT COALESCE(SUM(amount),0) AS ta, COALESCE(SUM(weight_ton),0) AS tw, COUNT(*) AS tc
                FROM ({_lot_agg_subquery()}) la
            """, (rd,)).fetchone()
            return {
                "ref_date": rd,
                "total_amount": fmt_amount(r["ta"], unit)["value"],
                "total_weight": round(r["tw"], 3),
                "total_count": int(r["tc"]),
            }

        return {
            "mode": mode, "unit": unit,
            "current": _summary(ref_date), "previous": _summary(prev_rd),
            "mode_label": {"month": "전월", "quarter": "전분기", "year": "전년도"}.get(mode, "전월"),
        }
    except Exception as e:
        logger.error(f"period_compare 오류: {e}", exc_info=True)
        raise HTTPException(500, f"비교 조회 실패: {e}")
    finally:
        conn.close()


# ══════════════════════════════════════════════
# 재고 목록 (조회전용) - 저장품 제외, LOT 집계 기준
# ══════════════════════════════════════════════
@router.get("/inventory")
async def inventory_list(
    ref_date: Optional[str] = Query(None), factory: Optional[str] = Query(None),
    item_type: Optional[str] = Query(None), item_code: Optional[str] = Query(None),
    lot_no: Optional[str] = Query(None), dept: Optional[str] = Query(None),
    plan_type: Optional[str] = Query(None), cost_center: Optional[str] = Query(None),
    item_name: Optional[str] = Query(None),
    page: int = Query(1, ge=1), page_size: int = Query(50, ge=1, le=500),
    unit: str = Query("HM"),
):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)

        conds = ["1=1"]
        params: list = []
        if factory:     conds.append("la.factory=?");          params.append(factory)
        if item_type:   conds.append("la.item_type=?");        params.append(item_type)
        if item_code:   conds.append("la.item_code LIKE ?");   params.append(f"%{item_code}%")
        if lot_no:      conds.append("la.lot_no LIKE ?");      params.append(f"%{lot_no}%")
        if item_name:   conds.append("la.item_name LIKE ?");   params.append(f"%{item_name}%")
        if cost_center: conds.append("(la.cost_center=? OR la.cost_center_name LIKE ?)"); params += [cost_center, f"%{cost_center}%"]
        if dept:        conds.append("p.dept=?");              params.append(dept)
        if plan_type:   conds.append("p.plan_type=?");         params.append(plan_type)
        where = " AND ".join(conds)
        offset = (page - 1) * page_size

        base_sql = f"""
            FROM ({_lot_agg_subquery()}) la
            LEFT JOIN depletion_plans p ON p.lot_no = la.lot_no
            LEFT JOIN (SELECT DISTINCT lot_no FROM depletion_actuals WHERE ref_date=?) ax
                   ON ax.lot_no = la.lot_no
            WHERE {where}
        """
        total = conn.execute(f"SELECT COUNT(*) {base_sql}", [ref_date, ref_date] + params).fetchone()[0]
        rows = conn.execute(f"""
            SELECT la.factory, la.item_type, la.item_code, la.item_name,
                   COALESCE(la.cost_center_name, la.cost_center) AS cc_name, la.cost_center,
                   la.lot_no, la.qty, la.weight_ton, la.amount,
                   la.base_date, la.months_label, la.is_new,
                   p.dept, p.reason, p.plan_type, p.plan_date, p.detail_plan,
                   p.created_by, p.created_by_name, p.created_at AS plan_created_at,
                   p.updated_by, p.updated_by_name, p.updated_at AS plan_updated_at,
                   CASE WHEN ax.lot_no IS NOT NULL THEN 1 ELSE 0 END AS has_actual
            {base_sql}
            ORDER BY la.amount DESC LIMIT ? OFFSET ?
        """, [ref_date, ref_date] + params + [page_size, offset]).fetchall()

        items = []
        for r in rows:
            d = dict(r)
            d["amount"] = fmt_amount(d["amount"], unit)["value"]
            d["weight_ton"] = round(d["weight_ton"], 3)
            items.append(d)
        return {"ref_date": ref_date, "unit": unit, "total": total, "page": page, "page_size": page_size, "items": items}
    except Exception as e:
        logger.error(f"재고 목록 오류: {e}", exc_info=True)
        raise HTTPException(500, f"재고 목록 조회 실패: {e}")
    finally:
        conn.close()


@router.get("/inventory/export")
async def export_inventory(
    ref_date: Optional[str] = Query(None),
    factory: Optional[str] = Query(None),
    item_type: Optional[str] = Query(None),
    unit: str = Query("HM"),
):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        if not ref_date:
            raise HTTPException(404, "업로드된 재고 데이터가 없습니다.")
        conds = ["1=1"]; params: list = []
        if factory:   conds.append("la.factory=?");   params.append(factory)
        if item_type: conds.append("la.item_type=?"); params.append(item_type)
        where = " AND ".join(conds)
        rows = conn.execute(f"""
            SELECT la.factory, la.item_type, la.item_code, la.item_name,
                   COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   la.lot_no, la.qty, la.weight_ton, la.amount,
                   la.base_date, la.months_label, la.is_new,
                   p.dept, p.reason, p.plan_type, p.plan_date, p.detail_plan,
                   p.created_by_name, p.created_at, p.updated_by_name, p.updated_at
            FROM ({_lot_agg_subquery()}) la
            LEFT JOIN depletion_plans p ON p.lot_no = la.lot_no
            WHERE {where}
            ORDER BY la.amount DESC
        """, [ref_date] + params).fetchall()
        conn.close()

        data = [dict(r) for r in rows]
        for d in data:
            d["amount"] = fmt_amount(d["amount"], unit)["value"]
        df = pd.DataFrame(data) if data else pd.DataFrame()
        COL_MAP = {
            "factory": "공장", "item_type": "품목구분", "item_code": "품목코드", "item_name": "품명",
            "cc_name": "원가중심점", "lot_no": "LOT NO", "qty": "수량", "weight_ton": "중량(ton)",
            "amount": f"금액({'억원' if unit=='HM' else '백만원' if unit=='MN' else '원'})",
            "base_date": "기준일자", "months_label": "개월", "is_new": "신규여부",
            "dept": "담당부서", "reason": "장기재고사유", "plan_type": "소진계획방안",
            "plan_date": "소진계획기한", "detail_plan": "세부계획",
            "created_by_name": "작성자", "created_at": "작성일시",
            "updated_by_name": "수정자", "updated_at": "수정일시",
        }
        if not df.empty:
            df = df.rename(columns={k: v for k, v in COL_MAP.items() if k in df.columns})
        else:
            df = pd.DataFrame(columns=list(COL_MAP.values()))

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False, sheet_name="장기재고현황")
            ws = w.sheets["장기재고현황"]
            hdr = w.book.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
            for ci, col in enumerate(df.columns):
                ws.write(0, ci, col, hdr)
                ws.set_column(ci, ci, max(12, len(str(col)) + 2))

        fname = f"장기재고현황_{ref_date}_{datetime.now().strftime('%H%M%S')}.xlsx"
        return _xlsx_response(buf, fname)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"재고 Excel 다운로드 오류: {e}", exc_info=True)
        raise HTTPException(500, f"Excel 생성 실패: {e}")


@router.get("/inventory/export-ppt")
async def export_inventory_ppt(
    ref_date: Optional[str] = Query(None),
    factory: Optional[str] = Query(None),
    item_type: Optional[str] = Query(None),
    unit: str = Query("HM"),
):
    """장기재고현황 조회 화면 PPT 다운로드 (요구사항 7)"""
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        if not ref_date:
            raise HTTPException(404, "업로드된 재고 데이터가 없습니다.")
        conds = ["1=1"]; params: list = []
        if factory:   conds.append("la.factory=?");   params.append(factory)
        if item_type: conds.append("la.item_type=?"); params.append(item_type)
        where = " AND ".join(conds)

        rows = conn.execute(f"""
            SELECT la.factory, la.item_type, la.item_code, la.item_name,
                   COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   la.lot_no, la.weight_ton, la.amount, la.months_label,
                   p.plan_type
            FROM ({_lot_agg_subquery()}) la
            LEFT JOIN depletion_plans p ON p.lot_no = la.lot_no
            WHERE {where}
            ORDER BY la.amount DESC
        """, [ref_date] + params).fetchall()

        summ = conn.execute(f"""
            SELECT COUNT(*) AS tc, COALESCE(SUM(weight_ton),0) AS tw, COALESCE(SUM(amount),0) AS ta,
                   SUM(CASE WHEN months_label='7개월이상' THEN 1 ELSE 0 END) AS cc_cnt,
                   COALESCE(SUM(CASE WHEN months_label='7개월이상' THEN amount ELSE 0 END),0) AS cc_amt
            FROM ({_lot_agg_subquery()}) la
            WHERE {where}
        """, [ref_date] + params).fetchone()

        plan_cnt = conn.execute(f"""
            SELECT COUNT(*) FROM ({_lot_agg_subquery()}) la
            JOIN depletion_plans p ON p.lot_no = la.lot_no
            WHERE {where}
        """, [ref_date] + params).fetchone()[0]
        conn.close()

        items = [dict(r) for r in rows]
        for d in items:
            d["amount"] = fmt_amount(d["amount"], unit)["value"]

        summary = {
            "total_count": summ["tc"], "total_weight": summ["tw"],
            "total_amount": fmt_amount(summ["ta"], unit)["value"],
            "critical_count": summ["cc_cnt"] or 0,
            "critical_amount": fmt_amount(summ["cc_amt"], unit)["value"],
            "plan_count": plan_cnt,
        }

        from backend.ppt_exporter import generate_inventory_ppt
        ppt_bytes = generate_inventory_ppt(
            items, summary, ref_date,
            unit_label="억원" if unit == "HM" else "백만원" if unit == "MN" else "원",
        )
        buf = io.BytesIO(ppt_bytes)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"장기재고현황_{ref_date}_{ts}.pptx"
        return _pptx_response(buf, fname)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"재고 PPT 생성 오류: {e}", exc_info=True)
        try:
            conn.close()
        except Exception:
            pass
        raise HTTPException(500, f"PPT 생성 실패: {str(e)}")


# ══════════════════════════════════════════════
# 소진계획: 경로 충돌 주의 - 구체적 경로를 먼저 등록
# ══════════════════════════════════════════════
@router.get("/plans/export-template")
async def export_plan_template(ref_date: Optional[str] = Query(None)):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        if not ref_date:
            raise HTTPException(404, "업로드된 재고 데이터가 없습니다.")
        rows = conn.execute(f"""
            SELECT la.factory, la.item_type, la.item_code, la.item_name,
                   COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   la.lot_no, la.qty, la.weight_ton, la.amount, la.base_date,
                   p.dept, p.reason, p.plan_type, p.plan_date, p.detail_plan
            FROM ({_lot_agg_subquery()}) la
            LEFT JOIN depletion_plans p ON p.lot_no = la.lot_no
            ORDER BY la.amount DESC
        """, (ref_date,)).fetchall()
        conn.close()

        COL_NAMES = ["공장", "품목구분", "품목코드", "품명", "원가중심점",
                     "LOT NO", "수량", "중량(ton)", "금액", "기준일자",
                     "담당부서", "장기재고사유", "소진계획방안", "소진계획기한", "세부계획"]
        READONLY_COLS = 10
        data = [dict(r) for r in rows]
        df = pd.DataFrame(data) if data else pd.DataFrame(columns=COL_NAMES)
        if not df.empty:
            df.columns = COL_NAMES

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False, sheet_name="소진계획입력")
            wb = w.book; ws = w.sheets["소진계획입력"]
            hdr = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1, "align": "center"})
            ro  = wb.add_format({"bg_color": "#F2F2F2", "border": 1})
            inp = wb.add_format({"bg_color": "#FFFFFF", "border": 1})
            for ci, cn in enumerate(COL_NAMES):
                ws.write(0, ci, cn, hdr)
                ws.set_column(ci, ci, 15)
            for ri in range(len(df)):
                for ci in range(READONLY_COLS):
                    val = df.iloc[ri, ci] if ci < len(df.columns) else ""
                    ws.write(ri + 1, ci, "" if str(val) == "nan" else val, ro)
                for ci in range(READONLY_COLS, len(COL_NAMES)):
                    val = df.iloc[ri, ci] if ci < len(df.columns) else ""
                    ws.write(ri + 1, ci, "" if str(val) == "nan" else val, inp)
            n = max(len(df) + 100, 200)
            ws.data_validation(1, 10, n, 10, {"validate": "list", "source": ["영업", "생산", "구매"]})
            ws.data_validation(1, 11, n, 11, {"validate": "list", "source": ["주문 변경", "주문 취소", "납품 후 잔량", "기타"]})
            ws.data_validation(1, 12, n, 12, {"validate": "list", "source": ["생산투입", "전환 판매", "폐기", "기타"]})

        fname = f"소진계획입력템플릿_{ref_date}.xlsx"
        return _xlsx_response(buf, fname)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"템플릿 다운로드 오류: {e}", exc_info=True)
        raise HTTPException(500, f"템플릿 생성 실패: {e}")


@router.post("/plans/bulk-upload")
async def bulk_upload_plans(request: Request, file: UploadFile = File(...)):
    u = cur_user(request)
    if not u:
        raise HTTPException(401, "로그인이 필요합니다.")
    try:
        content = await file.read()
        df = pd.read_excel(io.BytesIO(content), header=0, dtype=str)
    except Exception as e:
        raise HTTPException(400, f"엑셀 파싱 실패: {e}")

    if "LOT NO" not in df.columns:
        raise HTTPException(400, "필수 컬럼 'LOT NO' 없음. 템플릿을 다시 다운로드하여 사용하세요.")

    by = u["username"]; by_name = u.get("display_name", "") or by
    ok = fail = 0; errors = []; now = _now_str()
    conn = get_conn()
    try:
        for idx, row in df.iterrows():
            lot = str(row.get("LOT NO", "")).strip()
            if not lot or lot == "nan":
                continue
            try:
                dept  = str(row.get("담당부서", "")).strip()     or None
                rsn   = str(row.get("장기재고사유", "")).strip()  or None
                ptype = str(row.get("소진계획방안", "")).strip()  or None
                pdate = str(row.get("소진계획기한", "")).strip()  or None
                det   = str(row.get("세부계획", "")).strip()      or None
                if pdate and pdate not in ("nan", "None", ""):
                    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
                        try:
                            pdate = datetime.strptime(pdate[:10], fmt).strftime("%Y-%m-%d"); break
                        except Exception:
                            pass
                    else:
                        pdate = None
                else:
                    pdate = None

                ex = conn.execute("SELECT id FROM depletion_plans WHERE lot_no=?", (lot,)).fetchone()
                if ex:
                    conn.execute("""
                        UPDATE depletion_plans
                        SET dept=?,reason=?,plan_type=?,plan_date=?,detail_plan=?,
                            updated_by=?,updated_by_name=?,updated_at=?
                        WHERE lot_no=?
                    """, (dept, rsn, ptype, pdate, det, by, by_name, now, lot))
                else:
                    inv = conn.execute(
                        f"SELECT item_code,item_name,factory,cost_center,cost_center_name,item_type "
                        f"FROM inventory_items i WHERE {EXCLUDE_STORAGE} AND lot_no=? LIMIT 1",
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
                          inv["factory"] if inv else None, inv["cost_center"] if inv else None,
                          inv["cost_center_name"] if inv else None, inv["item_type"] if inv else None,
                          dept, rsn, ptype, pdate, det, by, by_name, by, by_name))
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
    finally:
        conn.close()


@router.get("/plans/no-plan")
async def inventory_no_plan(
    ref_date: Optional[str] = Query(None), factory: Optional[str] = Query(None),
    item_name: Optional[str] = Query(None), cost_center: Optional[str] = Query(None),
    lot_no: Optional[str] = Query(None), item_code: Optional[str] = Query(None),
    page: int = Query(1, ge=1), page_size: int = Query(50, ge=1, le=500),
    unit: str = Query("HM"),
):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        conds = ["NOT EXISTS (SELECT 1 FROM depletion_plans p WHERE p.lot_no = la.lot_no)"]
        params: list = []
        if factory:     conds.append("la.factory=?");        params.append(factory)
        if item_name:   conds.append("la.item_name LIKE ?"); params.append(f"%{item_name}%")
        if item_code:   conds.append("la.item_code LIKE ?"); params.append(f"%{item_code}%")
        if cost_center: conds.append("(la.cost_center=? OR la.cost_center_name LIKE ?)"); params += [cost_center, f"%{cost_center}%"]
        if lot_no:      conds.append("la.lot_no LIKE ?");     params.append(f"%{lot_no}%")
        where = " AND ".join(conds)
        offset = (page - 1) * page_size

        total = conn.execute(f"""
            SELECT COUNT(*) FROM ({_lot_agg_subquery()}) la WHERE {where}
        """, [ref_date] + params).fetchone()[0]
        rows = conn.execute(f"""
            SELECT la.factory, la.item_type, la.item_code, la.item_name,
                   COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   la.lot_no, la.weight_ton, la.amount, la.base_date, la.is_new, la.months_label
            FROM ({_lot_agg_subquery()}) la WHERE {where}
            ORDER BY la.amount DESC LIMIT ? OFFSET ?
        """, [ref_date] + params + [page_size, offset]).fetchall()

        items = []
        for r in rows:
            d = dict(r)
            d["amount"] = fmt_amount(d["amount"], unit)["value"]
            d["weight_ton"] = round(d["weight_ton"], 3)
            items.append(d)
        return {"ref_date": ref_date, "unit": unit, "total": total, "page": page, "items": items}
    finally:
        conn.close()


@router.get("/plans")
async def get_plans(
    ref_date: Optional[str] = Query(None), factory: Optional[str] = Query(None),
    dept: Optional[str] = Query(None), plan_type: Optional[str] = Query(None),
    lot_no: Optional[str] = Query(None), lot_no_exact: Optional[str] = Query(None),
    item_name: Optional[str] = Query(None), cost_center: Optional[str] = Query(None),
    item_code: Optional[str] = Query(None),
    page: int = Query(1, ge=1), page_size: int = Query(50, ge=1, le=500),
    unit: str = Query("HM"),
):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)

        if lot_no_exact:
            rows = conn.execute(f"""
                SELECT la.factory, la.item_type, la.item_code, la.item_name,
                       COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                       la.lot_no, la.weight_ton, la.amount, la.base_date, la.months_label,
                       p.dept, p.reason, p.plan_type, p.plan_date, p.detail_plan,
                       p.created_by, p.created_by_name, p.created_at,
                       p.updated_by, p.updated_by_name, p.updated_at
                FROM ({_lot_agg_subquery()}) la
                JOIN depletion_plans p ON p.lot_no = la.lot_no
                WHERE p.lot_no = ? LIMIT 1
            """, (ref_date, lot_no_exact)).fetchall()
            if not rows:
                rows = conn.execute("SELECT * FROM depletion_plans WHERE lot_no=? LIMIT 1", (lot_no_exact,)).fetchall()
            items = [dict(r) for r in rows]
            for d in items:
                if "amount" in d:
                    d["amount"] = fmt_amount(d["amount"], unit)["value"]
            return {"ref_date": ref_date, "unit": unit, "total": len(items), "page": 1, "items": items}

        conds = ["1=1"]; params: list = []
        if factory:     conds.append("la.factory=?");          params.append(factory)
        if dept:        conds.append("p.dept=?");              params.append(dept)
        if plan_type:   conds.append("p.plan_type=?");         params.append(plan_type)
        if lot_no:      conds.append("la.lot_no LIKE ?");      params.append(f"%{lot_no}%")
        if item_name:   conds.append("la.item_name LIKE ?");   params.append(f"%{item_name}%")
        if item_code:   conds.append("la.item_code LIKE ?");   params.append(f"%{item_code}%")
        if cost_center: conds.append("(la.cost_center=? OR la.cost_center_name LIKE ?)"); params += [cost_center, f"%{cost_center}%"]
        where = " AND ".join(conds)
        offset = (page - 1) * page_size

        total = conn.execute(f"""
            SELECT COUNT(*) FROM ({_lot_agg_subquery()}) la
            JOIN depletion_plans p ON p.lot_no = la.lot_no WHERE {where}
        """, [ref_date] + params).fetchone()[0]
        rows = conn.execute(f"""
            SELECT la.factory, la.item_type, la.item_code, la.item_name,
                   COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   la.lot_no, la.weight_ton, la.amount, la.base_date, la.months_label,
                   p.dept, p.reason, p.plan_type, p.plan_date, p.detail_plan,
                   p.created_by, p.created_by_name, p.created_at AS plan_created_at,
                   p.updated_by, p.updated_by_name, p.updated_at AS plan_updated_at
            FROM ({_lot_agg_subquery()}) la
            JOIN depletion_plans p ON p.lot_no = la.lot_no
            WHERE {where} ORDER BY la.amount DESC LIMIT ? OFFSET ?
        """, [ref_date] + params + [page_size, offset]).fetchall()

        items = []
        for r in rows:
            d = dict(r)
            d["amount"] = fmt_amount(d["amount"], unit)["value"]
            d["weight_ton"] = round(d["weight_ton"], 3)
            items.append(d)
        return {"ref_date": ref_date, "unit": unit, "total": total, "page": page, "items": items}
    finally:
        conn.close()


class PlanBody(BaseModel):
    dept: Optional[str] = None
    reason: Optional[str] = None
    plan_type: Optional[str] = None
    plan_date: Optional[str] = None
    detail_plan: Optional[str] = None


@router.post("/plans/{lot_no}")
async def upsert_plan(lot_no: str, body: PlanBody, request: Request):
    u = cur_user(request)
    if not u:
        raise HTTPException(401, "로그인이 필요합니다.")
    by = u["username"]; by_name = u.get("display_name", "") or by
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
            """, (body.dept, body.reason, body.plan_type, body.plan_date, body.detail_plan,
                  by, by_name, now, lot_no))
        else:
            inv = conn.execute(
                f"SELECT item_code,item_name,factory,cost_center,cost_center_name,item_type "
                f"FROM inventory_items i WHERE {EXCLUDE_STORAGE} AND lot_no=? LIMIT 1",
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
                  inv["factory"] if inv else None, inv["cost_center"] if inv else None,
                  inv["cost_center_name"] if inv else None, inv["item_type"] if inv else None,
                  body.dept, body.reason, body.plan_type, body.plan_date, body.detail_plan,
                  by, by_name, by, by_name))
        conn.commit()
        return {"ok": True, "lot_no": lot_no}
    except Exception as e:
        conn.rollback()
        logger.error(f"계획 저장 오류: {e}", exc_info=True)
        raise HTTPException(500, f"계획 저장 실패: {e}")
    finally:
        conn.close()


@router.delete("/plans/{lot_no}")
async def delete_plan(lot_no: str, request: Request):
    u = cur_user(request)
    if not u:
        raise HTTPException(401, "로그인이 필요합니다.")
    conn = get_conn()
    try:
        conn.execute("DELETE FROM depletion_plans WHERE lot_no=?", (lot_no,))
        conn.commit()
        return {"ok": True}
    except Exception as e:
        conn.rollback()
        raise HTTPException(500, f"계획 삭제 실패: {e}")
    finally:
        conn.close()


# ══════════════════════════════════════════════
# 계획/실적 비교
# ══════════════════════════════════════════════
@router.get("/compare")
async def compare_plan_actual(
    ref_date: Optional[str] = Query(None), factory: Optional[str] = Query(None),
    dept: Optional[str] = Query(None), cost_center: Optional[str] = Query(None),
    page: int = Query(1, ge=1), page_size: int = Query(50, ge=1, le=500),
    unit: str = Query("HM"),
):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        sql = _lot_agg_subquery(factory=factory, cost_center=cost_center)
        lot_params = _lot_agg_params(ref_date, factory, cost_center)

        conds = ["1=1"]; params: list = []
        if dept: conds.append("p.dept=?"); params.append(dept)
        where = " AND ".join(conds)
        offset = (page - 1) * page_size

        total = conn.execute(f"""
            SELECT COUNT(*) FROM ({sql}) la
            JOIN depletion_plans p ON p.lot_no = la.lot_no
            WHERE {where}
        """, lot_params + params).fetchone()[0]

        rows = conn.execute(f"""
            SELECT la.factory, la.item_type, la.item_code, la.item_name,
                   COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   la.lot_no, la.weight_ton, la.amount, la.base_date,
                   p.dept, p.plan_type, p.plan_date, p.reason,
                   ax.actual_type, ax.actual_type_manual, ax.process_date,
                   ax.weight_ton AS actual_weight,
                   CASE WHEN ax.lot_no IS NOT NULL THEN 1 ELSE 0 END AS has_actual,
                   CASE WHEN ax.lot_no IS NOT NULL THEN '조치' ELSE '미조치' END AS action_status,
                   ax.id AS actual_id
            FROM ({sql}) la
            JOIN depletion_plans p ON p.lot_no = la.lot_no
            LEFT JOIN (
                SELECT a.*, MIN(a.id) OVER (PARTITION BY a.lot_no) AS first_id
                FROM depletion_actuals a WHERE a.ref_date = ?
            ) ax ON ax.lot_no = la.lot_no AND ax.id = ax.first_id
            WHERE {where} ORDER BY la.amount DESC LIMIT ? OFFSET ?
        """, lot_params + [ref_date] + params + [page_size, offset]).fetchall()

        items = []
        for r in rows:
            d = dict(r)
            pt = d.get("plan_type") or ""
            # [요구사항7] 실적유형 원본 문자열을 Sales/WIP 규칙으로 매핑한 카테고리와
            # 계획유형을 비교해야 진짜 "계획대로 처리됐는지"를 판단할 수 있다.
            raw_at = d.get("actual_type_manual") or d.get("actual_type") or ""
            mapped_at = _actual_type_to_plan_type(raw_at)
            d["actual_type_mapped"] = mapped_at
            d["type_match"] = (pt == mapped_at) if (pt and mapped_at) else None
            d["amount"] = fmt_amount(d["amount"], unit)["value"]
            d["weight_ton"] = round(d["weight_ton"], 3)
            items.append(d)
        return {"ref_date": ref_date, "unit": unit, "total": total, "page": page, "items": items}
    except Exception as e:
        logger.error(f"비교 조회 오류: {e}", exc_info=True)
        raise HTTPException(500, f"비교 조회 실패: {e}")
    finally:
        conn.close()


@router.get("/compare/summary")
async def compare_summary(
    ref_date: Optional[str] = Query(None),
    factory: Optional[str] = Query(None),
    cost_center: Optional[str] = Query(None),
    dept: Optional[str] = Query(None),
    unit: str = Query("HM"),
):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)

        sql = _lot_agg_subquery(factory=factory, cost_center=cost_center)
        lot_params = _lot_agg_params(ref_date, factory, cost_center)

        conds = ["1=1"]; params: list = []
        if dept: conds.append("p.dept=?"); params.append(dept)
        where = " AND ".join(conds)

        r = conn.execute(f"""
            SELECT COUNT(*) AS plan_total,
                   SUM(CASE WHEN ax.lot_no IS NOT NULL THEN 1 ELSE 0 END) AS action_count,
                   SUM(CASE WHEN ax.lot_no IS NULL THEN 1 ELSE 0 END) AS no_action_count,
                   COALESCE(SUM(la.weight_ton),0) AS total_weight,
                   COALESCE(SUM(la.amount),0) AS total_amount,
                   COALESCE(SUM(CASE WHEN ax.lot_no IS NOT NULL THEN la.weight_ton ELSE 0 END),0) AS action_weight,
                   COALESCE(SUM(CASE WHEN ax.lot_no IS NOT NULL THEN la.amount ELSE 0 END),0) AS action_amount,
                   COALESCE(SUM(CASE WHEN ax.lot_no IS NULL THEN la.weight_ton ELSE 0 END),0) AS no_action_weight,
                   COALESCE(SUM(CASE WHEN ax.lot_no IS NULL THEN la.amount ELSE 0 END),0) AS no_action_amount
            FROM ({sql}) la
            JOIN depletion_plans p ON p.lot_no = la.lot_no
            LEFT JOIN (SELECT DISTINCT lot_no FROM depletion_actuals WHERE ref_date=?) ax
                   ON ax.lot_no = la.lot_no
            WHERE {where}
        """, lot_params + [ref_date] + params).fetchone()

        pt = max(r["plan_total"] or 1, 1)
        ac = r["action_count"] or 0
        tw = float(r["total_weight"]) or 1.0
        aw = float(r["action_weight"])
        achievement_rate_wt = round(aw / tw * 100, 1) if tw > 0 else 0.0
        achievement_rate_cnt = round(ac / pt * 100, 1)

        plan_rows = conn.execute(f"""
            SELECT p.plan_type, COUNT(*) AS plan_count,
                   COALESCE(SUM(la.weight_ton),0) AS plan_weight,
                   COALESCE(SUM(la.amount),0) AS plan_amount
            FROM depletion_plans p
            JOIN ({sql}) la ON la.lot_no = p.lot_no
            GROUP BY p.plan_type ORDER BY plan_amount DESC
        """, lot_params).fetchall()

        actual_rows = conn.execute("""
            SELECT COALESCE(a.actual_type_manual, a.actual_type, '기타') AS actual_type,
                   COUNT(DISTINCT a.lot_no) AS actual_count,
                   COALESCE(SUM(a.weight_ton),0) AS actual_weight
            FROM depletion_actuals a WHERE a.ref_date=?
            GROUP BY actual_type ORDER BY actual_weight DESC
        """, (ref_date,)).fetchall()

        return {
            "ref_date": ref_date, "unit": unit,
            "plan_total": int(pt), "action_count": int(ac),
            "no_action_count": int(r["no_action_count"] or 0),
            "action_rate": achievement_rate_cnt,
            "action_rate_weight": achievement_rate_wt,
            "total_weight": round(float(r["total_weight"]), 3),
            "total_amount": fmt_amount(r["total_amount"], unit)["value"],
            "action_weight": round(float(r["action_weight"]), 3),
            "action_amount": fmt_amount(r["action_amount"], unit)["value"],
            "no_action_weight": round(float(r["no_action_weight"]), 3),
            "no_action_amount": fmt_amount(r["no_action_amount"], unit)["value"],
            "plan_by_type": [dict(x) for x in plan_rows],
            "actual_by_type": [dict(x) for x in actual_rows],
        }
    except Exception as e:
        logger.error(f"compare_summary 오류: {e}", exc_info=True)
        raise HTTPException(500, f"비교 요약 실패: {e}")
    finally:
        conn.close()


@router.get("/compare/export")
async def export_compare(
    ref_date: Optional[str] = Query(None), unit: str = Query("HM"),
    factory: Optional[str] = Query(None), cost_center: Optional[str] = Query(None),
):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        if not ref_date:
            raise HTTPException(404, "데이터가 없습니다.")
        sql = _lot_agg_subquery(factory=factory, cost_center=cost_center)
        lot_params = _lot_agg_params(ref_date, factory, cost_center)
        rows = conn.execute(f"""
            SELECT la.factory, la.item_type, la.item_code, la.item_name,
                   COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   la.lot_no, la.weight_ton, la.amount, la.base_date,
                   p.dept, p.plan_type, p.plan_date,
                   COALESCE(ax.actual_type_manual, ax.actual_type, '') AS actual_type,
                   ax.process_date,
                   CASE WHEN ax.lot_no IS NOT NULL THEN '조치' ELSE '미조치' END AS action_status
            FROM ({sql}) la
            JOIN depletion_plans p ON p.lot_no = la.lot_no
            LEFT JOIN (
                SELECT a.*, MIN(a.id) OVER (PARTITION BY a.lot_no) AS first_id
                FROM depletion_actuals a WHERE a.ref_date=?
            ) ax ON ax.lot_no = la.lot_no AND ax.id = ax.first_id
            WHERE 1=1 ORDER BY la.amount DESC
        """, lot_params + [ref_date]).fetchall()
        conn.close()
        data = [dict(r) for r in rows]
        for d in data:
            d["amount"] = fmt_amount(d["amount"], unit)["value"]
        df = pd.DataFrame(data) if data else pd.DataFrame()
        if not df.empty:
            df.columns = ["공장", "품목구분", "품목코드", "품명", "원가중심점", "LOT NO",
                          "중량(ton)", "금액", "기준일자", "담당부서", "계획유형", "계획기한",
                          "실적유형", "처리일자", "조치여부"]
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False, sheet_name="계획실적비교")
            if not df.empty:
                ws = w.sheets["계획실적비교"]
                hdr = w.book.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
                for ci, col in enumerate(df.columns):
                    ws.write(0, ci, col, hdr); ws.set_column(ci, ci, 14)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"계획실적비교_{ref_date}_{ts}.xlsx"
        return _xlsx_response(buf, fname)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"비교 Excel 오류: {e}", exc_info=True)
        raise HTTPException(500, f"Excel 생성 실패: {e}")


@router.get("/compare/export-ppt")
async def export_compare_ppt(ref_date: Optional[str] = Query(None), unit: str = Query("HM")):
    conn = get_conn()
    try:
        if not ref_date:
            ref_date = _latest_ref(conn)
        if not ref_date:
            raise HTTPException(404, "데이터가 없습니다.")

        summ = conn.execute(f"""
            SELECT COUNT(*) AS pt,
                   SUM(CASE WHEN ax.lot_no IS NOT NULL THEN 1 ELSE 0 END) AS ac,
                   SUM(CASE WHEN ax.lot_no IS NULL THEN 1 ELSE 0 END) AS nc,
                   COALESCE(SUM(la.amount),0) AS ta,
                   COALESCE(SUM(CASE WHEN ax.lot_no IS NOT NULL THEN la.amount ELSE 0 END),0) AS aa,
                   COALESCE(SUM(CASE WHEN ax.lot_no IS NULL THEN la.amount ELSE 0 END),0) AS na,
                   COALESCE(SUM(la.weight_ton),0) AS tw
            FROM ({_lot_agg_subquery()}) la
            JOIN depletion_plans p ON p.lot_no = la.lot_no
            LEFT JOIN (SELECT DISTINCT lot_no FROM depletion_actuals WHERE ref_date=?) ax
                   ON ax.lot_no = la.lot_no
        """, (ref_date, ref_date)).fetchone()
        pt = summ["pt"] or 1; ac = summ["ac"] or 0

        plan_by = conn.execute(f"""
            SELECT p.plan_type, COUNT(*) AS plan_count,
                   COALESCE(SUM(la.weight_ton),0) AS plan_weight
            FROM depletion_plans p
            LEFT JOIN ({_lot_agg_subquery()}) la ON la.lot_no = p.lot_no
            GROUP BY p.plan_type ORDER BY plan_count DESC
        """, (ref_date,)).fetchall()

        actual_by = conn.execute("""
            SELECT COALESCE(actual_type_manual, actual_type, '기타') AS actual_type,
                   COUNT(DISTINCT lot_no) AS actual_count,
                   COALESCE(SUM(weight_ton),0) AS actual_weight
            FROM depletion_actuals WHERE ref_date=?
            GROUP BY actual_type ORDER BY actual_count DESC
        """, (ref_date,)).fetchall()

        items = conn.execute(f"""
            SELECT la.factory, la.item_name, la.lot_no, la.amount,
                   p.plan_type, CASE WHEN ax.lot_no IS NOT NULL THEN '조치' ELSE '미조치' END AS action_status
            FROM ({_lot_agg_subquery()}) la
            JOIN depletion_plans p ON p.lot_no = la.lot_no
            LEFT JOIN (SELECT DISTINCT lot_no FROM depletion_actuals WHERE ref_date=?) ax
                   ON ax.lot_no = la.lot_no
            ORDER BY la.amount DESC LIMIT 50
        """, (ref_date, ref_date)).fetchall()

        plan_trend_rows = conn.execute(f"""
            SELECT substr(p.plan_date,1,7) AS plan_month, COUNT(*) AS plan_count,
                   COALESCE(SUM(la.weight_ton),0) AS plan_weight_ton,
                   COALESCE(SUM(la.amount),0) AS plan_amount
            FROM depletion_plans p
            LEFT JOIN ({_lot_agg_subquery()}) la ON la.lot_no = p.lot_no
            WHERE p.plan_date IS NOT NULL AND p.plan_date != ''
            GROUP BY plan_month ORDER BY plan_month
        """, (ref_date,)).fetchall()

        cc_rows = conn.execute(f"""
            SELECT COALESCE(la.cost_center_name, la.cost_center) AS cc_name,
                   COUNT(*) AS item_count,
                   COALESCE(SUM(la.weight_ton),0) AS total_weight,
                   COALESCE(SUM(la.amount),0) AS total_amount,
                   COUNT(p.lot_no) AS plan_count,
                   COUNT(ax.lot_no) AS actual_count
            FROM ({_lot_agg_subquery()}) la
            LEFT JOIN depletion_plans p ON p.lot_no = la.lot_no
            LEFT JOIN (SELECT DISTINCT lot_no FROM depletion_actuals WHERE ref_date=?) ax
                   ON ax.lot_no = la.lot_no
            GROUP BY la.cost_center, la.cost_center_name
            ORDER BY total_amount DESC
        """, (ref_date, ref_date)).fetchall()

        # 소진완료금액 (단가방식) - PPT에도 동일 기준 반영
        unit_price_rows = conn.execute(f"SELECT lot_no, COALESCE(amount,0) AS amount, COALESCE(weight_ton,0) AS weight_ton FROM ({_lot_agg_subquery()}) la", (ref_date,)).fetchall()
        unit_price = {r["lot_no"]: (r["amount"]/r["weight_ton"]) for r in unit_price_rows if r["weight_ton"] and r["weight_ton"] > 0}
        actual_wt_rows = conn.execute("SELECT lot_no, SUM(weight_ton) wt FROM depletion_actuals WHERE ref_date=? GROUP BY lot_no", (ref_date,)).fetchall()
        consumed_completed = sum(abs(r["wt"]) * unit_price.get(r["lot_no"], 0) for r in actual_wt_rows if r["lot_no"] in unit_price)

        conn.close()

        summary = {
            "plan_total": int(pt), "action_count": int(ac), "no_action_count": int(summ["nc"] or 0),
            "action_rate": round(ac / pt * 100, 1), "total_amount": fmt_amount(summ["ta"], unit)["value"],
            "action_amount": fmt_amount(summ["aa"], unit)["value"], "no_action_amount": fmt_amount(summ["na"], unit)["value"],
            "consumed_amount": fmt_amount(consumed_completed, unit)["value"],
            "plan_by_type": [dict(r) for r in plan_by],
            "actual_by_type": [dict(r) for r in actual_by],
        }
        items_data = [dict(r) for r in items]
        for d in items_data:
            d["amount"] = fmt_amount(d["amount"], unit)["value"]

        from backend.ppt_exporter import generate_compare_ppt
        plan_trend_data = [dict(r) for r in plan_trend_rows]
        for d in plan_trend_data:
            d["plan_amount"] = fmt_amount(d["plan_amount"], unit)["value"]
        cc_data = [dict(r) for r in cc_rows]
        for d in cc_data:
            d["total_amount"] = fmt_amount(d["total_amount"], unit)["value"]

        ppt_bytes = generate_compare_ppt(
            summary, items_data, ref_date,
            plan_trend_data, cc_data,
            unit_label="억원" if unit == "HM" else "백만원" if unit == "MN" else "원",
        )
        buf = io.BytesIO(ppt_bytes)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fname = f"계획실적비교_{ref_date}_{ts}.pptx"
        return _pptx_response(buf, fname)

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"PPT 생성 오류: {e}", exc_info=True)
        try:
            conn.close()
        except Exception:
            pass
        raise HTTPException(500, f"PPT 생성 실패: {str(e)}")


@router.patch("/actuals/{actual_id}/type")
async def patch_actual_type(actual_id: int, body: dict, request: Request):
    u = cur_user(request)
    if not u:
        raise HTTPException(401, "로그인이 필요합니다.")
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
    finally:
        conn.close()
