"""
loader.py v3
- 장기재고현황 엑셀 업로드 파싱 + DB 저장
- [중요] item_type='저장품'은 장기재고관리 대상에서 제외 (요청사항)
- 파일 내 여러 '조회기준일'이 섞여 있으면 각각을 별도 스냅샷(ref_date)으로 저장
- 재고/재공 시트: weight는 kg 단위로 입력되어 있으므로 ÷1000 하여 weight_ton 저장
- 재고_상세/재공_상세 시트: 소진실적(actuals)로 저장. weight는 마이너스로 들어오는 경우가
  있는데, 이는 "소진(감소)"을 의미하므로 절대값으로 저장(주석 명시)
"""
import io, uuid, logging
from datetime import datetime
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)

EXCLUDED_ITEM_TYPES = {"저장품"}  # 장기재고관리 제외 품목구분

INV_SHEET_CANDIDATES  = ["장기재고현황_재고"]
WIP_SHEET_CANDIDATES  = ["장기재고현황_재공"]
ACT_INV_SHEET_CANDIDATES = ["장기재고현황_재고_상세"]
ACT_WIP_SHEET_CANDIDATES = ["장기재고현황_재공_상세"]


def _pick_sheet(sheet_names, candidates):
    for c in candidates:
        if c in sheet_names:
            return c
    return None


def _to_float(v) -> float:
    try:
        if v is None or v == "":
            return 0.0
        return float(v)
    except (ValueError, TypeError):
        return 0.0


def _norm_date(v) -> str:
    """조회기준일/기준일 등을 YYYYMMDD 또는 YYYY-MM-DD 문자열로 정규화하지 않고
    원본 8자리 숫자 문자열(YYYYMMDD)을 ref_date 키로 그대로 사용한다."""
    if v is None:
        return ""
    s = str(v).strip()
    # 20260531 같은 8자리 숫자
    if s.isdigit() and len(s) == 8:
        return s
    # datetime 객체로 들어온 경우
    try:
        if hasattr(v, "strftime"):
            return v.strftime("%Y%m%d")
    except Exception:
        pass
    return s


def _norm_base_date(v) -> str:
    """기준일(LOT 발생일) -> YYYY-MM-DD"""
    if v is None or v == "":
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if len(s) == 8 and s.isdigit():
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
    return s


def parse_inventory_file(content: bytes, filename: str, uploaded_by: str) -> dict:
    """엑셀 파일을 파싱하여 재고/재공/실적 레코드 딕셔너리 리스트로 반환.
    파일 내 여러 ref_date가 혼재된 경우 ref_date별로 그룹화하여 모두 반환한다.
    """
    try:
        xls = pd.ExcelFile(io.BytesIO(content))
    except Exception as e:
        return {"error": f"엑셀 파일을 열 수 없습니다: {e}"}

    sheet_names = xls.sheet_names
    inv_sheet = _pick_sheet(sheet_names, INV_SHEET_CANDIDATES)
    wip_sheet = _pick_sheet(sheet_names, WIP_SHEET_CANDIDATES)
    act_inv_sheet = _pick_sheet(sheet_names, ACT_INV_SHEET_CANDIDATES)
    act_wip_sheet = _pick_sheet(sheet_names, ACT_WIP_SHEET_CANDIDATES)

    if not inv_sheet and not wip_sheet:
        return {"error": "'장기재고현황_재고' 또는 '장기재고현황_재공' 시트를 찾을 수 없습니다."}

    inv_records: list = []
    wip_records: list = []
    act_records: list = []
    warnings: list = []
    excluded_count = 0

    # ── 재고 시트 ──
    if inv_sheet:
        df = pd.read_excel(xls, sheet_name=inv_sheet)
        for _, row in df.iterrows():
            item_type = str(row.get("품목구분", "") or "").strip()
            if item_type in EXCLUDED_ITEM_TYPES:
                excluded_count += 1
                continue  # 저장품 제외
            ref_date = _norm_date(row.get("조회기준일"))
            if not ref_date:
                continue
            weight_kg = _to_float(row.get("중량"))
            rec = dict(
                ref_date=ref_date,
                factory=row.get("공장"),
                item_type=item_type,
                item_group=row.get("품목군"),
                item_code=row.get("품목코드"),
                item_name=row.get("품명"),
                cost_center=row.get("원가중심점"),
                cost_center_name=row.get("원가중심점명"),
                lot_no=str(row.get("Lot No", "") or "").strip(),
                wo_no=row.get("WO No"),
                qty=_to_float(row.get("수량")),
                weight_kg=weight_kg,
                weight_ton=weight_kg / 1000.0,       # kg -> ton 변환
                amount=_to_float(row.get("재고금액")),  # '재고금액' 컬럼 사용 확정
                base_date=_norm_base_date(row.get("기준일")),
                months_label=row.get("개월"),
                is_new=1 if str(row.get("장기구분", "")).strip() == "신규" else 0,
                source_sheet="재고",
            )
            if rec["lot_no"]:
                inv_records.append(rec)

    # ── 재공 시트 ──
    if wip_sheet:
        df = pd.read_excel(xls, sheet_name=wip_sheet)
        for _, row in df.iterrows():
            item_type = str(row.get("품목구분", "") or "").strip()
            if item_type in EXCLUDED_ITEM_TYPES:
                excluded_count += 1
                continue
            ref_date = _norm_date(row.get("조회기준일"))
            if not ref_date:
                continue
            weight_kg = _to_float(row.get("중량"))
            lot_no = str(row.get("WO No", "") or "").strip()
            rec = dict(
                ref_date=ref_date,
                factory=row.get("공장"),
                item_type=item_type or "재공품",
                item_group=row.get("품목군"),
                item_code=row.get("품목코드"),
                item_name=row.get("품명"),
                cost_center=row.get("원가중심점"),
                cost_center_name=row.get("원가중심점명"),
                lot_no=lot_no,
                wo_no=lot_no,
                qty=_to_float(row.get("수량")),
                weight_kg=weight_kg,
                weight_ton=weight_kg / 1000.0,
                amount=_to_float(row.get("재고금액")),
                base_date=_norm_base_date(row.get("기준일")),
                months_label=row.get("개월"),
                is_new=1 if str(row.get("장기구분", "")).strip() == "신규" else 0,
                source_sheet="재공",
            )
            if rec["lot_no"]:
                wip_records.append(rec)

    # ── 재고_상세(실적) 시트 ──
    if act_inv_sheet:
        df = pd.read_excel(xls, sheet_name=act_inv_sheet)
        for _, row in df.iterrows():
            item_type = str(row.get("품목구분", "") or "").strip()
            if item_type in EXCLUDED_ITEM_TYPES:
                continue
            ref_date = _norm_date(row.get("조회기준일"))
            if not ref_date:
                continue
            weight_raw = _to_float(row.get("중량"))
            rec = dict(
                ref_date=ref_date,
                factory=row.get("공장"),
                item_type=item_type,
                item_group=row.get("품목군"),
                item_code=row.get("품목코드"),
                item_name=row.get("품명"),
                cost_center=row.get("원가중심점"),
                lot_no=str(row.get("Lot No", "") or "").strip(),
                wo_no=row.get("WO No"),
                qty=_to_float(row.get("수량")),
                weight_kg=abs(weight_raw),
                # 음수는 "소진(감소)"을 의미 -> 절대값으로 저장
                weight_ton=abs(weight_raw) / 1000.0,
                actual_type_raw=row.get("유형"),
                actual_type=row.get("유형"),
                process_date=_norm_base_date(row.get("처리일자")),
                processor=row.get("처리자"),
                source_sheet="재고_상세",
            )
            if rec["lot_no"]:
                act_records.append(rec)

    # ── 재공_상세(실적) 시트 ──
    if act_wip_sheet:
        df = pd.read_excel(xls, sheet_name=act_wip_sheet)
        for _, row in df.iterrows():
            item_type = str(row.get("품목구분", "") or "").strip()
            if item_type in EXCLUDED_ITEM_TYPES:
                continue
            ref_date = _norm_date(row.get("조회기준일"))
            if not ref_date:
                continue
            weight_raw = _to_float(row.get("중량"))
            lot_no = str(row.get("WO No", "") or "").strip()
            rec = dict(
                ref_date=ref_date,
                factory=row.get("공장"),
                item_type=item_type or "재공품",
                item_group=row.get("품목군"),
                item_code=row.get("품목코드"),
                item_name=row.get("품명"),
                cost_center=row.get("원가중심점"),
                lot_no=lot_no,
                wo_no=lot_no,
                qty=_to_float(row.get("수량")),
                weight_kg=abs(weight_raw),
                weight_ton=abs(weight_raw) / 1000.0,
                actual_type_raw=row.get("유형"),
                actual_type=row.get("유형"),
                process_date=_norm_base_date(row.get("처리일자")),
                processor=row.get("처리자"),
                source_sheet="재공_상세",
            )
            if rec["lot_no"]:
                act_records.append(rec)

    all_ref_dates = sorted(set(
        [r["ref_date"] for r in inv_records] +
        [r["ref_date"] for r in wip_records] +
        [r["ref_date"] for r in act_records]
    ))
    if not all_ref_dates:
        return {"error": "유효한 조회기준일을 찾을 수 없습니다."}

    latest_ref = all_ref_dates[-1]

    if excluded_count:
        warnings.append(f"저장품 {excluded_count}건은 장기재고관리 대상에서 제외되었습니다.")
    if len(all_ref_dates) > 1:
        warnings.append(
            f"파일 안에 {len(all_ref_dates)}개의 조회기준일({', '.join(all_ref_dates)})이 "
            f"포함되어 있어 각각 별도 스냅샷으로 저장됩니다."
        )

    return {
        "ref_date": latest_ref,
        "all_ref_dates": all_ref_dates,
        "inv_records": inv_records,
        "wip_records": wip_records,
        "act_records": act_records,
        "filename": filename,
        "uploaded_by": uploaded_by,
        "warnings": warnings,
        "excluded_count": excluded_count,
    }


def save_parsed_data(parsed: dict, uploaded_by: str) -> dict:
    """파싱된 데이터를 DB에 저장. 동일 ref_date의 기존 데이터는 먼저 삭제 후 재적재(재업로드 시 중복 방지)."""
    from backend.database import get_conn

    upload_id = uuid.uuid4().hex[:16]
    conn = get_conn()
    try:
        all_ref_dates = parsed["all_ref_dates"]

        # 같은 ref_date가 이미 있으면 삭제 후 재적재 (재업로드 안전성)
        for rd in all_ref_dates:
            conn.execute("DELETE FROM inventory_items WHERE ref_date=?", (rd,))
            conn.execute("DELETE FROM depletion_actuals WHERE ref_date=?", (rd,))

        inv_count = 0
        for rec in parsed["inv_records"] + parsed["wip_records"]:
            conn.execute("""
                INSERT INTO inventory_items
                (upload_id, ref_date, factory, item_type, item_group, item_code, item_name,
                 cost_center, cost_center_name, lot_no, wo_no, qty, weight_kg, weight_ton,
                 amount, qty_consumed, amount_consumed, base_date, months_label, is_new, source_sheet)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                upload_id, rec["ref_date"], rec["factory"], rec["item_type"], rec["item_group"],
                rec["item_code"], rec["item_name"], rec["cost_center"], rec["cost_center_name"],
                rec["lot_no"], rec["wo_no"], rec["qty"], rec["weight_kg"], rec["weight_ton"],
                rec["amount"], 0, 0, rec["base_date"], rec["months_label"], rec["is_new"],
                rec["source_sheet"],
            ))
            inv_count += 1

        act_count = 0
        for rec in parsed["act_records"]:
            conn.execute("""
                INSERT INTO depletion_actuals
                (upload_id, ref_date, factory, item_type, item_group, item_code, item_name,
                 cost_center, lot_no, wo_no, qty, weight_kg, weight_ton,
                 actual_type_raw, actual_type, process_date, processor, source_sheet)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                upload_id, rec["ref_date"], rec["factory"], rec["item_type"], rec["item_group"],
                rec["item_code"], rec["item_name"], rec["cost_center"], rec["lot_no"], rec["wo_no"],
                rec["qty"], rec["weight_kg"], rec["weight_ton"],
                rec["actual_type_raw"], rec["actual_type"], rec["process_date"], rec["processor"],
                rec["source_sheet"],
            ))
            act_count += 1

        latest_rd = parsed["ref_date"]
        total_amount = conn.execute(
            "SELECT COALESCE(SUM(amount),0) FROM inventory_items WHERE ref_date=?", (latest_rd,)
        ).fetchone()[0]
        inv_only = conn.execute(
            "SELECT COUNT(*) FROM inventory_items WHERE ref_date=? AND source_sheet='재고'", (latest_rd,)
        ).fetchone()[0]
        wip_only = conn.execute(
            "SELECT COUNT(*) FROM inventory_items WHERE ref_date=? AND source_sheet='재공'", (latest_rd,)
        ).fetchone()[0]
        act_only = conn.execute(
            "SELECT COUNT(*) FROM depletion_actuals WHERE ref_date=?", (latest_rd,)
        ).fetchone()[0]

        conn.execute("""
            INSERT INTO upload_history
            (upload_id, filename, ref_date, inv_count, wip_count, act_count, total_amount, uploaded_by)
            VALUES (?,?,?,?,?,?,?,?)
        """, (upload_id, parsed["filename"], latest_rd, inv_only, wip_only, act_only,
              total_amount, uploaded_by))

        conn.commit()
        return {
            "upload_id": upload_id,
            "inv_count": inv_only,
            "wip_count": wip_only,
            "act_count": act_only,
            "total_amount": float(total_amount),
        }
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
