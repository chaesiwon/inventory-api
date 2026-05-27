"""
sample_generator.py - 테스트용 샘플 Excel 생성
실제 업로드 파일(장기재고현황_NIJ_*.xlsx)과 동일한 시트/컬럼 구조
"""
import io, random
from datetime import date, timedelta
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

random.seed(42)

FACTORIES    = ["임실공장", "수원공장"]
ITEM_TYPES   = ["제품", "원자재", "반제품"]
COST_CENTERS = {
    "임실공장": ["CC-임실-001", "CC-임실-002", "CC-임실-003", "CC-임실-004"],
    "수원공장": ["CC-수원-001", "CC-수원-002", "CC-수원-003"],
}
CC_NAMES = {
    "CC-임실-001": "임실압연1팀", "CC-임실-002": "임실압연2팀",
    "CC-임실-003": "임실가공팀",  "CC-임실-004": "임실품질팀",
    "CC-수원-001": "수원조관팀",  "CC-수원-002": "수원가공팀",
    "CC-수원-003": "수원품질팀",
}
ITEM_GROUPS  = ["탄소강관", "STS관", "합금강관", "특수강관", "배관용강관"]
GRADES       = ["ASTM A106B","API 5L X65","SUS304","SUS316","SA213 T22","STPG370","STS410"]


def _make_lot(factory: str, idx: int) -> str:
    prefix = "ST" if factory == "임실공장" else "SW"
    return f"{prefix}2502{idx:06d}"


def _make_item_code(factory: str, idx: int) -> str:
    prefix = "1" if factory == "임실공장" else "2"
    suffixes = ["ERD","LSA","VAB","STS","PAI"]
    return f"{prefix}{random.choice(suffixes)}{idx:06d}A"


def _make_row(idx: int, ref_date_str: str, factory: str, months_label: str, is_new_label: str):
    cc = random.choice(COST_CENTERS[factory])
    item_type = random.choices(ITEM_TYPES, weights=[60,30,10])[0]
    item_code = _make_item_code(factory, idx)
    lot_no    = _make_lot(factory, idx)
    grade     = random.choice(GRADES)
    weight_kg = round(random.uniform(500, 80000), 3)   # kg 단위
    unit_p    = round(random.uniform(3000, 25000), 2)
    amount    = round(weight_kg * unit_p, 0)
    base_dt   = (date.today() - timedelta(days=random.randint(90, 500))).strftime("%Y%m%d")
    months    = months_label

    return {
        "조회기준일":   ref_date_str,
        "공장":        factory,
        "창고유형":    "완제품창고" if item_type == "제품" else "원자재창고",
        "품목구분":    item_type,
        "품목군":      random.choice(ITEM_GROUPS),
        "품목코드":    item_code,
        "품명":        f"{grade} {random.choice([21.7,34.0,48.6,76.1,114.3])}x{random.choice([2.3,3.0,4.5,6.35])}",
        "원가중심점":  cc,
        "원가중심점명": CC_NAMES.get(cc, cc),
        "Grade":       grade,
        "Lot No":      lot_no,
        "WO No":       f"WO{idx:07d}",
        "수량":        random.randint(1, 50),
        "중량":        weight_kg,        # kg 단위 (N열)
        "재고금액":    amount,           # 금액 (O열)
        "기준일":      base_dt,
        "개월":        months,
        "장기구분":    is_new_label,     # R열: 신규/기존
    }


def create_sample_xlsx(n_inv: int = 200, n_wip: int = 30) -> bytes:
    """
    실제 업로드 파일과 동일한 구조의 샘플 Excel 생성.
    시트: 장기재고현황_재고, 장기재고현황_재공, 장기재고현황_재고_상세, 장기재고현황_재공_상세
    """
    today_str = date.today().strftime("%Y%m%d")

    # ── Sheet1: 장기재고현황_재고
    inv_rows = []
    for i in range(1, n_inv + 1):
        factory = random.choice(FACTORIES)
        months  = random.choice(["6개월이상", "12개월이상", "24개월이상"])
        is_new  = random.choices(["신규", "기존", ""], weights=[25, 60, 15])[0]
        inv_rows.append(_make_row(i, today_str, factory, months, is_new))

    # ── Sheet2: 장기재고현황_재공 (재공품)
    wip_rows = []
    for i in range(n_inv + 1, n_inv + n_wip + 1):
        factory = random.choice(FACTORIES)
        cc      = random.choice(COST_CENTERS[factory])
        is_new  = random.choices(["신규", "기존", ""], weights=[30, 55, 15])[0]
        weight_kg = round(random.uniform(200, 30000), 3)
        amount    = round(weight_kg * random.uniform(4000, 20000), 0)
        wip_rows.append({
            "조회기준일":   today_str,
            "공장":        factory,
            "품목구분":    "재공품",
            "품목군":      random.choice(ITEM_GROUPS),
            "품목코드":    _make_item_code(factory, i),
            "품명":        f"재공-{random.choice(GRADES)}",
            "원가중심점":  cc,
            "원가중심점명": CC_NAMES.get(cc, cc),
            "WO No":       f"WO{i:07d}",
            "재공상태":    random.choice(["진행중", "대기"]),
            "공정":        random.choice(["압연", "절단", "검사"]),
            "공정명":      random.choice(["압연공정", "절단공정", "검사공정"]),
            "수량":        random.randint(1, 20),
            "중량":        weight_kg,
            "재고금액":    amount,
            "기준일":      (date.today() - timedelta(days=random.randint(90, 400))).strftime("%Y%m%d"),
            "개월":        random.choice(["6개월이상", "12개월이상"]),
            "장기구분":    is_new,
        })

    # ── Sheet3: 재고 상세 (소진실적)
    types3 = ["WIP Issue", "Account alias issue (INV_SG_22213)",
              "Account alias issue (INV_LOT_CHANGE)", "WIP Completion"]
    valid_inv = [r for r in inv_rows if r["장기구분"] in ("신규", "기존")]
    detail_inv_rows = []
    for r in random.sample(valid_inv, min(40, len(valid_inv))):
        raw_type = random.choice(types3)
        detail_inv_rows.append({
            "조회기준일": today_str,
            "공장": r["공장"],
            "품목구분": r["품목구분"],
            "품목군": r["품목군"],
            "품목코드": r["품목코드"],
            "품명": r["품명"],
            "원가중심점": r["원가중심점"],
            "원가중심점명": r["원가중심점명"],
            "Lot No": r["Lot No"],
            "WO No": r["WO No"],
            "수량": r["수량"],
            "중량": round(r["중량"] * random.uniform(0.3, 1.0), 3),
            "유형": raw_type,
        })

    # ── Sheet4: 재공 상세
    valid_wip = [r for r in wip_rows if r["장기구분"] in ("신규", "기존")]
    detail_wip_rows = []
    for r in random.sample(valid_wip, min(15, len(valid_wip))):
        detail_wip_rows.append({
            "조회기준일": today_str,
            "공장": r["공장"],
            "품목코드": r["품목코드"],
            "품명": r["품명"],
            "원가중심점": r["원가중심점"],
            "WO No": r["WO No"],
            "수량": r["수량"],
            "중량": round(r["중량"] * random.uniform(0.3, 1.0), 3),
            "유형": "WIP Completion",
        })

    # Excel 생성
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(inv_rows).to_excel(writer, sheet_name="장기재고현황_재고",    index=False)
        pd.DataFrame(wip_rows).to_excel(writer, sheet_name="장기재고현황_재공",    index=False)
        pd.DataFrame(detail_inv_rows).to_excel(writer, sheet_name="장기재고현황_재고_상세", index=False)
        pd.DataFrame(detail_wip_rows).to_excel(writer, sheet_name="장기재고현황_재공_상세", index=False)

    buf.seek(0)
    return buf.getvalue()
