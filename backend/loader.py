"""
loader.py v3
[v3 추가]
 - qty_consumed / amount_consumed: 전월 대비 감소 수량/금액 계산
 - LOT NO 기준으로 전월 재고 vs 당월 재고 비교
 - 원가중심점명 우선 사용
"""
import io
import pandas as pd
import numpy as np
from datetime import datetime

KG_TO_TON = 0.001

def map_actual_type(raw):
    s = str(raw).strip() if raw and str(raw).strip() not in ("","nan") else ""
    return "생산투입" if s in ("WIP Issue","WIP Completion") else "기타"

def _s(v, default=""):
    if v is None: return default
    if isinstance(v, float) and np.isnan(v): return default
    s = str(v).strip()
    return default if s in ("","nan","None","\t") else s

def _f(v, default=0.0):
    try:
        f = float(v); return default if np.isnan(f) else f
    except: return default


def parse_inventory_file(file_bytes, filename, uploaded_by="system"):
    upload_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    warnings = []
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception as e:
        return {"error": f"Excel 파일 열기 실패: {e}"}

    for sh in ["장기재고현황_재고","장기재고현황_재공"]:
        if sh not in xl.sheet_names:
            return {"error": f"필수 시트 없음: {sh}"}

    inventory, actuals, ref_date = [], [], ""

    # ── Sheet1: 재고
    df1 = pd.read_excel(xl, sheet_name="장기재고현황_재고", header=0, dtype=str).dropna(how="all")
    for _, row in df1.iterrows():
        rd = _s(row.get("조회기준일",""))
        if rd and not ref_date: ref_date = rd
        jg = _s(row.get("장기구분",""))
        if jg not in ("신규","기존"): continue
        lot = _s(row.get("Lot No",""))
        if not lot: continue
        wt_kg = _f(row.get("중량",0))
        cc    = _s(row.get("원가중심점",""))
        ccn   = _s(row.get("원가중심점명","")) or cc
        inventory.append({
            "ref_date":rd or ref_date, "factory":_s(row.get("공장","")),
            "item_type":_s(row.get("품목구분","")), "item_group":_s(row.get("품목군","")),
            "item_code":_s(row.get("품목코드","")), "item_name":_s(row.get("품명","")),
            "cost_center":cc, "cost_center_name":ccn,
            "lot_no":lot, "wo_no":_s(row.get("WO No","")),
            "qty":_f(row.get("수량",0)), "weight_kg":wt_kg,
            "weight_ton":wt_kg*KG_TO_TON,
            "amount":_f(row.get("재고금액",0)),
            "qty_consumed":0.0, "amount_consumed":0.0,
            "base_date":_s(row.get("기준일","")), "months_label":_s(row.get("개월","")),
            "is_new":1 if jg=="신규" else 0,
            "source_sheet":"재고", "upload_id":upload_id,
        })

    # ── Sheet2: 재공
    df2 = pd.read_excel(xl, sheet_name="장기재고현황_재공", header=0, dtype=str).dropna(how="all")
    for _, row in df2.iterrows():
        rd = _s(row.get("조회기준일",""))
        if rd and not ref_date: ref_date = rd
        jg = _s(row.get("장기구분",""))
        if jg not in ("신규","기존"): continue
        lot = _s(row.get("Lot No","")) or _s(row.get("WO No",""))
        if not lot: continue
        wt_kg = _f(row.get("중량",0))
        cc    = _s(row.get("원가중심점",""))
        ccn   = _s(row.get("원가중심점명","")) or cc
        inventory.append({
            "ref_date":rd or ref_date, "factory":_s(row.get("공장","")),
            "item_type":"재공품", "item_group":_s(row.get("품목군","")),
            "item_code":_s(row.get("품목코드","")), "item_name":_s(row.get("품명","")),
            "cost_center":cc, "cost_center_name":ccn,
            "lot_no":lot, "wo_no":_s(row.get("WO No","")),
            "qty":_f(row.get("수량",0)), "weight_kg":wt_kg,
            "weight_ton":wt_kg*KG_TO_TON,
            "amount":_f(row.get("재고금액",0)),
            "qty_consumed":0.0, "amount_consumed":0.0,
            "base_date":_s(row.get("기준일","")), "months_label":_s(row.get("개월","")),
            "is_new":1 if jg=="신규" else 0,
            "source_sheet":"재공", "upload_id":upload_id,
        })

    # ── Sheet3: 재고 상세 (소진실적)
    if "장기재고현황_재고_상세" in xl.sheet_names:
        df3 = pd.read_excel(xl, sheet_name="장기재고현황_재고_상세", header=0, dtype=str).dropna(how="all")
        for _, row in df3.iterrows():
            lot = _s(row.get("Lot No",""))
            if not lot: continue
            wt_kg = _f(row.get("중량",0)); raw_t = _s(row.get("유형",""))
            actuals.append({
                "ref_date":_s(row.get("조회기준일",ref_date)),
                "factory":_s(row.get("공장","")), "item_type":_s(row.get("품목구분","")),
                "item_group":_s(row.get("품목군","")), "item_code":_s(row.get("품목코드","")),
                "item_name":_s(row.get("품명","")), "cost_center":_s(row.get("원가중심점","")),
                "lot_no":lot, "wo_no":_s(row.get("WO No","")),
                "qty":_f(row.get("수량",0)), "weight_kg":wt_kg,
                "weight_ton":wt_kg*KG_TO_TON,
                "qty_consumed":0.0, "amount_consumed":0.0,
                "actual_type_raw":raw_t, "actual_type":map_actual_type(raw_t),
                "process_date":_s(row.get("처리일자","")), "processor":_s(row.get("처리자","")),
                "source_sheet":"재고_상세", "upload_id":upload_id,
            })

    # ── Sheet4: 재공 상세 (소진실적)
    if "장기재고현황_재공_상세" in xl.sheet_names:
        df4 = pd.read_excel(xl, sheet_name="장기재고현황_재공_상세", header=0, dtype=str).dropna(how="all")
        for _, row in df4.iterrows():
            lot = _s(row.get("WO No",""))
            if not lot: continue
            wt_kg = _f(row.get("중량",0)); raw_t = _s(row.get("유형",""))
            actuals.append({
                "ref_date":_s(row.get("조회기준일",ref_date)),
                "factory":_s(row.get("공장","")), "item_type":"재공품",
                "item_group":_s(row.get("품목군","")), "item_code":_s(row.get("품목코드","")),
                "item_name":_s(row.get("품명","")), "cost_center":_s(row.get("원가중심점","")),
                "lot_no":lot, "wo_no":lot,
                "qty":_f(row.get("수량",0)), "weight_kg":wt_kg,
                "weight_ton":wt_kg*KG_TO_TON,
                "qty_consumed":0.0, "amount_consumed":0.0,
                "actual_type_raw":raw_t, "actual_type":map_actual_type(raw_t),
                "process_date":_s(row.get("처리일자","")), "processor":_s(row.get("처리자","")),
                "source_sheet":"재공_상세", "upload_id":upload_id,
            })

    return {"upload_id":upload_id, "ref_date":ref_date,
            "inventory":inventory, "actuals":actuals,
            "warnings":warnings, "filename":filename}


def _calc_consumed(parsed, conn):
    """
    전월 대비 수량 감소 → 소진 수량/금액 계산
    LOT NO 기준: 전월 qty - 당월 qty = 소진 qty
    소진 금액 = 소진 qty / 당월 qty * 당월 amount
    """
    rd = parsed["ref_date"]
    # 이전 기준일 조회
    prev_row = conn.execute(
        "SELECT ref_date FROM inventory_items WHERE ref_date < ? ORDER BY ref_date DESC LIMIT 1",
        (rd,)
    ).fetchone()
    if not prev_row:
        return  # 전월 데이터 없으면 건너뜀

    prev_rd = prev_row["ref_date"]
    prev_data = {
        r["lot_no"]: {"qty": r["qty"], "amount": r["amount"]}
        for r in conn.execute(
            "SELECT lot_no, qty, amount FROM inventory_items WHERE ref_date=?", (prev_rd,)
        ).fetchall()
    }

    for item in parsed["inventory"]:
        lot = item["lot_no"]
        if lot in prev_data:
            prev_qty = prev_data[lot]["qty"]
            curr_qty = item["qty"]
            consumed = max(0, prev_qty - curr_qty)  # 감소분만
            if consumed > 0 and curr_qty > 0:
                # 단가 = 당월 금액 / 당월 수량
                unit_price = item["amount"] / curr_qty if curr_qty else 0
                item["qty_consumed"]    = consumed
                item["amount_consumed"] = round(consumed * unit_price, 0)
            elif consumed > 0 and curr_qty == 0:
                # 당월에 아예 없어진 경우
                prev_amt = prev_data[lot]["amount"]
                item["qty_consumed"]    = consumed
                item["amount_consumed"] = round(prev_amt * (consumed / prev_qty), 0)


def save_parsed_data(parsed, uploaded_by="system"):
    from backend.database import get_conn
    conn = get_conn()
    uid  = parsed["upload_id"]
    rd   = parsed["ref_date"]

    # 소진 수량/금액 계산 (전월 데이터 있을 때만)
    _calc_consumed(parsed, conn)

    conn.execute("DELETE FROM inventory_items    WHERE ref_date=?", (rd,))
    conn.execute("DELETE FROM depletion_actuals  WHERE ref_date=?", (rd,))

    for item in parsed["inventory"]:
        conn.execute("""INSERT INTO inventory_items
            (upload_id,ref_date,factory,item_type,item_group,item_code,item_name,
             cost_center,cost_center_name,lot_no,wo_no,qty,weight_kg,weight_ton,
             amount,qty_consumed,amount_consumed,base_date,months_label,is_new,source_sheet)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            [item.get(c) for c in
             ("upload_id","ref_date","factory","item_type","item_group","item_code","item_name",
              "cost_center","cost_center_name","lot_no","wo_no","qty","weight_kg","weight_ton",
              "amount","qty_consumed","amount_consumed","base_date","months_label","is_new","source_sheet")])

    for act in parsed["actuals"]:
        conn.execute("""INSERT INTO depletion_actuals
            (upload_id,ref_date,factory,item_type,item_group,item_code,item_name,
             cost_center,lot_no,wo_no,qty,weight_kg,weight_ton,
             qty_consumed,amount_consumed,actual_type_raw,actual_type,
             process_date,processor,source_sheet)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            [act.get(c) for c in
             ("upload_id","ref_date","factory","item_type","item_group","item_code","item_name",
              "cost_center","lot_no","wo_no","qty","weight_kg","weight_ton",
              "qty_consumed","amount_consumed","actual_type_raw","actual_type",
              "process_date","processor","source_sheet")])

    inv_cnt = sum(1 for i in parsed["inventory"] if i.get("source_sheet")=="재고")
    wip_cnt = sum(1 for i in parsed["inventory"] if i.get("source_sheet")=="재공")
    total_a = sum(i.get("amount",0) for i in parsed["inventory"])
    conn.execute("""INSERT OR REPLACE INTO upload_history
        (upload_id,filename,ref_date,inv_count,wip_count,act_count,total_amount,uploaded_by)
        VALUES (?,?,?,?,?,?,?,?)""",
        (uid,parsed.get("filename",""),rd,inv_cnt,wip_cnt,len(parsed["actuals"]),total_a,uploaded_by))
    conn.commit(); conn.close()
    return {"upload_id":uid,"inv_count":inv_cnt,"wip_count":wip_cnt,
            "act_count":len(parsed["actuals"]),"total_amount":total_a}
