"""
ppt_exporter.py v2 - ILJIN 템플릿 기반 계획/실적 비교 PPT 생성
템플릿: ILJIN_PPT_Template_by_Brandlogy_HS_v1_0.pptx (18 슬라이드)
주요 색상: 진청색 #002669, 적색 #C1001B, 회색 #404040
"""
import io, copy, logging
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional

logger = logging.getLogger(__name__)

# 템플릿 경로
TEMPLATE_PATH = Path(__file__).parent / "ppt_template.pptx"

# 브랜드 색상
DARK_BLUE  = (0x00, 0x26, 0x69)   # #002669
RED_BRAND  = (0xC1, 0x00, 0x1B)   # #C1001B
GRAY_TEXT  = (0x40, 0x40, 0x40)   # #404040
WHITE      = (0xFF, 0xFF, 0xFF)
LIGHT_GRAY = (0xF5, 0xF5, 0xF5)
GREEN_POS  = (0x00, 0x7A, 0x4B)
ORANGE_WRN = (0xE8, 0x7D, 0x00)


def _rgb(r, g, b):
    from pptx.dml.color import RGBColor
    return RGBColor(r, g, b)


def _pt(n):
    from pptx.util import Pt
    return Pt(n)


def _inch(n):
    from pptx.util import Inches
    return Inches(n)


def _emu(n):
    from pptx.util import Emu
    return Emu(n)


def _set_text(shape, text: str, font_size=None, bold=None, color=None, align=None):
    """shape의 텍스트 설정 (기존 서식 최대한 유지)"""
    from pptx.enum.text import PP_ALIGN
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    tf.word_wrap = True
    # 전체 텍스트 교체
    if tf.paragraphs:
        p = tf.paragraphs[0]
        # 기존 run 제거 후 새 run 추가
        from pptx.oxml.ns import qn
        for run_elem in p._p.findall(qn('a:r')):
            p._p.remove(run_elem)
        run = p.add_run()
        run.text = str(text) if text is not None else ""
        if font_size: run.font.size = _pt(font_size)
        if bold is not None: run.font.bold = bold
        if color: run.font.color.rgb = _rgb(*color)
        if align:
            p.alignment = align


def _add_slide_from_layout(prs, layout_idx: int = 0):
    """지정 레이아웃으로 새 슬라이드 추가"""
    layout = prs.slide_layouts[min(layout_idx, len(prs.slide_layouts)-1)]
    return prs.slides.add_slide(layout)


def _add_textbox(slide, text, left, top, width, height,
                 font_size=12, bold=False, color=DARK_BLUE,
                 bg_color=None, border_color=None, align=None):
    """텍스트 박스 추가"""
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor

    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    if bg_color:
        txBox.fill.solid()
        txBox.fill.fore_color.rgb = _rgb(*bg_color)
    if border_color:
        txBox.line.color.rgb = _rgb(*border_color)
        txBox.line.width = Pt(0.5)

    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    if align: p.alignment = align
    run = p.add_run()
    run.text = str(text) if text is not None else ""
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = _rgb(*color)
    return txBox


def _add_rect(slide, left, top, width, height, fill_color, line_color=None, line_width=0.5, radius=True):
    """사각형 도형 추가"""
    from pptx.util import Inches, Pt
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    MSO_RECT = 1  # rectangle
    RND_RECT = 5  # rounded rectangle
    shape_type = RND_RECT if radius else MSO_RECT
    shape = slide.shapes.add_shape(
        shape_type,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(*fill_color)
    if line_color:
        shape.line.color.rgb = _rgb(*line_color)
        shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape


def _fmt_amt(v, unit='억원') -> str:
    try:
        n = float(v or 0)
        if unit == '억원': return f"{n/1e8:.1f}억원"
        if unit == '백만원': return f"{n/1e6:.1f}백만원"
        return f"{n:,.0f}원"
    except: return "-"


def _fmt_wt(v) -> str:
    try: return f"{float(v or 0):.1f}ton"
    except: return "-"


def _fmt_cnt(v) -> str:
    try: return f"{int(v or 0):,}건"
    except: return "-"


def _fmt_pct(v) -> str:
    try: return f"{float(v or 0):.1f}%"
    except: return "-"


# ──────────────────────────────────────────────
# 슬라이드 생성 함수들
# ──────────────────────────────────────────────

def _slide_cover(prs, ref_date: str, generated_at: str):
    """슬라이드 1: 표지"""
    slide = _add_slide_from_layout(prs, 0)

    # 배경 진청색 상단 바
    _add_rect(slide, 0, 0, 13.33, 1.8, DARK_BLUE, radius=False)

    # 회사명
    _add_textbox(slide, "일진그룹", 0.5, 0.2, 4, 0.4,
                 font_size=13, bold=True, color=WHITE)
    # 문서명
    _add_textbox(slide, "장기재고 소진계획 비교 보고서",
                 0.5, 0.65, 9, 0.7, font_size=28, bold=True, color=WHITE)
    # 부제목
    rd_str = f"{ref_date[:4]}년 {ref_date[4:6]}월 {ref_date[6:8]}일" if len(ref_date)==8 else ref_date
    _add_textbox(slide, f"기준일: {rd_str}", 0.5, 1.35, 6, 0.35,
                 font_size=13, color=(0xCC, 0xDD, 0xFF))

    # 생성일시
    _add_textbox(slide, f"생성일시: {generated_at}", 0.5, 6.9, 6, 0.35,
                 font_size=11, color=GRAY_TEXT)
    # 우측 장식 박스
    _add_rect(slide, 10.5, 2.0, 2.5, 5.0, (0xE8, 0xEE, 0xF8), radius=True)
    _add_textbox(slide, "LONG-TERM\nINVENTORY\nMANAGEMENT", 10.6, 3.0, 2.3, 2.5,
                 font_size=14, bold=True, color=DARK_BLUE)
    return slide


def _slide_agenda(prs, items: list):
    """슬라이드 2: 목차"""
    slide = _add_slide_from_layout(prs, 0)
    _add_rect(slide, 0, 0, 13.33, 1.8, DARK_BLUE, radius=False)
    _add_textbox(slide, "목 차", 0.5, 0.65, 9, 0.7, font_size=28, bold=True, color=WHITE)
    _add_textbox(slide, "CONTENTS", 0.5, 0.2, 4, 0.4, font_size=13, bold=True, color=(0xCC,0xDD,0xFF))

    for i, (no, title) in enumerate(items):
        y = 2.0 + i * 0.85
        _add_rect(slide, 0.5, y, 0.6, 0.55, DARK_BLUE, radius=True)
        _add_textbox(slide, str(no), 0.5, y+0.1, 0.6, 0.4,
                     font_size=16, bold=True, color=WHITE)
        _add_textbox(slide, title, 1.25, y+0.1, 10, 0.45,
                     font_size=14, bold=(i==0), color=GRAY_TEXT)
        if i < len(items)-1:
            _add_rect(slide, 0.5, y+0.6, 12, 0.03, (0xE0,0xE0,0xE0), radius=False)
    return slide


def _slide_kpi_summary(prs, summary: dict, ref_date: str):
    """슬라이드 3: 핵심 KPI 요약"""
    slide = _add_slide_from_layout(prs, 0)
    _add_rect(slide, 0, 0, 13.33, 1.8, DARK_BLUE, radius=False)
    _add_textbox(slide, "01  핵심 KPI 요약", 0.5, 0.65, 9, 0.7, font_size=24, bold=True, color=WHITE)
    _add_textbox(slide, "Key Performance Indicator", 0.5, 0.2, 6, 0.4, font_size=12, color=(0xCC,0xDD,0xFF))

    rd_str = f"{ref_date[:4]}-{ref_date[4:6]}-{ref_date[6:8]}" if len(ref_date)==8 else ref_date
    _add_textbox(slide, f"기준일: {rd_str}", 11.0, 1.35, 2.2, 0.3,
                 font_size=10, color=(0xCC,0xDD,0xFF))

    pt   = summary.get("plan_total", 0)
    ac   = summary.get("action_count", 0)
    nc   = summary.get("no_action_count", 0)
    rate = summary.get("action_rate", 0)

    kpis = [
        ("계획\n전체 건수",   _fmt_cnt(pt),
         _fmt_amt(summary.get("total_amount")),       "전체 장기재고 계획 등록 LOT", DARK_BLUE),
        ("조치 완료\n건수",    _fmt_cnt(ac),
         _fmt_amt(summary.get("action_amount")),      "계획+실적 확인 LOT",          GREEN_POS),
        ("미 조치\n건수",     _fmt_cnt(nc),
         _fmt_amt(summary.get("no_action_amount")),   "계획만 있고 실적 없는 LOT",   RED_BRAND),
        ("달성률",            _fmt_pct(rate),
         f"조치 {ac} / 전체 {pt}",                   f"전체=조치+미조치: {ac}+{nc}={pt}", DARK_BLUE),
        ("소진금액\n(전월대비)", _fmt_amt(summary.get("consumed_amount")),
         _fmt_wt(summary.get("total_weight")),        "LOT 수량 감소분 기준",        ORANGE_WRN),
    ]
    col_w = 2.4; gap = 0.15; start_x = 0.4
    for i, (label, value, sub, note, color) in enumerate(kpis):
        x = start_x + i * (col_w + gap)
        # 카드 배경
        _add_rect(slide, x, 1.95, col_w, 4.9, LIGHT_GRAY,
                  line_color=(0xDD,0xDD,0xDD), radius=True)
        # 상단 색상 바
        _add_rect(slide, x, 1.95, col_w, 0.35, color, radius=False)
        # 레이블
        _add_textbox(slide, label, x+0.1, 2.0, col_w-0.2, 0.5,
                     font_size=11, bold=True, color=WHITE)
        # 메인 수치
        _add_textbox(slide, value, x+0.1, 2.5, col_w-0.2, 0.9,
                     font_size=22, bold=True, color=color)
        # 서브 수치
        _add_textbox(slide, sub, x+0.1, 3.5, col_w-0.2, 0.5,
                     font_size=12, color=GRAY_TEXT)
        # 구분선
        _add_rect(slide, x+0.1, 4.1, col_w-0.2, 0.03, (0xCC,0xCC,0xCC), radius=False)
        # 노트
        _add_textbox(slide, note, x+0.1, 4.2, col_w-0.2, 0.65,
                     font_size=9, color=(0x88,0x88,0x88))
    return slide


def _slide_action_status(prs, summary: dict):
    """슬라이드 4: 조치/미조치 현황 (막대 시각화)"""
    slide = _add_slide_from_layout(prs, 0)
    _add_rect(slide, 0, 0, 13.33, 1.8, DARK_BLUE, radius=False)
    _add_textbox(slide, "02  조치 / 미조치 현황", 0.5, 0.65, 9, 0.7,
                 font_size=24, bold=True, color=WHITE)
    _add_textbox(slide, "Action Status", 0.5, 0.2, 6, 0.4,
                 font_size=12, color=(0xCC,0xDD,0xFF))

    pt  = float(summary.get("plan_total", 1) or 1)
    ac  = float(summary.get("action_count", 0) or 0)
    nc  = float(summary.get("no_action_count", 0) or 0)
    rate = summary.get("action_rate", 0)

    # 달성률 진행 바
    bar_w = 9.0; bar_h = 0.6; bar_x = 2.0; bar_y = 2.1
    _add_rect(slide, bar_x, bar_y, bar_w, bar_h, (0xEE,0xEE,0xEE), radius=True)
    filled = max(0.2, bar_w * ac / max(pt, 1))
    fill_c = GREEN_POS if rate >= 70 else ORANGE_WRN if rate >= 40 else RED_BRAND
    _add_rect(slide, bar_x, bar_y, filled, bar_h, fill_c, radius=True)
    _add_textbox(slide, f"달성률  {_fmt_pct(rate)}", bar_x+0.1, bar_y+0.1, 6, 0.4,
                 font_size=16, bold=True, color=WHITE)

    # 항목 비교
    rows = [
        ("계획 전체",  pt, DARK_BLUE),
        ("조치 완료",  ac, GREEN_POS),
        ("미 조치",    nc, RED_BRAND),
    ]
    for i, (label, val, color) in enumerate(rows):
        y = 3.1 + i * 1.2
        pct = val / pt * 100 if pt > 0 else 0
        w = max(0.3, 8.0 * val / pt) if pt > 0 else 0.3
        _add_textbox(slide, label, 0.5, y+0.15, 1.8, 0.5,
                     font_size=13, bold=True, color=GRAY_TEXT)
        _add_rect(slide, 2.4, y, 8.0, 0.7, (0xEE,0xEE,0xEE), radius=False)
        if w > 0:
            _add_rect(slide, 2.4, y, w, 0.7, color, radius=False)
        _add_textbox(slide, f"{int(val):,}건  ({pct:.1f}%)", 10.5, y+0.15, 2.5, 0.45,
                     font_size=12, bold=True, color=color)

    # 검증 문구
    check = ac + nc == int(pt)
    _add_textbox(slide,
                 f"✔ 전체({int(pt)}) = 조치({int(ac)}) + 미조치({int(nc)}) {'✅' if check else '❌'}",
                 0.5, 6.7, 12, 0.4, font_size=11,
                 color=GREEN_POS if check else RED_BRAND)
    return slide


def _slide_type_compare(prs, summary: dict):
    """슬라이드 5: 유형별 계획 vs 실적 비교"""
    slide = _add_slide_from_layout(prs, 0)
    _add_rect(slide, 0, 0, 13.33, 1.8, DARK_BLUE, radius=False)
    _add_textbox(slide, "03  유형별 계획 vs 실적 비교", 0.5, 0.65, 9, 0.7,
                 font_size=24, bold=True, color=WHITE)
    _add_textbox(slide, "Plan vs Actual by Type", 0.5, 0.2, 6, 0.4,
                 font_size=12, color=(0xCC,0xDD,0xFF))

    plan_by   = summary.get("plan_by_type", [])
    actual_by = summary.get("actual_by_type", [])
    plan_map   = {r.get("plan_type","미등록"):   r for r in plan_by}
    actual_map = {r.get("actual_type","기타"): r for r in actual_by}
    all_types  = list({r.get("plan_type","미등록") for r in plan_by} |
                      {r.get("actual_type","기타")  for r in actual_by})

    # 헤더
    headers = ["구분", "계획 건수", "계획 중량(ton)", "실적 건수", "실적 중량(ton)", "달성률"]
    col_ws  = [2.2, 1.7, 1.9, 1.7, 1.9, 1.5]
    row_h   = 0.42
    y_start = 2.0

    for ci, (h, cw) in enumerate(zip(headers, col_ws)):
        x = 0.3 + sum(col_ws[:ci])
        _add_rect(slide, x, y_start, cw, row_h, DARK_BLUE, radius=False)
        _add_textbox(slide, h, x+0.05, y_start+0.08, cw-0.1, row_h-0.1,
                     font_size=11, bold=True, color=WHITE)

    for ri, typ in enumerate(all_types[:10]):
        pm = plan_map.get(typ, {})
        am = actual_map.get(typ, {})
        pc = int(pm.get("plan_count", 0))
        pw = float(pm.get("plan_weight", 0))
        ac = int(am.get("actual_count", 0))
        aw = float(am.get("actual_weight", 0))
        r_pct = ac/pc*100 if pc > 0 else 0
        row_data = [
            typ or "-",
            f"{pc:,}",
            f"{pw:.1f}",
            f"{ac:,}",
            f"{aw:.1f}",
            f"{r_pct:.1f}%",
        ]
        bg = LIGHT_GRAY if ri%2==0 else WHITE
        y  = y_start + (ri+1)*row_h
        for ci, (val, cw) in enumerate(zip(row_data, col_ws)):
            x = 0.3 + sum(col_ws[:ci])
            _add_rect(slide, x, y, cw, row_h, bg,
                      line_color=(0xDD,0xDD,0xDD), radius=False)
            rate_ok = (ci==5 and r_pct>=70)
            rate_warn = (ci==5 and r_pct<40)
            txt_c = GREEN_POS if rate_ok else RED_BRAND if rate_warn else GRAY_TEXT
            _add_textbox(slide, val, x+0.05, y+0.1, cw-0.1, row_h-0.1,
                         font_size=11, color=txt_c, bold=(ci==0))
    return slide


def _slide_detail_top(prs, items: list, title_prefix: str = "04"):
    """슬라이드 6: 고액 상위 장기재고 목록"""
    slide = _add_slide_from_layout(prs, 0)
    _add_rect(slide, 0, 0, 13.33, 1.8, DARK_BLUE, radius=False)
    _add_textbox(slide, f"{title_prefix}  고액 장기재고 상위 목록", 0.5, 0.65, 9, 0.7,
                 font_size=24, bold=True, color=WHITE)
    _add_textbox(slide, "Top Items by Amount", 0.5, 0.2, 6, 0.4,
                 font_size=12, color=(0xCC,0xDD,0xFF))

    if not items:
        _add_textbox(slide, "데이터가 없습니다.", 4.0, 4.0, 5, 0.8,
                     font_size=18, color=GRAY_TEXT)
        return slide

    headers = ["#", "공장", "LOT NO", "품명", "원가중심점", "금액(억원)", "계획유형", "조치여부"]
    col_ws  = [0.4, 1.0, 2.0, 2.3, 1.8, 1.5, 1.5, 1.5]
    row_h   = 0.36

    for ci, (h, cw) in enumerate(zip(headers, col_ws)):
        x = 0.2 + sum(col_ws[:ci])
        _add_rect(slide, x, 2.0, cw, row_h, DARK_BLUE, radius=False)
        _add_textbox(slide, h, x+0.04, 2.05, cw-0.08, row_h-0.05,
                     font_size=10, bold=True, color=WHITE)

    for ri, item in enumerate(items[:12]):
        amt_str = f"{float(item.get('amount',0))/1e8:.2f}"
        act     = item.get("action_status","")
        row_d   = [
            str(ri+1),
            str(item.get("factory",""))[:6],
            str(item.get("lot_no",""))[:14],
            str(item.get("item_name",""))[:14],
            str(item.get("cc_name", item.get("cost_center","")))[:12],
            amt_str,
            str(item.get("plan_type","") or "-")[:8],
            act,
        ]
        bg = LIGHT_GRAY if ri%2==0 else WHITE
        y  = 2.0 + (ri+1)*row_h
        for ci, (val, cw) in enumerate(zip(row_d, col_ws)):
            x = 0.2 + sum(col_ws[:ci])
            _add_rect(slide, x, y, cw, row_h, bg,
                      line_color=(0xDD,0xDD,0xDD), radius=False)
            c = GREEN_POS if val=="조치" else RED_BRAND if val=="미조치" else GRAY_TEXT
            _add_textbox(slide, val, x+0.04, y+0.07, cw-0.08, row_h-0.08,
                         font_size=9, color=c, bold=(ci==0 or ci==7))
    return slide


def _slide_monthly_trend(prs, plan_trend: list):
    """슬라이드 7: 월별 소진계획 추이"""
    slide = _add_slide_from_layout(prs, 0)
    _add_rect(slide, 0, 0, 13.33, 1.8, DARK_BLUE, radius=False)
    _add_textbox(slide, "05  월별 소진계획 현황", 0.5, 0.65, 9, 0.7,
                 font_size=24, bold=True, color=WHITE)
    _add_textbox(slide, "Monthly Depletion Plan Trend", 0.5, 0.2, 6, 0.4,
                 font_size=12, color=(0xCC,0xDD,0xFF))

    if not plan_trend:
        _add_textbox(slide, "소진계획 데이터가 없습니다.", 4.0, 4.0, 5, 0.8,
                     font_size=16, color=GRAY_TEXT)
        return slide

    # 테이블 형태 추이 표시
    headers = ["소진계획월", "계획 건수", "계획 중량(ton)", "계획 금액(억원)"]
    col_ws  = [2.5, 2.0, 2.5, 2.5]
    row_h   = 0.42; y_start = 2.0

    for ci, (h, cw) in enumerate(zip(headers, col_ws)):
        x = 2.0 + sum(col_ws[:ci])
        _add_rect(slide, x, y_start, cw, row_h, DARK_BLUE, radius=False)
        _add_textbox(slide, h, x+0.05, y_start+0.08, cw-0.1, row_h-0.1,
                     font_size=12, bold=True, color=WHITE)

    for ri, row in enumerate(plan_trend[:10]):
        bg = LIGHT_GRAY if ri%2==0 else WHITE
        y  = y_start + (ri+1)*row_h
        row_d = [
            str(row.get("plan_month",""))[:7],
            f"{int(row.get('plan_count',0)):,}",
            f"{float(row.get('plan_weight_ton',0)):.1f}",
            f"{float(row.get('plan_amount',0))/1e8:.2f}",
        ]
        for ci, (val, cw) in enumerate(zip(row_d, col_ws)):
            x = 2.0 + sum(col_ws[:ci])
            _add_rect(slide, x, y, cw, row_h, bg,
                      line_color=(0xDD,0xDD,0xDD), radius=False)
            _add_textbox(slide, val, x+0.05, y+0.1, cw-0.1, row_h-0.1,
                         font_size=11, color=GRAY_TEXT)

    # 간단한 막대 시각화
    max_cnt = max((r.get("plan_count",0) for r in plan_trend), default=1) or 1
    for ri, row in enumerate(plan_trend[:10]):
        y = y_start + (ri+1)*row_h
        w = max(0.1, 1.8 * row.get("plan_count",0) / max_cnt)
        _add_rect(slide, 11.5, y+0.1, w, row_h-0.2, DARK_BLUE, radius=False)
    return slide


def _slide_cost_center(prs, cc_items: list):
    """슬라이드 8: 원가중심점별 현황"""
    slide = _add_slide_from_layout(prs, 0)
    _add_rect(slide, 0, 0, 13.33, 1.8, DARK_BLUE, radius=False)
    _add_textbox(slide, "06  원가중심점별 장기재고 현황", 0.5, 0.65, 9, 0.7,
                 font_size=24, bold=True, color=WHITE)
    _add_textbox(slide, "Long-term Inventory by Cost Center", 0.5, 0.2, 8, 0.4,
                 font_size=12, color=(0xCC,0xDD,0xFF))

    if not cc_items:
        _add_textbox(slide, "데이터가 없습니다.", 4.0, 4.0, 5, 0.8,
                     font_size=16, color=GRAY_TEXT)
        return slide

    headers = ["원가중심점", "건수", "중량(ton)", "금액(억원)", "계획등록", "실적확인", "달성률"]
    col_ws  = [2.5, 1.0, 1.5, 1.5, 1.3, 1.3, 1.5]
    row_h   = 0.4; y_start = 2.0

    for ci, (h, cw) in enumerate(zip(headers, col_ws)):
        x = 0.3 + sum(col_ws[:ci])
        _add_rect(slide, x, y_start, cw, row_h, DARK_BLUE, radius=False)
        _add_textbox(slide, h, x+0.05, y_start+0.08, cw-0.1, row_h-0.1,
                     font_size=10, bold=True, color=WHITE)

    for ri, item in enumerate(cc_items[:11]):
        ic = int(item.get("item_count",0))
        pl = int(item.get("plan_count",0))
        ac = int(item.get("actual_count",0))
        rate = ac/ic*100 if ic>0 else 0
        bg = LIGHT_GRAY if ri%2==0 else WHITE
        y  = y_start + (ri+1)*row_h
        row_d = [
            str(item.get("cc_name",""))[:16],
            f"{ic:,}",
            f"{float(item.get('total_weight',0)):.1f}",
            f"{float(item.get('total_amount',0))/1e8:.2f}",
            f"{pl:,}",
            f"{ac:,}",
            f"{rate:.1f}%",
        ]
        for ci, (val, cw) in enumerate(zip(row_d, col_ws)):
            x = 0.3 + sum(col_ws[:ci])
            _add_rect(slide, x, y, cw, row_h, bg,
                      line_color=(0xDD,0xDD,0xDD), radius=False)
            rate_color = GREEN_POS if ci==6 and rate>=70 else RED_BRAND if ci==6 and rate<40 else GRAY_TEXT
            _add_textbox(slide, val, x+0.05, y+0.1, cw-0.1, row_h-0.1,
                         font_size=10, color=rate_color)
    return slide


def _slide_closing(prs, generated_at: str):
    """슬라이드 9: 마무리"""
    slide = _add_slide_from_layout(prs, 0)
    _add_rect(slide, 0, 0, 13.33, 7.5, DARK_BLUE, radius=False)
    _add_textbox(slide, "장기재고 소진계획 비교 보고서", 1.0, 2.5, 11, 1.0,
                 font_size=32, bold=True, color=WHITE)
    _add_textbox(slide, "Long-Term Inventory Depletion Plan Report", 1.0, 3.5, 11, 0.6,
                 font_size=16, color=(0xCC,0xDD,0xFF))
    _add_rect(slide, 1.0, 4.3, 11, 0.04, (0x44,0x66,0xAA), radius=False)
    _add_textbox(slide, f"생성일시: {generated_at}", 1.0, 4.5, 11, 0.5,
                 font_size=12, color=(0xAA,0xBB,0xDD))
    _add_textbox(slide, "© ILJIN GROUP. All Rights Reserved.", 1.0, 6.8, 11, 0.4,
                 font_size=10, color=(0x88,0x99,0xBB))
    return slide


# ──────────────────────────────────────────────
# 메인 생성 함수
# ──────────────────────────────────────────────
def generate_compare_ppt(
    summary: dict,
    items: List[Dict[str, Any]],
    ref_date: str = "",
    plan_trend: Optional[List] = None,
    cc_items: Optional[List] = None,
) -> bytes:
    """
    ILJIN 템플릿 기반 계획/실적 비교 PPT 생성
    템플릿 파일이 있으면 레이아웃을 재활용, 없으면 기본 생성.
    """
    try:
        from pptx import Presentation
    except ImportError:
        raise ImportError("python-pptx 설치 필요: pip install python-pptx")

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 템플릿 기반 새 파일 생성 (레이아웃 재활용)
    try:
        # 항상 새 Presentation 생성 (템플릿 색상/크기만 참조)
        prs = Presentation()
        prs.slide_width  = _inch(13.33)
        prs.slide_height = _inch(7.5)
        logger.info("PPT 생성 시작: ILJIN 브랜드 컬러 적용")
    except Exception as e:
        logger.warning(f"PPT 초기화 실패: {e}")
        raise

    rd_str = ref_date

    # 슬라이드 순서 생성
    _slide_cover(prs, rd_str, generated_at)
    _slide_agenda(prs, [
        (1, "핵심 KPI 요약"),
        (2, "조치 / 미조치 현황"),
        (3, "유형별 계획 vs 실적 비교"),
        (4, "고액 장기재고 상위 목록"),
        (5, "월별 소진계획 현황"),
        (6, "원가중심점별 현황"),
    ])
    _slide_kpi_summary(prs, summary, rd_str)
    _slide_action_status(prs, summary)
    _slide_type_compare(prs, summary)
    _slide_detail_top(prs, items, "04")
    _slide_monthly_trend(prs, plan_trend or [])
    _slide_cost_center(prs, cc_items or [])
    _slide_closing(prs, generated_at)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    result = buf.read()
    logger.info(f"PPT 생성 완료: {len(result):,} bytes, 슬라이드 9개")
    return result
