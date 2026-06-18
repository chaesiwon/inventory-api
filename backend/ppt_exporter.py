"""
ppt_exporter.py v5 - ILJIN 임원 보고용 PPT 생성
unit_label 파라미터로 억원/백만원/원 표기 통일
"""
import io, logging
from datetime import datetime
from typing import List, Dict, Any, Optional

logger = logging.getLogger(__name__)

NAVY   = (0x0E, 0x28, 0x41)
BLUE   = (0x15, 0x60, 0x82)
ORANGE = (0xE9, 0x71, 0x32)
RED    = (0xC1, 0x00, 0x1B)
GREEN  = (0x19, 0x6B, 0x24)
WHITE  = (0xFF, 0xFF, 0xFF)
LGRAY  = (0xF4, 0xF6, 0xF9)
MGRAY  = (0xD0, 0xD8, 0xE4)
GRAY   = (0x50, 0x50, 0x50)


def _rgb(r, g, b):
    from pptx.dml.color import RGBColor
    return RGBColor(r, g, b)


def _I(n):
    from pptx.util import Inches
    return Inches(n)


def _Pt(n):
    from pptx.util import Pt
    return Pt(n)


def _rect(slide, l, t, w, h, fill, line=None, lw=0.5, rnd=False):
    sh = slide.shapes.add_shape(5 if rnd else 1, _I(l), _I(t), _I(w), _I(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = _rgb(*fill)
    if line:
        sh.line.color.rgb = _rgb(*line); sh.line.width = _Pt(lw)
    else:
        sh.line.fill.background()
    return sh


def _txt(slide, text, l, t, w, h, sz=12, bold=False, clr=GRAY, align=None, italic=False):
    from pptx.enum.text import PP_ALIGN
    if align is None:
        align = PP_ALIGN.LEFT
    bx = slide.shapes.add_textbox(_I(l), _I(t), _I(w), _I(h))
    tf = bx.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = str(text) if text else ""
    run.font.size = _Pt(sz); run.font.bold = bold
    run.font.color.rgb = _rgb(*clr); run.font.italic = italic
    return bx


def _fmt_amt(v, unit_label='억원'):
    try:
        n = float(v or 0)
        return f"{n:,.2f}{unit_label}" if unit_label != '원' else f"{round(n):,}원"
    except Exception:
        return "-"


def _fmt_wt(v):
    try:
        return f"{float(v or 0):.1f}ton"
    except Exception:
        return "-"


def _fmt_cnt(v):
    try:
        return f"{int(v or 0):,}건"
    except Exception:
        return "-"


def _fmt_pct(v):
    try:
        return f"{float(v or 0):.1f}%"
    except Exception:
        return "-"


def _hbar(slide, title, sub=None):
    _rect(slide, 0, 0, 13.33, 1.52, NAVY)
    _rect(slide, 0, 1.48, 13.33, 0.06, ORANGE)
    _txt(slide, title, 0.48, 0.2, 11, 0.78, sz=24, bold=True, clr=WHITE)
    if sub:
        _txt(slide, sub, 0.48, 0.9, 11, 0.42, sz=11, clr=MGRAY, italic=True)


def _foot(slide, n=None):
    from pptx.enum.text import PP_ALIGN
    _rect(slide, 0, 7.2, 13.33, 0.3, NAVY)
    _txt(slide, "장기재고 소진계획 계획/실적 비교 보고서", 0.42, 7.22, 9, 0.26, sz=9, clr=MGRAY)
    if n:
        _txt(slide, str(n), 12.55, 7.22, 0.65, 0.26, sz=9, clr=MGRAY, align=PP_ALIGN.RIGHT)


def _slide_cover(prs, ref_date, generated_at):
    from pptx.enum.text import PP_ALIGN
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, NAVY)
    _rect(s, 9.85, 0, 3.48, 7.5, BLUE)
    _rect(s, 9.8, 0, 0.07, 7.5, ORANGE)
    _txt(s, "장기재고 소진계획", 0.65, 2.15, 9, 0.88, sz=40, bold=True, clr=WHITE)
    _txt(s, "계획/실적 비교 보고서", 0.65, 3.0, 9, 0.88, sz=32, bold=True, clr=WHITE)
    rd = f"{ref_date[:4]}-{ref_date[4:6]}-{ref_date[6:8]}" if len(ref_date) == 8 else ref_date
    _rect(s, 0.65, 3.95, 5.5, 0.07, ORANGE)
    _txt(s, f"기준일: {rd}", 0.65, 4.12, 9, 0.48, sz=15, clr=(0xCC, 0xDD, 0xFF))
    _txt(s, f"생성: {generated_at}", 0.65, 5.0, 9, 0.38, sz=12, clr=MGRAY)


def _slide_kpi(prs, summary, ref_date, unit_label):
    from pptx.enum.text import PP_ALIGN
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _hbar(s, "01  핵심 KPI 요약")
    pt = summary.get("plan_total", 0) or 1
    ac = summary.get("action_count", 0) or 0
    nc = summary.get("no_action_count", 0) or 0
    rate_cnt = summary.get("action_rate", 0)

    kpis = [
        ("계획 등록 LOT", _fmt_cnt(pt), _fmt_amt(summary.get("total_amount"), unit_label), NAVY),
        ("조치 완료", _fmt_cnt(ac), _fmt_amt(summary.get("action_amount"), unit_label), GREEN),
        ("미 조치", _fmt_cnt(nc), _fmt_amt(summary.get("no_action_amount"), unit_label), RED),
        ("달성률", _fmt_pct(rate_cnt), "건수 기준", BLUE),
        ("소진완료금액", _fmt_amt(summary.get("consumed_amount"), unit_label), "LOT단가×실적중량", ORANGE),
    ]
    for i, (lbl, val, sub, col) in enumerate(kpis):
        lx = 0.38 + i * 2.52
        _rect(s, lx, 1.65, 2.35, 1.5, col, rnd=True)
        _txt(s, lbl, lx + 0.1, 1.72, 2.15, 0.5, sz=11, bold=True, clr=(0xCC, 0xDD, 0xFF), align=PP_ALIGN.CENTER)
        _txt(s, val, lx + 0.1, 2.2, 2.15, 0.55, sz=20, bold=True, clr=WHITE, align=PP_ALIGN.CENTER)
        _txt(s, sub, lx + 0.1, 2.78, 2.15, 0.28, sz=9, clr=(0xAA, 0xBB, 0xCC), align=PP_ALIGN.CENTER)

    _rect(s, 0.38, 3.42, 12.55, 0.72, LGRAY, rnd=True)
    fill_w = max(0.3, 12.55 * ac / pt)
    fill_c = GREEN if rate_cnt >= 70 else (0xF5, 0x9E, 0x0B) if rate_cnt >= 40 else RED
    _rect(s, 0.38, 3.42, fill_w, 0.72, fill_c, rnd=True)
    _txt(s, f"달성률 {_fmt_pct(rate_cnt)}  |  조치 {ac:,}건 / 전체 {pt:,}건", 0.55, 3.55, 8, 0.45, sz=14, bold=True, clr=WHITE)
    _foot(s, 2)


def _slide_type_compare(prs, summary, unit_label):
    from pptx.enum.text import PP_ALIGN
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _hbar(s, "02  유형별 계획 vs 실적 비교")
    plan_by = summary.get("plan_by_type", [])
    actual_by = summary.get("actual_by_type", [])
    pM = {r.get("plan_type", "미등록"): r for r in plan_by}
    aM = {r.get("actual_type", "기타"): r for r in actual_by}
    all_types = list({r.get("plan_type", "미등록") for r in plan_by} | {r.get("actual_type", "기타") for r in actual_by})

    hdrs = ["구분", "계획 건수", "계획 중량(ton)", "실적 건수", "실적 중량(ton)", "달성률(중량)"]
    cws = [2.5, 1.7, 1.9, 1.7, 1.9, 1.5]
    y0 = 1.68; rh = 0.48
    xs = [0.3] + [0.3 + sum(cws[:i]) for i in range(1, len(cws))]
    for h, x, cw in zip(hdrs, xs, cws):
        _rect(s, x, y0, cw, rh, NAVY)
        _txt(s, h, x + 0.06, y0 + 0.08, cw - 0.1, rh - 0.1, sz=11, bold=True, clr=WHITE, align=PP_ALIGN.CENTER)

    for ri, typ in enumerate(all_types[:10]):
        pm = pM.get(typ, {}); am = aM.get(typ, {})
        pc = int(pm.get("plan_count", 0)); pw = float(pm.get("plan_weight", 0))
        ac = int(am.get("actual_count", 0)); aw = float(am.get("actual_weight", 0))
        pct = round(aw / pw * 100, 1) if pw > 0 else 0
        bg = LGRAY if ri % 2 == 0 else WHITE
        y = y0 + (ri + 1) * rh
        row_d = [typ or "-", f"{pc:,}", f"{pw:.1f}", f"{ac:,}", f"{aw:.1f}", f"{pct:.1f}%"]
        for ci, (val, x, cw) in enumerate(zip(row_d, xs, cws)):
            _rect(s, x, y, cw, rh, bg, line=MGRAY)
            clr = GREEN if ci == 5 and pct >= 70 else RED if ci == 5 and pct < 40 else GRAY
            _txt(s, val, x + 0.06, y + 0.1, cw - 0.1, rh - 0.1, sz=11, clr=clr, bold=(ci == 0 or ci == 5),
                 align=PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT)
    _foot(s, 3)


def _slide_top_items(prs, items, unit_label):
    from pptx.enum.text import PP_ALIGN
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _hbar(s, "03  고액 상위 장기재고 현황")
    if not items:
        _txt(s, "데이터가 없습니다.", 4, 4, 5, 0.8, sz=18, clr=GRAY)
        _foot(s, 4); return
    hdrs = ["#", "공장", "LOT NO", "품명", f"금액({unit_label})", "계획유형", "조치여부"]
    cws = [0.38, 1.05, 1.85, 3.2, 1.45, 1.55, 1.45]
    y0 = 1.68; rh = 0.4
    xs = [0.28] + [0.28 + sum(cws[:i]) for i in range(1, len(cws))]
    for h, x, cw in zip(hdrs, xs, cws):
        _rect(s, x, y0, cw, rh, NAVY)
        _txt(s, h, x + 0.04, y0 + 0.08, cw - 0.06, rh - 0.1, sz=10, bold=True, clr=WHITE, align=PP_ALIGN.CENTER)
    for ri, item in enumerate(items[:14]):
        amt = f"{float(item.get('amount', 0)):,.2f}"
        act = item.get("action_status", "")
        bg = LGRAY if ri % 2 == 0 else WHITE
        y = y0 + (ri + 1) * rh
        row_d = [str(ri + 1), str(item.get("factory", ""))[:6], str(item.get("lot_no", ""))[:14],
                 str(item.get("item_name", ""))[:18], amt, str(item.get("plan_type", "") or "-")[:10], act]
        for ci, (val, x, cw) in enumerate(zip(row_d, xs, cws)):
            _rect(s, x, y, cw, rh, bg, line=MGRAY)
            ac_clr = GREEN if val == "조치" else RED if val == "미조치" else GRAY
            _txt(s, val, x + 0.04, y + 0.08, cw - 0.06, rh - 0.1, sz=9,
                 clr=ac_clr if ci == 6 else GRAY, bold=(ci == 0 or ci == 6), align=PP_ALIGN.CENTER if ci != 3 else PP_ALIGN.LEFT)
    _foot(s, 4)


def _slide_plan_trend(prs, plan_trend, unit_label):
    from pptx.enum.text import PP_ALIGN
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _hbar(s, "04  월별 소진계획 현황")
    if not plan_trend:
        _txt(s, "소진계획 데이터가 없습니다.", 4, 4, 5, 0.8, sz=16, clr=GRAY)
        _foot(s, 5); return
    hdrs = ["소진계획월", "계획 건수", "계획 중량(ton)", f"계획 금액({unit_label})"]
    cws = [2.6, 2.0, 2.5, 2.5]
    y0 = 1.68; rh = 0.48
    xs = [2.0] + [2.0 + sum(cws[:i]) for i in range(1, len(cws))]
    for h, x, cw in zip(hdrs, xs, cws):
        _rect(s, x, y0, cw, rh, BLUE)
        _txt(s, h, x + 0.06, y0 + 0.08, cw - 0.1, rh - 0.1, sz=12, bold=True, clr=WHITE, align=PP_ALIGN.CENTER)
    for ri, row in enumerate(plan_trend[:10]):
        bg = LGRAY if ri % 2 == 0 else WHITE
        y = y0 + (ri + 1) * rh
        row_d = [str(row.get("plan_month", ""))[:7], f"{int(row.get('plan_count', 0)):,}",
                 f"{float(row.get('plan_weight_ton', 0)):.1f}", f"{float(row.get('plan_amount', 0)):,.2f}"]
        for ci, (val, x, cw) in enumerate(zip(row_d, xs, cws)):
            _rect(s, x, y, cw, rh, bg, line=MGRAY)
            _txt(s, val, x + 0.06, y + 0.1, cw - 0.1, rh - 0.1, sz=11, clr=GRAY, align=PP_ALIGN.CENTER)
    _foot(s, 5)


def _slide_cost_center(prs, cc_items, unit_label):
    from pptx.enum.text import PP_ALIGN
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _hbar(s, "05  원가중심점별 장기재고 현황")
    if not cc_items:
        _txt(s, "데이터가 없습니다.", 4, 4, 5, 0.8, sz=16, clr=GRAY)
        _foot(s, 6); return
    hdrs = ["원가중심점", "건수", "중량(ton)", f"금액({unit_label})", "계획등록", "실적확인", "달성률"]
    cws = [2.6, 0.95, 1.4, 1.4, 1.1, 1.1, 1.35]
    y0 = 1.68; rh = 0.42
    xs = [0.3] + [0.3 + sum(cws[:i]) for i in range(1, len(cws))]
    for h, x, cw in zip(hdrs, xs, cws):
        _rect(s, x, y0, cw, rh, NAVY)
        _txt(s, h, x + 0.04, y0 + 0.06, cw - 0.07, rh - 0.08, sz=10, bold=True, clr=WHITE, align=PP_ALIGN.CENTER)
    for ri, item in enumerate(cc_items[:12]):
        ic = int(item.get("item_count", 0)); pl = int(item.get("plan_count", 0)); ac = int(item.get("actual_count", 0))
        rate = round(ac / ic * 100, 1) if ic > 0 else 0
        bg = LGRAY if ri % 2 == 0 else WHITE
        y = y0 + (ri + 1) * rh
        row_d = [str(item.get("cc_name", ""))[:18], f"{ic:,}", f"{float(item.get('total_weight', 0)):.1f}",
                 f"{float(item.get('total_amount', 0)):,.2f}", f"{pl:,}", f"{ac:,}", f"{rate:.1f}%"]
        for ci, (val, x, cw) in enumerate(zip(row_d, xs, cws)):
            _rect(s, x, y, cw, rh, bg, line=MGRAY)
            rate_clr = GREEN if ci == 6 and rate >= 70 else RED if ci == 6 and rate < 40 else GRAY
            _txt(s, val, x + 0.04, y + 0.08, cw - 0.07, rh - 0.1, sz=10, clr=rate_clr, bold=(ci == 6),
                 align=PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT)
    _foot(s, 6)


def _slide_closing(prs, generated_at):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, 13.33, 7.5, NAVY)
    _rect(s, 9.85, 0, 3.48, 7.5, BLUE)
    _rect(s, 9.8, 0, 0.07, 7.5, ORANGE)
    _txt(s, "감사합니다", 0.65, 2.5, 9, 1.05, sz=52, bold=True, clr=WHITE)
    _txt(s, f"생성일시: {generated_at}", 0.65, 4.0, 6, 0.35, sz=11, clr=(0x60, 0x70, 0x80))


def generate_compare_ppt(
    summary: dict,
    items: List[Dict[str, Any]],
    ref_date: str = "",
    plan_trend: Optional[List] = None,
    cc_items: Optional[List] = None,
    unit_label: str = "억원",
) -> bytes:
    try:
        from pptx import Presentation
        from pptx.util import Inches
    except ImportError:
        raise ImportError("python-pptx 설치 필요")

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    _slide_cover(prs, ref_date, generated_at)
    _slide_kpi(prs, summary, ref_date, unit_label)
    _slide_type_compare(prs, summary, unit_label)
    _slide_top_items(prs, items, unit_label)
    _slide_plan_trend(prs, plan_trend or [], unit_label)
    _slide_cost_center(prs, cc_items or [], unit_label)
    _slide_closing(prs, generated_at)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    result = buf.read()
    logger.info(f"PPT 생성 완료: {len(result):,} bytes, 7슬라이드")
    return result
