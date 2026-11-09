# -*- coding: utf-8 -*-
"""sheet.json → 심사평가표(1쪽) + 심사의견요약표(2쪽) .docx

표 스타일·저수준 헬퍼는 검토의견서프레임워크 render_docx.py를 그대로 재사용한다.
사용법: python build.py <sheet.json> [-o 출력.docx] [--draft]
"""
import argparse
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from sheetcalc import LOAN_FW, compute_sheet, gaps_sheet, merge_loans  # noqa: E402

sys.path.insert(0, str(LOAN_FW))
from calc import comma, empty, get, rate_str  # noqa: E402
from render_docx import (CENTER, FONT, HEAD_BG, LEFT, RIGHT, W, borders,  # noqa: E402
                         cell_text, fixed, font, no_borders, para)
from docx.enum.text import WD_TAB_ALIGNMENT  # noqa: E402

BLUE = RGBColor(0x00, 0x00, 0xC0)

# (항목, [(착안사항, 값키)]) — 값키는 sheet.평가표의 키이거나 calc가 만든 값
FORM = [
    ("신용등급", [("신용조사 구분", "신용조사구분"), ("신용등급", "@신용등급")]),
    ("적법성 심사", [("법적 제한 사항(대리행위 또는\n이사회의결 적정성 등)", "법적제한사항"),
                 ("업무방법 상 채무관계자 자격", "채무관계자자격")]),
    ("대출조건\n심사", [("대출기간", "@대출기간"), ("금리(요율)", "@금리")]),
    ("채권보전\n심사", [("담보물 적정 여부", "담보물적정"), ("보증인 적정 여부", "보증인적정")]),
    ("소요자금\n심사", [("자금용도", "@자금용도"), ("소요자금 한도", "소요자금한도")]),
    ("대출한도", [("동일인당 대출한도 초과 또는\n편중대출 여부", "편중대출여부")]),
    ("수익성", [("향후 거래 전망 및\n조합수지 기여 정도", "@이자수익")]),
    ("사업계획\n심사", [("사업계획", "사업계획"), ("자금수지계획", "자금수지계획")]),
    ("안정성", [("업계현황 및 전망", "업계현황전망"), ("현금흐름의 상태", "현금흐름상태")]),
    ("주석사항 심사", [("주석사항 확인결과", "주석사항")]),
    ("기  타", [("기타 특이사항", "기타특이사항")]),
]
BLUE_KEYS = {"@대출기간", "@금리", "@자금용도"}


def approval_box(doc, meta):
    appr = meta.get("결재선") or ["계원", "실장", "본부장", "상임이사"]
    t = doc.add_table(rows=1, cols=2)
    no_borders(t)
    fixed(t, [W - 16 * len(appr) - 2, 16 * len(appr) + 2])
    at = t.rows[0].cells[1].add_table(rows=2, cols=len(appr))
    borders(at, sz=4, outer=6)
    fixed(at, [16] * len(appr))
    for i, a in enumerate(appr):
        cell_text(at.rows[0].cells[i], a, 8.5)
    at.rows[1].height, at.rows[1].height_rule = Mm(10), WD_ROW_HEIGHT_RULE.EXACTLY
    return t


def title_block(doc, meta, page):
    p = para(doc, space_before=2, space_after=1, align=CENTER)
    mark = {"평가": ("■", "□"), "요약": ("□", "■")}.get(meta.get("체크", ""), ("□", "□"))
    font(p.add_run(f"{mark[0]} 심사평가표   {mark[1]} 심사의견요약표"), 15, True)
    p2 = para(doc, f"( {fmt_kdate(meta.get('작성일'))} )", size=10.5, align=CENTER, space_after=2)
    return p2


def fmt_kdate(s):
    y, m, d = (s or "").replace(".", "-").split("-")[:3]
    return f"{int(y):04d}년 {int(m):02d}월 {int(d):02d}일"


def cell_val(cell, text, size=9.5, bold=False, align=CENTER, blue=False, bg=None):
    cell_text(cell, text, size, bold, align, bg)
    if blue:
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.color.rgb = BLUE


def put_table(cell, obj, take, cell_w):
    """셀 안 표. 제목·단위는 셀 폭 기준으로 우측 정렬한다."""
    if obj.get("title"):
        p = take()
        p.paragraph_format.tab_stops.add_tab_stop(Mm(cell_w - 4), WD_TAB_ALIGNMENT.RIGHT)
        p.paragraph_format.space_before = Pt(2)
        font(p.add_run(obj["title"]), 9.5)
        if obj.get("unit"):
            font(p.add_run("	" + obj["unit"]), 8)
    widths, rows, head = obj["widths"], obj["rows"], obj.get("header")
    t = cell.add_table(rows=(1 if head else 0) + len(rows), cols=len(widths))
    borders(t, sz=4)
    fixed(t, widths)
    aligns = obj.get("aligns") or [CENTER] * len(widths)
    off = 0
    if head:
        for i, h in enumerate(head):
            cell_text(t.rows[0].cells[i], h, 8, False, CENTER, HEAD_BG)
        off = 1
    for r, row in enumerate(rows):
        t.rows[r + off].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        t.rows[r + off].height = Mm(4.2)
        for i, v in enumerate(row):
            cell_text(t.rows[r + off].cells[i], v, 8, False, aligns[i])
    return t


def fill(cell, lines, size=8, cell_w=148):
    """셀 안에 ○ 글줄·표·이미지를 채운다(검토의견서 box()와 같은 방식)."""
    state = {"slot": cell.paragraphs[0]}

    def take():
        p = state["slot"] or cell.add_paragraph()
        state["slot"] = None
        return p

    for ln in lines:
        if ln is None:
            continue
        if isinstance(ln, dict):
            if ln["type"] == "image":
                take().add_run().add_picture(ln["path"], width=Mm(ln.get("width", cell_w - 8)))
            else:
                put_table(cell, ln, take, cell_w)
            state["slot"] = cell.paragraphs[-1]
            state["slot"].paragraph_format.space_before = Pt(2)
            continue
        for sub in str(ln).split("\n"):
            p = take()
            pf = p.paragraph_format
            pf.space_before = pf.space_after = Pt(0.2)
            pf.line_spacing = 1.0
            if sub.startswith(("○", "■", "※", "☞")):
                pf.left_indent, pf.first_line_indent = Mm(4.5), Mm(-4.5)
            elif sub.startswith(" "):
                pf.left_indent = Mm(4.5)
            font(p.add_run(sub), size)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP


# ---------- 1쪽: 심사평가표 ----------
def page_eval(doc, sh, d):
    meta = sh.get("meta", {})
    approval_box(doc, meta)
    title_block(doc, meta, "평가")

    # 신청영업점 / 심사담당자
    outer = doc.add_table(rows=1, cols=3)
    no_borders(outer)
    fixed(outer, [52, 36, 86])
    t1 = outer.rows[0].cells[0].add_table(rows=1, cols=2)
    borders(t1, sz=4, outer=6)
    fixed(t1, [24, 28])
    cell_text(t1.rows[0].cells[0], "신청영업점", 9, True, CENTER, HEAD_BG)
    cell_val(t1.rows[0].cells[1], meta.get("신청영업점"), 9.5, blue=True)
    t2 = outer.rows[0].cells[2].add_table(rows=1, cols=5)
    borders(t2, sz=4, outer=6)
    fixed(t2, [24, 10, 20, 12, 20])
    for i, (k, v, hd) in enumerate([("심사담당자", None, True), ("직", None, True), (None, meta.get("심사담당자_직"), False),
                                    ("성명", None, True), (None, meta.get("심사담당자_성명"), False)]):
        cell_text(t2.rows[0].cells[i], k if hd else v, 9, hd, CENTER, HEAD_BG if hd else None)
    para(doc, size=2)

    # 채무자 / 대출과목 / 금액
    s = sh.get("신청", {})
    t = doc.add_table(rows=1, cols=6)
    borders(t, sz=4, outer=12)
    fixed(t, [22, 32, 22, 26, 16, 56])
    t.rows[0].height, t.rows[0].height_rule = Mm(8), WD_ROW_HEIGHT_RULE.AT_LEAST
    cells = t.rows[0].cells
    cell_text(cells[0], "채무자명", 9.5, True, CENTER, HEAD_BG)
    cell_val(cells[1], get(sh, "채무자.성명"), 10, True, blue=True)
    cell_text(cells[2], "대출과목", 9.5, True, CENTER, HEAD_BG)
    cell_val(cells[3], s.get("대출과목"), 9.5)
    cell_text(cells[4], "금액", 9.5, True, CENTER, HEAD_BG)
    cell_val(cells[5], f"{comma(s.get('금액'))}원", 10, True, blue=True)
    para(doc, size=2)

    # 본표
    vals = eval_values(sh, d)
    nrows = 1 + sum(len(items) for _, items in FORM) + 1
    t = doc.add_table(rows=nrows, cols=3)
    borders(t, sz=4, outer=12)
    fixed(t, [30, 62, 82])
    for i, h in enumerate(["항  목", "착  안  사  항", "평가내용 또는 심사의견"]):
        cell_text(t.rows[0].cells[i], h, 9.5, True, CENTER, HEAD_BG)
    r = 1
    for label, items in FORM:
        start = r
        for 착안, key in items:
            t.rows[r].height, t.rows[r].height_rule = Mm(6.8), WD_ROW_HEIGHT_RULE.AT_LEAST
            cell_text(t.rows[r].cells[1], 착안, 9, False, LEFT)
            cell_val(t.rows[r].cells[2], vals.get(key, ""), 9.5, blue=(key in BLUE_KEYS))
            r += 1
        m = t.rows[start].cells[0].merge(t.rows[r - 1].cells[0]) if r - 1 > start else t.rows[start].cells[0]
        for p in m.paragraphs[1:]:
            p._p.getparent().remove(p._p)
        cell_text(m, label, 9.5, True, CENTER, HEAD_BG)
    # 특인 내용 및 사유
    t.rows[r].height, t.rows[r].height_rule = Mm(14), WD_ROW_HEIGHT_RULE.AT_LEAST
    cell_text(t.rows[r].cells[0], "특인 내용 및\n사유", 9.5, True, CENTER, HEAD_BG)
    box = t.rows[r].cells[1].merge(t.rows[r].cells[2])
    for p in box.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    fill(box, [f"○ {x}" for x in (get(sh, "평가표.특인내용") or ["특인 없음"])], 9.5)


def eval_values(sh, d):
    s, ev = sh.get("신청", {}), sh.get("평가표", {})
    v = dict(ev)
    v["@신용등급"] = d["신용등급표시"]
    v["@대출기간"] = f"{s.get('기간', '')}({s.get('상환방식', '')})"
    v["@금리"] = f"{rate_str(s.get('금리'))}%({s.get('금리종류', '')} {s.get('변동고정', '변동')})"
    v["@자금용도"] = s.get("자금용도구분", "")
    v["@이자수익"] = d["이자수익문장"]
    v.setdefault("신용조사구분", "소호CSS평가" if get(sh, "신용.유형") == "CSS" else "기업신용평가")
    for k, default in [("법적제한사항", "없음"), ("채무관계자자격", "적정"), ("담보물적정", "적정"),
                       ("보증인적정", "해당없음"), ("소요자금한도", "적정"), ("편중대출여부", "부"),
                       ("사업계획", "적정"), ("자금수지계획", "적정"), ("업계현황전망", "적정"),
                       ("현금흐름상태", "적정"), ("주석사항", "해당없음"), ("기타특이사항", "")]:
        if empty(v.get(k)):
            v[k] = default
    return v


# ---------- 2쪽: 심사의견요약표 ----------
def page_summary(doc, sh, d):
    meta = sh.get("meta", {})
    doc.add_page_break()
    approval_box(doc, meta)
    title_block(doc, meta, "요약")

    t = doc.add_table(rows=2, cols=2)
    borders(t, sz=4, outer=12)
    fixed(t, [26, 148])
    t.rows[0].height, t.rows[0].height_rule = Mm(60), WD_ROW_HEIGHT_RULE.AT_LEAST
    for row in t.rows:            # 행이 페이지 경계에서 쪼개지지 않게
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    cell_text(t.rows[0].cells[0], "종합의견", 10, True, CENTER, HEAD_BG)
    t.rows[0].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    fill(t.rows[0].cells[1], summary_lines(sh, d), 9.5)

    cell_text(t.rows[1].cells[0], "승인조건 및\n지도사항", 10, True, CENTER, HEAD_BG)
    t.rows[1].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cond = sh.get("승인조건") or {}
    lines = ["■ 실행조건"] + [f"○ {x}" for x in (cond.get("실행조건") or ["해당없음"])]
    lines += ["■ 관리조건"] + [f"○ {x}" for x in (cond.get("관리조건") or ["해당없음"])]
    fill(t.rows[1].cells[1], lines, 9.5)


def summary_lines(sh, d):
    s, h, r = sh.get("신청", {}), sh.get("담보한도", {}), sh.get("요약표", {})
    c = sh.get("채무자", {})
    who = c.get("성명", "")
    if c.get("상호"):
        who += f"({c['상호']})"
    tail = ", ".join(x for x in [c.get("조합원"), c.get("하나로등급"), c.get("신규여부")] if x)
    L = [f"○ 채무자 : {who}" + (f" {tail}" if tail else "")]

    loans = merge_loans(sh.get("기대출") or [])
    if loans:
        rows = [[comma(x.get("잔액", 0) / 1000) + (f"({comma(x['한도'] / 1000)})" if x.get("한도") else ""),
                 x.get("자금용도", ""), rate_str(x.get("금리")), x.get("담보구분", ""), x.get("비고", "")]
                for x in loans]
        rows.append([comma(d["기대출합계"] / 1000), "", "", "", ""])
        L.append({"type": "table", "title": "○ 우리농협 기대출 현황", "unit": "(단위 : 천원, %)",
                  "header": ["대출잔액(한도)", "자금용도", "금리", "담보구분", "비고"], "rows": rows,
                  "widths": [32, 26, 18, 22, 28], "aligns": [RIGHT, CENTER, CENTER, CENTER, CENTER]})

    L.append(f"○ 대출신청금액 : {comma(s.get('금액'))}원({s.get('자금용도구분', '')})")
    if not empty(h.get("감정가")):
        appr = f"○ 감정가 : {comma(h['감정가'])}원"
        if h.get("감정일") or h.get("감정기관"):
            appr += f" ({(h.get('감정일') or '').replace('-', '.')}. {h.get('감정기관', '')})"
        L.append(appr)
        if h.get("소재지"):
            L.append(f" ({h['소재지']})")
        L.append("○ " + d["대출가능_산식"][0])
        L += [" " + x for x in d["대출가능_산식"][1:]]
    if d.get("소요자금한도"):
        L.append("○ " + d["소요자금한도"])
    L.append(f"○ 적용금리 : {rate_str(s.get('금리'))}% (가산 {rate_str(d['금리']['가산'])}%)")
    L.append(" " + d["금리"]["산출근거"])
    L += [f"○ {x}" for x in (r.get("조항") or [])]
    if r.get("RTI"):
        L.append(f"○ RTI산출 : {r['RTI']}")
    L += [f"○ {x}" for x in (r.get("추가줄") or [])]
    if r.get("소득"):
        L.append(f"○ {r['소득']}")
    if r.get("재산세"):
        L.append(f"○ {r['재산세']}")
    if r.get("낙찰통계"):   # 제목 줄만 — 표는 사용자가 결재 전에 붙인다(차주10 결재본)
        L.append(f"○ {r['낙찰통계'].split(' — ')[0]}")
    return L


def build(sh, out_path):
    d = compute_sheet(sh)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(18)
    sec.top_margin, sec.bottom_margin = Mm(12), Mm(10)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = FONT, Pt(9.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    ppr = st.element.get_or_add_pPr()
    for tag in ("w:autoSpaceDE", "w:autoSpaceDN"):
        e = OxmlElement(tag)
        e.set(qn("w:val"), "0")
        ppr.append(e)
    page_eval(doc, sh, d)
    page_summary(doc, sh, d)
    doc.save(out_path)
    return d


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("sheet")
    ap.add_argument("-o", "--out")
    ap.add_argument("--draft", action="store_true", help="A등급 누락이 있어도 초안으로 출력")
    a = ap.parse_args()
    p = Path(a.sheet)
    sh = json.loads(p.read_text(encoding="utf-8"))
    blockers = [x for x in gaps_sheet(sh, compute_sheet(sh)) if x[0] == "A"]
    if blockers and not a.draft:
        print("BLOCKED: 작성 전 질문으로 해소해야 할 A등급 항목")
        for _, k, m in blockers:
            print(f"  - ({k}) {m}")
        sys.exit(2)
    name = get(sh, "채무자.성명", "채무자")
    out = Path(a.out) if a.out else p.parent / f"심사평가표({name}){'_초안' if blockers else ''}.docx"
    d = build(sh, out)
    print(f"OK {out}")
    for lvl, cid, msg in d["checks"]:
        print(f"  검산 {lvl} {cid} {msg}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
