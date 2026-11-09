# -*- coding: utf-8 -*-
"""case.json + calc 결과 → 심사담당자 검토의견서 .docx (결재본 PDF 양식 재현)

사용법: python render_docx.py <case.json> [-o 출력.docx]
숫자는 calc.compute()에서만 가져온다. 서술 문장은 case.json 에 쓴 그대로 넣는다.
"""
import argparse
import re
import sys
from datetime import date
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
import json  # noqa: E402
from calc import D, comma, compute, empty, gaps, get, mil, rate_str, trunc  # noqa: E402

FONT = "맑은 고딕"
HEAD_BG = "F2E9D8"
W = 174  # 본문 폭(mm)
DRAFT = {"on": False}  # --draft 시 빈 상자에 【미확보】 표시
CENTER, LEFT, RIGHT = WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT

NOTE = {
    "가": "대출거래자격 및 차입행위의 제한여부, 여신규제대상 여부, 채무자의 대출의사결정상의 하자 여부",
    "나": "구체적인 자금용도와 여신신청 금액은 적정한지, 소요자금이 적정한지 여부와 자금조달 계획을 참고하여 실현가능성과 자기자본과 차입금 구성의 적정성 여부",
    "다": "차입금 구성 및 만기구조의 적정성, 상환재원에 따른 상환 가능성 등 검토",
    "라": "향후 거래 전망 및 조합수지기여 정도",
    "마": "동일인에 대한 편중대출 여부 및 조합자기자본 규모 대비 대출금액의 적정성",
}
LAND_TXT = ("본 건 토지와 유사한 이용가치를 지닌 인근 및 인접 지역 내 표준지의 공시지가를 기준으로 해당 토지의 위치, 형상, "
            "환경, 이용 상황 등 토지의 객관적 가치에 영향을 미치는 제반요인을 종합 참작하여 평가함.")
# 기본 문구 = 결재본 최빈(2026-10-20 집계: 토지 "인근 및 인접 지역…종합 참작" 6건 / 건물 "표준단가" 5건 > 원가법 2건).
# 유형별 대안(집합건물·구분건물·분양대금 등)은 rules/06_문구은행.md — 감정평가서 문구가 있으면 그게 우선.
BLDG_TXT = "대상건물과 유사한 용도, 구조 건물의 표준단가를 참작하되, 사용자재, 관리상태, 장래효용성 및 부대설비 등을 고려하여 평가함."


# ---------------- 저수준 헬퍼 ----------------
def font(run, size=9.5, bold=False, underline=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.underline = underline
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), FONT)


def para(container, text="", size=9.5, bold=False, align=LEFT, space_before=0, space_after=0, indent=None, underline=False):
    p = container.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after = Pt(space_before), Pt(space_after)
    pf.line_spacing = 1.15
    p.alignment = align
    if indent is not None:
        pf.left_indent, pf.first_line_indent = Mm(indent), Mm(-indent)
    if text:
        font(p.add_run(text), size, bold, underline)
    return p


def shade(cell, fill=HEAD_BG):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def borders(table, sz=4, outer=None, inner=True):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        is_inner = edge.startswith("inside")
        if is_inner and not inner:
            e.set(qn("w:val"), "nil")
        else:
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), str(sz if is_inner else (outer or sz)))
            e.set(qn("w:color"), "000000")
        b.append(e)
    tblPr.append(b)


def no_borders(table):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "nil")
        b.append(e)
    tblPr.append(b)


def fixed(table, widths):
    table.autofit = False
    tblPr = table._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    for row in table.rows:
        for i, c in enumerate(row.cells):
            if i < len(widths):
                c.width = Mm(widths[i])


def cell_text(cell, text, size=9, bold=False, align=CENTER, bg=None):
    lines = text if isinstance(text, list) else str("" if text is None else text).split("\n")
    first = True
    for ln in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        if ln:
            font(p.add_run(ln), size, bold)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    if bg:
        shade(cell, bg)


def grid(container, header, rows, widths, size=9, head_size=9, aligns=None, outer=12, row_h=6.5):
    t = container.add_table(rows=1 + len(rows), cols=len(widths))
    borders(t, sz=4, outer=outer)
    fixed(t, widths)
    for i, h in enumerate(header):
        cell_text(t.rows[0].cells[i], h, head_size, True, CENTER, HEAD_BG)
    for r, row in enumerate(rows, start=1):
        t.rows[r].height = Mm(row_h)
        t.rows[r].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        for i, v in enumerate(row):
            al = aligns[i] if aligns else CENTER
            cell_text(t.rows[r].cells[i], v, size, False, al)
    return t


def heading(doc, text, unit=None, size=11.5, lead=0):
    p = para(doc, space_before=3, space_after=1)
    p.paragraph_format.tab_stops.add_tab_stop(Mm(W), WD_TAB_ALIGNMENT.RIGHT)
    font(p.add_run(" " * lead + text), size, True)
    if unit:
        font(p.add_run("\t" + unit), 9, False)
    return p


def box(doc, lines, size=10):
    """테두리 상자 안에 ○ bullet. '- '로 시작하면 하위줄, '(' 로 시작하면 이어쓰기 줄. dict = 상자 안 표/이미지."""
    def new_box():
        t = doc.add_table(rows=1, cols=1)
        borders(t, sz=12, outer=12, inner=False)
        fixed(t, [W])
        return t, t.rows[0].cells[0]

    t, c = new_box()
    if DRAFT["on"] and not any((isinstance(x, dict) or str(x).strip()) for x in lines if x is not None):
        lines = ["○ 【미확보】"]
    state = {"slot": c.paragraphs[0]}   # 다음 글줄이 들어갈 빈 문단(표 뒤 자동 생성 문단 재사용)

    def take():
        p = state["slot"] or c.add_paragraph()
        state["slot"] = None
        return p

    for ln in lines:
        if ln is None:
            continue
        if isinstance(ln, dict) and ln.get("새쪽"):
            # 상자 안에서는 Word가 쪽나눔을 무시 → 상자를 닫고 다음 쪽에 이어지는 상자를 연다
            para(doc).add_run().add_break(WD_BREAK.PAGE)
            t, c = new_box()
            state["slot"] = c.paragraphs[0]
        if isinstance(ln, dict):
            _box_obj(c, ln, take, state, size)
            continue
        for sub in str(ln).split("\n"):
            p = take()
            pf = p.paragraph_format
            pf.space_before = pf.space_after = Pt(0.5)
            pf.line_spacing = 1.1
            if sub.startswith(("○", "※", "☞")):
                pf.left_indent, pf.first_line_indent = Mm(4), Mm(-4)
            elif sub.startswith(("-", " -", "(")):
                pf.left_indent = Mm(2)
            font(p.add_run(sub), size)
    return t


def _box_obj(cell, obj, take, state, size):
    if obj["type"] == "table":
        if obj.get("title"):
            p = take()
            p.paragraph_format.tab_stops.add_tab_stop(Mm(W - 6), WD_TAB_ALIGNMENT.RIGHT)
            p.paragraph_format.space_before = Pt(1)
            font(p.add_run(obj["title"]), size)
            if obj.get("unit"):
                font(p.add_run("\t" + obj["unit"]), 8.5)
        widths = obj["widths"]
        rows = obj["rows"]
        head = obj.get("header")
        t = cell.add_table(rows=(1 if head else 0) + len(rows), cols=len(widths))
        borders(t, sz=4)
        fixed(t, widths)
        aligns = obj.get("aligns") or [CENTER] * len(widths)
        off = 0
        if head:
            for i, h in enumerate(head):
                cell_text(t.rows[0].cells[i], h, 9, False, CENTER, HEAD_BG)
            off = 1
        for r, row in enumerate(rows):
            for i, v in enumerate(row):
                cell_text(t.rows[r + off].cells[i], v, 9, False, aligns[i])
        # 제목·머리행이 쪽 끝에 홀로 남지 않게: 제목~마지막 직전 행까지 다음 줄과 묶기
        if obj.get("title"):
            p.paragraph_format.keep_with_next = True
        for row in t.rows[:-1]:
            for c in row.cells:
                for p in c.paragraphs:
                    p.paragraph_format.keep_with_next = True
        if state["slot"] is not None and state["slot"] is cell.paragraphs[0] and not state["slot"].runs:
            # 상자 첫 요소가 제목 없는 표면 맨 앞 빈 문단 제거
            state["slot"]._p.getparent().remove(state["slot"]._p)
        state["slot"] = cell.paragraphs[-1]          # python-docx가 표 뒤에 만든 빈 문단
        state["slot"].paragraph_format.space_before = Pt(2)
    elif obj["type"] == "image":
        p = take()
        p.add_run().add_picture(obj["path"], width=Mm(obj.get("width", W - 8)))


def note(doc, key):
    p = para(doc, space_before=1, space_after=2, size=8.5)
    p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = Mm(6), Mm(-6)
    font(p.add_run(f"주) 착안사항 : {NOTE[key]}"), 8.5)


def page_number_footer(section):
    p = section.footer.paragraphs[0]
    p.alignment = CENTER
    font(p.add_run("- "), 9)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    p._p.append(fld)
    font(p.add_run(" -"), 9)


def fmt_date(s):
    y, m, d = (s or date.today().isoformat()).replace(".", "-").split("-")[:3]
    return f"({int(y):04d}. {int(m):02d}. {int(d):02d}.)"


def fmt_area(v):
    if empty(v):
        return ""
    x = D(v)
    return f"{x:,.0f}" if x == x.to_integral() else f"{x:,.2f}".rstrip("0").rstrip(".")


def office_label(s):
    m = re.match(r"(.+?)(지점|본점|지소)$", s or "")
    return f"{m.group(1)}\n{m.group(2)}" if m else (s or "")


# ---------------- 섹션 렌더링 ----------------
def _nn(x, top=True):
    """None → '' (초안의 빈 값 대응). 단 보증인/신용위험의 None은 '섹션 판단 미정' 의미라 보존."""
    if isinstance(x, dict):
        return {k: (v if top and k in ("보증인", "신용위험") and v is None else _nn(v, False)) for k, v in x.items()}
    if isinstance(x, list):
        return [_nn(v, False) for v in x]
    return "" if x is None else x


def render(case, out_path):
    d = compute(case)
    case = _nn(case)
    s, h, u = case.get("신청", {}), case.get("담보한도", {}), case.get("업체", {})
    meta = case.get("meta", {})
    is_inc = s.get("구분") == "증액"
    joint = s.get("공동대출")
    credit = d["신용등급표시"]
    who = u.get("표시명") or u.get("업체명", "")

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(18)
    sec.top_margin, sec.bottom_margin = Mm(15), Mm(15)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = FONT, Pt(9.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    ppr = st.element.get_or_add_pPr()
    for tag in ("w:autoSpaceDE", "w:autoSpaceDN"):
        e = OxmlElement(tag)
        e.set(qn("w:val"), "0")
        ppr.append(e)
    page_number_footer(sec)

    # ---- 제목 + 결재란 ----
    appr = meta.get("결재선") or ["계원", "실장", "본부장", "상임이사"]
    top = doc.add_table(rows=1, cols=2)
    no_borders(top)
    fixed(top, [W - 16 * len(appr) - 2, 16 * len(appr) + 2])
    lc = top.rows[0].cells[0]
    cell_text(lc, "", 10)
    p = lc.paragraphs[0]
    p.alignment = RIGHT
    font(p.add_run("심사담당자 검토의견서"), 18, True, underline=True)
    p2 = lc.add_paragraph()
    p2.alignment = CENTER
    p2.paragraph_format.left_indent = Mm(40)
    font(p2.add_run(fmt_date(meta.get("작성일"))), 10)
    rc = top.rows[0].cells[1]
    at = rc.add_table(rows=2, cols=len(appr))
    borders(at, sz=4, outer=8)
    fixed(at, [16] * len(appr))
    for i, a in enumerate(appr):
        cell_text(at.rows[0].cells[i], a, 8.5)
    at.rows[1].height, at.rows[1].height_rule = Mm(13), WD_ROW_HEIGHT_RULE.EXACTLY

    # ---- 1. 승인 신청 내역 ----
    heading(doc, "1. 승인 신청 내역", "(단위 : 백만원)", size=12)
    amt = comma(mil(s.get("대출금액")))
    if is_inc:
        amt += f"\n({comma(mil(D(s['대출금액']) - D(s.get('증액전금액'))))} 증액)"
    if joint:
        amt += f"\n(총 {comma(mil(joint.get('총액')))})"
    elif s.get("대출금액_부기"):  # 공동대출 아닌 동순위 개별취급 총액 표시 (2026-10-16 업체B)
        amt += "\n" + s["대출금액_부기"]
    rate_kind = s.get("금리종류", "")
    period = f"{s.get('기간', '')}\n({s.get('상환방식', '')})"
    fy = case.get("양식", {})
    headA = ["사무소", "구분", "채무자", "사업자번호", "대출금액", "금 리", "기 간", "대출과목", "자금용도"]
    rowA = [office_label(s.get("사무소")), s.get("구분", ""), u.get("표A명") or who, u.get("사업자번호", ""), amt,
            f"{rate_str(s.get('금리'))}%\n({rate_kind})", period, s.get("대출과목", ""), s.get("자금용도구분", "")]
    widA = [14, 12, 20, 24, 16, 20, 22, 22, 24]
    if fy.get("표A_취급수수료"):  # 직원 결재본 양식: 금리 다음에 취급수수료 칸 (2026-10-12 업체A)
        headA.insert(6, "취급\n수수료")
        rowA.insert(6, fy["표A_취급수수료"])
        widA = [13, 10, 19, 23, 15, 18, 12, 18, 20, 26]
    grid(doc, headA, [rowA], widA, size=8.5, row_h=12)
    para(doc, size=3)
    b = d["표B"]
    ltv_txt = f"{b['LTV']}%\n({h.get('담보종류표시', '')})"
    senior_txt = comma(b["선순위_백만"]) + ("주)" if isinstance(h.get("주석"), str) and h.get("주석") else "")
    if fy.get("표B") == "보증":  # HUG 보증서 사업비대출: 매매금액(총사업비)|보증금액|… (2026-11-03 업체Q건설 결재본)
        grid(doc, [fy.get("표B_라벨", "매매금액"), "보증금액", "담보인정비율\n및 적용가액", "대출가능금액", "총여신액/\n총가능금액", "비 고"],
             [[comma(mil(fy.get("표B_매매") or 0)) if fy.get("표B_매매") else "", comma(b["감정가_백만"]), ltv_txt,
               comma(b["대출가능_백만"]), b["총여신비율"], h.get("비고", "")]],
             [28, 28, 32, 30, 30, 26], row_h=12)
    elif fy.get("표B") == "매매":  # 공동대출 총액 기준: 매매금액 × LTV = 총 대출가능, 총여신/총가능 (2026-10-12 업체A)
        grid(doc, [fy.get("표B_라벨", "매매금액"), "담보인정비율\n및 적용가액", "대출가능금액", "총여신액/\n총가능금액", "비 고"],
             [[comma(b["감정가_백만"]), ltv_txt, comma(b["적용가액_백만"]), b["총여신비율"], h.get("비고", "")]],
             [34, 40, 36, 36, 28], row_h=12)
    else:
        grid(doc, ["감정가", "담보인정비율\n및 적용가액", "선순위차감액", "기대출", "대출가능금액", "총여신액/\n가능금액", "비 고"],
             [[f"{comma(b['감정가_백만'])}\n({h.get('감정기관', '')})", ltv_txt, senior_txt, comma(b["기대출_백만"]),
               comma(b["대출가능_백만"]), b["총여신비율"], h.get("비고", "")]],
             [26, 30, 28, 24, 27, 24, 15], row_h=12)
    # 주) = 선순위 차감 근거(임금채권·임차보증금) 한 줄. LTV 예외·대주 정보처럼 둘 이상이면 리스트 → 주1), 주2) (2026-10-20 사용자)
    notes = h.get("주석")
    if isinstance(notes, list):
        for k, n in enumerate(notes, 1):
            para(doc, n if n.startswith("주") else f"주{k}) " + n, size=8)
    elif notes:
        para(doc, notes if notes.startswith("주") else "주) " + notes, size=8)

    # ---- 2. 사업체 개요 ----
    heading(doc, "2. 사업체 개요", size=12)
    second = fy.get("담보제공인")
    heading(doc, "가. 채무자현황" if second else "가. 사업체현황", lead=1, size=11)
    _company_table(doc, u, credit)
    if second:  # 담보제공인 겸 연대보증인(공동사업자) 절 (2026-10-12 업체A)
        heading(doc, f"나. {second.get('제목', '담보제공인 겸 연대보증인')}", lead=1, size=11)
        _company_table(doc, second, second.get("신용등급", ""))
    letters = iter("다라마바사" if second else "나다라마바사")
    if not fy.get("재무금융생략"):
        _fin_loans(doc, case, s, next(letters))
        eq = case.get("출자현황")
        if eq:  # 사업비대출(SPC·리츠): 금융거래 표 대신 출자현황 (2026-11-03)
            heading(doc, f"{next(letters)}. 출자현황", eq.get("unit", "(단위 : 주, 백만원, %)"), lead=1, size=11)
            grid(doc, eq["header"], eq["rows"], eq["widths"], size=9, row_h=7)
        else:
            _loans(doc, case, s, next(letters))
    heading(doc, f"{next(letters)}. 업 체 전 망", lead=1, size=11)
    box(doc, _bul(get(case, "전망.업체", [])))
    heading(doc, f"{next(letters)}. 업 계 전 망", lead=1, size=11)
    box(doc, _bul(get(case, "전망.업계", [])))
    if get(case, "전망.영업위험"):  # 사업비대출 결재본의 '바. 영업위험 분석'
        heading(doc, f"{next(letters)}. 영업위험 분석", lead=1, size=11)
        box(doc, _bul(get(case, "전망.영업위험", [])))
    _section3(doc, case, d, s, who, credit, meta)
    doc.save(out_path)
    return d


def _company_table(doc, u, credit):
    corp = u.get("기업형태") == "법인"
    t = doc.add_table(rows=5, cols=6)
    borders(t, sz=4, outer=12)
    fixed(t, [22, 36, 22, 36, 22, 36])
    rows = [
        [("업 체 명", u.get("업체명")), ("대 표 자", u.get("대표자")), ("기업규모", u.get("기업규모"))],
        [("업    종", u.get("업종")), ("주요제품", u.get("주요제품")), ("기업형태", u.get("기업형태"))],
        [("신용등급", credit), None, ("설 립 일", u.get("설립일"))],
        [("주    소", u.get("주소")), None, ("전화번호", u.get("전화"))],
        [("사업자번호", u.get("사업자번호")), ("법인번호" if corp else "주민등록번호", u.get("주민법인번호")), None],
    ]
    for r, spec in enumerate(rows):
        cells = t.rows[r].cells
        t.rows[r].height, t.rows[r].height_rule = Mm(7), WD_ROW_HEIGHT_RULE.AT_LEAST
        if r in (2, 3):
            cell_text(cells[0], spec[0][0], 9.5, True, CENTER, HEAD_BG)
            m = cells[1].merge(cells[3])
            cell_text(m, spec[0][1], 9.5)
            cell_text(cells[4], spec[2][0], 9.5, True, CENTER, HEAD_BG)
            cell_text(cells[5], spec[2][1], 9.5)
        elif r == 4:
            cell_text(cells[0], spec[0][0], 9.5, True, CENTER, HEAD_BG)
            m = cells[1].merge(cells[2])
            cell_text(m, spec[0][1], 9.5)
            cell_text(cells[3], spec[1][0], 9.5, True, CENTER, HEAD_BG)
            m2 = cells[4].merge(cells[5])
            cell_text(m2, spec[1][1], 9.5)
        else:
            for i, (k, v) in enumerate(spec):
                cell_text(cells[i * 2], k, 9.5, True, CENTER, HEAD_BG)
                cell_text(cells[i * 2 + 1], v, 9.5)


def _fin_loans(doc, case, s, letter):
    heading(doc, f"{letter}. 재무현황", "(단위 : 백만원, %)", lead=1, size=11)
    fin_head = ["구분", "총자산", "자기\n자본", "매출액", "영업\n이익", "당기\n순이익", "감가\n상각비", "EBITDA", "금융\n비용",
                "EBITDA\n이자보\n상배율", "부채\n(차입금)", "부채\n비율", "유동\n비율"]
    fin_keys = ["구분", "총자산", "자기자본", "매출액", "영업이익", "당기순이익", "감가상각비", "EBITDA", "금융비용",
                "이자보상배율", "부채", "부채비율", "유동비율"]
    fin_w = [11, 14, 13, 15, 13, 13, 13, 14, 13, 14, 15, 13, 13]
    fin = case.get("재무") or []
    if fin:
        grid(doc, fin_head, [[_num(r.get(k)) for k in fin_keys] for r in fin], fin_w, size=8.5, head_size=8.5)
    else:
        ft = grid(doc, fin_head, [[""] * 13], fin_w, size=9, head_size=8.5, row_h=14)
        m = ft.rows[1].cells[0].merge(ft.rows[1].cells[12])
        cell_text(m, "- 재무제표 없음 -", 9.5)


def _loans(doc, case, s, letter):
    heading(doc, f"{letter}. 금융거래 상황", "(단위 : 천원, %)", lead=1, size=11)
    loans = case.get("금융거래") or []
    has_period = any(l.get("기간") for l in loans)
    head = ["금융기관", "대출금액", "신용·담보", "금리"] + (["대출기간"] if has_period else []) + ["비 고"]
    widths = [34, 30, 30, 30] + ([22] if has_period else []) + [50 - (22 if has_period else 0)]
    body = []
    for l in loans:
        row = [l.get("기관", ""), comma(trunc(D(l.get("금액")) / 1000)), l.get("신용담보", ""),
               (rate_str(l["금리"]) + "%") if not empty(l.get("금리")) else "-"]
        if has_period:
            row.append(l.get("기간", ""))
        row.append(l.get("비고", ""))
        body.append(row)
    if s.get("구분") == "신규" and not any(l.get("기관") == "회천농협" for l in loans):
        body.append(["회천농협", "0", "", ""] + ([""] if has_period else []) + ["신규"])
    total = trunc(sum((D(l.get("금액")) for l in loans), D(0)) / 1000)
    body.append(["합    계", comma(total), "", ""] + ([""] if has_period else []) + [""])
    lt = grid(doc, head, body, widths, size=9.5, row_h=7)
    _vmerge_same(lt, 0, 1, len(body))


def _section3(doc, case, d, s, who, credit, meta):
    # ---- 3. 주요 항목별 검토 사항 ----
    heading(doc, "3. 주요 항목별 검토 사항", size=12)
    heading(doc, "가. 적법성 심사 의견", lead=1, size=11)
    jl = [f"○ 채무자 {who} {credit}"] + [_o(x["문구"]) for x in d["적법성"]] + _bul(get(case, "적법성.추가문구", []))
    box(doc, jl)
    note(doc, "가")

    heading(doc, "나. 타당성(자금용도 및 소요자금의 적정성 등) 심사 의견", lead=1, size=11)
    tl = [f"○ 자금용도 : {s.get('자금용도상세', '')}"]
    if d["타당성"]["문구"]:
        tl.append(_o(d["타당성"]["문구"]))
    tl += _bul(get(case, "타당성.추가문구", []))
    box(doc, tl)
    note(doc, "나")

    heading(doc, "다. 상환능력 등 안정성 심사 의견", lead=1, size=11)
    box(doc, _blocks(get(case, "상환능력.블록", []), d))
    note(doc, "다")

    heading(doc, "라. 수지기여에 대한 심사 의견", lead=1, size=11)
    box(doc, ["○ " + d["수지기여"]["문장"]] + ([d["수지기여"]["표"]] if d["수지기여"].get("표") else [])
        + ([d["고정금리분석"]["표"], "- " + d["고정금리분석"]["1bp"]] if d.get("고정금리분석") else [])
        + _bul(case.get("수지기여_추가")))
    note(doc, "라")

    heading(doc, "마. 동일인에 대한 편중대출 여부 심사 의견", lead=1, size=11)
    box(doc, ["○ " + d["편중"]["문장"]])
    note(doc, "마")

    # ---- 바. 채권보전 ----
    heading(doc, "바. 채권보전 심사 의견", lead=1, size=11)
    for col in case.get("담보물") or []:
        dt = (col.get("감정일") or "").replace("-", ".")
        label = ". ".join(x for x in [dt, col.get("감정기관", "")] if x)
        heading(doc, f"(1) 담보물 내역({col.get('제목') or label})", "(단위 : ㎡, 백만원)", size=10.5)
        if col.get("표"):  # 매매계약 기준 자유 표 (2026-10-12 업체A)
            tb = col["표"]
            if col.get("소재지"):
                para(doc, f"○ 소재지 : {col['소재지']}", size=9.5)
            grid(doc, tb["header"], tb["rows"], tb["widths"], size=9, row_h=9)
            if col.get("주석"):
                para(doc, col["주석"], size=8)
        else:
            _collateral(doc, col)
    ov = case.get("담보개요", {})
    heading(doc, "(2) 담보물 개요 및 현황", size=10.5)
    box(doc, [f"○ 입지현황 : {ov.get('입지', '')}", f"○ 교통상황 : {ov.get('교통', '')}", f"○ 공법상 제한사항 : {ov.get('공법', '')}"])

    heading(doc, "(3) 평가액 산출근기", size=10.5)
    base = case.get("산출근거", {})
    has_bldg = any(r.get("종류") == "건물" and not empty(r.get("면적")) for c in case.get("담보물", []) for r in c.get("행", []))
    if d.get("산출근거_보증"):  # 보증서 담보: 보증금액 × 100%(HUG)
        bl = ["○ " + d["산출근거_보증"]]
    else:
        bl = [f"○ 토지 : {base.get('토지') or LAND_TXT}"]
        if base.get("건물") or has_bldg:
            bl.append(f"○ 건물 : {base.get('건물') or BLDG_TXT}")
    box(doc, bl)

    heading(doc, "(4) 평가액의 적정성 및 담보물의 환가성, 기타 사항 검토", size=10.5)
    box(doc, _adequacy(case))

    heading(doc, "(5) 보증인 내역", size=10.5)
    g = case.get("보증인")
    box(doc, _bul(g) if g else ["○ 해당없음"])

    if case.get("신용위험") is not None and meta.get("신용위험_섹션", True):
        heading(doc, "사. 주요 신용위험(key credit risks)", lead=1, size=11)
        box(doc, _bul(case["신용위험"]) or ["○ 해당없음"])

    # ---- 4. 종합심사의견 ----
    heading(doc, "4. 종합심사의견", size=12)
    z = case.get("종합", {})
    heading(doc, "가. 긍정적인 사항", lead=1, size=11)
    box(doc, _bul(z.get("긍정", [])))
    heading(doc, "나. 부정적인 사항", lead=1, size=11)
    box(doc, _bul(z.get("부정", [])))
    heading(doc, "다. 종합 의견", lead=1, size=11)
    box(doc, _summary(case, d))

    # ---- 5. 승인조건 ----
    heading(doc, "5. 승인조건 및 관리조건", size=12)
    heading(doc, "가. 승인 조건", lead=1, size=11)
    box(doc, _bul(d["승인조건"]))
    heading(doc, "나. 관리 조건", lead=1, size=11)
    box(doc, _bul(case.get("관리조건") or ["해당없음"]))


# ---------------- 섹션 부품 ----------------
def _num(v):
    if empty(v):
        return ""
    if isinstance(v, str):
        return v
    x = D(v)
    return f"{x:,.0f}" if x == x.to_integral() else f"{x:,.2f}".rstrip("0").rstrip(".")


def _o(text):
    return text if str(text).startswith(("○", "※", "-", "(")) else "○ " + text


def _bul(items):
    # dict(표/이미지)는 상자 안 객체로 그대로 통과 (2026-10-16 업체B: 타당성·적정성에 표·그림)
    return [x if isinstance(x, dict) else _o(x) for x in (items or [])]


def _blocks(blocks, d):
    """상환능력 블록: 문자열(○자동) 또는 {"type":"table",...}. {금융비용}/{이자수익} 치환."""
    out = []
    for b in blocks or []:
        if isinstance(b, dict):
            out.append(b)
        else:
            out.append(_o(b.replace("{금융비용}", d["금융비용"]["문장"])))
    return out


def _vmerge_same(t, col, start, end):
    i = start
    while i < end:
        j = i
        txt = t.rows[i].cells[col].text
        while j + 1 < end and t.rows[j + 1].cells[col].text == txt and txt:
            j += 1
        if j > i:
            m = t.rows[i].cells[col].merge(t.rows[j].cells[col])
            for p in m.paragraphs[1:]:
                p._p.getparent().remove(p._p)
            cell_text(m, "", 9.5) if not m.paragraphs[0].runs else None
        i = j + 1


def _collateral(doc, col):
    rows = col.get("행", [])
    n = len(rows)
    t = doc.add_table(rows=n + 4, cols=7)
    borders(t, sz=4, outer=12)
    widths = [13, 22, 26, 20, 20, 54, 19]
    fixed(t, widths)
    m = t.rows[0].cells[0].merge(t.rows[0].cells[6])
    cell_text(m, f"# 소재지 :  {col.get('소재지', '')}", 9.5, False, LEFT)
    for i, hname in enumerate(["종류", "지 목\n용 도", "구 조", "면 적", "평가액", "권리사항", "담보제공자"]):
        cell_text(t.rows[1].cells[i], hname, 9, False, CENTER, HEAD_BG)
    tot_area, tot_val = D(0), D(0)
    for r, row in enumerate(rows, start=2):
        t.rows[r].height, t.rows[r].height_rule = Mm(9), WD_ROW_HEIGHT_RULE.AT_LEAST
        cell_text(t.rows[r].cells[0], row.get("종류", ""), 9)
        cell_text(t.rows[r].cells[1], row.get("지목용도", ""), 8.5)
        cell_text(t.rows[r].cells[2], row.get("구조지역", ""), 8.5)
        cell_text(t.rows[r].cells[3], fmt_area(row.get("면적")), 9, False, RIGHT)
        val = row.get("평가액")
        cell_text(t.rows[r].cells[4], comma(D(val), 0) if not empty(val) else "-", 9, False, RIGHT)
        tot_area += D(row.get("면적"))
        tot_val += D(val)
    ref = n + 2
    mm = t.rows[ref].cells[0].merge(t.rows[ref].cells[4])
    cell_text(mm, "- 감정평가명세표 참고 -", 8.5)
    rights = t.rows[2].cells[5].merge(t.rows[ref].cells[5])
    cell_text(rights, col.get("권리사항", []), 9)
    owner = t.rows[2].cells[6].merge(t.rows[ref].cells[6])
    cell_text(owner, col.get("담보제공자", []), 9)
    last = t.rows[n + 3].cells
    cell_text(last[0], "계", 9.5)
    cell_text(last[3], fmt_area(tot_area), 9, False, RIGHT)
    # 필지별 평가액이 없고 총평가액만 있는 경우(사전검토 감정평가 정보) — 2026-10-06 업체A
    total = col.get("평가액합계")
    cell_text(last[4], comma(D(total)) if not empty(total) else comma(tot_val), 10, True, RIGHT)


def _adequacy(case):
    a = case.get("적정성", {})
    meta = case.get("meta", {})
    lines = []
    # 탁상감정: 받은 기관만큼 칸 생성(감정기관은 건마다 다르므로 고정하지 않음). 금액 없으면 공란
    desk = a.get("탁상감정") or []
    if desk:
        n = len(desk)
        w = (W - 8) / n
        lines.append({"type": "table", "title": "○ 탁상감정 가격비교", "unit": "(단위 : 천원)",
                      "header": [x.get("기관", "") for x in desk],
                      "rows": [[comma(x.get("금액")) if not empty(x.get("금액")) else "" for x in desk]],
                      "widths": [w] * n, "aligns": [RIGHT] * n})
    deals = a.get("거래사례") or []
    if deals:
        lines.append({"type": "table", "title": "○ 인근 거래사례", "unit": "(단위 : ㎡, 천원/㎡)",
                      "header": ["소재지", "이용상황", "토지면적", "토지단가", "거래시점", "비고"],
                      "rows": [[x.get("소재지", ""), x.get("이용상황", ""), _num(x.get("토지면적")), _num(x.get("토지단가")),
                                x.get("거래시점", ""), x.get("비고", "")] for x in deals],
                      "widths": [36, 30, 30, 30, 24, 16], "aligns": [CENTER, CENTER, RIGHT, RIGHT, CENTER, CENTER]})
    levels = a.get("지가수준") or []
    if levels:
        lines.append({"type": "table", "title": "○ 인근 지가수준", "unit": "(단위 : ㎡, 천원/㎡)",
                      "header": ["용도지역", "토지용도", "도로조건", "가격수준", "비고"],
                      "rows": [[x.get("용도지역", ""), x.get("토지용도", ""), x.get("도로조건", ""), _num(x.get("가격수준")),
                                x.get("비고", "")] for x in levels],
                      "widths": [34, 34, 34, 32, 32], "aligns": [CENTER, CENTER, CENTER, RIGHT, CENTER]})
    if a.get("감정단가"):
        lines.append(f"  ※ 본건 감정단가 : {a['감정단가']}")
    lines += _bul(a.get("기타"))
    auc = a.get("낙찰사례")
    if auc:  # None·""(=_nn 변환)이면 낙찰사례 줄 생략
        lines.append(f"○ 낙찰사례({auc.get('제목', '인포케어, 최근1년')})")
        if auc.get("이미지") and Path(auc["이미지"]).exists():
            lines.append({"type": "image", "path": auc["이미지"]})
        elif auc.get("표"):
            tb = auc["표"]  # {"지역":["경기","양주시","광적면"], "행":[["1년간 평균", "48.04%", ...9개], ...]}
            hdr = ["지역통계"] + [f"{r}\n{k}" for r in tb["지역"] for k in ("낙찰가율", "낙찰률", "건수")]
            lines.append({"type": "table", "header": hdr, "rows": tb["행"], "widths": [22] + [16] * 9})
    return lines


def _summary(case, d):
    s, u, z = case.get("신청", {}), case.get("업체", {}), case.get("종합", {})
    who = u.get("표시명") or u.get("업체명", "")
    extra = " ".join(x for x in [u.get("조합원"), u.get("하나로등급")] if x)
    L = [f"○ 채무자 : {who} {extra + ' ' if extra else ''}{d['신용등급표시']}",
         f"○ {z.get('목적문구') or s.get('자금용도상세', '') + ' 마련을 위한 ' + s.get('자금용도구분', '') + ' 요청'}",
         "○ 제규정에 의한 적법성, 상환능력, 채권보전 등 검토 결과 적정"]
    note_amt = s.get("대출예정_부기")
    L.append(f"○ 대출예정금액 : {comma(s.get('대출금액'))}원({s.get('자금용도구분', '')}{', ' + note_amt if note_amt else ''})")
    L.append(f"○ 대출기간 : {s.get('종합_기간표시') or (s.get('기간', '') + ' ' + s.get('상환방식', ''))}")
    L.append(f"○ 대출금리: {rate_str(s.get('금리'))}%{' ' + s['금리_추가문구'] if s.get('금리_추가문구') else ''}")
    if not s.get("금리_기준줄생략"):
        line = f" - 기준금리 {s.get('금리종류', '')} {s.get('변동고정', '변동')} {rate_str(s.get('기준금리'))}%, 예상 가산금리 {rate_str(s.get('가산금리'))}%"
        if not empty(s.get("대심위우대금리")):
            line += f" → 대출심사위원회 우대금리 {rate_str(s['대심위우대금리'])}%"
        if not empty(s.get("기준금리")):
            L.append(line)
    L += [" - " + x.lstrip("- ") for x in s.get("우대금리내역", [])]
    if case.get("meta", {}).get("결정금리상담", True):
        L.append(" - 결정금리 상담 건으로 실행일 기준금리 변동 시 결정금리 따름")
    cond = z.get("기타대출조건") or []
    if cond:
        L.append("")
        L.append("○ 기타 대출조건")
        L += [" - " + x.lstrip("- ") for x in cond]
    joint = s.get("공동대출")
    if joint and joint.get("참여기관"):
        rows = [[x["기관"], comma(mil(x["금액"])), x.get("비고", "")] for x in joint["참여기관"]]
        rows.append(["합계", comma(mil(joint["총액"])), ""])
        L.append({"type": "table", "title": "○ 참여 금융기관", "unit": "(단위 : 백만원)", "header": ["금융기관", "취급금액", "비고"],
                  "rows": rows, "widths": [70, 50, 46], "aligns": [CENTER, RIGHT, CENTER]})
    if joint:  # 공동대출일 때만 (2026-10-20 사용자) — 문구는 결재본(업체A) 그대로, z.유효기간으로 바꿀 수 있음
        L.insert(len(L) - (1 if joint.get("참여기관") else 0),
                 f"○ 대출심사위원회 유효기간 : {z.get('유효기간') or '승인조건 충족 시 실행일까지 유효'}")
    L += [x if isinstance(x, dict) else _o(x) for x in (z.get("추가줄") or [])]  # 표(dict)도 허용
    return L


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("case")
    ap.add_argument("-o", "--out")
    ap.add_argument("--draft", action="store_true", help="A등급 누락이 있어도 초안으로 출력(파일명에 _초안)")
    a = ap.parse_args()
    p = Path(a.case)
    case = json.loads(p.read_text(encoding="utf-8"))
    name = get(case, "업체.업체명", "채무자")
    all_gaps = gaps(case, compute(case))
    blockers = [r for r in all_gaps if r[0] == "A"]
    if blockers and not a.draft:
        print("BLOCKED: 작성 전 질문으로 해소해야 할 A등급 항목")
        for _, key, msg in blockers:
            print(f"  - ({key}) {msg}")
        sys.exit(2)
    suffix = "_초안" if blockers else ""
    DRAFT["on"] = bool(a.draft)
    out = Path(a.out) if a.out else p.parent / f"심사담당자 검토의견서({name}){suffix}.docx"
    d = render(case, out)
    # 작성 시점 질문 개수 스냅샷 — 나중에 결재본과 비교할 때(learn.py record) 성적표에 쓴다
    counts = {lvl: sum(1 for r in all_gaps if r[0] == lvl) for lvl in "ABCD"}
    (p.parent / ".gaps_snapshot.json").write_text(json.dumps(counts, ensure_ascii=False), encoding="utf-8")
    print(f"OK {out}")
    for lvl, cid, msg in d["checks"]:
        print(f"  검산 {lvl} {cid} {msg}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
