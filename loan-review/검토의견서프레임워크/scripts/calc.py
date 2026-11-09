# -*- coding: utf-8 -*-
"""심사담당자 검토의견서 — 결정론적 계산 / 규정 트리거 / 검산 / 작성전 질문(누락) 탐지

사용법
  python calc.py <case.json>          계산 + 검산 결과 출력, 같은 폴더에 derived.json 저장
  python calc.py <case.json> --gaps   작성 전에 사용자에게 물어야 할 항목만 출력 (A필수/B충돌/C보완/D판단)

원칙: 숫자는 여기서만 만든다. 금액은 case.json에 전부 '원' 단위로 넣는다.
규정 문구를 고치면 rules/02_규정조항_라이브러리.md 도 함께 고친다.
"""
import argparse
import json
import re
import sys
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP
from pathlib import Path

MIL = Decimal(1_000_000)

LIB = {
    "L01": "동일인당 10억원 초과로 대출심사위원회 심의대상",
    "L02": "동일인당 30억원 이상 대출로 3인심사역 협의체 합의대상 ☞",
    "L02b": "공동대출로 3인심사역 협의체 합의대상 ☞",
    "L03": "여신업무방법(예) 제2편제1장제2절제3관제3조4호에 의거 임금채권 전액 미차감",
    "L04": "여신업무방법(예) 제2편제7장제4조에 의거 부동산담보신탁 임금채권 미차감",
    "L05": "여신업무방법(예) 제2편별첨3.나.3)에 의거 비거치 분할상환 대출로 담보인정비율 {n}% 상향",
    "L06": "여신업무방법(예) 제1편제7장제1절제6조제2항에 의거 한도증액인 경우 당초금리 적용가능",
    "L07": "여신업무방법(예) 제1편제7장제3절제7조6항에 의거 대출심사위원회 우대금리 승인",
    "L08": "여신업무방법(예) 제4편제6장제2절제4관제2조2항에 의거 1억원 이하 대출로 운전자금 자금용도 사후점검 생략",
    "L09": "여신업무방법(예) 제4편제2장제1절제2조제4항에 의거 경기민감업종 6A 이상으로 적정",
    # L10: 공동대출 총액이 제4조② 기본한도 초과일 때만. 2026.07.01 개정 제4조③④ 기준 (등급 충족 여부는 사람이 확인)
    "L10": "여신업무방법(예) 제2편제2장제7절제4조제3항에 의거 기업신용등급 5C등급 이상으로 공동대출 {n}억원 이내 취급",
    "L10b": "여신업무방법(예) 제2편제2장제7절제4조제4항에 의거 여신심사·관리 우수조합 및 기업신용등급 4B등급 이상으로 공동대출 {n}억원 이내 취급",
    # 2026-10-28 사용자: 규정 개정으로 13억(제1관제5조) 폐지 → 제2관제4조 개인 5억·법인 10억
    "T1": "여신업무방법(예) 제4편제4장제2절제2관제4조에 의거 동일인당 운전자금 순여신액 {한도}억원 이하 중소기업 운전자금으로 소요자금 산출 생략",
    "T2": "여신업무방법(예) 제4편제4장제2절제1관제3조제2항에 의거 운전자금차입금이 최근 결산년도 매출액 범위 안에 해당되면 지원 가능",
    "T3": "여신업무방법(예) 제4편제4장제2절제2관제3조에 의거 시설자금 소요한도 산출(80% 이내)",
    "T4": "여신업무방법(예) 제4편제4장제2절제2관제3조제4항에 의거 타 금융기관 대환자금 소요자금 산출 생략",
    "T5": "여신업무방법(예) 제4편제4장제2절제3관제4조에 의거 타 금융기관 시설자금 잔액 이내 대환",
    "T6": "여신업무방법(예) 제4편제4장제2절제2관제3조제1항제2호에 의거 부동산개발업종으로 총 소요자금에서 당해 공사관련 차입금을 차감한 범위 내 소요자금 산출",
    # 사업비대출(HUG 주택사업금융보증 보증서 담보) — 2026-11-03 업체Q건설 결재본, rules/07_사업비대출.md
    "P01": "여신업무방법(예) 제1편제1장제1절제5조제3항제4호나목에 의거 「공공기관의 운영에 관한 법률」에 따른 공공기관(주택도시보증공사)의 보증된 채권으로 동일인대출한도 산정 시 제외함",
    "P02": "여신업무방법(예) 제2편제1장제2절제1관제1조제1항에 의거 각종 보증서(주택도시보증공사) 담보 대출로 임차보증금·임금채권·조세채권 차감하지 않음\n- 여신업무방법(예) 제2편제1장제2절제3관제3조4호에 의거 임금채권 미차감 검토의견서 첨부",
    "P03": "여신업무방법(예) 제4편제2장제1절제1조제2항에 의거 주택도시보증공사(HUG) 보증 기업대출로 「주택도시보증공사 보증대출 취급기준 매뉴얼」의 채무자 자격(기업신용등급 6C 이상) 적정\n- 경기민감업종 신용등급 유의 : 개별 상품취급기준(매뉴얼) 상 6C 이상이므로 취급기준 적정(매뉴얼에서 정하지 않은 사항은 본편 준용)",
    "P04": "여신업무방법(예) 제1편제6장제1절제6조제1항에 의거 동일인당 30억원 이상 대출로 3인심사역 협의체 합의대상",
    "T7": "여신업무방법(예) 제4편제4장제2절제1관제5조9호 및 <별표10>에 의거 소요자금 한도산출 생략 여신으로 보증금액 이내에서 취급 - 보증서발급금액 : {보증}원",
    "T7b": "여신업무방법(예) 제4편제6장제2절제4관제2조2항8호에 의거 사후점검 생략 대상",
    "T7c": "여신업무방법(예) 제4편제4장제3절제4관제1조제1항 및 보증기간 이내로 운전자금 대출기간 산정",
}
APPROVAL = {
    "임금": "임금채권 미차감 검토의견서 첨부",
    "당초금리": "MOR금리 승인 검토의견서 첨부",
    "고액": "고액여신거래자로 정기 신용평가 등 관리 철저",
    "고액신설": "신설법인·고액여신거래자로 정기 신용평가 등 관리 철저",
    "보증서류": "HUG보증서담보대출 제반서류 일체 징구",
}
# 동일인 총대출한도(백만원) — 상호금융업감독규정 제6조⑦(2026-11-03 사용자 제공 원문):
# 자기자본 20% 최고한도 50억. 직전 사업연도말 자기자본 500억 이상 조합이 법인 조합원 또는 법인 준조합원
# (건설업·부동산업 법인 준조합원 제외)에 대출하면 100억. 회천은 500억 이상(사용자 '법인은 100억까지 가능').
SAME_PERSON_CAP_DEFAULT = 5000
SAME_PERSON_CAP_CORP = 10000


def same_person_cap(case):
    """(한도_백만, 판정근거). 조합원 구분이 비어 있으면 None 근거 → 5,000 초과 시 확인 질문"""
    if (get(case, "업체.기업형태", "") or "") != "법인":
        return SAME_PERSON_CAP_DEFAULT, "개인"
    mem = get(case, "업체.조합원", "") or ""
    biz = (get(case, "업체.업종", "") or "") + " " + (get(case, "업체.주요제품", "") or "")
    if mem == "조합원":
        return SAME_PERSON_CAP_CORP, "법인 조합원"
    if mem == "준조합원":
        if re.search(r"건설|부동산", biz):
            return SAME_PERSON_CAP_DEFAULT, "건설·부동산업 법인 준조합원(100억 특례 제외)"
        return SAME_PERSON_CAP_CORP, "법인 준조합원"
    if mem:
        return SAME_PERSON_CAP_DEFAULT, f"법인 {mem}(100억 특례 대상 아님)"
    return SAME_PERSON_CAP_DEFAULT, None
SEOUL_GU = "종로구 중구 용산구 성동구 광진구 동대문구 중랑구 성북구 강북구 도봉구 노원구 은평구 서대문구 마포구 양천구 강서구 구로구 금천구 영등포구 동작구 관악구 서초구 강남구 송파구 강동구".split()


# ---------- 공통 유틸 ----------
def D(x):
    return Decimal(str(x)) if x not in (None, "") else Decimal(0)


def trunc(x, nd=0):
    return D(x).quantize(Decimal(1).scaleb(-nd), rounding=ROUND_DOWN)


def mil(won):
    """원 → 백만원 절사 (int)"""
    return int(trunc(D(won) / MIL))


def comma(n, nd=0):
    if n is None or n == "":
        return ""
    v = trunc(n, nd)
    return f"{v:,.{nd}f}"


def rate_str(r):
    """4.70 → '4.7', 5.17 → '5.17'"""
    if r in (None, ""):
        return ""
    s = f"{D(r):.2f}".rstrip("0").rstrip(".")
    return s


def get(d, path, default=None):
    cur = d
    for k in path.split("."):
        if isinstance(cur, dict) and k in cur:
            cur = cur[k]
        else:
            return default
    return cur


def empty(v):
    return v is None or v == "" or v == [] or v == {}


# ---------- 계산 ----------
def compute(case):
    s = case.get("신청", {})
    h = case.get("담보한도", {})
    loans = case.get("금융거래", [])
    out = {"checks": [], "notes": []}
    chk = out["checks"]

    amount = D(s.get("대출금액"))
    is_inc = s.get("구분") == "증액"
    before = D(s.get("증액전금액")) if is_inc else Decimal(0)
    increment = amount - before if is_inc else amount
    rate = D(s.get("금리"))
    joint = s.get("공동대출") or None
    guar = case.get("상품") == "사업비대출"  # HUG 보증서 담보 사업비대출 (rules/07_사업비대출.md)

    # ---- 표B ----
    appraisal = D(h.get("감정가"))
    ltv = D(h.get("LTV"))
    senior = D(h.get("선순위차감"))
    prior_calc = sum((D(l.get("금액")) for l in loans
                      if l.get("기관") == "회천농협" and l.get("동일담보") and not l.get("증액대상")
                      and not l.get("대환대상")), Decimal(0))  # 대환대상은 금차로 상환되므로 기대출 아님 (2026-10-16 업체B)
    prior = D(h["기대출"]) if not empty(h.get("기대출")) else prior_calc
    if not empty(h.get("기대출")) and prior_calc and prior != prior_calc and not joint:
        chk.append(("B", "C14", f"표B 기대출 {comma(prior)}원 ≠ 금융거래 중 동일담보(증액대상 제외) 합계 {comma(prior_calc)}원 — 어느 쪽이 맞는지 확인"))
    cap = appraisal * ltv / 100 - senior           # 선순위 차감 후 담보 한도
    avail = cap - prior
    # 총여신비율: 소수 셋째 자리 반올림(2026-10-20 사용자 지정 — 절사보다 보수적. 이전 결재본 관행은 절사)
    total_lend = D(s.get("대주단총액") or amount) if guar else prior + amount  # 보증 건: 대주단 총액 ÷ 보증금액
    ratio = (total_lend / cap * 100).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) if cap > 0 else None
    out["표B"] = {
        "감정가_백만": mil(appraisal), "LTV": rate_str(ltv), "선순위_백만": mil(senior),
        "기대출_백만": mil(prior), "대출가능_원": int(trunc(avail)), "대출가능_백만": mil(avail),
        "적용가액_백만": mil(cap),  # 감정가(매매가) × LTV − 선순위 = 공동대출 총액 기준 대출가능금액
        "총여신비율": f"{ratio:.2f}%" if ratio is not None else "",
    }
    if amount > avail:
        chk.append(("B", "C01", f"신청금액 {comma(amount)}원 > 대출가능금액 {comma(avail)}원"))
    if ratio is not None and ratio > 100:
        chk.append(("B", "C02", f"총여신비율 {ratio}% > 100%"))

    # ---- 편중 (회천농협만, 증액이면 증액분만) ----
    hc_exist = sum((D(l.get("금액")) for l in loans if l.get("기관") == "회천농협"), Decimal(0))
    hc_refi = sum((D(l.get("금액")) for l in loans if l.get("기관") == "회천농협" and l.get("대환대상")), Decimal(0))
    concentration = hc_exist - hc_refi + (increment if is_inc else amount)
    cap_sp, cap_why = same_person_cap(case)
    if cap_why is None and concentration > SAME_PERSON_CAP_DEFAULT * MIL and not guar:
        chk.append(("B", "C24", "법인 동일인 5,000백만원 초과 — 조합원/준조합원 여부(100억 특례 요건) 확인"))
    out["편중"] = {"합계_원": int(concentration), "합계_백만": mil(concentration), "한도_백만": cap_sp,
                 "문장": (LIB["P01"] + "." if guar else
                          f"금차대출 포함 총 대출금액 {comma(mil(concentration))}백만원으로 우리농협 동일인당 총대출한도 {comma(cap_sp)}백만원 이내로 적정함.")}
    if concentration > cap_sp * MIL and not guar:  # 공공기관 보증서 담보는 동일인 한도 산정 안 함(P01)
        chk.append(("A", "C15", f"동일인 총대출 {comma(mil(concentration))}백만원 > {comma(cap_sp)}백만원 한도 초과 — 작성 중단"))

    # ---- 수지기여 ----
    base = increment if is_inc else amount
    income = trunc(base * rate / 100)
    out["수지기여"] = {
        "이자수익_원": int(income),
        "문장": (f"금차 증액대출 실행 시 연간 약 {comma(income)}원의 추가 이자수익 예상({rate_str(rate)}%)" if is_inc
                 else f"금차 대출실행 시 연간 약 {comma(income)}원의 이자수익 예상({rate_str(rate)}%)"),
    }
    first = s.get("최초인출액")
    if guar and not empty(first):  # 분할인출: 결재본은 1회차 기표 기준 (2026-11-03 업체Q건설)
        inc1 = trunc(D(first) * rate / 100)
        out["수지기여"]["최초인출_이자_원"] = int(inc1)
        out["수지기여"]["문장"] = (f"금차 대출 최초 인출({s.get('최초인출_표시') or '1회차'}) 시 연간 약 {comma(inc1)}원({rate_str(rate)}%)의 "
                               f"이자수익이 발생될 것으로 예상되며, 추후 사업 진행에 따라 추가 인출 예정")
        out["수지기여"]["표"] = {"type": "table", "title": "", "unit": "(단위 : 원)",
                              "header": ["구분", "실행(인출)금액", "금리", "연간 이자수익"],
                              "rows": [["최초 인출 예정", comma(first), f"{rate_str(rate)}%", comma(inc1)],
                                       ["약정금액 전액 인출", comma(amount), f"{rate_str(rate)}%", comma(income)]],
                              "widths": [40, 44, 30, 52]}

    # ---- 금융비용: 본건 + 기존대출 전체(타행·가계 포함). 증액·대환 대상은 본건으로 대체되므로 제외 ----
    # (2026-09-14 사용자 지정. 범위를 달리할 일이 있으면 그때 별도 요청)
    cost_base = [l for l in loans if not l.get("증액대상") and not l.get("대환대상")]
    missing_rate = [l for l in cost_base if empty(l.get("금리"))]
    fin_cost = sum((trunc(D(l["금액"]) * D(l["금리"]) / 100) for l in cost_base if not empty(l.get("금리"))), Decimal(0))
    fin_cost += trunc(amount * rate / 100)
    out["금융비용"] = {"원": int(fin_cost), "문장": f"본건 포함 연 금융비용(차입) {comma(fin_cost)}원"}
    if missing_rate:
        names = ", ".join(f"{l.get('기관')} {comma(trunc(D(l.get('금액')) / 1000))}천원" for l in missing_rate)
        chk.append(("B", "C16", f"금리 미기재 {len(missing_rate)}건({names}) → 금리를 주시면 금융비용에 반영, 모르면 제외하고 계산"))

    # ---- 운전자금 순여신 ----
    purpose = s.get("자금용도구분") or ""
    working_exist = sum((D(l.get("금액")) for l in loans
                         if l.get("기관") == "회천농협" and l.get("용도") == "운전" and not l.get("증액대상")), Decimal(0))
    working_net = working_exist + (amount if "운전" in purpose else Decimal(0))
    out["운전자금순여신_백만"] = mil(working_net)

    # ---- 적법성 트리거 ----
    L = []
    j = case.get("적법성", {})
    if concentration > 1000 * MIL:
        L.append(("L01", LIB["L01"]))
    if joint:
        L.append(("L02b", LIB["L02b"]))
    elif guar:
        if concentration >= 3000 * MIL:
            L.append(("P04", LIB["P04"]))
        L += [("P03", LIB["P03"]), ("P01", LIB["P01"])]  # P02(임금채권 미차감)는 넣지 않음 — 보증서는 차감 대상 아님(2026-11-03 사용자)
    elif concentration >= 3000 * MIL:
        L.append(("L02", LIB["L02"]))
    if s.get("당초금리") and is_inc:
        L.append(("L06", LIB["L06"]))
    if not empty(s.get("대심위우대금리")):
        L.append(("L07", LIB["L07"]))
    if j.get("경기민감업종") and not guar:  # 보증 건은 HUG 매뉴얼 6C 기준(P04)이 우선
        L.append(("L09", LIB["L09"]))
    if joint:
        # 공동대출 총한도 특례: 총액이 제4조② 기본한도(법인 100억·개인사업자 30억·개인 10억)를 넘을 때만 (2026-10-12 사용자 지정)
        form = get(case, "업체.기업형태", "") or ""
        base, tier3, tier4 = {"법인": (100, 200, 500), "개인사업자": (30, 90, None), "개인": (10, 30, None)}.get(form, (100, 200, 500))
        total = D(joint.get("총액") or 0)
        if total > base * 100 * MIL:
            if total <= tier3 * 100 * MIL:
                L.append(("L10", LIB["L10"].format(n=tier3)))
            else:
                L.append(("L10", LIB["L10b"].format(n=tier4 or tier3)))
                if tier4 is None or total > tier4 * 100 * MIL:
                    chk.append(("B", "C23", f"공동대출 총액 {comma(mil(total))}백만원 — 제4조 특례 한도 초과 여부 확인"))
    if D(h.get("LTV상향")) > 0:
        L.append(("L05", LIB["L05"].format(n=rate_str(h.get("LTV상향")))))
        if "거치" in (s.get("상환방식") or "") and "비거치" not in (s.get("상환방식") or ""):
            chk.append(("B", "C13", f"L05 문구는 '비거치 분할상환'인데 상환방식이 '{s.get('상환방식')}' — 조항·문구 확인"))
    if j.get("부동산담보신탁"):
        L.append(("L04", LIB["L04"]))
    elif j.get("임금채권미차감"):
        txt = LIB["L03"]
        if j.get("임금채권사유"):
            txt += "\n(" + j["임금채권사유"].strip("()") + ")"
        L.append(("L03", txt))
    if "운전" in purpose and amount <= 100 * MIL and working_net <= 500 * MIL:  # 건당 = 증액 후 한도 기준
        L.append(("L08", LIB["L08"]))
    excl = set(j.get("제외ID", []))
    triggered = [(i, t) for i, t in L if i not in excl]
    src = j.get("자료조항") or []
    if src:
        # 자료에 적힌 조항(번호 포함)을 그대로 쓴다. 라이브러리는 누락 탐지·의문 제기에만 쓴다.
        out["적법성"] = [{"id": x.get("id", ""), "문구": x["문구"]} for x in src]
        src_ids = {x.get("id") for x in src if x.get("id")}
        fired = {i for i, _ in triggered}
        out["적법성_누락후보"] = [{"id": i, "문구": t} for i, t in triggered if i not in src_ids]
        for x in src:
            if x.get("id") in LIB and x["id"] not in fired:
                chk.append(("B", "C20", f"자료에는 {x['id']} 조항이 있으나 계산상 요건 미해당 — {x['문구'][:40]}…"))
    else:
        out["적법성"] = [{"id": i, "문구": t} for i, t in triggered]
        out["적법성_누락후보"] = []

    # ---- 타당성 패턴 ----
    t_id, t_txt = "T0", ""
    refi_other = s.get("대환대상") == "타행"
    fin = case.get("재무") or []
    biz = (get(case, "업체.업종", "") or "") + " " + (get(case, "업체.주요제품", "") or "")
    is_dev = bool(re.search(r"부동산\s*개발|건설", biz))
    t1_cap = 1000 if "법인" in (get(case, "업체.기업형태", "") or "") else 500  # 백만원
    if guar and "운전" in purpose:
        t_id = "T7"
        t_txt = (LIB["T7"].format(보증=comma(appraisal)).replace("보증서발급금액", s.get("보증금액_라벨") or "보증서발급금액") + "\n○ " + LIB["T7b"] + "\n○ " + LIB["T7c"])
    elif "운전" in purpose:
        if refi_other:
            t_id, t_txt = "T4", LIB["T4"]
        elif is_dev:
            # 부동산개발업·건설업 운전자금: 총 소요자금 − 당해 공사관련 기차입금 범위 내 (2026-10-06 업체A)
            t_id, t_txt = "T6", LIB["T6"]
            tot, prev, borrow = s.get("총소요자금"), s.get("당해기차입금") or 0, s.get("사업차입금")
            if empty(tot) or empty(borrow):
                chk.append(("A", "C21", "부동산개발·건설업 운전자금(T6) → 총 소요자금·금차 사업 차입금 합계 필요"))
            else:
                room = D(tot) - D(prev)
                eq = D(tot) - D(prev) - D(borrow)
                t_txt += (f"\n(총 소요자금 {comma(mil(tot))}백만원 − 당해 공사관련 기차입금 {comma(mil(prev))}백만원 = "
                          f"{comma(mil(room))}백만원 ≥ 금차 사업 차입금 {comma(mil(borrow))}백만원, "
                          f"자기자금 {comma(mil(eq))}백만원({trunc(eq / D(tot) * 100, 2)}%))")
                if D(borrow) > room:
                    chk.append(("B", "C22", f"사업 차입금 {comma(mil(borrow))}백만원 > 소요자금 범위 {comma(mil(room))}백만원"))
                if s.get("T6_부기"):  # 단계별 조달(본PF 차환 등)처럼 자동 산식 문장이 맞지 않을 때 설명문 대체 (2026-10-13 업체A)
                    note6 = s["T6_부기"]
                    t_txt = LIB["T6"] + "\n" + (note6 if note6.startswith("(") else f"({note6})")
        elif working_net <= t1_cap * MIL:
            t_id, t_txt = "T1", LIB["T1"].format(한도=t1_cap // 100)
        else:
            t_id = "T2"
            last = fin[-1] if fin else {}
            sales, debt = last.get("매출액"), last.get("부채")
            if empty(sales) or empty(debt):
                chk.append(("A", "C17", f"운전자금 순여신 {t1_cap // 100}억 초과(T2) → 최근 결산 매출액·차입금 필요"))
                t_txt = LIB["T2"]
            else:
                cap_w = D(sales) - D(debt)
                t_txt = LIB["T2"] + f"\n(여신가능액 : 매출액 {comma(sales)}백만원 – 기차입금 {comma(debt)}백만원 = {comma(cap_w)}백만원)"
                if mil(working_net) > cap_w:
                    chk.append(("B", "C18", f"운전자금 순여신 {comma(mil(working_net))}백만원 > 여신가능액 {comma(cap_w)}백만원"))
    elif "시설" in purpose:
        if refi_other:
            t_id, t_txt = "T5", LIB["T5"]
        elif not empty(s.get("매매금액")):
            t_id = "T3"
            price = D(s["매매금액"])
            lim = price * Decimal("0.8")
            t_txt = LIB["T3"] + f"\n({s.get('매매대상') or ''}매매대금 {comma(mil(price))}백만원 × 80% = {comma(mil(lim))}백만원)"
            if amount > lim:
                chk.append(("B", "C08", f"시설자금 신청 {comma(mil(amount))}백만원 > 매매대금 80% {comma(mil(lim))}백만원"))
        else:
            chk.append(("A", "C19", "시설자금인데 매매(공사)금액이 없음 → 소요한도 산출 불가"))
    out["타당성"] = {"id": t_id, "문구": t_txt}
    fx = s.get("고정금리분석")
    if guar and fx:  # 고정금리 조달원가 민감도 (2026-11-04 업체P) — 연간 마진 = 잔액 × (대출금리 − 조달원가)
        avg = D(fx["평균잔액"])
        rows = []
        for c_rate, label in fx["조달원가"]:
            cr = D(c_rate)
            inc_a, cost_a = trunc(avg * rate / 100), trunc(avg * cr / 100)
            rows.append([f"{rate_str(cr)}%", label, comma(mil(inc_a)), comma(mil(cost_a)), comma(mil(inc_a - cost_a)),
                         comma(mil(trunc(amount * (rate - cr) / 100)))])
        out["고정금리분석"] = {
            "표": {"type": "table", "title": f"○ 조달원가별 연간 이자마진 (평균 실행잔액 {comma(mil(avg))}백만원 가정)", "unit": "(단위 : 백만원)",
                  "header": ["조달원가", "시나리오", "이자수익", "조달비용", "이자마진", "전액 인출 시 마진"],
                  "rows": rows, "widths": [18, 46, 24, 24, 24, 30]},
            "1bp": f"금리 1%p 변동 시 연간 이자 영향 : 약정 전액 기준 {comma(mil(trunc(amount / 100)))}백만원, "
                   f"평균 실행잔액 기준 {comma(mil(trunc(avg / 100)))}백만원",
        }
    if guar:
        out["산출근거_보증"] = f"{comma(appraisal)}원(보증금액) x {rate_str(ltv)}%(HUG) = {comma(trunc(cap))}원"

    # ---- 승인조건 ----
    ap = [APPROVAL["보증서류"]] if guar else []
    if j.get("임금채권미차감") and not j.get("부동산담보신탁") and not guar:
        ap.append(APPROVAL["임금"])
    if s.get("당초금리") and is_inc:
        ap.append(APPROVAL["당초금리"])
    if concentration >= 3000 * MIL and not joint and not guar:
        ap.insert(0, APPROVAL["고액신설"] if get(case, "업체.신설법인") else APPROVAL["고액"])
    ap = [a for a in ap if a not in case.get("승인조건_제외", [])] + case.get("승인조건_추가", [])
    out["승인조건"] = ap or ["해당없음"]

    # ---- 신용등급 표기 ----
    out["신용등급표시"] = credit_label(case.get("신용", {}))

    # ---- 기타 검산 ----
    run_checks(case, out)
    return out


def credit_label(c):
    if c.get("유형") == "기업":
        return f"기업 신용등급 {c.get('기업등급') or ''}".strip()
    parts = []
    if not empty(c.get("CSS등급")):
        parts.append(f"CSS {c['CSS등급']}등급")
    kind = c.get("점수구분") or "소호"  # "CB"면 개인CB 점수 (2026-10-28 차주09)
    sep = " " if kind == "CB" else ""
    if not empty(c.get("NICE소호점수")):
        parts.append(f"NICE{sep}{kind} {c['NICE소호점수']}점({c.get('NICE소호등급') or ''})")
    if not empty(c.get("KCB소호점수")):
        parts.append(f"KCB{sep}{kind} {c['KCB소호점수']}점({c.get('KCB소호등급') or ''})")
    return ", ".join(parts)


def run_checks(case, out):
    chk = out["checks"]
    s, h, raw = case.get("신청", {}), case.get("담보한도", {}), case.get("원자료", {})
    cols = case.get("담보물", [])

    # C03 담보물 평가액 합 vs 감정가
    if cols and not empty(h.get("감정가")):
        tot = sum((D(r.get("평가액")) for c in cols for r in c.get("행", [])), Decimal(0))
        if tot and mil(tot * MIL) != mil(D(h["감정가"])) and abs(tot - D(h["감정가"]) / MIL) >= 1:
            chk.append(("B", "C03", f"담보물 표 평가액 합 {comma(tot, 1)}백만원 ≠ 감정가 {comma(D(h['감정가']) / MIL, 1)}백만원"))

    # C04/C05 감정일·감정기관 일관성
    dates = {c.get("감정일") for c in cols if c.get("감정일")}
    if raw.get("감정시점"):
        dates.add(raw["감정시점"])
    if len(dates) > 1:
        chk.append(("B", "C04", f"감정일 불일치: {sorted(dates)}"))
    orgs = {re.sub(r"감정(평가)?(법인)?$", "", o or "") for o in
            [h.get("감정기관"), raw.get("감정기관")] + [c.get("감정기관") for c in cols] if o}
    if len(orgs) > 1:
        chk.append(("B", "C05", f"감정기관 불일치: {sorted(orgs)}"))

    # C06 신용등급 원자료 대조
    if raw.get("신용등급") and raw["신용등급"] != out["신용등급표시"]:
        chk.append(("B", "C06", f"신용등급 표기 '{out['신용등급표시']}' ≠ 원자료 '{raw['신용등급']}' — 신용평가 화면 기준 확인"))

    # C07 기존대출 합계
    if not empty(raw.get("기존대출합계")):
        tot = sum((D(l.get("금액")) for l in case.get("금융거래", [])), Decimal(0))
        if tot != D(raw["기존대출합계"]):
            chk.append(("B", "C07", f"금융거래 명세 합계 {comma(tot)}원 ≠ 사전검토 합계 {comma(raw['기존대출합계'])}원"))

    # C09 금리 = 기준 + 가산
    if not empty(s.get("기준금리")) and not empty(s.get("가산금리")) and not empty(s.get("금리")):
        if abs(D(s["기준금리"]) + D(s["가산금리"]) - D(s["금리"])) > Decimal("0.005"):
            chk.append(("B", "C09", f"금리 {s['금리']} ≠ 기준 {s['기준금리']} + 가산 {s['가산금리']}"))

    # C10 날짜
    m = case.get("meta", {})
    if m.get("작성일") and m.get("사전검토일") and m["작성일"] < m["사전검토일"]:
        chk.append(("B", "C10", f"작성일 {m['작성일']} < 사전검토일 {m['사전검토일']}"))

    # C11 입지현황 문장의 읍면동리 vs 소재지
    ip = get(case, "담보개요.입지", "") or ""
    loc_all = " ".join([c.get("소재지") or "" for c in cols] + [raw.get("담보물주소") or ""])
    tok = lambda t: set(re.findall(r"([가-힣]{1,4}(?:읍|면|동|리))(?=[\s,·]|$)", t))
    t_ip, t_loc = tok(ip), tok(loc_all)
    t_ip = {t for t in t_ip if t not in ("주변", "측면", "전면", "후면", "이면", "정면", "마을리")}
    if t_ip and t_loc and not (t_ip & t_loc):
        chk.append(("B", "C11", f"입지현황 지명 {sorted(t_ip)} 이 담보물 소재지 {sorted(t_loc)} 와 다름 — 다른 물건 문구 복붙 의심"))

    # C12 시도 표기
    for c in cols:
        loc = c.get("소재지") or ""
        if loc.startswith("경기") and any(g in loc for g in SEOUL_GU):
            chk.append(("B", "C12", f"소재지 '{loc}' — 서울 자치구명인데 '경기'로 표기"))


# ---------- 누락 탐지 (작성 전 질문) ----------
REQUIRED_A = [
    ("업체.업체명", "채무자(업체)명"), ("업체.기업형태", "기업형태(법인/개인사업자)"),
    ("신청.사무소", "신청 사무소(지점)"), ("신청.구분", "구분(신규/증액/대환/재대출/기한연기)"),
    ("신청.대출금액", "대출금액"), ("신청.금리", "최종금리"), ("신청.금리종류", "기준금리 종류(MOR N개월/고정)"),
    ("신청.기간", "대출기간"), ("신청.상환방식", "상환방식"), ("신청.대출과목", "대출과목"),
    ("신청.자금용도구분", "자금용도 구분(기업운전자금/기업시설자금 등)"), ("신청.자금용도상세", "구체적 자금용도"),
    ("담보한도.감정가", "감정가"), ("담보한도.LTV", "담보인정비율"), ("담보한도.감정기관", "감정기관"),
    ("신용.유형", "신용등급 유형(기업/CSS)"),
]
SUPPLEMENT_C = [
    ("업체.사업자번호", "사업자번호", "사업자등록증"), ("업체.설립일", "설립일", "사업자등록증/등기부"),
    ("업체.주소", "사업장 주소", "사업자등록증"), ("업체.주요제품", "주요제품", "사업자등록증 종목"),
    ("업체.주민법인번호", "주민/법인번호", "사업자등록증"), ("업체.대표자", "대표자", "사업자등록증"),
    ("업체.기업규모", "기업규모", "신용평가 화면"),
    ("담보개요.입지", "입지현황", "사전검토 인근현황"), ("담보개요.교통", "교통상황", "사전검토 도로상황"),
    ("담보개요.공법", "공법상 제한사항", "사전검토 도시계획"),
    ("적정성.감정단가", "본건 감정단가", "감정평가서"),
    ("전망.업체", "업체전망", "사전검토 사업성 분석"),
    ("종합.긍정", "긍정적인 사항", "사전검토 사업성 분석"), ("종합.부정", "부정적인 사항", "사전검토 사업성 분석"),
]


def gaps(case, out):
    rows = []
    s = case.get("신청", {})
    for p, name in REQUIRED_A:
        if empty(get(case, p)):
            rows.append(("A", p, f"{name} 없음"))
    if s.get("구분") == "증액" and empty(s.get("증액전금액")):
        rows.append(("A", "신청.증액전금액", "증액 전 기존 대출금액(증액대상 계좌) 없음"))
    if empty(case.get("금융거래")) and s.get("구분") != "신규":
        rows.append(("A", "금융거래", "기존 여신 명세 없음"))
    for l in case.get("금융거래", []):
        if l.get("기관") == "회천농협" and empty(l.get("용도")):
            rows.append(("A", "금융거래.용도", f"회천농협 {comma(l.get('금액'))}원 건의 용도(운전/시설/가계) 없음 → 금융비용·순여신 계산 불가"))
    c = case.get("신용", {})
    if c.get("유형") == "기업" and empty(c.get("기업등급")):
        rows.append(("A", "신용.기업등급", "기업신용등급 없음"))
    if c.get("유형") == "CSS" and empty(c.get("CSS등급")):
        rows.append(("A", "신용.CSS등급", "CSS 등급 없음"))
    if c.get("유형") == "CSS" and (empty(c.get("NICE소호점수")) or empty(c.get("KCB소호점수"))):
        rows.append(("C", "신용.소호", "NICE/KCB 소호 점수·등급 — 신용평가 결과 화면"))

    for lvl, cid, msg in out["checks"]:
        rows.append((lvl, cid, msg))

    guar = case.get("상품") == "사업비대출"
    skip = {"적정성.감정단가"} if guar else set()
    for p, name, src in SUPPLEMENT_C:
        if p not in skip and empty(get(case, p)):
            rows.append(("C", p, f"{name} — 필요자료: {src}"))
    if get(case, "업체.기업형태") == "개인사업자" and empty(get(case, "상환능력.소득")):
        rows.append(("C", "상환능력.소득", "연소득·재산세·매출액 — 소득금액증명원, 지방세세목별과세증명, 부가세과세표준증명"))
    if get(case, "업체.기업형태") == "법인" and not get(case, "업체.신설법인") and empty(case.get("재무")):
        rows.append(("C", "재무", "결산 재무(3개년) — 표준재무제표증명원"))
    cols = case.get("담보물", [])
    if guar:
        pass  # 보증서 담보: 감정·탁상·거래사례·낙찰 항목 대상 아님 (rules/07_사업비대출.md)
    elif not cols:
        rows.append(("C", "담보물", "담보물 내역(지목·구조·면적·평가액·권리사항) — 감정평가서·등기부"))
    if not guar and empty(get(case, "적정성.탁상감정")):
        rows.append(("C", "적정성.탁상감정", "탁상감정 가격비교 — 탁상감정 회신"))
    if not guar and empty(get(case, "적정성.거래사례")) and empty(get(case, "적정성.지가수준")):
        rows.append(("C", "적정성.거래사례", "인근 거래사례/지가수준 — 감정평가서"))
    if not guar and empty(get(case, "적정성.낙찰사례.이미지")) and empty(get(case, "적정성.낙찰사례.표")):
        rows.append(("C", "적정성.낙찰사례", "낙찰사례 — 인포케어 캡처 이미지"))
    if case.get("보증인") is None:
        rows.append(("C", "보증인", "보증인 여부 (없으면 [] 로 '해당없음')"))

    # D: 판단 질문
    if not get(case, "적법성.자료조항"):
        rows.append(("C", "적법성.자료조항", "자료(사전검토 적법성심사 등)에서 조항을 뽑지 않음 → 라이브러리 트리거로 대체 작성됨"))
    if out["적법성_누락후보"]:
        rows.append(("D", "적법성", "자료에는 없지만 요건상 해당될 수 있는 조항 — 추가할지: "
                     + " / ".join(f"{x['id']} {x['문구'][:30]}…" for x in out["적법성_누락후보"])))
    if empty(get(case, "전망.업계")):
        rows.append(("D", "전망.업계", "업계전망 근거 없음 — 일반지식으로 작성할지 / 웹검색 허용할지"))
    if case.get("신용위험") is None:
        rows.append(("D", "신용위험", "주요 신용위험 섹션: '해당없음'으로 넣을지 / 뺄지"))
    if not guar and (s.get("공동대출") or re.search(r"임대", get(case, "업체.업종", "") or "")):
        if get(case, "업체.기업형태") == "법인":
            rows.append(("D", "상환능력.현금흐름", "임대업/공동대출 법인 → DSCR 현금흐름표 넣을지 (가정값 필요)"))
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("case")
    ap.add_argument("--gaps", action="store_true")
    a = ap.parse_args()
    path = Path(a.case)
    case = json.loads(path.read_text(encoding="utf-8"))
    out = compute(case)

    if a.gaps:
        rows = gaps(case, out)
        order = {"A": 0, "B": 1, "C": 2, "D": 3}
        label = {"A": "A 필수(없으면 작성불가)", "B": "B 충돌·확인", "C": "C 보완자료(없으면 공란)", "D": "D 판단(사용자 선택)"}
        cur = None
        for lvl, key, msg in sorted(rows, key=lambda r: order[r[0]]):
            if lvl != cur:
                print(f"\n[{label[lvl]}]")
                cur = lvl
            print(f"  - ({key}) {msg}")
        n_a = sum(1 for r in rows if r[0] == "A")
        print(f"\nSUMMARY A={n_a} B={sum(1 for r in rows if r[0]=='B')} C={sum(1 for r in rows if r[0]=='C')} D={sum(1 for r in rows if r[0]=='D')}")
        print("READY" if n_a == 0 else "BLOCKED: A등급 해소 전 작성 금지")
        return

    (path.parent / "derived.json").write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in out.items() if k != "checks"}, ensure_ascii=False, indent=2))
    print("\n[검산]")
    for lvl, cid, msg in out["checks"] or [("-", "-", "이상 없음")]:
        print(f"  {lvl} {cid} {msg}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
