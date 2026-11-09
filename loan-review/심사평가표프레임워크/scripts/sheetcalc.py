# -*- coding: utf-8 -*-
"""심사평가표 / 심사의견요약표 — 계산 어댑터

대출가능금액·이자수익·편중·검산은 **검토의견서프레임워크의 calc.py를 그대로 재사용**한다(산식 이중관리 금지).
이 파일이 하는 일은 셋뿐이다.
  ① sheet.json → calc.py 입력 형식으로 변환
  ② 이 양식 전용 계산(금리 산출근거 조립, 대출가능금액 산식 문장, 소요자금한도, CB식 신용등급 표기)
  ③ 작성 전 누락·충돌 탐지(--gaps)

사용법
  python sheetcalc.py <sheet.json>          계산 결과 출력
  python sheetcalc.py <sheet.json> --gaps   작성 전 질문할 항목(A필수/B충돌/C보완/D판단)
"""
import argparse
import json
import sys
from pathlib import Path

LOAN_FW = Path(__file__).resolve().parents[2] / "검토의견서프레임워크" / "scripts"  # 형제 폴더 기준(폴더 이동에 안전)
if not (LOAN_FW / "calc.py").exists():
    raise SystemExit(f"검토의견서프레임워크 calc.py를 찾을 수 없습니다: {LOAN_FW}")
sys.path.insert(0, str(LOAN_FW))

from calc import D, comma, compute, empty, get, mil, rate_str, trunc  # noqa: E402

MIL = 1_000_000


# ---------- sheet.json → calc.py 입력 ----------
def to_case(sh):
    """심사평가표 스키마를 검토의견서 calc.py가 아는 모양으로 바꾼다."""
    s, h = sh.get("신청", {}), sh.get("담보한도", {})
    loans = []
    for r in sh.get("기대출") or []:
        용도 = r.get("용도")
        if not 용도:
            t = (r.get("자금용도") or "") + (r.get("비고") or "")
            용도 = "가계" if "가계" in t else ("시설" if "시설" in t else "운전")
        loans.append({"기관": "회천농협", "금액": r.get("잔액"), "금리": r.get("금리"), "용도": 용도,
                      "신용담보": r.get("담보구분"), "동일담보": bool(r.get("동일담보") or "동일담보" in (r.get("비고") or "")),
                      "증액대상": bool(r.get("증액대상")), "대환대상": bool(r.get("대환대상")), "비고": r.get("비고")})
    return {
        "meta": {"작성일": get(sh, "meta.작성일")},
        "신청": {"구분": s.get("구분", "신규"), "대출금액": s.get("금액"), "증액전금액": s.get("증액전금액"),
               "금리": s.get("금리"), "금리종류": s.get("금리종류"), "기준금리": s.get("기준금리"),
               "가산금리": s.get("가산금리"), "기간": s.get("기간"), "상환방식": s.get("상환방식"),
               "대출과목": s.get("대출과목"), "자금용도구분": s.get("자금용도구분"),
               "자금용도상세": s.get("자금용도상세"), "매매금액": s.get("매매금액"),
               "대환대상": s.get("대환대상"), "대심위우대금리": s.get("대심위우대금리")},
        "담보한도": {"감정가": h.get("감정가"), "감정기관": h.get("감정기관"), "LTV": h.get("LTV"),
                  "LTV상향": h.get("LTV상향", 0), "담보종류표시": h.get("담보종류표시"),
                  "선순위차감": 차감합(h) - D(h.get("기대출")), "기대출": h.get("기대출")},
        "업체": {"기업형태": sh.get("채무자", {}).get("기업형태", "개인사업자")},
        "신용": sh.get("신용", {}),
        "금융거래": loans,
        "적법성": {"자료조항": [{"id": "", "문구": x} for x in (get(sh, "요약표.조항") or [])]},
        "재무": sh.get("재무") or [],
    }


def merge_loans(loans, limit=6):
    """기대출이 limit행 이상이면 자금용도가 같은 계좌끼리 묶는다(금액 합산, 금리는 최대값).
    2026-09-16 사용자 지정 — 칸이 많아 표가 길어지는 것을 막기 위함."""
    loans = [x for x in (loans or []) if not empty(x.get("잔액"))]
    if len(loans) < limit:
        return loans
    groups = {}
    for x in loans:
        k = x.get("자금용도") or ""
        g = groups.setdefault(k, {"잔액": D(0), "한도": D(0), "자금용도": k, "금리": None,
                                  "담보구분": [], "비고": [], "_n": 0})
        g["잔액"] += D(x.get("잔액"))
        g["한도"] += D(x.get("한도"))
        g["_n"] += 1
        if not empty(x.get("금리")):
            g["금리"] = max(g["금리"], D(x["금리"])) if g["금리"] is not None else D(x["금리"])
        for f in ("담보구분", "비고"):
            v = x.get(f)
            if v and v not in g[f]:
                g[f].append(v)
    out = []
    for g in groups.values():
        out.append({"잔액": int(g["잔액"]), "한도": int(g["한도"]) or None, "자금용도": g["자금용도"],
                    "금리": g["금리"], "담보구분": "/".join(g["담보구분"]),
                    "비고": "/".join(g["비고"]) + (f" ({g['_n']}건 합산)" if g["_n"] > 1 else "")})
    return out


def 차감합(h):
    """기대출 + 선순위 차감내역 합계(원)"""
    tot = D(h.get("기대출"))
    for x in h.get("차감내역") or []:
        tot += D(x.get("금액"))
    return tot


# ---------- 이 양식 전용 계산 ----------
def credit_label_cb(c):
    """이 양식은 'NICE CB' 표기를 쓴다(검토의견서의 'NICE소호'와 다름)."""
    if c.get("유형") == "기업":
        return f"기업 신용등급 {c.get('기업등급') or ''}".strip()
    parts = []
    if not empty(c.get("CSS등급")):
        parts.append(f"CSS {c['CSS등급']}등급")
    if not empty(c.get("NICE점수")):
        parts.append(f"NICE CB {c['NICE점수']}점({c.get('NICE등급') or ''})")
    if not empty(c.get("KCB점수")):
        parts.append(f"KCB CB {c['KCB점수']}점({c.get('KCB등급') or ''})")
    return ", ".join(parts)


def 금리계산(s):
    """기준금리 + 원가가산 − 우대 = 결정금리. 산출근거 문자열까지 만든다."""
    base, cost = D(s.get("기준금리")), D(s.get("원가가산"))
    plus = sum((D(x.get("값")) for x in s.get("가산내역") or []), D(0))
    minus = sum((D(x.get("값")) for x in s.get("우대내역") or []), D(0))
    spread = cost + plus - minus
    final = base + spread
    # 산출근거_기준명: 금리화면의 기준금리 명칭이 대출조건 금리종류와 다르게 적힌 경우(차주10 결재본)만 쓴다
    parts = [f"{rate_str(base)}%({s.get('산출근거_기준명') or s.get('금리종류', '')})", f"+{rate_str(cost)}%(원가가산)"]
    parts += [f"+{rate_str(x.get('값'))}%({x.get('명목', '')})" for x in s.get("가산내역") or []]
    parts += [f"-{rate_str(x.get('값'))}%({x.get('명목', '')})" for x in s.get("우대내역") or []]
    return {"가산": spread, "결정": final, "산출근거": " ".join(parts)}


def 대출가능_산식(h, avail):
    """'감정가 859,180,000원 × 70%(답) - 기대출 500,000,000원 - 소액임차 10,000,000원(제시외)'"""
    ltv = f"{rate_str(h.get('LTV'))}%" + (f"({h.get('담보종류표시')})" if h.get("담보종류표시") else "")
    line1 = f"감정가 {comma(h.get('감정가'))}원 × {ltv}"
    subs = []
    if not empty(h.get("기대출")) and D(h.get("기대출")) > 0:
        subs.append(f"- 기대출 {comma(h['기대출'])}원")
    for x in h.get("차감내역") or []:
        note = f"({x['비고']})" if x.get("비고") else ""
        subs.append(f"- {x.get('명목', '차감')} {comma(x.get('금액'))}원{note}")
    if h.get("산식_한줄") and subs:   # 차주10 결재본: '× 60%(아파트임대) -소액임차 25,000,000원' 한 줄
        return [f"대출가능금액 : {comma(avail)}원", line1 + " " + " ".join(x.replace("- ", "-", 1) for x in subs)]
    return [f"대출가능금액 : {comma(avail)}원", line1] + ([" ".join(subs)] if subs else [])


def compute_sheet(sh):
    case = to_case(sh)
    out = compute(case)
    s, h = sh.get("신청", {}), sh.get("담보한도", {})
    chk = out["checks"]

    out["신용등급표시"] = credit_label_cb(sh.get("신용", {}))
    out["금리"] = 금리계산(s)
    if not empty(s.get("금리")) and abs(out["금리"]["결정"] - D(s["금리"])) > D("0.005"):
        chk.append(("B", "S01", f"금리 산출근거 합계 {rate_str(out['금리']['결정'])}% ≠ 기재 금리 {rate_str(s['금리'])}%"))
    if not empty(s.get("가산금리")) and abs(out["금리"]["가산"] - D(s["가산금리"])) > D("0.005"):
        chk.append(("B", "S02", f"계산 가산 {rate_str(out['금리']['가산'])}% ≠ 기재 가산 {rate_str(s['가산금리'])}%"))

    avail = D(h.get("감정가")) * D(h.get("LTV")) / 100 - 차감합(h)
    out["대출가능금액"] = int(trunc(avail))
    out["대출가능_산식"] = 대출가능_산식(h, avail)
    if not empty(h.get("감정가")) and D(s.get("금액")) > avail:
        chk.append(("B", "S03", f"신청금액 {comma(s.get('금액'))}원 > 대출가능금액 {comma(avail)}원"))

    if not empty(s.get("매매금액")):
        lim = D(s["매매금액"]) * D("0.8")
        out["소요자금한도"] = f"소요자금한도 : 매매계약 {comma(mil(s['매매금액']))}백만원 × 80% = {comma(mil(lim))}백만원"
        if D(s.get("금액")) > lim:
            chk.append(("B", "S04", f"시설자금 신청 {comma(mil(s['금액']))}백만원 > 매매금액 80% {comma(mil(lim))}백만원"))
    else:
        out["소요자금한도"] = None

    out["기대출합계"] = int(sum((D(r.get("잔액")) for r in sh.get("기대출") or []), D(0)))
    out["이자수익문장"] = (f"금차 대출 실행 시 연간 약 {comma(trunc(D(s.get('금액')) * D(s.get('금리')) / 100))}원의 "
                     f"이자수익 예상({rate_str(s.get('금리'))}%)")
    return out


# ---------- 누락 탐지 ----------
REQ = [("meta.작성일", "작성일"), ("meta.신청영업점", "신청영업점"), ("채무자.성명", "채무자명"),
       ("신청.대출과목", "대출과목"), ("신청.금액", "대출금액"), ("신청.기간", "대출기간"),
       ("신청.상환방식", "상환방식"), ("신청.금리", "금리"), ("신청.금리종류", "기준금리 종류"),
       ("신청.기준금리", "기준금리"), ("신청.원가가산", "원가가산금리"), ("신청.자금용도구분", "자금용도"),
       ("신용.유형", "신용조사 구분(기업/CSS)")]
SUP = [("담보한도.소재지", "담보물 소재지", "감정평가서"), ("담보한도.감정일", "감정일", "감정평가서"),
       ("담보한도.감정기관", "감정기관", "감정평가서"),
       ("요약표.소득", "소득금액증명원 연소득", "소득금액증명원"),
       ("요약표.재산세", "지방세세목별과세증명 재산세", "지방세 과세증명"),
       ("요약표.조항", "적용 규정 조항", "사전검토 적법성심사·규정 발췌")]


def gaps_sheet(sh, out):
    rows = []
    for p, name in REQ:
        if empty(get(sh, p)):
            rows.append(("A", p, f"{name} 없음"))
    h = sh.get("담보한도", {})
    if sh.get("담보대출", True) and (empty(h.get("감정가")) or empty(h.get("LTV"))):
        rows.append(("A", "담보한도", "감정가·담보인정비율 없음 → 대출가능금액 산출 불가"))
    c = sh.get("신용", {})
    if c.get("유형") == "CSS" and empty(c.get("CSS등급")):
        rows.append(("A", "신용.CSS등급", "CSS 등급 없음"))
    if c.get("유형") == "CSS" and (empty(c.get("NICE점수")) or empty(c.get("KCB점수"))):
        rows.append(("C", "신용.CB", "NICE/KCB CB 점수·등급 — 신용평가 결과 화면"))
    for lvl, cid, msg in out["checks"]:
        rows.append((lvl, cid, msg))
    for p, name, src in SUP:
        if empty(get(sh, p)):
            rows.append(("C", p, f"{name} — 필요자료: {src}"))
    if sh.get("평가표", {}).get("특인내용") is None:
        rows.append(("D", "평가표.특인내용", "특인 사항: '특인 없음'인지 / 조항을 적을지"))
    if sh.get("승인조건") is None:
        rows.append(("D", "승인조건", "실행조건·관리조건: '해당없음'인지 / 조건을 적을지"))
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("sheet")
    ap.add_argument("--gaps", action="store_true")
    a = ap.parse_args()
    sh = json.loads(Path(a.sheet).read_text(encoding="utf-8"))
    out = compute_sheet(sh)
    if a.gaps:
        order = {"A": 0, "B": 1, "C": 2, "D": 3}
        label = {"A": "A 필수(없으면 작성불가)", "B": "B 충돌·확인", "C": "C 보완자료(없으면 공란)", "D": "D 판단(사용자 선택)"}
        rows = gaps_sheet(sh, out)
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
    show = {k: v for k, v in out.items() if k not in ("checks", "적법성", "적법성_누락후보", "타당성", "승인조건", "표B", "notes")}
    show["금리"] = {k: (rate_str(v) if k != "산출근거" else v) for k, v in out["금리"].items()}
    print(json.dumps(show, ensure_ascii=False, indent=2, default=str))
    print("\n[검산]")
    for lvl, cid, msg in out["checks"] or [("-", "-", "이상 없음")]:
        print(f"  {lvl} {cid} {msg}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
