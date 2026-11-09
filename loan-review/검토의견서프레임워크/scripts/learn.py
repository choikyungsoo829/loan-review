# -*- coding: utf-8 -*-
"""검토의견서 학습 로그 — 결재 최종본과 초안을 비교해 수정 패턴을 쌓고, 반복되면 규칙 후보로 올린다.
자동으로 규정(rules/*.md)을 고치지 않는다. candidates.md는 사람이 보고, 승인하면 Claude가 손으로 규칙을 고친다.

사용법
  python learn.py record <case_dir> <final.docx|final.pdf>   결재본과 초안을 비교해 diffs.jsonl·scorecard.jsonl에 적재
  python learn.py candidates                                  2건 이상 반복된 수정 패턴을 candidates.md에 갱신
  python learn.py close <id> approved|rejected [메모]          후보를 closed.md로 옮겨 재제안 방지(승인 시 규칙은 손으로 고칠 것)
  python learn.py score                                        누적 스코어카드 요약
"""
import argparse
import difflib
import hashlib
import json
import re
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

from docx import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent.parent
LEARN = ROOT / "learning"
DIFFS = LEARN / "diffs.jsonl"
SCORE = LEARN / "scorecard.jsonl"
CANDIDATES = LEARN / "candidates.md"
CLOSED = LEARN / "closed.md"
HEAD_RE = re.compile(r"^(\d\.\s|[가나다라마바사]\.\s|\(\d\)\s|[○※])")
PAGE_MARK_RE = re.compile(r"^-\s*\d+\s*-$")  # PDF 쪽번호 꼬리("- 4 -") — 내용 신호 아님


# ---------- 텍스트 추출 ----------
def _iter_blocks(doc):
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def _row_line(cells_text):
    dedup = []
    for c in cells_text:
        if not dedup or dedup[-1] != c:
            dedup.append(c)
    return " | ".join(x for x in dedup if x)


def _table_lines(t):
    """행이 1칸(상자형 서술 블록)이면 문단 단위로 쪼개 세밀하게, 여러 칸(표)이면 행 단위로 합쳐서."""
    out = []
    for row in t.rows:
        cells = row.cells
        if len(cells) == 1:
            out += [s.strip() for s in cells[0].text.split("\n") if s.strip()]
        else:
            line = _row_line([c.text.replace("\n", "/").strip() for c in cells])
            if line:
                out.append(line)
        for c in cells:
            for nested in c.tables:
                out += _table_lines(nested)
    return out


def docx_text(path):
    doc = Document(str(path))
    lines = []
    for block in _iter_blocks(doc):
        if isinstance(block, Paragraph):
            t = block.text.strip()
            if t:
                lines.append(t)
        else:
            lines += _table_lines(block)
    return lines


def pdf_text(path):
    import pdfplumber
    lines = []
    with pdfplumber.open(str(path)) as pdf:
        for pg in pdf.pages:
            for ln in (pg.extract_text() or "").splitlines():
                ln = ln.strip()
                if ln and not PAGE_MARK_RE.match(ln):
                    lines.append(ln)
    return lines


def text_lines(path):
    p = Path(path)
    if p.suffix.lower() == ".pdf":
        return pdf_text(p)
    if p.suffix.lower() == ".docx":
        return docx_text(p)
    raise SystemExit(f"지원 형식 아님(.docx/.pdf만): {p}  (.hwp면 PDF로 저장해서 다시)")


def find_draft(case_dir):
    docs = sorted(Path(case_dir).glob("심사담당자 검토의견서(*).docx"))
    non_draft = [d for d in docs if "_초안" not in d.stem]
    picked = non_draft or docs
    if not picked:
        raise SystemExit(f"draft .docx 없음: {case_dir}")
    return picked[-1]


# ---------- diff ----------
def section_of(lines, idx):
    for i in range(idx, -1, -1):
        if HEAD_RE.match(lines[i]):
            return lines[i][:40]
    return "(머리글 없음)"


def diff_hunks(a, b):
    sm = difflib.SequenceMatcher(a=a, b=b, autojunk=False)
    hunks = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            continue
        removed, added = a[i1:i2], b[j1:j2]
        if norm(removed) == norm(added):
            continue  # 위치만 밀린 것(내용은 동일) — 노이즈이므로 제외
        anchor = j1 if j1 < len(b) else i1
        anchor_lines = b if j1 < len(b) else a
        hunks.append({"section": section_of(anchor_lines, anchor), "removed": removed, "added": added})
    return hunks


def norm(lines):
    """숫자·콤마를 지워 패턴만 남긴다 — 같은 수정을 값만 다르게 반복해도 하나로 묶기 위함."""
    return re.sub(r"[\d,.]+", "#", " ¶ ".join(lines))


# ---------- 명령 ----------
def cmd_record(args):
    case_dir = Path(args.case_dir)
    case_id = case_dir.name
    draft = find_draft(case_dir)
    a, b = text_lines(draft), text_lines(args.final)
    hunks = diff_hunks(a, b)
    meta_p = case_dir / ".gaps_snapshot.json"
    gaps_meta = json.loads(meta_p.read_text(encoding="utf-8")) if meta_p.exists() else {}
    today = date.today().isoformat()
    LEARN.mkdir(exist_ok=True)
    with DIFFS.open("a", encoding="utf-8") as f:
        for h in hunks:
            f.write(json.dumps({"case": case_id, "date": today, **h}, ensure_ascii=False) + "\n")
    with SCORE.open("a", encoding="utf-8") as f:
        f.write(json.dumps({"case": case_id, "date": today, "gaps_at_render": gaps_meta, "hunks": len(hunks),
                             "sections_changed": len({h["section"] for h in hunks})}, ensure_ascii=False) + "\n")
    print(f"OK {case_id}: 결재본 대비 수정 {len(hunks)}건(섹션 {len({h['section'] for h in hunks})}개) 기록")
    for h in hunks[:25]:
        print(f"  [{h['section']}]")
        for r in h["removed"]:
            print(f"    - {r}")
        for a2 in h["added"]:
            print(f"    + {a2}")
    print(f"→ {DIFFS.name}, {SCORE.name} 갱신. 2회 이상 반복되면 `python learn.py candidates`로 후보 확인")


def cmd_candidates(args):
    if not DIFFS.exists():
        print("diffs.jsonl 없음 — 먼저 record")
        return
    closed_ids = set()
    if CLOSED.exists():
        closed_ids = set(re.findall(r"^- `(\w+)`", CLOSED.read_text(encoding="utf-8"), re.M))
    groups = defaultdict(list)  # sig -> [(case,date,hunk)]
    for line in DIFFS.read_text(encoding="utf-8").splitlines():
        h = json.loads(line)
        sig = hashlib.sha1((h["section"] + "|" + norm(h["added"]) + "|" + norm(h["removed"])).encode()).hexdigest()[:8]
        groups[sig].append(h)
    rows = []
    for sig, hs in groups.items():
        cases = {h["case"] for h in hs}
        if len(cases) < 2 or sig in closed_ids:
            continue
        rows.append((sig, hs))
    rows.sort(key=lambda x: -len({h["case"] for h in x[1]}))
    lines = ["# 규칙 후보 (2건 이상 반복된 수정 패턴)",
              "",
              "사람이 검토 후 승인하면 Claude가 `rules/01_작성규칙.md` 등을 손으로 고치고,",
              "`python learn.py close <id> approved` 로 닫는다. 반려는 `rejected`.",
              ""]
    for sig, hs in rows:
        cases = sorted({h["case"] for h in hs})
        ex = hs[0]
        lines.append(f"## `{sig}` — {ex['section']}  ({len(cases)}건: {', '.join(cases)})")
        if ex["removed"]:
            lines.append("- 삭제/변경 전: " + " / ".join(ex["removed"][:3]))
        if ex["added"]:
            lines.append("- 결재본 최종: " + " / ".join(ex["added"][:3]))
        lines.append("")
    CANDIDATES.write_text("\n".join(lines), encoding="utf-8")
    print(f"OK 후보 {len(rows)}건 → {CANDIDATES}")
    for sig, hs in rows:
        print(f"  {sig}  {hs[0]['section']}  ({len({h['case'] for h in hs})}건)")


def cmd_close(args):
    CLOSED.parent.mkdir(exist_ok=True)
    with CLOSED.open("a", encoding="utf-8") as f:
        f.write(f"- `{args.id}` {args.status} ({date.today().isoformat()}) {args.note or ''}\n")
    print(f"OK {args.id} → {args.status}로 닫음")


def cmd_score(args):
    if not SCORE.exists():
        print("scorecard.jsonl 없음")
        return
    rows = [json.loads(l) for l in SCORE.read_text(encoding="utf-8").splitlines()]
    print(f"{'건':<28} {'날짜':<10} {'A':>2} {'B':>2} {'C':>2} {'D':>2}  수정건수  섹션수")
    for r in rows:
        g = r.get("gaps_at_render", {})
        print(f"{r['case']:<28} {r['date']:<10} {g.get('A', '-'):>2} {g.get('B', '-'):>2} "
              f"{g.get('C', '-'):>2} {g.get('D', '-'):>2}  {r['hunks']:>6}  {r['sections_changed']:>4}")
    if len(rows) >= 2:
        h = [r["hunks"] for r in rows]
        print(f"\n평균 수정건수: 전체 {sum(h)/len(h):.1f}  |  최근5건 {sum(h[-5:])/len(h[-5:]):.1f}  |  이전5건 {sum(h[:-5])/max(len(h[:-5]),1):.1f}"
              if len(h) > 5 else f"\n평균 수정건수: {sum(h)/len(h):.1f} ({len(h)}건)")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("record")
    p1.add_argument("case_dir")
    p1.add_argument("final")
    p1.set_defaults(func=cmd_record)
    p2 = sub.add_parser("candidates")
    p2.set_defaults(func=cmd_candidates)
    p3 = sub.add_parser("close")
    p3.add_argument("id")
    p3.add_argument("status", choices=["approved", "rejected"])
    p3.add_argument("note", nargs="?", default="")
    p3.set_defaults(func=cmd_close)
    p4 = sub.add_parser("score")
    p4.set_defaults(func=cmd_score)
    a = ap.parse_args()
    a.func(a)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
