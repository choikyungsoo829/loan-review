# -*- coding: utf-8 -*-
"""심사평가표 학습 로그 — 검토의견서프레임워크 learn.py를 그대로 쓰되 저장 위치만 이 프레임워크로 돌린다.

  python learn.py record <case_dir> <결재최종본.docx|.pdf>
  python learn.py candidates
  python learn.py close <id> approved|rejected [메모]
  python learn.py score

규칙은 자동으로 바뀌지 않는다. candidates.md를 사람이 보고 승인해야 rules/*.md를 고친다.
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from sheetcalc import LOAN_FW  # noqa: E402

sys.path.insert(0, str(LOAN_FW))
import learn as base  # noqa: E402

ROOT = HERE.parent
base.LEARN = ROOT / "learning"
base.DIFFS = base.LEARN / "diffs.jsonl"
base.SCORE = base.LEARN / "scorecard.jsonl"
base.CANDIDATES = base.LEARN / "candidates.md"
base.CLOSED = base.LEARN / "closed.md"
base.find_draft = lambda d: _find(Path(d))


def _find(case_dir):
    docs = sorted(case_dir.glob("심사평가표(*).docx"))
    picked = [x for x in docs if "_초안" not in x.stem] or docs
    if not picked:
        raise SystemExit(f"심사평가표 .docx 없음: {case_dir}")
    return picked[-1]


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    base.LEARN.mkdir(exist_ok=True)
    base.main()
