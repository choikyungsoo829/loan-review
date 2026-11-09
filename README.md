# 론리뷰 (loan-review) — 여신 심사서류 자동작성 프레임워크

Claude Code 스킬 2종과 그 프레임워크(규칙·양식·계산·렌더·학습 루프).

👉 **처음이면 [사용설명서.md](사용설명서.md)부터 읽는다.**

| 스킬 | 산출물 | 프레임워크 |
|---|---|---|
| `loan-review-1` | 심사담당자 검토의견서 (대심위 부의용, 4쪽 .docx/.pdf) | `loan-review/검토의견서프레임워크` |
| `loan-review-2` | 심사평가표·심사의견요약표 (2쪽 .docx) | `loan-review/심사평가표프레임워크` |

## 구성

```
skills/loan-review-1/SKILL.md      스킬 정의 (작성 절차)
skills/loan-review-2/SKILL.md
loan-review/<프레임워크>/
  rules/        작성규칙·규정조항 라이브러리·질문목록·오류노트·문구은행·사례집
  scripts/      calc(검산·한도·트리거) / render_docx·build(양식 렌더) / learn(결재본 학습) / to_pdf
  cases/_template/   입력 양식(case.json, sheet.json)
  learning/     학습 루프 골격 (diffs.jsonl·scorecard.jsonl·candidates.md·closed.md·style_notes.md)
  tests/        회귀 테스트 (python tests/regression.py → ALL PASS)
```

## 설치 (새 PC)

1. `skills/*` → `%USERPROFILE%\.claude\skills\` 로 복사
2. `loan-review/` → `%USERPROFILE%\.claude\loan-review\` 로 복사 (두 프레임워크는 **같은 폴더에 나란히** — 심사평가표가 형제 폴더의 calc를 쓴다)
3. 각 SKILL.md의 `프레임워크 루트:`·`산출물 폴더:` 줄을 새 PC 경로로 고친다
4. 의존성: `pip install python-docx` (PDF 변환은 MS Word 필요)

## 공개본에서 뺀 것

- **실제 케이스(cases/)와 학습 기록(learning/ 데이터)은 고객 신용정보라 제외**했다. `_template`만 남김.
- 규칙·스크립트 주석·회귀 테스트에 있던 실명·상호는 가명(`차주01`, `업체A` 등)으로 바꿨다. 문서 안의 `cases/예시_…` 경로는 공개본에 없다.
- `.gitignore`가 새 케이스와 학습 데이터를 커밋에서 막는다.
