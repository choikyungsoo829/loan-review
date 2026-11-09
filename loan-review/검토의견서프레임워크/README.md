# 심사담당자 검토의견서 자동작성 프레임워크

여신자료를 올리면 **작성 전에 질문**하고, 결재본과 같은 양식의 `.docx`로 작성한다.
Claude Code에서 자료를 첨부하고 `/loan-review-1`이라고 입력하거나 "검토의견서 써줘"라고 하면 된다.

## 흐름

```
자료 첨부 ─▶ ① 재료 목록 ─▶ ② case.json 초안 ─▶ ③ calc.py --gaps ─▶ 질문(A필수·B충돌·C보완·D판단)
                                                        ▲                     │
                                                        └──── 답변 반영 ◀─────┘
                                   READY ─▶ ④ render_docx.py ─▶ to_pdf.py 검수 ─▶ 보고(미확보 목록 포함)
```

## 폴더

| 경로 | 내용 |
|---|---|
| `rules/01_작성규칙.md` | 섹션별 규칙 v2 (사용자 프롬프트 v1 + 결재본 대조로 고친 부분 【v2】) |
| `rules/02_규정조항_라이브러리.md` | 트리거 → 적법성·타당성·승인조건 문구 |
| `rules/03_질문목록.md` | 작성 전 질문 프로토콜과 자료별 필요 칸 |
| `rules/05_샘플오류노트.md` | 결재본에서 발견한 오류 (모방 금지) |
| `scripts/calc.py` | 계산·트리거·검산·누락탐지 (숫자는 여기서만) |
| `scripts/render_docx.py` | case.json → docx. A등급 누락 시 작성 거부(`--draft`로 초안) |
| `scripts/to_pdf.py` | docx → pdf (MS Word) 검수용 |
| `tests/regression.py` | 결재본 5건 수치 재현 테스트 — 규칙 수정 후 필수 |
| `cases/_template/case.json` | 빈 입력 틀 |
| `cases/예시_업체J/` | 결재본을 그대로 옮긴 완성 예시 |
| `cases/예시_업체O_사전검토만/` | 사전검토의견서만으로 채운 초안 → 질문 기능 시연 |

## 수동 실행

```
python scripts/calc.py cases/<건>/case.json --gaps
python scripts/render_docx.py cases/<건>/case.json
python scripts/to_pdf.py "cases/<건>/심사담당자 검토의견서(<업체>).docx"
python tests/regression.py
```

## 검산 ID

| ID | 내용 | 등급 |
|---|---|---|
| C01 | 신청금액 > 대출가능금액 | B |
| C02 | 총여신비율 > 100% | B |
| C03 | 담보물 표 평가액 합 ≠ 감정가 | B |
| C04 | 감정일 불일치 | B |
| C05 | 감정기관 불일치 | B |
| C06 | 신용등급 표기 ≠ 원자료 | B |
| C07 | 금융거래 명세 합계 ≠ 사전검토 합계 | B |
| C08 | 시설자금 > 매매대금 80% | B |
| C09 | 금리 ≠ 기준 + 가산 | B |
| C10 | 작성일 < 사전검토일 | B |
| C11 | 입지현황 지명 ≠ 소재지 (복붙 의심) | B |
| C12 | 서울 자치구인데 '경기' 표기 | B |
| C13 | '비거치' 문구 vs 거치식 상환 | B |
| C14 | 표B 기대출 ≠ 동일담보 대출 합계 | B |
| C15 | 동일인 5,000백만원 초과 | A |
| C16 | 금리 미기재 대출 → 금융비용 누락 | B |
| C17 | T2인데 결산 매출·차입금 없음 | A |
| C18 | 운전자금 순여신 > 매출−차입금 | B |
| C19 | 시설자금인데 매매금액 없음 | A |
| C20 | 자료 조항이 계산상 요건 미해당 | B |

## 학습 로그 (건이 쌓일수록 최적화)

결재 최종본을 주면 초안과 비교해 쌓는다. 규칙은 **자동으로 바뀌지 않는다** — 후보만 올라오고, 사람이 승인해야 반영된다.

```
python scripts/learn.py record <case_dir> <결재최종본.docx|.pdf>   # 초안↔최종본 비교 기록
python scripts/learn.py candidates                                # 2건 이상 반복된 수정 패턴 → candidates.md
python scripts/learn.py close <id> approved|rejected [메모]        # 검토 완료 → 재제안 방지
python scripts/learn.py score                                     # 건별 질문 수·수정 건수 추이
```

| 경로 | 내용 |
|---|---|
| `learning/diffs.jsonl` | 건별 수정 내역(섹션·삭제·추가) 원본 로그 |
| `learning/scorecard.jsonl` | 건별 작성시점 질문 수(A/B/C/D)·최종 수정 건수 |
| `learning/candidates.md` | 반영 검토 대기 중인 규칙 후보 |
| `learning/closed.md` | 승인/반려로 닫힌 후보(재제안 방지) |

## 한계 (현재 버전)

- `.hwp`는 직접 읽지 못한다 → PDF로 저장해서 올린다.
- 출력은 `.docx`다. 한글에서 열어 `.hwp`로 저장할 수 있다. 원본 hwp 서식과 픽셀 단위로 같지는 않다.
- 조항은 받은 자료 기준으로 쓴다. 라이브러리는 누락 탐지용이다.
