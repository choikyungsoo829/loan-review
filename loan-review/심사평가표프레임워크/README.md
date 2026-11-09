# 심사평가표 / 심사의견요약표 자동작성 프레임워크

여신자료를 올리면 **작성 전에 질문**하고, 결재본과 같은 **2쪽 양식**(1쪽 심사평가표 / 2쪽 심사의견요약표) `.docx`로 뽑는다.
Claude Code에서 자료를 첨부하고 `/loan-review-2` 또는 "심사평가표 써줘"라고 하면 된다.

> 검토의견서(대심위용 '심사담당자 검토의견서')는 다른 프레임워크다 → `../검토의견서프레임워크` / `/loan-review-1`

## 흐름

```
자료 첨부 ─▶ sheet.json 초안 ─▶ sheetcalc.py --gaps ─▶ 질문(A필수·B충돌·C보완·D판단)
                                      ▲                    │
                                      └──── 답변 반영 ◀────┘
                      READY ─▶ build.py ─▶ to_pdf.py 검수 ─▶ 결재 후 learn.py record
```

## 재사용 구조 (산식 이중관리 금지)

| 이 프레임워크 | 가져다 쓰는 것 |
|---|---|
| `scripts/sheetcalc.py` | 검토의견서 `calc.py`의 대출가능금액·이자수익·편중·검산 |
| `scripts/build.py` | 검토의견서 `render_docx.py`의 표·셀·테두리 헬퍼 |
| `scripts/learn.py` | 검토의견서 `learn.py`의 학습 루프(저장 위치만 이쪽으로) |
| 검수 | 검토의견서 `to_pdf.py` 그대로 |

## 폴더

| 경로 | 내용 |
|---|---|
| `rules/01_작성규칙.md` | 1쪽·2쪽 항목별 작성 규칙 |
| `rules/02_문구_라이브러리.md` | 표준 평가값·조항·승인조건·금리 우대 명목 |
| `rules/03_질문목록.md` | 작성 전 질문 프로토콜 |
| `rules/04_샘플오류노트.md` | 결재본 오타 이력 (따라 쓰지 말 것) |
| `rules/05_유형별_사례집.md` | 특수 건 규정 조합(유사 건 검토용 후보) |
| `scripts/sheetcalc.py` | 계산·검산·누락탐지 (숫자는 여기서만) |
| `scripts/build.py` | sheet.json → 2쪽 docx. A등급 누락 시 작성 거부(`--draft`로 초안) |
| `scripts/learn.py` | 결재본 대비 수정 이력 적재 → 규칙 후보 |
| `cases/_template/sheet.json` | 빈 입력 틀 |
| `cases/예시_차주14/` | 운전자금·기대출 2건 |
| `cases/예시_차주13/` | 시설자금·신규고객·소요자금한도 |
| `cases/예시_차주15/` | 타행대환·RTI |
| `cases/예시_차주11/` | 자립예탁금·기대출 3건 |
| `cases/예시_차주12/` | 공장담보·상임이사우대 |
| `cases/261014_차주10/` | 주택임대업·임차보증금 반환·RTI 추정임대소득 전결권 상향 (결재본 재현) |

결재본 6건 모두 재현 확인 완료(차주10은 문구까지).

## 수동 실행

```
python scripts/sheetcalc.py cases/<건>/sheet.json --gaps
python scripts/build.py cases/<건>/sheet.json
python "../검토의견서프레임워크/scripts/to_pdf.py" "cases/<건>/심사평가표(<채무자>).docx"
python scripts/learn.py record cases/<건> <결재최종본.pdf>
```

## 검산 ID

`S01` 금리 산출근거 ≠ 기재 금리 / `S02` 계산 가산 ≠ 기재 가산 / `S03` 신청금액 > 대출가능금액 /
`S04` 시설자금 > 매매금액 80% + 검토의견서 `calc.py`의 C01~C20을 그대로 물려받는다.

## 한계

- 결재본이 스캔 이미지면 learn.py가 텍스트 비교를 못 한다 → PDF/docx 필요.
- 종합의견 분량이 많으면 2쪽을 넘긴다. 넘치면 글자 크기(`build.py`의 `fill(size=8)`)를 조정한다.
- 낙찰통계는 넣지 않는다(사용자가 자료 가공 시 따로 첨부).
- 기대출 6건 이상이면 자금용도별로 묶어 합산금액·최대금리로 표시한다.
