---
name: dsd-review
description: DART DSD/정산표(xlsx)의 AI 검토 — 오타·띄어쓰기, 주기(주석 참조) 적절성, 페이지 분할, 대사쌍 제안 정제. 사용자가 "/dsd-review", "DSD 검토해줘", "오타·주기 검사"류를 요청할 때 사용. dsdxlsx check --json 산출(review.json)을 읽고 findings.json을 만들어 annotate-ai로 워크북 'AI검토' 시트에 기록한다.
---

# DSD AI 검토 (/dsd-review)

룰베이스 검증(dsdxlsx check C1~C8)이 **못 하는** 판단형 검토를 수행한다.
도구가 결정론 데이터(review.json)를 내놓고, 이 스킬(AI)이 소견을 만들며,
기록은 다시 도구(annotate-ai)가 한다. **모든 소견은 advisory** — 본문 수치·텍스트를
직접 수정하지 않는다. 반영은 사람이 셀을 고친 뒤 check 재실행으로 검산.

## 전제

- 도구 위치(`DSDXLSX_HOME`, 환경이 다르면 조정):
  `C:\Users\PC\Desktop\02. Projects\01. Playground\01. 실무 AX 프로젝트\04. 업무도구\01. DSD-XLSX 변환 자동화`
- 대상은 **dsdxlsx export로 만든 .xlsx** 권장(시트 기록 가능). `.dsd`가 오면 먼저
  `export`를 실행하거나, 기록 없이 콘솔 보고만 할지 사용자에게 확인.
- 고객 실데이터: 산출물(review.json·findings.json)은 스크래치패드에만 두고 repo에 넣지 않는다.

## 절차

1. **검증 + 리뷰 JSON 생성** (파일은 엑셀에서 닫혀 있어야 함):
   ```powershell
   uv run --project "<DSDXLSX_HOME>" --no-editable dsdxlsx check "<파일.xlsx>" --json "<스크래치>\review.json"
   ```
2. **review.json 읽기** — 주요 필드:
   - `document.paragraphs[]` `{id, section, text}` — 문단·헤딩 전문
   - `document.tables[]` `{id, section, headers, rows:[{id,label}], height, unit}`
   - `document.fs_rows[]` `{section, row_id, label, refs:[주석번호], amounts:{당기,전기}, unit}`
   - `document.notes[]` `{no, title}` / `document.pgbrks[]` `{offset, after_block}`
   - `ties.suggested[]` — 대사쌍 제안(양쪽 위치·금액 포함), `ties.confirmed[]` — 판정
   - `checks.findings[]` — 룰베이스 결과(C4 '참조번호 오류 의심' 등), `coverage[]`
3. **체크리스트 4종 수행** (아래) → findings 작성.
4. **findings.json 저장** (스크래치패드) 후 기록:
   ```powershell
   uv run --project "<DSDXLSX_HOME>" --no-editable dsdxlsx annotate-ai "<파일.xlsx>" "<스크래치>\findings.json"
   ```
5. **사용자 보고**: warn 우선 순서로 위치·내용·근거 요약. 소견 0건이면 0건이라고 명확히.

## findings.json 계약

```json
{"schema": 1, "findings": [
  {"kind": "오타|주기|페이지|대사|기타",
   "severity": "warn|info",
   "row_id": "t014.r004",        // 선택 — 표 행이면 하이퍼링크 생성됨
   "section": "주석4",            // 선택 — 위치 설명 (review.json의 section 라벨)
   "message": "발견 내용 (필수, 근거 포함)",
   "suggestion": "수정 제안 (선택)"}
]}
```

severity 기준: **warn** = 공시 오류 가능성(잘못된 주석 참조, 명백한 오탈자, 전기 문안
잔재) / **info** = 개선 제안·확인 권고. 확신이 없으면 info.

## 체크리스트

### A. 오타·띄어쓰기 (kind: 오타)
`paragraphs[].text` + `tables[].rows[].label` + `tables[].headers` 대상:
- 명백한 한글 오탈자·조사 오류·중복 단어("당기당기")·미완성 문장.
- **기수·날짜 잔재** (가장 가치 큼): 문단 속 "제 N 기"·연도가 문서의 당기와 안 맞으면
  전기 문안 복사 잔재 의심 → warn. (당기 기수는 fs_rows/헤더의 "제 N(당) 기"에서 파악)
- 계정과목 표기 비일관: 같은 계정이 문서 내 다른 표기("현금및현금성자산" vs
  "현금 및 현금성자산")로 혼용되면 info로 묶어 지적.
- 오인 금지: 자간 벌림 제목("재 무 상 태 표")은 관례. K-IFRS/세법 전문용어를 오타로
  단정하지 않는다(확신 없으면 info + "확인 요망").

### B. 주기(주석 참조) 적절성 (kind: 주기)
`fs_rows[].refs` ↔ `notes[]` 대조 (참조 열이 없는 양식이면 이 절 생략):
- `checks.findings`의 C4 "참조번호 오류 의심"을 우선 인용하고 원인 추정을 보강.
- 계정↔주석 상식 매핑으로 **빠진 참조** 후보: 재고자산→재고자산 주석, 유형자산→유형자산
  (+담보제공 있으면 그 주석), 차입금→차입금·담보·약정, 매출채권→매출채권·신용위험,
  퇴직급여부채→퇴직급여, 법인세→법인세비용·이연법인세, 자본금→자본금·주당이익,
  특수관계자 채권채무→특수관계자 주석. 주석 제목(notes[].title)과 대조해 판단.
- 존재하지 않는 주석 번호 참조(refs에 있는데 notes에 없음) → warn.
- 회계기준서가 요구하는 주석의 존재 여부까지는 판단하지 않는다(범위 밖).

### C. 페이지 분할 (kind: 페이지, 전부 info)
`tables[].height`(TR HEIGHT 합, pt 근사)와 `pgbrks[]`:
- 높이 ~650pt 초과 표는 한 페이지를 넘길 가능성 — 표 직전에 PGBRK가 있으면 "제목/서두만
  앞 페이지에 남을 수 있음", 없으면 "표가 중간에서 잘릴 수 있음" 후보로.
- 헤딩 문단(주석 제목) 직후 PGBRK → 제목 고아(orphan) 패턴 지적.
- 반드시 "DART 편집기 미리보기로 확인 요망" 문구를 suggestion에 포함 — 높이는 근사치다.

### D. 대사 제안 정제 (kind: 대사)
`ties.suggested[]` 대상 — 확정은 사람이 시트에서 하고, 여기서는 **추천/보류 판단**만:
- 라벨 이형이 같은 계정인지 회계 지식으로 판단: "매출채권" ↔ "매출채권및기타채권",
  "차감후 장부금액" ↔ FS 순액 표시 등 → "확정 추천" + 근거.
- 우연 일치 의심: 자릿수가 작거나(수만 원 단위) 계정 성격이 무관하면 "보류 권고" + 사유.
- confidence: low 항목을 우선 검토. 이미 confirmed에 있는 쌍과 중복 판단하지 않는다.

## 원칙

- 소견은 위치(row_id/section)를 반드시 붙인다 — review.json의 ID를 그대로 사용.
- 같은 원인의 반복 소견은 하나로 묶고 대표 위치 + "외 N곳"으로.
- 검토했으나 소견이 없는 체크리스트도 보고에 "이상 없음"으로 명시(침묵 금지).
- findings.json 기록 후 워크북의 'AI검토' 시트 위치를 사용자에게 안내한다.
