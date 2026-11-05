# JSON 형식 - inspect·apply·결과 봉투

LLM 에이전트나 스크립트가 엑셀 없이 DSD를 읽고 고칠 때 쓰는 형식이다.
읽기는 `inspect`, 쓰기는 `apply`, 명령 결과는 공통 봉투(`--json`)로 주고받는다.

## 작업 순서

1. `dsdxlsx inspect 보고서.dsd --section 4`로 고칠 구획을 읽는다(마크다운이 짧아 LLM 문맥에 맞다. 정확한 값이 필요하면 `--json`).
2. 편집 목록을 JSON으로 쓴다.
3. `dsdxlsx apply 보고서.dsd edits.json --dry-run --json`으로 바뀔 내용을 확인한다.
4. `dsdxlsx apply 보고서.dsd edits.json --json`으로 반영한다(`보고서_수정.dsd`).
5. `dsdxlsx check 보고서_수정.dsd`로 검산하고, 불일치가 있으면 1로 돌아간다.

행·문단·표를 넣거나 지우면 ID가 다시 매겨진다. 그 뒤 편집은 새 파일을 다시 `inspect`해서 쓴다.

## inspect (`dsdxlsx.inspect/1`)

엑셀로 내보내는 구획(재무제표별, 주석 번호별)만 담는다. `apply`가 고칠 수 있는 범위와 같다.

```json
{"schema": "dsdxlsx.inspect/1", "file": "보고서.dsd", "sha256": "…",
 "sections": [
  {"section": "4", "kind": "note", "title": "현금및현금성자산",
   "blocks": [
    {"type": "para", "id": "p003", "text": "4. 현금및현금성자산\n…"},
    {"type": "table", "id": "t003", "ncols": 3, "unit": "백만원",
     "columns": [{"col": 0, "header": "구  분", "period": null},
                 {"col": 1, "header": "당기말", "period": "당기"},
                 {"col": 2, "header": "전기말", "period": "전기"}],
     "rows": [
      {"id": "t003.r002", "label": "구  분", "head": true, "cells": […]},
      {"id": "t003.r004", "label": "보통예금",
       "cells": [{"id": "t003.r004.c01", "col": 0, "text": "보통예금"},
                 {"id": "t003.r004.c02", "col": 1, "text": "1,100", "value": 1100},
                 {"id": "t003.r004.c03", "col": 2, "text": "750", "value": 750}]},
      {"id": "t003.r005", "label": "합  계", "total": true, "cells": […]}]}]}]}
```

| 필드 | 뜻 |
|---|---|
| `section` | 구획 키. 재무제표는 명칭(`재무상태표` 등), 주석은 번호 문자열(`"4"`), 주석 머리는 `"주석"`. `apply`의 `section`에 그대로 쓴다 |
| `kind` | `statement`, `note`, `notes_intro` |
| 블록 `type` | `para`(고칠 수 있는 문단), `label`(제목, 고칠 수 없음), `table` |
| `columns[].col` | 격자 열 번호(0부터). 병합 칸을 펼친 기준이다 |
| `columns[].period` | 머리글로 판정한 `당기`·`전기`, 모르면 `null` |
| `rows[].label` | 행 이름(첫 글자 칸). 앞 공백은 계층 들여쓰기다 |
| `rows[].head` / `total` | 머리 행 / 합계류 행 |
| `cells[].id` | 칸 ID. 끝의 `c02`는 행 안 칸 순번(1부터)이라 `col`과 다르다 |
| `cells[].value` | 숫자로 읽히는 칸의 값. 괄호는 음수, `-`는 0 |
| `cells[].readonly` | 자동 필드(TU). 고칠 수 없다 |
| `notes` | 표 잠금 사유(고정 표, 불규칙 격자, 자동합계) |

마크다운(`--format` 없이 기본)은 같은 내용을 표로 보여 준다. 열 머리글 앞 `[1]`이 `col` 번호다.

## apply 편집 목록

JSON 배열이나 `{"edits": [...]}`. 위에서부터 차례로 적용하고, 하나라도 실패하면 아무것도 쓰지 않는다.

### 칸·행 가리키기

| 대상 | 쓰는 법 |
|---|---|
| 칸 | `"cell": "t003.r004.c02"` 또는 `"row"` + `"col"` |
| 행 | 행 ID `"t003.r004"` 또는 `{"section": "4", "row": "보통예금"}`(+ `"row_nth": 2`, `"table": "t003"`) |
| 열 | 번호 `1`, 머리글 `"당기말"`(공백 무시), `"당기"`·`"전기"` |

`set`은 `section`을 바깥에 둘 수도 있다: `{"op": "set", "section": "4", "row": "보통예금", "col": "당기", "value": 950}`.
행 이름이 여러 행에 걸리거나 같은 기간 열이 여럿이면 고르지 않고 후보와 함께 오류를 낸다.

### 편집 종류

| op | 필드 | 하는 일 |
|---|---|---|
| `set` | 칸, `value` | 칸 값을 바꾼다. `null`이나 `""`는 비움 |
| `set_text` | `id`(p###), `text` | 문단 글을 바꾼다 |
| `insert_row` | `after`(행), `values` | 행 바로 아래에 새 행. `values`는 열 순서 배열이나 `{열: 값}` |
| `delete_row` | `row` | 행을 지운다 |
| `insert_para` | `after`(p### 또는 t###), `text` | 블록 바로 아래에 새 문단(주석 구획만) |
| `delete_para` | `id` | 문단을 지운다(주석 구획만) |
| `insert_table` | `after`, `rows`, `header_rows`(기본 1), `template`(t###) | 같은 열 수의 기존 표를 본떠 새 표(주석 구획만) |
| `delete_table` | `id` | 표를 지운다(주석 구획만) |

숫자는 JSON 숫자로 준다. 괄호 음수, 소수 자리, '주' 같은 단위는 원래 칸의 표기를 따르고, 0은 `-`로 쓴다.
글자로 준 값은 그대로 들어간다.

### 막는 것

`import`와 같은 안전장치를 쓴다. 막히면 종료코드 2와 이유를 낸다.

- 고정 표(AFIXTABLE), 불규칙 격자, 세로 병합에 걸린 행, 머리 행의 행 추가·삭제
- 자동 필드(TU) 수정, 행 재배열
- 재무제표 본문 표의 신설·삭제, 열 수가 다른 표 신설
- 새 주석 추가·삭제(아직 `import`의 `#새주석`·`#주석삭제`로만)

## 결과 봉투 (`dsdxlsx.result/1`)

`--json`을 주면 결과를 표준출력에 이 형식으로 낸다.

```json
{"schema": "dsdxlsx.result/1", "command": "apply", "version": "0.2.0",
 "ok": true, "exit_code": 0, "outputs": ["보고서_수정.dsd"],
 "warnings": [], "errors": [],
 "dry_run": false, "changed": true, "structural": false,
 "changes": {"cell_edits": ["t003.r004.c02"], "row_inserts": ["t003 r005 위치에 신규 행 [구획 4 새 줄]"]}}
```

| 종료코드 | 뜻 |
|---|---|
| 0 | 정상 |
| 1 | 정상 실행했고 확인할 발견 사항이 있음 |
| 2 | 입력 오류(파일, JSON, 위치 지정, 안전장치). 아무것도 쓰지 않음 |
| 3 | 안전 장치로 멈춤(출력 파일이 이미 있음, 입력 파일 자리에 쓰기). 아무것도 쓰지 않음 |

`changes`의 키: `cell_edits`, `row_inserts`, `row_deletes`, `table_inserts`, `table_deletes`, `para_inserts`, `para_deletes`, `autosum_tables`.

### dart 결과

```json
{"schema": "dsdxlsx.result/1", "command": "dart", "ok": true, "exit_code": 0,
 "outputs": ["샘플전자_감사보고서_2025.dsd", "샘플전자_감사보고서_2025.xlsx"],
 "warnings": ["이미지 1개는 공시 원문에 파일이 없어 자리만 남았다 - …"], "errors": [],
 "company": "샘플전자 주식회사", "corp_code": "00999001", "year": 2025,
 "source": {"rcept_no": "20260310000001", "rcept_dt": "20260310", "report_nm": "감사보고서 (2025.12)",
            "period": "2025.12", "kind": "F", "url": "https://dart.fss.or.kr/dsaf001/main.do?rcpNo=…",
            "file": "20260310000001.xml", "acode": "00760", "document": "감사보고서"},
 "fixes": {"crlf": 7618}, "images": ["….jpg"],
 "rollforward": {"자동확정": 49, "확인필요": 34, "미처리": 6, "painted": 355}}
```

`source.kind`는 `F`(외부감사관련 공시) 또는 `A`(사업보고서 첨부)다. 접수번호로 받으면 공시 정보가 없어 `rcept_no`가 `null`이다.
`rollforward`는 `--rollforward`를 줬을 때만 있다. `--list --json`은 `corp`(고유번호·회사명·종목코드)와 `filings`(공시 목록)를 낸다.
키가 없거나 회사·공시를 못 찾으면 종료코드 2, 출력 파일이 이미 있으면 3이다.
`structural`이 `true`면 ID가 다시 매겨졌으니 다음 편집 전에 `inspect`를 다시 한다.
