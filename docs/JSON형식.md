# JSON 형식 - inspect·apply·결과 봉투

LLM 에이전트나 스크립트가 엑셀 없이 DSD를 읽고 고칠 때 쓰는 형식이다.
읽기는 `inspect`, 쓰기는 `apply`, 명령 결과는 공통 봉투(`--json`)로 주고받는다.

## 작업 순서

1. `dsdxlsx inspect 보고서.dsd --outline`으로 구획과 표 목록을 본다. 찾는 글이 있으면 `--find 글`로 위치(ID)를 바로 얻는다.
2. `dsdxlsx inspect 보고서.dsd --section 4`로 고칠 구획을 읽는다(마크다운이 짧아 LLM 문맥에 맞다. 정확한 값이 필요하면 `--json`).
3. 편집 목록을 JSON으로 쓴다.
4. `dsdxlsx apply 보고서.dsd edits.json --dry-run --json`으로 바뀔 내용을 확인한다.
5. `dsdxlsx apply 보고서.dsd edits.json --json`으로 반영한다(`보고서_수정.dsd`).
6. `dsdxlsx check 보고서_수정.dsd --json`으로 검산한다. 종료코드가 1이면 `mismatches`의 행 ID와 열 번호로 2번부터 다시 한다.

행·문단·표·주석을 넣거나 지우면 ID가 다시 매겨진다. 그 뒤 편집은 새 파일을 다시 `inspect`해서 쓴다.

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

마크다운(`--json` 없이 기본)은 같은 내용을 표로 보여 준다. 열 머리글 앞 `[1]`이 `col` 번호다.

### 목차 (`--outline`, `dsdxlsx.outline/1`)

칸 내용 없이 구획마다 블록 목록만 준다. 주석이 수십 개인 문서에서 어디를 읽을지 먼저 고를 때 쓴다.

```json
{"schema": "dsdxlsx.outline/1", "file": "보고서.dsd", "sha256": "…",
 "sections": [
  {"section": "4", "kind": "note", "title": "현금및현금성자산", "tables": 1, "paras": 1,
   "blocks": [
    {"type": "para", "id": "p003", "chars": 38, "text": "4. 현금및현금성자산 현금및현금성자산의 내역은…"},
    {"type": "table", "id": "t003", "ncols": 3, "nrows": 5, "unit": "백만원",
     "headers": ["구  분", "당기말", "전기말"], "labels": ["보유현금", "보통예금", "합  계"]}]}]}
```

`labels`는 본문 행 이름 앞 4개다(표를 알아보는 용도).

### 찾기 (`--find 글`, `dsdxlsx.find/1`)

글이 든 문단·행·칸의 위치만 준다. 공백과 대소문자는 무시한다. 숫자를 주면 값으로도 찾는다(`1200`으로 `1,200` 칸을 찾음).
`--find`를 여러 번 주면 어느 하나라도 맞는 곳을 모으고, `--section`으로 범위를 줄일 수 있다.

```json
{"schema": "dsdxlsx.find/1", "file": "보고서.dsd", "queries": ["보통예금"], "total": 1, "truncated": false,
 "hits": [{"section": "4", "type": "row", "id": "t003.r004", "table": "t003", "query": "보통예금", "label": "보통예금"}]}
```

| `type` | 뜻 | 위치 필드 |
|---|---|---|
| `row` | 행 이름이 맞음 | `id`(행 ID), `table`, `label` |
| `cell` | 그 밖의 칸이 맞음(머리 행이면 `head: true`) | `id`(칸 ID), `row`, `table`, `col`, `header`, `text` |
| `para`, `label` | 문단, 제목 | `id`, `text`(맞은 곳 앞뒤) |

결과가 `--limit`(기본 200)을 넘으면 `truncated`가 `true`다.

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
| `insert_note` | `no`, `title`, `blocks` | 새 주석. `blocks`의 글자는 문단, `{template, header_rows, rows}` 객체는 표(본뜰 표 `template` 필수) |
| `delete_note` | `no` | 주석을 통째로 지운다 |

숫자는 JSON 숫자로 준다. 괄호 음수, 소수 자리, '주' 같은 단위는 원래 칸의 표기를 따르고, 0은 `-`로 쓴다.
글자로 준 값은 그대로 들어간다.

주석 추가와 삭제:

- `insert_note`의 `no`가 마지막 번호 + 1이면 맨 끝에 붙는다. 이미 있는 번호면 그 앞에 끼우고 뒤 주석의 제목 번호와 재무제표의 주석 참조 칸을 하나씩 민다. 본문 글에서 다른 주석을 가리키는 번호는 고치지 않고 `warnings`에 적는다.
- `delete_note`는 번호를 당기지 않는다. 지운 번호를 아직 가리키는 곳은 `warnings`에 나오고, 번호 정리는 `dsdxlsx renumber 파일.dsd --close-gap`으로 한다.
- 한 번의 `apply`에 `insert_note`는 하나만 쓸 수 있고 `delete_note`와 함께 쓸 수 없다(차례로 나눠 실행).

### 막는 것

`import`와 같은 안전장치를 쓴다. 막히면 종료코드 2와 이유를 낸다.

- 고정 표(AFIXTABLE), 불규칙 격자, 세로 병합에 걸린 행, 머리 행의 행 추가·삭제
- 자동 필드(TU) 수정, 행 재배열
- 재무제표 본문 표의 신설·삭제, 열 수가 다른 표 신설
- 번호를 건너뛰는 새 주석

## 결과 봉투 (`dsdxlsx.result/1`)

모든 명령은 `--json`을 주면 사람용 글 대신 이 형식 하나를 표준출력에 낸다(`inspect`의 `--json`은 문서 내용을 JSON으로 낸다는 뜻이라 예외).

```json
{"schema": "dsdxlsx.result/1", "command": "apply", "version": "0.6.1",
 "ok": true, "exit_code": 0, "outputs": ["보고서_수정.dsd"],
 "warnings": [], "errors": [],
 "dry_run": false, "changed": true, "structural": false,
 "changes": {"cell_edits": ["t003.r004.c02"], "row_inserts": ["t003 r005 위치에 신규 행 [구획 4 새 줄]"]}}
```

| 공통 필드 | 뜻 |
|---|---|
| `ok` | 명령이 끝까지 돌았는가(종료코드 0이나 1) |
| `exit_code` | 프로세스 종료코드와 같다 |
| `outputs` | 만들었거나 고친 파일 경로 |
| `warnings` / `errors` | 확인할 것 / 실패 이유 |
| `log` | 사람용 출력 줄(구조화한 필드가 충분한 `apply`, `dart`, `check`에는 없다) |

| 종료코드 | 뜻 |
|---|---|
| 0 | 정상 |
| 1 | 정상 실행했고 확인할 것이 있음(검산 불일치, 대사 확인 필요, 못 채운 칸 등) |
| 2 | 입력 오류(파일, JSON, 위치 지정, 안전장치). 아무것도 쓰지 않음 |
| 3 | 안전장치로 멈춤(사람이 고친 출력 파일, 입력 파일 자리에 쓰기). 아무것도 쓰지 않음 |

`check`, `recon`, `fill`은 `--json`을 줬을 때만 확인할 것이 있으면 1을 낸다. `--json` 없이 부르면 예전처럼 0이다(이 종료코드에 기대는 다른 도구가 있다).
`check`와 `recon`의 `--json`은 뒤에 `.json` 경로를 주면 예전 방식(그 파일에 리뷰용·요약 JSON 저장)으로 돈다.

### 덮어쓰기 규칙

사람이 이어서 고치는 결과물(`export`의 엑셀, `import`·`renumber`·`fill`의 DSD, `rollforward`·`roll-interim`의 DSD와 정산표)은 출력 파일이 이미 있을 때 이렇게 한다.

- dsdxlsx가 만든 그대로면 덮어쓴다.
- 만든 뒤에 바뀌었거나, 이 PC에서 dsdxlsx가 만든 기록이 없으면 멈춘다(종료코드 3). 다른 `-o`를 주거나 `--force`를 준다.
- 입력 파일과 같은 경로에는 `--force`를 줘도 쓰지 않는다.

`apply`와 `dart`는 더 엄격해서 출력 파일이 있기만 하면 멈춘다. 다시 만들 때마다 새로 쓰는 보고서(`check -o`, `recon`·`fill`의 결과 엑셀, `diff`의 HTML)는 그냥 덮어쓴다.
기록은 `%LOCALAPPDATA%\dsdxlsx\outputs.json`에 둔다(`DSDXLSX_STATE`로 폴더를 바꿀 수 있다).

### 명령별 필드

| 명령 | 더해지는 필드 |
|---|---|
| `apply` | `dry_run`, `changed`, `structural`, `changes` |
| `import` | `changed`, `structural`, `changes`. 풋팅 워크북이면 `kind: "foot"`, `source`, `changed`, `edits`(cell·node·old·new) |
| `export` | `sheets`, `tables`, `paras` |
| `check` | `counts`, `findings`, `mismatches`, `suppressed`, `tie_suggestions`, `coverage` |
| `foot` | `files`: 파일마다 `file`, `output`, `company`, `sheets`, `exit_code`, `footing`(total·pass·rounding·fail), `cross`(total·pass·fail), `mismatches`(kind·sheet·cell·label·diff). 실패한 파일은 `error`. `--dart-tie`면 `dart`(status·reason·total·pass·fail·sources·unmatched)가 붙고 `mismatches`에 `kind: "dart"`(current·dart 값 포함)가 더해진다. 건너뛴 사유는 `warnings`. `--review`면 `review`(path·items·note_sheets·texts·texts_truncated) |
| `recon` | `period`, `side`, `items`, `issues`, `counts`, `to_check`, `sheets`, `config_written` |
| `fill` | `period`, `side`, `counts`, `unfilled`, `map`(칸별 출처 기록 파일 경로 - `outputs`에도 들어 있다) |
| `build` | `side`, `steps`, `map`, `map_counts`(cells·overridden·missing·unfilled) |
| `tb` | `erp`(더존·세무사랑), `period`(cur·pri), `accounts`, `no_code`(계정코드를 못 찾아 코드 자리에 `*이름`을 둔 계정), `files`(읽은 파일), `rules`(쓴 규칙 파일, 안 썼으면 `null`). 경고가 있으면 종료코드 1, 자료가 모자라거나 부호합이 안 맞으면 2 |
| `link make` | `side`, `sheets`, `accounts`, `note_tables`, `to_check`(sheet·label·issue·prior), `unlinked_tb`(DSD 행에 짝이 없는 TB 과목), TB가 표일 때 `tb_columns`(읽은 당기·전기 열의 머리글). `to_check`가 있으면 종료코드 1 |
| `link test` | `side`, `same`, `diff`, `blank`, `note_same`, `note_diff`, `note_blank`, `steps`. `diff`·`blank` 항목은 `section`, `label`, `built`(결과 DSD 당기 칸), `expected`(뼈대 전기 칸). 하나라도 있으면 종료코드 1. `-o 결과.xlsx`를 주면 같은 내용을 엑셀 한 장(시트 `왕복시험` - 같음·다름·빈칸 칸 수와 어긋난 칸 목록)으로도 남기고 `outputs`에 싣는다 |
| `link refresh` | `accounts`, `added`(새로 붙인 계정 코드). 새 계정이 있으면 종료코드 1 |
| `status` | `in_sync`, `files`(dsd·skeleton·workbook·recipe마다 `file`, `state`: same·changed·missing, 기록이 없으면 `null`), `dsd_changes`(DSD가 바뀌었을 때 `link back`의 종류별 칸 수), `unmapped`(출처 기록에 없는 칸의 변경 수), `workbook_changes`(cells: 값이 달라진 출처 칸, no_values: 계산값이 없는 수식 칸), `next`(다음에 할 일), `made_by`. 맞으면 종료코드 0, 아니면 1 |
| `link back` | `changes`(아래), `counts`(종류별 수), `unchanged`, `applied`(정산표에 쓴 칸 수), `backup`(쓰기 전 사본), `unmapped`(출처 기록에 없는 칸의 변경: cell·label·old·new), `map`. 사람이 볼 것이 남으면 종료코드 1 |
| `rollforward`, `roll-interim` | `counts`(자동확정·확인필요·미처리), `painted`, `todo`. `roll-interim`은 `period`, `dart`도 |
| `renumber` | `notes`, `mapping`, `mentions`, `headings`, `ref_cells` |
| `refoot` | `applied`, `items`, `skipped` |
| `diff` | `total`, `by_kind` |
| `verify` | `same`, `differences` |
| `disc` | `form`, `listed`, `interim`, `counts` |
| `dart` | 아래 |

`changes`의 키: `cell_edits`, `row_inserts`, `row_deletes`, `table_inserts`, `table_deletes`, `para_inserts`, `para_deletes`, `note_inserts`, `note_deletes`, `renumbered`, `autosum_tables`.
`structural`이 `true`면 ID가 다시 매겨졌으니 다음 편집 전에 `inspect`를 다시 한다.

### check 결과

```json
{"schema": "dsdxlsx.result/1", "command": "check", "ok": true, "exit_code": 1,
 "counts": {"C1": 0, "C2": 2, "C3": 10, "C7": 0, "C4": 14, "C5": 0, "C8": 0, "C6": 86},
 "findings": [{"check": "C3", "level": "warn",
               "where": "[주석4 현금및현금성자산] t003 '합  계' × '당기말'",
               "message": "합계행 기재 1,200 ≠ 계산 1,300 (차이 (100))"}],
 "mismatches": [{"check": "C3", "section": "주석4 현금및현금성자산", "table": "t003",
                 "row": "t003.r005", "label": "합  계", "col": 1, "header": "당기말",
                 "stated": 1200, "computed": 1300, "note": ""}],
 "suppressed": 0, "tie_suggestions": 5,
 "coverage": {"tables": 6, "checked": 5, "unchecked": ["t001"]}}
```

`mismatches`의 `row`와 `col`은 `apply`의 `{"op": "set", "row": …, "col": …}`에 그대로 쓸 수 있다.
`stated`는 적힌 값, `computed`는 구성 칸으로 다시 계산한 값이다. 어느 쪽이 틀렸는지는 도구가 판단하지 않는다. 합계가 틀렸을 수도, 구성 칸이 틀렸을 수도 있으니 조서나 원자료로 확인한 뒤 고친다.
`findings`의 `level`이 `warn`이나 `error`인 것이 하나라도 있으면 종료코드가 1이다. `info`는 참고다.
`mismatches`에는 참고(`info`) 수준이거나 프로파일로 억제한 불일치도 들어 있어서, 종료코드가 0인데 `mismatches`가 비어 있지 않을 수 있다.

### 풋팅 워크북 검토 (`foot --review` → `annotate-ai`)

`foot --review`는 검토 재료(`dsdxlsx.foot-review/1`)를 파일로 낸다. 판정은 사람이나 에이전트가 하고, `annotate-ai`가 판정을 다시 검산해 워크북에 적는다.

검토 재료:

```json
{"schema": "dsdxlsx.foot-review/1", "workbook": "보고서_풋팅.xlsx", "source": "보고서.dsd",
 "source_sha256": "…", "company": "샘플전자(주)", "detection_profile": "standard",
 "guide": {"caution": "…", "verdicts": ["수정필요", "무시가능", "도구오탐", "확인필요"],
           "verdict": "판정 기준 …", "xref": "대사 제안 기준 …", "typo": "오탈자 기준 …"},
 "items": [{"id": "F1", "sheet": "2", "cell": "C14", "target": "2!C13", "kind": "total_row",
            "kind_desc": "합계류 라벨 행 = 위쪽 구성 행 합", "label": "당기말 합계",
            "relation": "C13-(C11+C12)  [기재 1,200 / 항 값: C11=200, C12=1,100]",
            "diff": "-100", "pretag": "금액 불일치", "stated": "1200", "expected": "1300",
            "context": "B11=보유현금  C11=200 …"}],
 "xref": {"statements": [{"sheet": "BS", "lines": ["B11 Ⅰ.유동자산 | D11=8,500 | E11=7,000"], "truncated": false}],
          "notes": [{"sheet": "2", "lines": ["B11 보유현금 | C11=200 | D11=150"], "truncated": false}],
          "linked": ["'2'!C13"]},
 "texts": [{"cell": "B12", "text": "4. 현금및현금성자산⏎당기말과 …"}], "texts_truncated": false}
```

| 필드 | 뜻 |
|---|---|
| `items[].id` | `F…`는 풋팅, `X…`는 대사 차이. 판정의 `id`에 그대로 쓴다 |
| `items[].target` | 대상 칸(`시트!칸`). `수정대상`에 그대로 쓴다 |
| `items[].stated` / `expected` | 적힌 값 / 도구가 세운 관계식으로 계산한 값(대사 차이는 빈칸) |
| `items[].pretag` | 도구의 사전 분류(금액 불일치, 절사차이, 비율 차이, 대사 차이) |
| `xref.*.lines` | 시트의 값 행. `칸 라벨 \| 칸=값` 꼴이고 `(단위: …)`, `[머리글 N행]` 줄이 섞여 있다 |
| `xref.linked` | 이미 대사된 대상 칸. 제안에서 뺀다 |
| `texts[].text` | 원문 시트의 글. `⏎`는 편집기의 자동 줄바꿈이다 |

판정 JSON(`annotate-ai`의 둘째 인자). 세 부분 모두 생략할 수 있다:

```json
{"reviewer": "판정한 쪽 이름",
 "verdicts": [{"id": "F1", "판정": "수정필요", "근거": "구성 합 1,300과 100 차이",
               "수정대상": "2!C13", "제안값": "1300", "수정제안": "합계를 1,300으로"}],
 "xref": [{"src": "BS!D12", "dst": "'2'!C13", "이유": "현금및현금성자산 주석 합계 = 재무상태표"}],
 "typos": [{"셀": "B12", "원문": "재무재표", "수정": "재무제표", "사유": "오타"}]}
```

다시 검산하는 규칙:

- `수정필요`는 `수정대상`이 대상 칸과 같고 `제안값`이 `expected`와 같을 때만 유지한다. 아니면 `확인필요`로 낮춘다. 대사 차이는 어느 쪽이 틀렸는지 정할 수 없어 늘 `확인필요`다.
- 판정이 빠졌거나 `id`가 겹친 항목은 `확인필요`로 남는다. 모르는 판정어도 `확인필요`가 된다.
- `xref`는 워크북 값으로 맞춰 본다. 일치하면 주석 시트 오른쪽에 링크와 비교 수식을 심고(`반영(일치)`), 금액이 다르면 `불일치`, 형식이 틀렸거나 받을 수 없으면 `거부`와 사유를 낸다.
- `typos`는 그 칸이 있고 `원문`이 그 칸의 글에 들어 있을 때만 받는다.

`annotate-ai --json`(풋팅 워크북)의 결과 필드: `kind: "foot"`, `status`(complete·partial), `items`, `verdicts`(최종 판정별 건수), `missing_ids`, `unknown_ids`, `duplicate_ids`, `to_check`(id·target·label·verdict·given·validation·reason·fix), `xref`(counts, proposals), `typos`(accepted·rejected·rejected_rows).
사람이 봐야 할 항목(`to_check`)이 있으면 `--json`일 때 종료코드 1이다.

### 칸별 출처 기록 (`dsdxlsx.map/1`)

`fill`과 `build`가 출력 DSD 옆 `_build/<DSD 이름>.map.json`에 쓴다. DSD 칸마다 값을 가져온 정산표 칸을 적어 두어, DSD에서 고친 숫자를 정산표의 어느 칸에 반영할지 찾는 데 쓴다.

```json
{
  "schema": "dsdxlsx.map/1",
  "side": "별도",
  "made_by": "build",
  "dsd": {"file": "보고서.dsd", "path": "…", "sha256": "…"},
  "skeleton": {"file": "뼈대.dsd", "path": "…", "sha256": "…"},
  "workbook": {"file": "정산표.xlsx", "path": "…", "sha256": "…"},
  "recipe": null,
  "counts": {"cells": 320, "overridden": 0, "missing": 0, "unfilled": 12},
  "cells": [
    {"cell": "t003.r004.c02", "section": "재무상태표", "table": "t003", "row": "t003.r004", "label": "대손충당금",
     "col": "제 2(당) 기말", "value": -120.0, "sources": [{"sheet": "BS", "cell": "C11"}], "scale": 1.0, "sign": -1,
     "how": "앵커(전기 값 일치)"}
  ]
}
```

| 필드 | 뜻 |
|---|---|
| `cells[].cell` | DSD 칸 ID(`inspect`의 칸 ID와 같다) |
| `value` | 채울 때 쓴 값. `sign × (sources 칸의 합) × scale`을 DSD 자릿수로 반올림한 것이다 |
| `sources` | 값을 가져온 정산표 칸. 여러 개면 그 합(하위 계정 합). 비어 있으면 정산표 칸 없이 계산한 값(위 칸의 합)이다 |
| `scale`, `sign` | 정산표 값에 곱한 배율(원과 천원, 율은 100)과 부호(차감 계정을 양수로 적는 정산표는 -1) |
| `row_label` | 최종 DSD에서 그 칸이 놓인 행의 이름. 나중에 칸 번호가 밀렸는지 볼 때 맞춰 본다 |
| `state` | `overridden`이면 채운 뒤 레시피나 글 바꾸기가 그 칸을 다시 고쳤다(`final`이 최종 값). `missing`이면 최종 DSD에 그 칸이 없다. 없으면 기록대로다 |
| `dsd`, `skeleton`, `workbook`, `recipe` | 만든 때의 파일 이름(`file`), 경로(`path`), 내용(`sha256`). `status`가 그 뒤 파일이 바뀌었는지 볼 때 쓴다 |
| `made_by` | 기록을 남긴 명령(`build`나 `fill`) |

### link back의 반영 목록

`changes`의 항목은 DSD에서 바뀐 칸 하나다.

| 필드 | 뜻 |
|---|---|
| `cell`, `section`, `label`, `col` | DSD 칸 |
| `old`, `new` | 채울 때 쓴 값과 지금 DSD의 값 |
| `kind` | 아래 표 |
| `sheet`, `ref` | 출처인 정산표 칸(없으면 `null`) |
| `proposed` | 정산표 칸에 넣을 값(단위와 부호를 되돌린 것). 수식 칸은 `null` |
| `applied` | `--apply`로 정산표에 썼으면 `true` |
| `tb_delta`, `accounts` | `kind`가 `tb`일 때: TB 금액(차변 +)을 얼마 바꾸면 되는지와 그 행에 걸린 TB 계정들 |
| `also` | 같은 정산표 칸에서 오는 다른 DSD 칸의 수(다시 build하면 같이 바뀐다) |
| `note` | 사람용 설명 |

| `kind` | 뜻 | `--apply` |
|---|---|---|
| `value` | 출처가 값 칸이다 | 쓴다 |
| `done` | 정산표에 이미 그 값이 있다 | - |
| `scaled` | 값 칸이지만 단위가 달라(원과 천원, 율) 정확한 값을 모른다 | 제안 값만 |
| `conflict` | 정산표의 그 칸도 바뀌었거나, 같은 정산표 칸에서 오는 DSD 칸들을 서로 다른 값으로 고쳤다 | 안 쓴다 |
| `tb` | TB를 끌어오는 수식이다(`link make` 정산표) | 안 쓴다 |
| `formula` | 합계나 다른 칸을 참조하는 수식이다 | 안 쓴다 |
| `sum` | 정산표 여러 칸의 합이다 | 안 쓴다 |
| `tie` | 정산표에 값이 같은 칸이 둘 이상이라 그중 하나를 골랐던 칸이다 | 안 쓴다 |
| `calc` | 정산표 칸 없이 계산한 값이다(위 칸의 합) | 안 쓴다 |
| `recipe` | 레시피가 넣은 칸이다 | 안 쓴다 |
| `blank` | 고친 칸이 비었거나 숫자가 아니다 | 안 쓴다 |
| `moved` | 칸 번호가 밀렸다(DSD에서 행이나 표를 넣고 뺌). 하나라도 있으면 아무것도 쓰지 않는다 | 안 쓴다 |

### link의 TB (`dsdxlsx.tb/1`)

`link make --tb`와 `link refresh --tb`가 받는 계정별 금액이다. 금액은 차변이 +, 대변이 -다.

```json
{
  "schema": "dsdxlsx.tb/1",
  "period": {"cur": "2026.12.31", "pri": "2025.12.31"},
  "rows": [
    {"code": "101", "name": "현금", "settle": "현금및현금성자산", "cur": 150, "aje": 0, "pri": 100},
    {"code": "109", "name": "대손충당금", "settle": "대손충당금(매출채권)", "cur": -120, "pri": -100}
  ]
}
```

| 필드 | 뜻 |
|---|---|
| `code` | 계정 코드. 없으면 `name`을 쓴다. `link refresh`가 이 값으로 줄을 찾는다 |
| `name` | 계정과목(필수) |
| `settle` | 정산 과목 - 재무제표에 나오는 이름. DSD 행 이름과 짝짓는 열쇠다. 없으면 `name`. 차감 계정은 `대손충당금(매출채권)`처럼 짝 자산을 괄호에 적는다 |
| `sub` | 공시용 과목(선택). DSD 행 이름이 이것과 같으면 그 아래 정산 과목을 모두 그 행에 묶는다 |
| `cur`, `aje`, `pri` | 당기 회사 제시 금액, 수정분개, 전기 금액. 정산표의 당기 = `cur` + `aje` |

`tb` 명령이 만든 파일에는 `group`(대분류), `sub_group`(소분류), `disclosure`(공시 그룹), `erp`, `files`도 들어 있다. `link`는 이 값들을 읽지 않는다.

짝짓기는 전기 금액으로 확인하므로 `pri`가 뼈대 DSD의 전기 칸과 맞아야 한다. 같은 내용을 엑셀이나 CSV 표로 줘도 된다(머리글 `계정과목` 필수, `코드`, `정산 과목`, 금액 열은 `당기`·`전기`나 기준일).

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
