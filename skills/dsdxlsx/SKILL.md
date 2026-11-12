---
name: dsdxlsx
description: DART DSD(전자공시 문서)를 dsdxlsx 명령으로 읽고, 고치고, 검산하는 일반 작업 흐름. "이 DSD 읽어줘 / 주석 N 숫자 고쳐줘 / 행·표·주석 넣어줘(지워줘) / 합계 맞는지 봐줘 / 두 DSD 비교해줘 / DART에서 감사보고서 받아줘 / 전기 DSD로 당기 뼈대 만들어줘" 요청에 사용. 분·반기 보고서를 조서에서 채우는 일은 dsd-interim, 오타·주기 같은 판단형 검토는 dsd-review.
---

# dsdxlsx - DSD 읽기·고치기·검산

도구가 정해진 일(읽기, 쓰기, 검산)을 하고, 이 스킬은 순서와 판단을 맡는다.
형식 전체는 저장소의 `docs/JSON형식.md`, 원칙은 `AGENTS.md`.

## 전제

- 도구 찾기: PATH의 `dsdxlsx`(`uv tool install git+https://github.com/procpalee/dsdxlsx`), 없으면 환경변수 `DSDXLSX_HOME`(저장소 폴더)에서 `uv run --no-editable --project "$DSDXLSX_HOME" dsdxlsx …`.
- 모든 명령에 `--json`을 붙여 결과 봉투를 읽는다. 종료코드: 0 정상 / 1 확인할 것 있음 / 2 입력 오류 / 3 안전장치로 멈춤.
- 작업 파일은 사용자의 건 폴더에 둔다. 원본 DSD는 덮어쓰지 않는다.

## 지킬 것

1. 숫자를 지어내지 않는다. 넣는 값은 사용자가 준 자료나 지시에서 온 것만.
2. 검산 불일치에서 어느 칸이 틀렸는지 근거 없이 정하지 않는다. `stated`(적힌 값)와 `computed`(다시 계산한 값)를 보여 주고 사용자에게 확인한다. 사용자가 "합계를 계산값으로 맞춰" 같이 정해 주면 그대로 한다.
3. 종료코드 3(멈춤)에서 `--force`를 스스로 붙이지 않는다. 멈춘 파일은 사람이 고친 파일일 수 있으니 사용자에게 묻는다.
4. 결과 보고에 "제출 전에 DART 편집기로 열어 확인"을 넣는다.

## 절차

### A. 읽기

```
dsdxlsx inspect <DSD> --outline          # 구획·표 목록부터
dsdxlsx inspect <DSD> --find <글>        # 찾는 글이나 숫자의 위치
dsdxlsx inspect <DSD> --section <값>     # 필요한 구획만(재무제표 명칭 또는 주석 번호)
```

문서를 통째로 읽지 않는다. 주석이 많은 문서는 `--outline` 결과로 구획을 고른다.

### B. 고치기

1. A로 고칠 자리의 행 ID와 열 번호를 얻는다.
2. 편집 목록을 JSON 파일로 쓴다(사용자 건 폴더나 임시 폴더).
3. `dsdxlsx apply <DSD> edits.json --dry-run --json`으로 바뀔 내용을 확인하고 사용자에게 요약한다.
4. `dsdxlsx apply <DSD> edits.json --json`으로 반영한다. 출력은 `<이름>_수정.dsd`.
5. C로 검산한다.
6. `dsdxlsx diff <원본> <수정본>`으로 비교 화면(HTML)을 만들어 경로를 알려 준다.

편집 종류: `set`, `set_text`, `insert_row`, `delete_row`, `insert_para`, `delete_para`, `insert_table`, `delete_table`, `insert_note`, `delete_note`.
구조를 바꾸는 편집(`structural: true`) 뒤에는 새 파일을 다시 `inspect`한 다음 이어서 고친다.
주석을 지운 뒤 번호를 당기려면 `dsdxlsx renumber <DSD> --close-gap`.

### C. 검산

```
dsdxlsx check <DSD> --json
```

- 종료코드 1이면 `mismatches`를 표로 정리해 보여 준다(구획, 행 이름, 열, 적힌 값, 계산값, 차이).
- `findings`의 `info`는 참고로만 전한다.
- `coverage.unchecked`의 표는 검산하지 못한 표다. 숨기지 말고 알린다.

### D. 다음 기간 뼈대

- 연차: `dsdxlsx rollforward <전기 DSD> --json` → `_차기기초.dsd`와 정산표(엑셀).
- 분·반기: `dsdxlsx roll-interim <직전 중간보고서 DSD> --prior-same <전년 동기 DSD> --json`.
- `todo`(확인필요·미처리)를 사용자에게 그대로 전한다.
- TB: 더존·세무사랑 자료 폴더가 있으면 `dsdxlsx tb <자료 폴더> --json` → `TB_<꼬리>.json`. `warnings`와 `no_code`를 사용자에게 전한다. 종료코드 2면 모자란 파일이나 부호합 불일치이니 메시지를 그대로 전하고 멈춘다.
- 정산표로 만든 DSD 를 다시 손댈 때는 먼저 `dsdxlsx status <출력 DSD> --json` 으로 어느 쪽이 바뀌었는지 보고 `next` 를 따른다.
- 예전 정산표(0.6.1 이하로 만든 것)의 두 줄 머리글: `dsdxlsx link fix-headers <뼈대 DSD> <정산표> --json` 의 `tables` 를 보여 주고 사용자가 좋다고 하면 `--apply`. 왕복 시험 원인에 `주석 머리글에 묶음 이름이 빠짐` 이 나오면 먼저 이것을 권한다.
- DSD 에서 고친 숫자를 정산표로: `dsdxlsx link back <고친 DSD> <정산표> --json` 으로 `changes` 를 사용자에게 보여 주고, 사용자가 좋다고 하면 `--apply`(값 칸만 쓰고 사본을 남긴다). `tb` · `formula` · `conflict` 같은 나머지 종류는 `note` 를 그대로 전한다.
- 정산표: `dsdxlsx link make <뼈대 DSD> --tb <TB> --json` → `<뼈대>_정산표.xlsx`. `to_check`와 `unlinked_tb`를 사용자에게 전하고, `dsdxlsx link test <뼈대 DSD> <정산표> --notes --json`의 `diff`·`blank`가 비었는지 확인한다. TB 형식은 `docs/JSON형식.md`의 `dsdxlsx.tb/1`.

### E. DART에서 받기

```
dsdxlsx dart <회사명|종목코드|접수번호|DART 주소> --list --json    # 고를 수 있는 공시
dsdxlsx dart <…> --year 2025 --xlsx --json                        # DSD + 편집용 엑셀
```

OpenDART 키(`OPENDART_API_KEY`)가 필요하다. 이름이 여러 회사에 걸리면 후보가 오류로 나오니 사용자에게 고르게 한다.

### F. 사람이 엑셀로 고칠 때

`dsdxlsx export <DSD>`로 엑셀을 만들어 넘기고, 사람이 고친 뒤 `dsdxlsx import <엑셀>`로 되돌린다.
사람이 고친 엑셀은 다시 `export`하면 종료코드 3으로 멈춘다. 그대로 두고 사용자에게 알린다.

## 보고

- 무엇을 어디서 어디로 바꿨는지(구획, 행 이름, 열, 전 → 후)를 표로.
- 검산 결과(불일치 건수, 검산 못 한 표).
- 만든 파일 경로와 "제출 전에 DART 편집기로 열어 확인".
