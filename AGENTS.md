# AGENTS.md - AI 에이전트용 안내

dsdxlsx는 DART 전자공시 문서(DSD)를 읽고, 고치고, 검산하는 명령줄 도구다.
읽는 쪽은 둘이다. 이 도구로 DSD 작업을 하는 에이전트(1절)와 이 저장소의 코드를 고치는 에이전트(2절).
형식 전체는 [docs/JSON형식.md](docs/JSON형식.md), 사람용 설명은 [README.md](README.md)에 있다.

## 1. 도구를 쓰는 에이전트

### 지킬 것

- 숫자를 지어내지 않는다. 칸에 넣는 값은 사용자가 준 자료(조서, 정산표, 지시)에서 온 것만 쓴다. 검산 결과의 `computed`는 "구성 칸을 더하면 이 값"이라는 뜻일 뿐 정답이 아니다. 어느 칸이 틀렸는지 근거 없이 정하지 말고 사용자에게 묻는다.
- 원본 DSD를 덮어쓰지 않는다. 결과는 새 파일로 받고(`_수정.dsd` 등), 도구가 종료코드 3으로 멈추면 `--force`를 붙이기 전에 사용자에게 확인한다. 멈춘 파일은 사람이 고친 파일일 수 있다.
- 실제 회사 자료를 저장소에 넣지 않는다. 작업 파일은 저장소 밖에 두고, 회사명이나 금액을 커밋·이슈·로그에 옮기지 않는다.
- 만든 DSD는 제출 전에 사람이 DART 편집기로 열어 확인해야 한다. 이 말을 결과 보고에 넣는다.

### 기본 흐름

```
dsdxlsx inspect 보고서.dsd --outline            # 1. 구획·표 목록
dsdxlsx inspect 보고서.dsd --find 보통예금      # 2. 찾는 글의 위치(ID)
dsdxlsx inspect 보고서.dsd --section 4          # 3. 고칠 구획 읽기
dsdxlsx apply 보고서.dsd edits.json --dry-run --json   # 4. 바뀔 내용 확인
dsdxlsx apply 보고서.dsd edits.json --json             # 5. 반영 → 보고서_수정.dsd
dsdxlsx check 보고서_수정.dsd --json                   # 6. 검산(불일치가 있으면 종료코드 1)
dsdxlsx diff 보고서.dsd 보고서_수정.dsd                # 7. 사람에게 보여 줄 비교 화면(HTML)
```

- 큰 문서를 통째로 읽지 않는다. `--outline`과 `--find`로 자리를 찾고 `--section`으로 필요한 구획만 읽는다.
- 칸은 `{"row": "t003.r004", "col": 1}`처럼 행 ID + 열 번호로 가리키는 것이 가장 안전하다. 행 이름으로 가리킬 때 같은 이름이 여럿이면 도구가 후보를 내며 멈춘다.
- 행·문단·표·주석을 넣거나 지우면(`structural: true`) ID가 다시 매겨진다. 다음 편집 전에 새 파일을 다시 `inspect`한다.
- 숫자는 JSON 숫자로 준다(`950`). 괄호 음수나 단위 표기는 도구가 원래 칸을 따라 맞춘다.

### 결과 읽기

`--json`을 주면 모든 명령이 같은 봉투를 표준출력에 낸다.

| 종료코드 | 뜻 | 할 일 |
|---|---|---|
| 0 | 정상 | `outputs`의 파일로 다음 단계 |
| 1 | 끝까지 돌았고 확인할 것이 있음 | `mismatches`, `to_check`, `unfilled`, `todo`를 사용자에게 알리거나 고친다 |
| 2 | 입력 오류. 아무것도 쓰지 않음 | `errors`의 이유대로 입력을 고친다 |
| 3 | 안전장치로 멈춤. 아무것도 쓰지 않음 | 사용자에게 확인. 다른 `-o`를 쓰거나 승인받고 `--force` |

### 하는 일별 명령

| 하는 일 | 명령 |
|---|---|
| 읽기, 목차, 찾기 | `inspect` (`--outline`, `--find`, `--section`, `--json`) |
| 칸·문단·행·표·주석 고치기 | `apply` |
| 합계·대사·표기 검산 | `check` |
| 사람이 볼 풋팅·대사 검증 워크북 | `foot` |
| 풋팅 검출 항목 판정, 대사 제안, 오탈자 | `foot --review` → 판정 JSON → `annotate-ai` (스킬 `skills/dsd-foot-review`) |
| 두 판 비교 | `diff`(사람용 HTML), `verify`(구조 차이 목록) |
| 전기 DSD → 당기 기초 DSD + 정산표 | `rollforward`(연차), `roll-interim`(분·반기) |
| DSD 숫자 ↔ 조서 대사, 당기 칸 채움 | `recon`, `fill`, `build` |
| 뼈대 DSD 행 그대로 정산표 만들기, 연결 확인 | `link make` → `link test`(왕복 시험), TB 가 바뀌면 `link refresh` |
| DART 공시 → DSD | `dart` (OpenDART 키 필요) |
| 주석 번호 정리 | `renumber` |
| 사람이 엑셀로 고칠 때 | `export` → (엑셀 편집) → `import`. 풋팅 워크북(`foot`)에서 고친 칸 값도 `import`로 되돌린다 |

### 자주 하는 실수

- `check --json 보고서.dsd`처럼 `--json`을 앞에 써도 되지만, 뒤에 `.json` 경로를 붙이면 옛 방식(파일로 저장)으로 돈다.
- 칸 ID 끝의 `c02`는 행 안에서 몇 번째 칸인지다. 열 번호(`col`)와 다르므로 `col`에 쓰지 않는다.
- 엑셀(`export` 결과)과 DSD를 헷갈리지 않는다. `apply`와 `inspect`는 DSD만 받는다.
- 재무제표 본문에는 표를 새로 만들거나 지울 수 없다. 표·문단·주석의 추가와 삭제는 주석 구획에서만 된다.

## 2. 이 저장소를 고치는 에이전트

- 시험: `uv run --no-editable python -m pytest -q` (한글 경로에서는 `uv run pytest`가 실패한다). 소스를 고친 뒤 설치본으로 돌려 보려면 `uv sync --no-editable --reinstall-package dsdxlsx`.
- 시험 자료는 `tests/fixtures/factory.py`의 합성 문서만 쓴다. 실제 DSD, 회사명, 실제 금액을 저장소에 넣지 않는다(`sample/`, `profiles/`는 git 제외).
- 쓰기 경로는 하나다. DSD 본문을 바꾸는 기능은 `splice.apply`(바뀐 바이트만 끼워 넣기)를 거치고, 고치지 않은 곳은 원본 바이트 그대로여야 한다. 편집이 없으면 출력이 원본과 바이트가 같아야 한다.
- 새 명령이나 옵션은 `cli.py`의 `INPUTS`(입력 파일 형식 점검)와 `_note`(봉투에 실을 결과)에 맞춘다. 사람이 이어서 고치는 출력은 `guard.check_out`으로 덮어쓰기를 막고 쓴 뒤 `guard.remember`로 기록한다.
- 기존 옵션과 `--json` 없는 종료코드는 바꾸지 않는다. 다른 도구(엑셀 리본, 스킬)가 기대고 있다.
- 도구 안에서 LLM API를 부르지 않는다. 판단이 필요한 일은 도구가 후보를 JSON으로 내고, 에이전트가 판단한 결과를 다시 JSON으로 받는 방식으로 만든다.
- 커밋 메시지는 한국어 본문에 `feat:`·`fix:`·`docs:` 머리말. 변경 내역은 `CHANGELOG.md`의 "다음 판"에 적는다.
