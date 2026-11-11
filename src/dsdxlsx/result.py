"""명령 결과 공통 봉투(--json) - LLM·스크립트가 명령마다 다른 출력을 해석하지 않게 한다.

{"schema": "dsdxlsx.result/1", "command": "apply", "version": "0.6.0", "ok": true,
 "exit_code": 0, "outputs": ["…_수정.dsd"], "warnings": [], "errors": [], …명령별 필드}

종료코드 규약
  0 OK          정상
  1 FINDINGS    정상 실행했고 확인할 발견 사항이 있음(검산 불일치 등)
  2 INPUT_ERROR 입력 오류(파일·JSON·위치 지정·가드 위반) - 아무것도 쓰지 않음
  3 STOPPED     안전 장치로 멈춤(기존 파일 덮어쓰기 등) - 아무것도 쓰지 않음
"""
from __future__ import annotations

import json
import sys

from . import __version__

SCHEMA = "dsdxlsx.result/1"
OK, FINDINGS, INPUT_ERROR, STOPPED = 0, 1, 2, 3


def envelope(command: str, exit_code: int, *, outputs=None, warnings=None, errors=None,
             **extra) -> dict:
    d = {"schema": SCHEMA, "command": command, "version": __version__,
         "ok": exit_code in (OK, FINDINGS), "exit_code": exit_code,
         "outputs": list(outputs or []), "warnings": list(warnings or []),
         "errors": list(errors or [])}
    d.update(extra)
    return d


def dumps(d: dict) -> str:
    return json.dumps(d, ensure_ascii=False, indent=1)


def emit(d: dict) -> int:
    """봉투를 표준출력에 쓰고 그 종료코드를 돌려준다."""
    sys.stdout.write(dumps(d) + "\n")
    return d["exit_code"]
