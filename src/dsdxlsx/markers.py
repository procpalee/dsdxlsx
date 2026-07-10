"""엑셀 구조 편집 마커 문법 — 파싱만, 부수효과 없음.

규약 (README·DART작성규칙 문서와 계약):
- 마커는 **B열**에 쓰고 A열은 비운다(A열은 행 ID 전용). `#`로 시작.
- 셀 텍스트를 `#`로 시작하고 싶으면 `##`로 이스케이프 (import가 한 겹 벗김).
- 알 수 없는 `#…`는 에러 — 오타가 셀 텍스트로 둔갑해 침묵하는 것을 막는다.

    #새표 [헤더=k] [원본=t014] [단위=원]   ← 다음 행부터 빈 행/#표끝 전까지가 표 내용
    #표끝                                   ← (선택) 새 표 블록 명시 종료
    #표삭제                                 ← 지우려는 표의 첫 행 B열에 덮어쓰기
    #문단 이하 전체가 본문…                 ← 새 문단 삽입 (직전 블록 뒤)
    #문단삭제                               ← 기존 p### 행의 B열에 덮어쓰기
    #새주석 35 제목 / #주석삭제             ← A2에서 지원 (현재는 안내 에러)
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

KINDS = {"새표", "표끝", "표삭제", "문단", "문단삭제", "새주석", "주석삭제"}
_ARG_RE = re.compile(r"^([^\s=]+)=(.+)$")


class MarkerError(Exception):
    pass


@dataclass
class Marker:
    kind: str
    args: dict[str, str] = field(default_factory=dict)
    text: str = ""  # 문단 본문 / 새주석 제목


def is_marker(value) -> bool:
    """B열 값이 마커 후보인가 ('##'는 이스케이프된 일반 텍스트)."""
    return (isinstance(value, str) and value.startswith("#")
            and not value.startswith("##"))


def unescape(value):
    """'##…' → '#…' (마커가 아닌 일반 텍스트로 강등)."""
    if isinstance(value, str) and value.startswith("##"):
        return value[1:]
    return value


def parse_marker(value: str) -> Marker:
    body = value[1:].strip()
    if not body:
        raise MarkerError("빈 마커 '#' — 텍스트가 #으로 시작하면 ##으로 쓰세요")
    kind = body.split(None, 1)[0]
    rest = body[len(kind):].strip()
    if kind not in KINDS:
        raise MarkerError(
            f"알 수 없는 마커 '#{kind}' — 사용 가능: "
            + ", ".join(f"#{k}" for k in sorted(KINDS))
            + " (텍스트가 #으로 시작하면 ##으로 이스케이프)")
    m = Marker(kind=kind)
    if kind == "문단":
        m.text = rest
        if not m.text.strip():
            raise MarkerError("#문단 뒤에 본문 텍스트가 없음")
        return m
    if kind == "새주석":
        parts = rest.split(None, 1)
        if not parts or not parts[0].isdigit():
            raise MarkerError("#새주석 형식: '#새주석 35 제목'")
        m.args["번호"] = parts[0]
        m.text = parts[1].strip() if len(parts) > 1 else ""
        return m
    # 나머지: 공백 구분 key=value 인자 (대괄호 표기 허용)
    for token in rest.split():
        token = token.strip("[]")
        if not token:
            continue
        am = _ARG_RE.match(token)
        if am is None:
            raise MarkerError(f"#{kind} 인자 형식 오류: '{token}' (key=value 형태)")
        m.args[am.group(1)] = am.group(2)
    return m
