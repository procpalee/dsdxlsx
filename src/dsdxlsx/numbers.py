"""DSD 숫자 표기 판별·변환.

문서 관례 (샘플 실측): 콤마 3자리 그룹, 음수는 괄호 "(232,116,145)",
플레이스홀더 "-", 주석 참조 "6,28,31"은 숫자가 아님(3자리 그룹 불일치로 배제).
"""
from __future__ import annotations

import re
from decimal import Decimal

_INT_RE = re.compile(r"^\d{1,3}(?:,\d{3})*$")
_DEC_RE = re.compile(r"^\d{1,3}(?:,\d{3})*\.\d+$")
# 수량·비율 접미 단위: "2,978,480주", "62.12%" — 합계 검증 대상 (주석1 주주현황 등)
_SUFFIXES = ("주", "%", "명", "건", "좌", "구좌", "개")

# Excel double 정밀도(15자리) 초과 정수는 텍스트로 취급
_MAX_SAFE = 10**15


def suffix_of(text: str) -> str:
    """숫자 뒤 단위 접미('주', '%' 등). 없으면 빈 문자열."""
    s = text.strip()
    for suf in _SUFFIXES:
        if s.endswith(suf) and classify(s[: -len(suf)]) is not None:
            return suf
    return ""


def classify(text: str) -> int | Decimal | None:
    """셀 텍스트가 숫자 표기이면 값 반환, 아니면 None. 전체 일치만 인정.
    접미 단위('2,978,480주', '62.12%')는 단위를 뗀 수치로 인정."""
    s = text.strip()
    if not s:
        return None
    for suf in _SUFFIXES:
        if s.endswith(suf):
            s = s[: -len(suf)].strip()
            break
    if not s:
        return None
    negative = s.startswith("(") and s.endswith(")")
    core = s[1:-1].strip() if negative else s
    if _INT_RE.match(core):
        value = int(core.replace(",", ""))
        if value >= _MAX_SAFE:
            return None
        return -value if negative else value
    if _DEC_RE.match(core):
        value = Decimal(core.replace(",", ""))
        return -value if negative else value
    return None


def to_dsd_text(value: int | float | Decimal) -> str:
    """숫자 → 문서 관례 표기 (콤마, 음수는 괄호)."""
    if isinstance(value, float):
        value = Decimal(str(value))
    if isinstance(value, Decimal) and value == value.to_integral_value():
        value = int(value)
    negative = value < 0
    magnitude = -value if negative else value
    if isinstance(magnitude, Decimal):
        text = f"{magnitude:,f}"
    else:
        text = f"{magnitude:,}"
    return f"({text})" if negative else text


def values_equal(xlsx_value, node_text: str) -> bool:
    """xlsx 셀 값과 원본 노드 텍스트의 의미 동등 비교 (무수정 판정)."""
    original = classify(node_text)
    if isinstance(xlsx_value, (int, float, Decimal)) and not isinstance(xlsx_value, bool):
        if original is None:
            # 플레이스홀더 "-"는 숫자 0으로 내보냄(수식 호환) — 0이면 무수정
            return node_text.strip() == "-" and Decimal(str(xlsx_value)) == 0
        return Decimal(str(xlsx_value)) == Decimal(original)
    text = "" if xlsx_value is None else str(xlsx_value)
    return text == node_text


def decimal_places(text: str) -> int:
    s = text.strip()
    if (suf := suffix_of(s)):
        s = s[: -len(suf)].strip()
    s = s.strip("()")
    return len(s.split(".")[1]) if "." in s else 0
