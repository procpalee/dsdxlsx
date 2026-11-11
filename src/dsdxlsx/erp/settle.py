"""정산 과목 규칙 - 회사 계정 이름을 재무제표에 나오는 과목 이름으로 묶는다.

회사 장부의 '보통예금' · '외상매출금'은 재무제표에서 '현금및현금성자산' · '매출채권'으로 나온다. DSD 행과 짝지으려면
TB 의 정산 과목이 재무제표 이름이어야 하므로, 규칙 파일(TOML)로 이름을 묶는다.

규칙 파일을 찾는 순서: 지정한 경로 → 환경변수 DSDXLSX_SETTLE_RULES → 패키지의 기본 규칙(data/settle_rules.toml).
기본 규칙은 흔한 묶음 몇 가지뿐이다. 회사나 법인의 과목 체계는 따로 만든 파일로 준다.
"""
from __future__ import annotations

import os
import re
import tomllib
from dataclasses import dataclass, field
from importlib import resources

from ..guard import InputError

ENV = "DSDXLSX_SETTLE_RULES"


@dataclass
class Rules:
    rules: list = field(default_factory=list)          # [(정규식, 대분류, 그룹 또는 None, 제외 정규식 또는 None, 과목)]
    contra: re.Pattern | None = None                   # '차감 계정(짝 자산)' 모양
    source: str = ""


def key(name) -> str:
    return re.sub(r"[\s.·]", "", str(name or ""))


def _parse(data: dict, source: str) -> Rules:
    out = []
    for i, r in enumerate(data.get("rule") or [], 1):
        try:
            out.append((re.compile(r["match"]), str(r["big"]), r.get("group") or None,
                        re.compile(r["except"]) if r.get("except") else None, str(r["settle"])))
        except (KeyError, re.error, TypeError) as e:
            raise InputError("정산 과목 규칙 %d번째가 잘못됐습니다(%s): %s" % (i, e, source))
    names = [str(n) for n in (data.get("contra") or {}).get("names") or []]
    contra = re.compile(r"^(%s)\((.+)\)$" % "|".join(re.escape(n) for n in names)) if names else None
    return Rules(out, contra, source)


def load(path: str | None = None) -> Rules:
    """규칙 파일을 읽는다. path 가 없으면 환경변수, 그것도 없으면 기본 규칙."""
    path = path or os.environ.get(ENV, "").strip() or None
    if path:
        if not os.path.exists(path):
            raise InputError("정산 과목 규칙 파일을 찾을 수 없습니다: %s" % path)
        with open(path, "rb") as f:
            try:
                return _parse(tomllib.load(f), path)
            except tomllib.TOMLDecodeError as e:
                raise InputError("정산 과목 규칙 파일을 읽지 못했습니다(%s): %s" % (e, path))
    data = resources.files("dsdxlsx").joinpath("data/settle_rules.toml").read_bytes()
    return _parse(tomllib.loads(data.decode("utf-8")), "기본 규칙")


def settle(name, big, sub: str = "", disc: str = "", rules: Rules | None = None) -> str:
    """정산 과목 이름 하나 → 묶은 이름(맞는 규칙이 없으면 그대로).

    big = 대분류(자산 · 부채 · 자본 · 손익), sub = 소분류, disc = 공시 그룹(당좌자산 · 유동부채 · 판매비와관리비 …).
    매출원가 아래 계정은 묶지 않는다. 차감 계정 '대손충당금(외상매출금)'은 괄호 안 짝 자산을 묶는다.
    """
    n = str(name or "").strip()
    if rules is None or not n:
        return n
    if big == "손익" and "매출원가" in (sub, disc):
        return n
    m = rules.contra.match(key(n)) if rules.contra else None
    if m:
        head = settle(m.group(1), big, sub, disc, rules) if big == "부채" else m.group(1)
        return "%s(%s)" % (head, settle(m.group(2), big, sub, disc, rules))
    k = key(n)
    for pat, b, grp, exc, std in rules.rules:
        if b == big and (grp is None or grp in (disc, sub)) and pat.fullmatch(k) and not (exc and exc.search(k)):
            return std
    return n
