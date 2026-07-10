"""검증 프로파일 — 규칙은 코드가 아니라 데이터.

기본값은 패키지 동봉 data/_default.toml (v1 하드코딩과 동일 — golden 회귀로 고정).
회사별 편차는 profiles/<회사명>/profile.toml 로 override 한다:

    profiles/
      E사(주)/
        profile.toml      # 여기 있는 키만 기본값을 덮어씀 (부분 지정)
        ties.json         # 대사맵 (ties.py가 관리)

회사 프로파일 탐색 경로: --profile 로 받은 경로/이름 → 실행 폴더의 profiles/.
이름 해석은 meta.xml DOCUMENT-HEADER regname 우선, 폴더명은 정규화(공백·(주) 제거)
매치도 허용한다. 고객사명이 든 프로파일은 커밋 금지(.gitignore) — 로컬 전용.
"""
from __future__ import annotations

import re
import tomllib
from dataclasses import dataclass, field
from functools import lru_cache
from importlib import resources
from pathlib import Path

PROFILES_DIR = "profiles"


@dataclass(frozen=True)
class SuppressRule:
    check: str
    reason: str
    where_contains: str = ""
    message_contains: str = ""

    def matches(self, check: str, where: str, message: str) -> bool:
        return (
            self.check == check
            and (not self.where_contains or self.where_contains in where)
            and (not self.message_contains or self.message_contains in message)
        )


@dataclass(frozen=True)
class Profile:
    name: str
    # [labels]
    total_re: re.Pattern
    grand_total_re: re.Pattern
    total_header_re: re.Pattern
    deduct_total_re: re.Pattern
    flow_end_re: re.Pattern
    generic_labels: frozenset[str]
    # [periods]
    period_current: tuple[str, ...]
    period_prior: tuple[str, ...]
    # [statement]
    anchor_re: re.Pattern
    pl_minus_re: re.Pattern
    flow_minus_re: re.Pattern
    inventory_re: re.Pattern
    nonsum_col_re: re.Pattern
    # [sections]
    skip_section_re: re.Pattern
    # [units]
    unit_re: re.Pattern
    unit_mult: dict[str, int]
    # [tolerance]
    footing_tol: int
    global_search_min: int
    # [[suppress]]
    suppress: tuple[SuppressRule, ...] = field(default=())


def _deep_merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in over.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def _build(name: str, d: dict) -> Profile:
    labels, periods = d["labels"], d["periods"]
    stmt, units, tol = d["statement"], d["units"], d["tolerance"]
    return Profile(
        name=name,
        total_re=re.compile(labels["total"]),
        grand_total_re=re.compile(labels["grand_total"]),
        total_header_re=re.compile(labels["total_header"]),
        deduct_total_re=re.compile(labels["deduct_total"]),
        flow_end_re=re.compile(labels["flow_end"]),
        generic_labels=frozenset(labels["generic"]),
        period_current=tuple(periods["current"]),
        period_prior=tuple(periods["prior"]),
        anchor_re=re.compile(stmt["anchor"]),
        pl_minus_re=re.compile(stmt["pl_minus"]),
        flow_minus_re=re.compile(stmt["flow_minus"]),
        inventory_re=re.compile(stmt["inventory"]),
        nonsum_col_re=re.compile(stmt["nonsum_col"]),
        skip_section_re=re.compile(d["sections"]["skip"]),
        unit_re=re.compile(units["pattern"]),
        unit_mult=dict(units["mult"]),
        footing_tol=int(tol["footing"]),
        global_search_min=int(tol["global_search_min"]),
        suppress=tuple(
            SuppressRule(
                check=s["check"], reason=s["reason"],
                where_contains=s.get("where_contains", ""),
                message_contains=s.get("message_contains", ""),
            )
            for s in d.get("suppress", [])
        ),
    )


@lru_cache(maxsize=1)
def _default_dict() -> dict:
    data = resources.files("dsdxlsx").joinpath("data/_default.toml").read_bytes()
    return tomllib.loads(data.decode("utf-8"))


def default_profile() -> Profile:
    return _build("_default", _default_dict())


def _norm_company(name: str) -> str:
    """폴더명 매칭용 정규화: 공백·법인 접미(주식회사/(주)) 제거."""
    s = re.sub(r"\s+", "", name)
    s = re.sub(r"(주식회사|\(주\)|㈜)", "", s)
    return s


def _company_toml(name: str, base_dir: Path) -> Path | None:
    root = base_dir / PROFILES_DIR
    if not root.is_dir():
        return None
    exact = root / name / "profile.toml"
    if exact.is_file():
        return exact
    want = _norm_company(name)
    for child in root.iterdir():
        if child.is_dir() and _norm_company(child.name) == want:
            p = child / "profile.toml"
            if p.is_file():
                return p
    return None


def company_dir(name: str, base_dir: Path | None = None) -> Path:
    """회사 프로파일 폴더 (없으면 생성하지 않고 경로만) — ties.json 등이 공유."""
    root = (base_dir or Path.cwd()) / PROFILES_DIR
    if root.is_dir():
        want = _norm_company(name)
        for child in root.iterdir():
            if child.is_dir() and _norm_company(child.name) == want:
                return child
    return root / name


def load_profile(spec: str, base_dir: Path | None = None) -> Profile:
    """spec = toml 경로 또는 회사명. 기본 프로파일 위에 부분 override 로 병합."""
    base_dir = base_dir or Path.cwd()
    path = Path(spec)
    if not (path.is_file() and path.suffix == ".toml"):
        found = _company_toml(spec, base_dir)
        if found is None:
            raise FileNotFoundError(
                f"프로파일을 찾지 못함: '{spec}' — {base_dir / PROFILES_DIR}/<회사명>/profile.toml"
            )
        path = found
    over = tomllib.loads(path.read_text(encoding="utf-8"))
    merged = _deep_merge(_default_dict(), over)
    name = over.get("meta", {}).get("name") or path.parent.name
    return _build(name, merged)


def regname_of(meta: bytes | None) -> str | None:
    """meta.xml DOCUMENT-HEADER의 회사명."""
    if not meta:
        return None
    m = re.search(rb'<DOCUMENT-HEADER\b[^>]*?regname="([^"]*)"', meta)
    if not m:
        return None
    return m.group(1).decode("utf-8", errors="replace") or None


def resolve_profile(meta: bytes | None = None, explicit: str | None = None,
                    base_dir: Path | None = None) -> Profile:
    """--profile 명시 > meta regname 자동 탐색 > 기본 프로파일."""
    if explicit:
        return load_profile(explicit, base_dir)
    name = regname_of(meta)
    if name:
        try:
            return load_profile(name, base_dir)
        except FileNotFoundError:
            pass  # 회사 프로파일 없음 — 기본으로
    return default_profile()
