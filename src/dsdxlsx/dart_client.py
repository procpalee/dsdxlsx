"""OpenDART 재무제표 조회 — 당전기 밀기의 '전기 = 공시 기준' 대사·채움 원천.

- fnlttSinglAcntAll.json (단일회사 전체 재무제표): 상장법인만 제공. 비상장·미공시는
  status 013 → 예외가 아니라 '생략 사유'로 돌려준다 (파일 기반 롤은 계속 진행).
- 키: 환경변수 OPENDART_API_KEY → repo 루트 .env → 실행 폴더 .env.
- 캐시: %LOCALAPPDATA%\\dsdxlsx\\cache\\dart_*.json (24시간). 네트워크 없이 재실행 가능.
- 표준 라이브러리만 사용(urllib) — dsdxlsx 의존성 추가 없음.
"""
from __future__ import annotations

import io
import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

API_BASE = "https://opendart.fss.or.kr/api"
CACHE_TTL = 24 * 3600
_KEY_NAME = "OPENDART_API_KEY"

# sj_div → 문서 시트 종류
SJ_KIND = {"BS": "BS", "IS": "PL", "CIS": "PL", "CF": "CF", "SCE": "SCE"}


class DartUnavailable(Exception):
    """조회 불가 사유(키 없음·비상장·네트워크) — 호출자는 대사를 생략하고 진행한다."""


@dataclass
class Item:
    sj_div: str
    account_id: str
    account_nm: str
    ord: int
    thstrm: int | None          # 당기 (분기 보고서: 3개월)
    thstrm_add: int | None      # 당기 누적 (분·반기 손익)
    frmtrm: int | None          # 전기 (연차) / 전기말 (BS)
    frmtrm_q: int | None        # 전기 같은 분기 3개월
    frmtrm_add: int | None      # 전기 누적
    detail: str = ""            # SCE 자본 구성요소 등

    @property
    def norm(self) -> str:
        return norm_label(self.account_nm)


@dataclass
class DartData:
    corp_code: str
    year: int
    reprt_code: str
    fs_div: str
    items: list[Item] = field(default_factory=list)
    rcept_no: str = ""
    thstrm_nm: str = ""
    frmtrm_nm: str = ""
    source: str = "cache"

    def by_kind(self, kind: str) -> list[Item]:
        return [it for it in self.items if SJ_KIND.get(it.sj_div) == kind]


# ── 키·캐시 ──────────────────────────────────────────────────────────

def _read_env_file(path: Path) -> str | None:
    try:
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith(_KEY_NAME + "="):
                v = line.split("=", 1)[1].strip().strip('"').strip("'")
                if v:
                    return v
    except OSError:
        return None
    return None


def api_key() -> str | None:
    v = os.environ.get(_KEY_NAME, "").strip()
    if v:
        return v
    here = Path(__file__).resolve()
    candidates = [Path.cwd() / ".env"]
    for parent in here.parents[:4]:  # src/dsdxlsx → repo root
        candidates.append(parent / ".env")
    for c in candidates:
        if c.is_file():
            v = _read_env_file(c)
            if v:
                return v
    return None


def cache_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_CACHE_HOME") or str(Path.home())
    d = Path(base) / "dsdxlsx" / "cache"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _cache_get(name: str) -> dict | None:
    p = cache_dir() / name
    try:
        if p.is_file() and time.time() - p.stat().st_mtime < CACHE_TTL:
            return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return None


def _cache_put(name: str, obj: dict) -> None:
    try:
        (cache_dir() / name).write_text(json.dumps(obj, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def _http_get(url: str, params: dict, timeout: int = 60) -> bytes:
    q = urllib.parse.urlencode(params)
    req = urllib.request.Request(f"{url}?{q}", headers={"User-Agent": "dsdxlsx"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


# ── corp_code 보조 조회 (AREGCIK 없는 문서용) ────────────────────────

def resolve_corp_code(company_name: str) -> str | None:
    """corpCode.xml(약 10만사, 24h 캐시)에서 회사명 → corp_code. 상장사 정확 일치 우선."""
    key = api_key()
    if not key or not company_name:
        return None
    cached = _cache_get("corp_codes.json")
    entries = cached.get("list") if cached else None
    if entries is None:
        try:
            raw = _http_get(f"{API_BASE}/corpCode.xml", {"crtfc_key": key})
            zf = zipfile.ZipFile(io.BytesIO(raw))
            root = ET.fromstring(zf.read(zf.namelist()[0]))
        except Exception:  # noqa: BLE001 — 보조 경로: 실패해도 진행
            return None
        entries = [
            {"c": (el.findtext("corp_code") or "").strip(),
             "n": (el.findtext("corp_name") or "").strip(),
             "s": (el.findtext("stock_code") or "").strip()}
            for el in root.iter("list")
        ]
        _cache_put("corp_codes.json", {"list": entries})
    want = _norm_company(company_name)
    hits = [e for e in entries if _norm_company(e["n"]) == want]
    if not hits:
        return None
    hits.sort(key=lambda e: (0 if e["s"] else 1))
    return hits[0]["c"]


def _norm_company(name: str) -> str:
    s = re.sub(r"\s+", "", name)
    return re.sub(r"(주식회사|\(주\)|㈜)", "", s)


# ── 재무제표 조회 ────────────────────────────────────────────────────

def _to_int(v) -> int | None:
    s = str(v if v is not None else "").replace(",", "").strip()
    if s in ("", "-"):
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def fetch_statements(corp_code: str, year: int, reprt_code: str,
                     fs_div: str = "OFS") -> DartData:
    """단일회사 전체 재무제표. 데이터 없음(013)·키 없음·네트워크 오류 → DartUnavailable."""
    if not corp_code:
        raise DartUnavailable("DART corp_code 없음 (문서 AREGCIK 비어 있음·회사명 조회 실패)")
    key = api_key()
    if not key:
        raise DartUnavailable("OPENDART_API_KEY 없음 (.env 또는 환경변수)")
    name = f"dart_{corp_code}_{year}_{reprt_code}_{fs_div}.json"
    js = _cache_get(name)
    source = "cache"
    if js is None:
        try:
            raw = _http_get(f"{API_BASE}/fnlttSinglAcntAll.json", {
                "crtfc_key": key, "corp_code": corp_code, "bsns_year": str(year),
                "reprt_code": reprt_code, "fs_div": fs_div,
            })
            js = json.loads(raw)
        except (urllib.error.URLError, OSError, ValueError) as e:
            raise DartUnavailable(f"OpenDART 호출 실패: {e}") from e
        source = "api"
        if js.get("status") in ("000", "013"):
            _cache_put(name, js)
    status = js.get("status")
    if status == "013":
        raise DartUnavailable(
            f"DART에 재무제표 데이터 없음 (corp {corp_code} / {year} / {reprt_code} / {fs_div} — "
            "비상장 또는 미공시)")
    if status != "000":
        raise DartUnavailable(f"OpenDART 오류 {status}: {js.get('message', '')}")
    items: list[Item] = []
    for it in js.get("list", []):
        try:
            ordv = int(it.get("ord") or 0)
        except ValueError:
            ordv = 0
        items.append(Item(
            sj_div=it.get("sj_div", ""), account_id=it.get("account_id", ""),
            account_nm=it.get("account_nm", ""), ord=ordv,
            thstrm=_to_int(it.get("thstrm_amount")),
            thstrm_add=_to_int(it.get("thstrm_add_amount")),
            frmtrm=_to_int(it.get("frmtrm_amount")),
            frmtrm_q=_to_int(it.get("frmtrm_q_amount")),
            frmtrm_add=_to_int(it.get("frmtrm_add_amount")),
            detail=it.get("account_detail", "") or "",
        ))
    first = js.get("list", [{}])[0] if js.get("list") else {}
    return DartData(corp_code, year, reprt_code, fs_div, items,
                    rcept_no=first.get("rcept_no", ""), thstrm_nm=first.get("thstrm_nm", ""),
                    frmtrm_nm=first.get("frmtrm_nm", ""), source=source)


# ── 라벨 정규화·매칭 ────────────────────────────────────────────────

_ROMAN_RE = re.compile(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVX]+\.?\s*")
_NUMBERING_RE = re.compile(r"^(\(\d+\)|\d+\.|[가-힣]\.|\(?[가-힣]\))\s*")


def norm_label(s: str) -> str:
    """계정명 비교용 정규화: 번호·로마자·공백·구두점 제거, 손익/이익 표기 통일."""
    t = s.strip()
    t = re.sub(r"\(\s*단위\s*[:：][^)]*\)", "", t)   # '기본주당이익 (단위 : 원)'
    t = re.sub(r"\(\s*주\s*\d+\s*\)|\(\*+\)|\*+", "", t)   # '기중 증가액(주1)', '(*)' 각주 표시
    t = _ROMAN_RE.sub("", t)
    t = _NUMBERING_RE.sub("", t)
    t = re.sub(r"[\s　·ㆍ,\.:;'\"\-–—_]+", "", t)
    t = t.replace("및", "")
    t = re.sub(r"\((손실|이익|손익)\)", "", t)      # 순이익(손실) → 순이익
    t = t.replace("총계", "합계") if t.endswith("총계") and t not in ("자본총계", "부채총계", "자산총계") else t
    t = t.replace("당기순", "순").replace("반기순", "순").replace("분기순", "순")
    t = t.replace("당기", "").replace("반기", "").replace("분기", "")
    return t


# 동의어: DSD 라벨 ↔ DART 계정명이 체계적으로 다른 것들 (정규화 후 비교)
_SYNONYMS = {
    "매출액": "수익(매출액)", "매출": "수익(매출액)", "영업수익": "수익(매출액)",
    "기타수익": "기타이익", "기타비용": "기타손실",
    "법인세수익(비용)": "법인세비용(수익)", "법인세비용": "법인세비용(수익)",
    "법인세비용차감전순이익": "법인세비용차감전순이익(손실)",
    "총포괄손익": "총포괄손익", "총포괄이익": "총포괄손익",
    "영업이익": "영업이익(손실)", "영업손실": "영업이익(손실)",
    "순이익": "순이익(손실)", "순손실": "순이익(손실)",
    "기본주당손익": "기본주당이익(손실)", "기본주당이익": "기본주당이익(손실)",
    "희석주당손익": "희석주당이익(손실)", "희석주당이익": "희석주당이익(손실)",
    "매출채권및기타채권": "매출채권및기타유동채권",
    "재고자산": "유동재고자산", "금융비용": "금융원가", "이자비용": "금융원가",
    "매입채무및기타채무": "매입채무및기타유동채무",
    "유형자산": "유형자산", "무형자산": "무형자산",
    "이익잉여금": "이익잉여금(결손금)", "결손금": "이익잉여금(결손금)",
}


def canon(label: str) -> str:
    n = norm_label(label)
    return norm_label(_SYNONYMS.get(n, n))


def match_items(labels: list[str], items: list[Item]) -> dict[int, Item]:
    """문서 행 라벨 목록 → DART 항목. 정규화 라벨 일치(유일할 때만).
    같은 라벨이 여러 번이면 순서대로 짝짓는다(BS의 '기타' 등)."""
    by_key: dict[str, list[Item]] = {}
    for it in sorted(items, key=lambda x: x.ord):
        by_key.setdefault(canon(it.account_nm), []).append(it)
    out: dict[int, Item] = {}
    used: set[int] = set()
    for i, lab in enumerate(labels):
        k = canon(lab)
        if not k:
            continue
        for it in by_key.get(k, []):
            if id(it) not in used:
                out[i] = it
                used.add(id(it))
                break
    return out


def anchor_match(values: list[int | None], items: list[Item],
                 field_name: str, unit: int = 1) -> dict[int, Item]:
    """값 앵커 매칭: 문서 행 값(원 환산) == DART 항목의 특정 필드 값(절대값, 유일)."""
    by_val: dict[int, list[Item]] = {}
    for it in items:
        v = getattr(it, field_name)
        if v is not None and v != 0:
            by_val.setdefault(abs(v), []).append(it)
    out: dict[int, Item] = {}
    for i, v in enumerate(values):
        if v is None or v == 0:
            continue
        cands = by_val.get(abs(int(v) * unit), [])
        # 같은 값의 DART 항목이 여럿(합계=단일 항목 등)이면 라벨 매칭에 맡긴다
        if len(cands) == 1:
            out[i] = cands[0]
    return out
