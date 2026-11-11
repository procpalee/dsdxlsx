"""OpenDART 원천 — 직전 보고서 찾기·원문(document.xml) 내려받기·파싱.

키: 환경변수 OPENDART_API_KEY 또는 .env (dart_client.api_key).
캐시: dart_client.cache_dir() - 원문 zip(dart 명령과 같은 파일)·목록 JSON. 같은 접수번호는 다시 받지 않는다.

DART 공시 원문(공시서류원본파일 API)은 dart4.xsd DOCUMENT XML 이라 같은 파서로 읽힌다.
다만 '&' 가 이스케이프 없이 들어오는 경우가 있어 파싱 전에 sanitize 한다.

식별자는 DSD 안에 있다: meta.xml regcik(=corp_code), DOCUMENT-NAME(보고서 종류),
TU PERIODTO(회계기간 말) — 회사 검색·연도 입력 없이 직전 보고서를 고른다.
"""
from __future__ import annotations

import io
import json
import re
import urllib.error
import zipfile
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from .. import dart_client as dc
from ..parse import parse_contents

_AMP = re.compile(rb"&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)")

# 정기보고서 코드 (fnlttSinglAcntAll reprt_code)
REPRT_ANNUAL, REPRT_HALF, REPRT_Q1, REPRT_Q3 = "11011", "11012", "11013", "11014"


class DartError(Exception):
    pass


class DartUnavailable(DartError):
    """키 없음·비공시사·데이터 없음 — 조용히 skipped 처리할 사유."""


# ── 현재 DSD 식별 ─────────────────────────────────────────────────────

@dataclass
class DocIdent:
    corp_code: str            # regcik 8자리
    company: str
    doc_name: str             # DOCUMENT-NAME 텍스트 (감사보고서 / 반기검토보고서 …)
    consolidated: bool        # 연결 여부
    period_to: date | None    # 회계기간 말
    period_from: date | None

    @property
    def kind(self) -> str:
        """annual | half | q1 | q3 | unknown"""
        n = self.doc_name.replace(" ", "")
        # 중간보고서는 회계기간 말 월이 우선 — 3분기 DSD 를 반기 문서를 뼈대로 만들면
        # 문서명이 '반기검토보고서' 로 남는다(분기 DSD 실측).
        if ("반기" in n or "분기" in n) and self.period_to:
            by_month = {3: "q1", 6: "half", 9: "q3"}.get(self.period_to.month)
            if by_month:
                return by_month
        if "반기" in n:
            return "half"
        if "분기" in n and self.period_to:
            return "q1" if self.period_to.month <= 4 else "q3"
        if "감사보고서" in n or "검토보고서" not in n:
            return "annual"
        return "unknown"

    @property
    def fiscal_year(self) -> int | None:
        return self.period_to.year if self.period_to else None


def _date_of(s: bytes | None) -> date | None:
    if not s:
        return None
    t = s.decode("ascii", "ignore")
    if len(t) == 8 and t.isdigit():
        try:
            return date(int(t[:4]), int(t[4:6]), int(t[6:8]))
        except ValueError:
            return None
    return None


def identify(contents: bytes, meta: bytes | None) -> DocIdent:
    m = re.search(rb'regcik="(\d{8})"', meta or b"") or re.search(rb'AREGCIK="(\d{8})"', contents)
    corp = m.group(1).decode() if m else ""
    m = re.search(rb'regname="([^"]*)"', meta or b"")
    company = m.group(1).decode("utf-8", "replace") if m else ""
    m = re.search(rb"<DOCUMENT-NAME[^>]*>([^<]*)<", contents)
    doc_name = m.group(1).decode("utf-8", "replace").strip() if m else ""
    m = re.search(rb'AUNIT="PERIODTO"[^>]*AUNITVALUE="(\d{8})"', contents)
    pto = _date_of(m.group(1)) if m else None
    m = re.search(rb'AUNIT="PERIODFROM"[^>]*AUNITVALUE="(\d{8})"', contents)
    pfrom = _date_of(m.group(1)) if m else None
    # 연결 여부: 문서명 또는 TITLE '(첨부)연 결 재 무 제 표' (자간 벌림 허용)
    titles = b" ".join(re.findall(rb"<TITLE[^>]*>([^<]*)<", contents))
    titles_n = re.sub(rb"\s+", b"", titles).decode("utf-8", "replace")
    consolidated = "연결" in doc_name or "연결재무제표" in titles_n
    return DocIdent(corp, company, doc_name, consolidated, pto, pfrom)


# ── HTTP ─────────────────────────────────────────────────────────────

class DartClient:
    """OpenDART 호출 + 캐시. fetch(endpoint, params) -> bytes 를 주면 통신을 그것으로 바꾼다(시험용)."""

    def __init__(self, key: str | None = None, cache_dir: Path | None = None, fetch=None):
        self.key = (dc.api_key() or "") if key is None else key.strip()
        self.cache_dir = Path(cache_dir) if cache_dir else dc.cache_dir()
        self._fetch = fetch or self._http

    def require_key(self) -> None:
        if not self.key:
            raise DartUnavailable("OPENDART_API_KEY 미설정 — DART 대사 건너뜀")

    def _http(self, endpoint: str, params: dict) -> bytes:
        try:
            return dc._http_get(f"{dc.API_BASE}/{endpoint}", {"crtfc_key": self.key, **params}, timeout=120)
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            raise DartError(f"OpenDART 호출 실패: {e}") from e

    def _get(self, endpoint: str, **params) -> bytes:
        self.require_key()
        return self._fetch(endpoint, params)

    def _json(self, endpoint: str, **params) -> dict:
        try:
            js = json.loads(self._get(endpoint, **params).decode("utf-8"))
        except ValueError as e:
            raise DartError(f"OpenDART {endpoint}: 응답을 읽지 못함 ({e})") from e
        st = js.get("status")
        if st == "013":
            raise DartUnavailable("조회 데이터 없음")
        if st != "000":
            raise DartError(f"OpenDART {endpoint}: {js.get('message')} (status {st})")
        return js

    # 공시 목록
    def filings(self, corp_code: str, bgn: date, end: date, pblntf_ty: str) -> list[dict]:
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        cp = self.cache_dir / f"dart_list_{corp_code}_{bgn:%Y%m%d}_{end:%Y%m%d}_{pblntf_ty}.json"
        if cp.exists() and end < date.today():
            return json.loads(cp.read_text("utf-8"))
        out: list[dict] = []
        page = 1
        while True:
            try:
                js = self._json("list.json", corp_code=corp_code, bgn_de=f"{bgn:%Y%m%d}",
                                end_de=f"{end:%Y%m%d}", pblntf_ty=pblntf_ty,
                                page_no=page, page_count=100)
            except DartUnavailable:
                break
            out.extend(js.get("list", []))
            if page >= int(js.get("total_page", 1)):
                break
            page += 1
        cp.write_text(json.dumps(out, ensure_ascii=False), "utf-8")
        return out

    # 공시 원문 zip → {파일명: bytes}
    def document(self, rcept_no: str) -> dict[str, bytes]:
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        cp = self.cache_dir / f"dart_doc_{rcept_no}.zip"
        if cp.exists():
            raw = cp.read_bytes()
        else:
            raw = self._get("document.xml", rcept_no=rcept_no)
            if not raw.startswith(b"PK"):
                # 오류 응답은 XML 로 온다
                m = re.search(rb"<message>([^<]*)</message>", raw)
                raise DartUnavailable(f"원문 없음: {m.group(1).decode('utf-8','replace') if m else raw[:80]!r}")
            cp.write_bytes(raw)
        z = zipfile.ZipFile(io.BytesIO(raw))
        return {n: z.read(n) for n in z.namelist()}


# ── 직전 보고서 선택 ─────────────────────────────────────────────────

@dataclass
class PriorDoc:
    label: str          # '사업보고서(2025.12) 첨부 감사보고서' 등 — 사람이 읽는 원천 표시
    rcept_no: str
    rcept_dt: str
    filename: str
    doc: object         # parsed Document
    contents: bytes
    kind: str           # annual | half | q1 | q3
    consolidated: bool


_LT = re.compile(rb"<(?![A-Za-z/!?])")   # 태그가 아닌 '<' (본문의 '<이사ㆍ감사 …>' 등)


def sanitize_xml(data: bytes) -> bytes:
    """DART 뷰어 XML 의 느슨한 이스케이프 보정 — '&' 와 태그 아닌 '<' 만 손본다.
    (문서 바이트를 고쳐 쓰는 용도가 아니라 읽기 전용 파싱용.)"""
    return _LT.sub(b"&lt;", _AMP.sub(b"&amp;", data))


def _doc_kind_of_name(name: str) -> str:
    n = name.replace(" ", "")
    if "반기" in n:
        return "half"
    if "1분기" in n or "분기(03" in n or "2026.03" in n:
        return "q1"
    if "3분기" in n or "분기(09" in n:
        return "q3"
    if "분기" in n:
        return "q"
    return "annual"


def fs_subdoc(doc, consolidated: bool):
    """정기보고서 본문에서 '재무제표'(별도) 또는 '연결재무제표' TITLE 구획만 남긴 Document.
    구획 = 해당 TITLE 부터 다음 번호 TITLE('N. …') 직전까지. 없으면 None."""
    from ..model import Document

    def is_num_title(b) -> bool:
        return b.kind == "label" and b.item.tag == "TITLE" and re.match(r"^\d+\.", b.item.text.strip()) is not None

    start = None
    for i, b in enumerate(doc.blocks):
        if not is_num_title(b):
            continue
        n = re.sub(r"\s+", "", b.item.text)
        n = re.sub(r"^\d+\.", "", n)
        if consolidated and n == "연결재무제표":
            start = i
            break
        if not consolidated and n == "재무제표":
            start = i
            break
    if start is None:
        return None
    end = len(doc.blocks)
    for j in range(start + 1, len(doc.blocks)):
        if is_num_title(doc.blocks[j]):
            end = j
            break
    blocks = doc.blocks[start:end]
    if not any(b.kind == "table" for b in blocks):
        return None
    sub = Document(data=doc.data, blocks=blocks)
    for b in blocks:
        if b.kind == "table":
            sub.tables[b.item.id] = b.item
            for r in b.item.rows:
                sub.rows[r.id] = r
                for c in r.cells:
                    sub.nodes[c.id] = c
        else:
            sub.nodes[b.item.id] = b.item
    return sub


def _attachment_candidates(files: dict[str, bytes], consolidated: bool) -> list[tuple[str, bytes, str]]:
    """zip 안의 감사/검토보고서 XML — (파일명, bytes, 문서명). 연결/별도 구분."""
    out = []
    for name, data in files.items():
        m = re.search(rb"<DOCUMENT-NAME[^>]*>([^<]*)<", data[:4000])
        dn = m.group(1).decode("utf-8", "replace").strip() if m else ""
        if not ("감사보고서" in dn or "검토보고서" in dn):
            continue
        is_con = "연결" in dn
        if is_con != consolidated:
            continue
        out.append((name, data, dn))
    return out


def find_prior_reports(client: DartClient, ident: DocIdent) -> list[PriorDoc]:
    """현재 문서 종류에 따라 대사에 쓸 직전 보고서 원문 1~2개.

    - 연차 감사보고서 → 전기 감사보고서 (사업보고서 첨부 또는 독립 공시)
    - 반기/분기 검토보고서 → ① 전기 감사보고서(BS 전기말) ② 전년 동기 검토보고서(PL·CF·CE)
    """
    client.require_key()
    if not ident.corp_code or not ident.fiscal_year:
        raise DartUnavailable("DSD 에서 회사코드/회계기간을 읽지 못함")
    fy = ident.fiscal_year
    wanted: list[tuple[str, str]] = []   # (kind, 설명)
    if ident.kind == "annual":
        wanted.append(("annual", "전기 감사보고서"))
    elif ident.kind in ("half", "q1", "q3"):
        wanted.append(("annual", "전기 감사보고서"))
        wanted.append((ident.kind, "전년 동기 검토보고서"))
    else:
        raise DartUnavailable(f"보고서 종류를 판별하지 못함: {ident.doc_name!r}")

    # 후보 공시: 전기 사업연도(fy-1)의 정기공시(A) — 접수는 fy 년 3~4월(연차) / fy-1 년(분·반기)
    #            + 외부감사관련(F) 독립 감사보고서 (비상장)
    bgn, end = date(fy - 1, 1, 1), date(fy, 12, 31)
    periodic = client.filings(ident.corp_code, bgn, end, "A")
    audit = client.filings(ident.corp_code, bgn, end, "F")
    out: list[PriorDoc] = []
    for kind, desc in wanted:
        pd = _pick(client, ident, periodic, audit, kind, fy - 1)
        if pd is not None:
            pd.label = f"{desc} — {pd.label}"
            out.append(pd)
    if not out:
        raise DartUnavailable(f"DART 에서 {ident.company or ident.corp_code} 의 전기 보고서를 찾지 못함 "
                              f"(정기공시 {len(periodic)}건 · 외부감사 {len(audit)}건)")
    return out


def _pick(client: DartClient, ident: DocIdent, periodic: list[dict], audit: list[dict],
          kind: str, prior_fy: int) -> PriorDoc | None:
    def period_ok(nm: str) -> bool:
        m = re.search(r"\((\d{4})\.(\d{2})\)", nm)
        if not m:
            return False
        y, mo = int(m.group(1)), int(m.group(2))
        if kind == "annual":
            return y == prior_fy and (ident.period_to is None or mo == ident.period_to.month or True)
        if kind == "half":
            return y == prior_fy and mo in (6, 5, 7)
        if kind == "q1":
            return y == prior_fy and mo in (3, 2, 4)
        if kind == "q3":
            return y == prior_fy and mo in (9, 8, 10)
        return False

    name_kw = {"annual": "사업보고서", "half": "반기보고서", "q1": "분기보고서", "q3": "분기보고서"}[kind]
    # 1) 정기보고서 첨부 — 최신 접수(정정 포함) 우선
    cands = [f for f in periodic if name_kw in f.get("report_nm", "") and period_ok(f.get("report_nm", ""))]
    cands.sort(key=lambda f: f.get("rcept_dt", ""), reverse=True)
    for f in cands:
        try:
            files = client.document(f["rcept_no"])
        except DartUnavailable:
            continue
        att = _attachment_candidates(files, ident.consolidated)
        if att:
            name, data, dn = att[0]
            doc = parse_contents(sanitize_xml(data))
            return PriorDoc(f"{f['report_nm']} 첨부 {dn} (접수 {f['rcept_no']})", f["rcept_no"],
                            f.get("rcept_dt", ""), name, doc, data, kind, ident.consolidated)
        # 첨부 감사·검토보고서가 없는 경우(분·반기보고서 실측) → 본문 'III. 재무에 관한 사항'의
        # 재무제표 구획만 잘라 쓴다 (연결/별도 구획 선택)
        name, data = next(iter(files.items()))
        try:
            main = parse_contents(sanitize_xml(data))
        except Exception:  # noqa: BLE001
            continue
        sub = fs_subdoc(main, ident.consolidated)
        if sub is not None:
            return PriorDoc(f"{f['report_nm']} 본문 {'연결' if ident.consolidated else ''}재무제표 (접수 {f['rcept_no']})",
                            f["rcept_no"], f.get("rcept_dt", ""), name, sub, data, kind, ident.consolidated)
    # 2) 독립 감사보고서 공시 (비상장 외감) — 연차만
    if kind == "annual":
        for f in sorted(audit, key=lambda f: f.get("rcept_dt", ""), reverse=True):
            nm = f.get("report_nm", "")
            if "감사보고서" not in nm or nm.startswith("회계법인"):
                continue
            if ("연결" in nm) != ident.consolidated:
                continue
            # 접수일이 prior_fy+1 년(=현재 fy) 상반기여야 전기 보고서
            dt = f.get("rcept_dt", "")
            if not dt.startswith(str(prior_fy + 1)):
                continue
            try:
                files = client.document(f["rcept_no"])
            except DartUnavailable:
                continue
            att = _attachment_candidates(files, ident.consolidated)
            if not att:
                # 원문 자체가 감사보고서 — DOCUMENT-NAME 없이 올 수 있음
                name, data = next(iter(files.items()))
                att = [(name, data, nm)]
            name, data, dn = att[0]
            try:
                doc = parse_contents(sanitize_xml(data))
            except Exception:  # noqa: BLE001
                continue
            return PriorDoc(f"{nm} (접수 {f['rcept_no']})", f["rcept_no"], dt, name, doc, data,
                            kind, ident.consolidated)
    return None
