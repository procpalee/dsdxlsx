"""dart - DART 에 공시된 감사보고서를 DSD 로 만든다(OpenDART 공시서류원본파일).

공시 원문 ZIP(document.xml) 안의 감사보고서 XML 은 DSD 본문(contents.xml)과 같은 dart4.xsd
문서다. 외부감사관련 공시(감사보고서 제출)는 그 자체가, 사업보고서는 첨부
'<접수번호>_00760.xml'(연결은 _00761) 이 감사보고서다. 그 XML 에 meta.xml 을 붙여 DSD 로 묶는다.

원문을 고치는 것은 두 가지뿐이다.
- 느슨한 이스케이프 보정: 공시 원문에 가끔 섞인 맨 '&' 와 태그가 아닌 '<'.
- 줄끝 LF → CRLF: DART 편집기가 쓰는 DSD 는 CRLF 다(--keep-lf 로 끔).
이미지는 원문 ZIP 에 없어 자리(<IMG>파일명</IMG>)만 남는다 - 결과에 목록으로 알린다.

키: 환경변수 OPENDART_API_KEY → 저장소 루트 .env → 실행 폴더 .env (dart_client 와 같다).
캐시: dart_client.cache_dir() - 회사 목록 24시간, 원문 ZIP 은 접수번호별로 계속 둔다(공시는 바뀌지 않음).
"""
from __future__ import annotations

import io
import json
import re
import time
import urllib.error
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from datetime import date

from . import dart_client as dc

FORMS = {"감사": "00760", "연결감사": "00761"}
ACODE_NAMES = {
    "00760": "감사보고서", "00761": "연결감사보고서",
    "00860": "감사전재무제표", "00861": "감사전연결재무제표",
    "90020": "반기검토보고서", "90320": "반기연결검토보고서",
    "90220": "분기검토보고서", "90420": "분기연결검토보고서",
}

_AMP = re.compile(rb"&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)")
_LT = re.compile(rb"<(?![A-Za-z/!?])")
_RCEPT_RE = re.compile(r"(?:rcpNo=|rcept_no=)?(\d{14})")


class DartError(Exception):
    """키 없음·회사·공시를 찾지 못함·원문 없음 - 사용자가 고칠 수 있는 사유."""


# ── HTTP (시험에서 바꿔 끼운다) ──────────────────────────────────────

def _get(endpoint: str, **params) -> bytes:
    key = dc.api_key()
    if not key:
        raise DartError("OpenDART 키가 없다 - 환경변수 OPENDART_API_KEY 나 저장소 루트·실행 폴더의 "
                        ".env 에 OPENDART_API_KEY=... 를 넣으라(키 발급: opendart.fss.or.kr)")
    try:
        return dc._http_get(f"{dc.API_BASE}/{endpoint}", {"crtfc_key": key, **params}, timeout=120)
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        raise DartError(f"OpenDART 호출 실패({endpoint}): {e}") from e


def _json(endpoint: str, **params) -> dict:
    js = json.loads(_get(endpoint, **params).decode("utf-8"))
    st = js.get("status")
    if st == "013":  # 조회된 데이터 없음
        return {"list": [], "total_page": 0}
    if st != "000":
        raise DartError(f"OpenDART {endpoint}: {js.get('message')} (status {st})")
    return js


# ── 회사 찾기 ────────────────────────────────────────────────────────

@dataclass
class Corp:
    code: str
    name: str
    stock: str = ""


def corp_list() -> list[Corp]:
    cached = dc._cache_get("corp_codes.json")
    entries = cached.get("list") if cached else None
    if entries is None:
        raw = _get("corpCode.xml")
        try:
            zf = zipfile.ZipFile(io.BytesIO(raw))
            root = ET.fromstring(zf.read(zf.namelist()[0]))
        except (zipfile.BadZipFile, ET.ParseError) as e:
            raise DartError(f"OpenDART 회사 목록을 읽지 못했다: {raw[:120]!r}") from e
        entries = [{"c": (el.findtext("corp_code") or "").strip(),
                    "n": (el.findtext("corp_name") or "").strip(),
                    "s": (el.findtext("stock_code") or "").strip()}
                   for el in root.iter("list")]
        dc._cache_put("corp_codes.json", {"list": entries})
    return [Corp(e["c"], e["n"], e.get("s", "")) for e in entries]


def find_corp(query: str) -> Corp:
    q = query.strip()
    corps = corp_list()
    if re.fullmatch(r"\d{8}", q):
        hit = [c for c in corps if c.code == q]
    elif re.fullmatch(r"\d{6}", q):
        hit = [c for c in corps if c.stock == q]
    else:
        want = dc._norm_company(q)
        hit = [c for c in corps if dc._norm_company(c.name) == want]
        if not hit:
            part = [c for c in corps if want and want in dc._norm_company(c.name)]
            if part:
                show = ", ".join(f"{c.name}({c.stock or c.code})" for c in part[:10])
                more = f" 외 {len(part) - 10}곳" if len(part) > 10 else ""
                raise DartError(f"'{q}' 와 이름이 같은 회사가 없다 - 비슷한 회사: {show}{more}")
    if not hit:
        raise DartError(f"회사 '{q}' 를 OpenDART 회사 목록에서 찾지 못했다")
    hit.sort(key=lambda c: 0 if c.stock else 1)  # 같은 이름이면 상장사 먼저
    if len(hit) > 1 and not hit[0].stock:
        show = ", ".join(f"{c.name}({c.code})" for c in hit[:10])
        raise DartError(f"'{q}' 이름의 회사가 {len(hit)}곳이다 - 고유번호(8자리)로 고르라: {show}")
    return hit[0]


# ── 공시 찾기 ────────────────────────────────────────────────────────

@dataclass
class Filing:
    rcept_no: str
    rcept_dt: str
    report_nm: str
    corp_code: str
    corp_name: str
    kind: str  # 'F' 외부감사관련 | 'A' 정기공시

    @property
    def period(self) -> str:
        m = re.search(r"\((\d{4})\.(\d{2})\)", self.report_nm)
        return f"{m.group(1)}.{m.group(2)}" if m else ""

    @property
    def year(self) -> int | None:
        m = re.search(r"\((\d{4})\.\d{2}\)", self.report_nm)
        return int(m.group(1)) if m else None

    def to_dict(self) -> dict:
        return {"rcept_no": self.rcept_no, "rcept_dt": self.rcept_dt, "report_nm": self.report_nm,
                "period": self.period, "kind": self.kind,
                "url": f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={self.rcept_no}"}


def _report_matches(report_nm: str, form: str, kind: str) -> bool:
    n = re.sub(r"\s+", "", report_nm)
    if kind == "A":
        return "사업보고서" in n
    if "감사보고서" not in n:
        return False
    return ("연결" in n) == (form == "연결감사")


def filings(corp: Corp, form: str, year: int | None = None) -> list[Filing]:
    """감사보고서가 들어 있을 공시 - 외부감사관련(F) 먼저, 그다음 사업보고서(A). 최신순."""
    today = date.today()
    if year:  # 사업연도 Y 의 감사보고서는 보통 Y+1 년에 공시된다
        bgn, end = date(year, 1, 1), min(date(year + 1, 12, 31), today)
    else:
        bgn, end = date(today.year - 2, 1, 1), today
    out: list[Filing] = []
    for kind in ("F", "A"):
        page = 1
        while True:
            js = _json("list.json", corp_code=corp.code, bgn_de=f"{bgn:%Y%m%d}",
                       end_de=f"{end:%Y%m%d}", pblntf_ty=kind, page_no=page, page_count=100)
            for it in js.get("list", []):
                if _report_matches(it.get("report_nm", ""), form, kind):
                    out.append(Filing(it["rcept_no"], it.get("rcept_dt", ""), it["report_nm"].strip(),
                                      corp.code, it.get("corp_name", corp.name), kind))
            if page >= int(js.get("total_page", 1) or 1):
                break
            page += 1
    if year:
        out = [f for f in out if f.year == year]
    out.sort(key=lambda f: (f.year or 0, f.kind == "F", f.rcept_dt), reverse=True)
    return out


# ── 원문 ─────────────────────────────────────────────────────────────

def document(rcept_no: str) -> dict[str, bytes]:
    """공시 원문 ZIP → {파일명: bytes}. 접수번호별로 캐시(공시는 바뀌지 않는다)."""
    cp = dc.cache_dir() / f"dart_doc_{rcept_no}.zip"
    if cp.is_file():
        raw = cp.read_bytes()
    else:
        raw = _get("document.xml", rcept_no=rcept_no)
        if not raw.startswith(b"PK"):
            m = re.search(rb"<message>([^<]*)</message>", raw)
            msg = m.group(1).decode("utf-8", "replace") if m else raw[:80].decode("utf-8", "replace")
            raise DartError(f"접수번호 {rcept_no} 의 원문을 받지 못했다: {msg}")
        try:
            cp.write_bytes(raw)
        except OSError:
            pass
    z = zipfile.ZipFile(io.BytesIO(raw))
    return {n: z.read(n) for n in z.namelist()}


@dataclass
class Part:
    filename: str
    acode: str
    name: str
    data: bytes


def parts(files: dict[str, bytes]) -> list[Part]:
    """원문 ZIP 의 XML 마다 문서 코드(DOCUMENT-NAME ACODE)와 이름."""
    out = []
    for fn, data in files.items():
        m = re.search(rb'<DOCUMENT-NAME\s+ACODE="(\d+)"[^>]*>([^<]*)<', data[:6000])
        if m:
            out.append(Part(fn, m.group(1).decode(), m.group(2).decode("utf-8", "replace").strip(), data))
    return out


# ── DSD 로 묶기 ──────────────────────────────────────────────────────

@dataclass
class Built:
    contents: bytes
    meta: bytes
    acode: str
    company: str
    corp_code: str
    images: list[str] = field(default_factory=list)
    fixes: dict = field(default_factory=dict)


def _xml_attr(s: str) -> str:
    return s.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;")


def build_dsd(xml: bytes, *, crlf: bool = True) -> Built:
    fixes = {}
    n_amp, n_lt = len(_AMP.findall(xml)), len(_LT.findall(xml))
    data = _LT.sub(b"&lt;", _AMP.sub(b"&amp;", xml))
    if n_amp or n_lt:
        fixes["escaped"] = n_amp + n_lt
    if crlf:
        n_lf = len(re.findall(rb"(?<!\r)\n", data))
        if n_lf:
            data = re.sub(rb"(?<!\r)\n", b"\r\n", data)
            fixes["crlf"] = n_lf
    m = re.search(rb'<DOCUMENT-NAME\s+ACODE="(\d+)"', data)
    if not m:
        raise DartError("감사보고서 XML 에 DOCUMENT-NAME ACODE 가 없다")
    acode = m.group(1).decode()
    cm = re.search(rb'<COMPANY-NAME\b[^>]*?AREGCIK="(\d*)"[^>]*>([^<]*)<', data)
    corp_code = cm.group(1).decode() if cm else ""
    company = cm.group(2).decode("utf-8", "replace").strip() if cm else ""
    images = [x.decode("utf-8", "replace").strip()
              for x in re.findall(rb"<IMG\b[^>]*>([^<]*)</IMG>", data)]
    meta = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<METAINFO>\r\n'
            '\t<GENERATOR schema="dart4.xsd" editver="5.107"/>\r\n'
            f'\t<DOCUMENT-HEADER regcik="{_xml_attr(corp_code)}" regname="{_xml_attr(company)}"/>\r\n'
            f'\t<DOCUMENT-INFO docver="6.0" bsn-id="{acode}" rpt-id="{acode}" doc-id="{acode}" '
            'iscorrection="N"/>\r\n'
            '\t<FILE-INFO date="" size="" rnnumber="" rncheck="" rnimgfile=""/>\r\n'
            "</METAINFO>").encode("utf-8")
    return Built(data, meta, acode, company, corp_code, images, fixes)


def write_dsd(path: str, built: Built) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("contents.xml", built.contents)
        zf.writestr("meta.xml", built.meta)


# ── 찾기부터 묶기까지 ────────────────────────────────────────────────

@dataclass
class Fetched:
    filing: Filing | None
    part: Part
    built: Built
    tried: list[str] = field(default_factory=list)


def fetch(query: str, form: str = "감사", year: int | None = None, *, crlf: bool = True) -> Fetched:
    """회사명·종목코드·고유번호·공시 주소·접수번호 → 감사보고서 DSD."""
    acode = FORMS[form]
    m = _RCEPT_RE.search(query.strip())
    if m and (len(query.strip()) == 14 or "rcpNo" in query or "rcept_no" in query):
        rcept = m.group(1)
        found = [p for p in parts(document(rcept)) if p.acode == acode]
        if not found:
            names = ", ".join(f"{p.name}({p.acode})" for p in parts(document(rcept))) or "없음"
            raise DartError(f"접수번호 {rcept} 원문에 {ACODE_NAMES[acode]}({acode})가 없다 - 들어 있는 문서: {names}")
        return Fetched(None, found[0], build_dsd(found[0].data, crlf=crlf))
    corp = find_corp(query)
    fl = filings(corp, form, year)
    if not fl:
        when = f"{year}년 " if year else "최근 3년 "
        raise DartError(f"{corp.name}: {when}{ACODE_NAMES[acode]} 공시를 찾지 못했다 "
                        "(외부감사관련 공시와 사업보고서를 봤다)")
    tried = []
    for f in fl[:6]:  # 최신 몇 건만 - 원문 받기는 느리다
        hit = [p for p in parts(document(f.rcept_no)) if p.acode == acode]
        if hit:
            return Fetched(f, hit[0], build_dsd(hit[0].data, crlf=crlf), tried)
        tried.append(f"{f.report_nm}({f.rcept_no})")
        time.sleep(0.2)
    raise DartError(f"{corp.name}: 찾은 공시 원문에 {ACODE_NAMES[acode]}({acode})가 없다 - 본 공시: "
                    + ", ".join(tried))
