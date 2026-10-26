"""문서 보고기간 판정 + 분·반기 전환용 어휘·날짜 치환 규칙.

DSD 한 벌에서 읽는 것:
- DOCUMENT-NAME ACODE — 00760 감사보고서(연차). 검토보고서 코드는 회사가 템플릿을
  재사용해 3분기 파일에 반기 코드(90020)가 남는 실측이 있어 분기 판정에는 쓰지 않는다.
- TU AUNIT="PERIODTO"/"PERIODTO2" AUNITVALUE=YYYYMMDD — 당기 보고기간 종료일.
  분기 = 종료월 (03→1Q, 06→2Q, 09→3Q, 12→연차).
- COMPANY-NAME AREGCIK — DART corp_code(8자리). 회사 검색 없이 API 직접 호출.
- 제목/DOCUMENT-NAME 에 '연결' → 연결재무제표(fs_div=CFS).
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from .model import Document

QUARTER_END = {1: (3, 31), 2: (6, 30), 3: (9, 30), 4: (12, 31)}
# 분기 → OpenDART reprt_code
REPRT_CODE = {1: "11013", 2: "11012", 3: "11014", 4: "11011"}
QUARTER_WORD = {1: "1분기", 2: "반기", 3: "3분기", 4: "연차"}

_DOCNAME_RE = re.compile(rb'<DOCUMENT-NAME\s+ACODE="(\d+)"[^>]*>([^<]*)</DOCUMENT-NAME>')
_CIK_RE = re.compile(rb'<COMPANY-NAME\b[^>]*\bAREGCIK="(\d+)"')
_TU_TO_RE = re.compile(rb'<TU\b[^>]*\bAUNIT="PERIODTO2?"[^>]*\bAUNITVALUE="(\d{8})"')
_TU_ANY_RE = re.compile(rb'<TU\b[^>]*\bAUNITVALUE="(\d{8})"')


@dataclass(frozen=True)
class Period:
    year: int
    quarter: int  # 1..4 (4 = 연차)
    corp_code: str = ""
    consolidated: bool = False
    doc_code: str = ""
    doc_name: str = ""

    @property
    def is_annual(self) -> bool:
        return self.quarter == 4

    @property
    def end(self) -> tuple[int, int, int]:
        m, d = QUARTER_END[self.quarter]
        return (self.year, m, d)

    @property
    def reprt_code(self) -> str:
        return REPRT_CODE[self.quarter]

    @property
    def label(self) -> str:
        return f"{self.year}년 {QUARTER_WORD[self.quarter]}"

    def next_interim(self) -> "Period":
        """다음 중간보고 기간: 1Q→2Q, 2Q→3Q, 3Q→(다음 해)1Q. 연차는 다음 해 1Q."""
        if self.quarter in (1, 2):
            return Period(self.year, self.quarter + 1, self.corp_code,
                          self.consolidated, self.doc_code, self.doc_name)
        return Period(self.year + 1, 1, self.corp_code,
                      self.consolidated, self.doc_code, self.doc_name)

    def with_quarter(self, quarter: int) -> "Period":
        year = self.year + 1 if quarter <= self.quarter else self.year
        return Period(year, quarter, self.corp_code, self.consolidated,
                      self.doc_code, self.doc_name)


def detect_period(doc: Document) -> Period | None:
    """contents.xml 바이트에서 보고기간을 읽는다. TU 종료일이 없으면 None."""
    data = doc.data
    m = _DOCNAME_RE.search(data)
    doc_code = m.group(1).decode() if m else ""
    doc_name = m.group(2).decode("utf-8", "replace").strip() if m else ""
    cik = _CIK_RE.search(data)
    corp = cik.group(1).decode() if cik else ""
    tos = [x.group(1).decode() for x in _TU_TO_RE.finditer(data)]
    if not tos:
        # PERIODTO 라벨이 없는 문서 — 8자리 TU 값 중 최대(종료일)로 추정
        tos = [x.group(1).decode() for x in _TU_ANY_RE.finditer(data)]
    if not tos:
        return None
    end = max(tos)
    year, month = int(end[:4]), int(end[4:6])
    quarter = {3: 1, 6: 2, 9: 3, 12: 4}.get(month)
    if quarter is None:
        # 비12월 결산·이례적 종료월 — 3개월 단위로 근사
        quarter = max(1, min(4, (month + 2) // 3))
    if doc_code == "00760":
        quarter = 4
    # 별도 보고서 본문에도 '연결재무제표' 언급은 흔하므로 재무상태표 제목으로만 판정
    consolidated = "연결" in doc_name or bool(
        re.search(r"연\s*결\s*(?:반\s*기\s*|분\s*기\s*)?재\s*무\s*상\s*태\s*표",
                  data.decode("utf-8", "replace")))
    return Period(year, quarter, corp, consolidated, doc_code, doc_name)


def parse_target(text: str | None, base: Period) -> Period:
    """--target '1Q|2Q|3Q|2026Q2' 파싱. 없으면 base.next_interim()."""
    if not text:
        return base.next_interim()
    m = re.fullmatch(r"(?:(\d{4})\s*[Qq]?\s*)?([123])\s*(?:[Qq]|분기)?", text.strip())
    if not m:
        raise ValueError(f"--target 형식 오류: {text!r} (예: 2Q, 2026Q3)")
    q = int(m.group(2))
    if m.group(1):
        return Period(int(m.group(1)), q, base.corp_code, base.consolidated,
                      base.doc_code, base.doc_name)
    return base.with_quarter(q)


# ── 어휘 치환 ────────────────────────────────────────────────────────

# 분·반기 어휘 표. 순서 중요(긴 것 먼저). 각 규칙 = (정규식, 치환).
# 공백이 벌어진 자간('반 기 재 무 상 태 표')은 별도 처리한다.
def vocab_rules(src_q: int, dst_q: int) -> list[tuple[re.Pattern, str]]:
    if src_q == dst_q:
        return []
    rules: list[tuple[re.Pattern, str]] = []
    if src_q == 1 and dst_q == 2:      # 1분기 → 반기
        rules += [
            (re.compile(r"1\s*분기"), "반기"),
            (re.compile(r"당\s*분\s*기"), "당반기"),
            (re.compile(r"전\s*분\s*기"), "전반기"),
            # 나머지 '분기'(분기순이익·분기재무상태표·분기검토보고서…) → '반기'
            # 예외: 숫자 뒤(2분기·4분기), '분기별'·'매분기'
            (re.compile(r"(?<![0-9매])분\s*기(?!별)"), "반기"),
        ]
    elif src_q == 2 and dst_q == 3:    # 반기 → 3분기
        rules += [
            (re.compile(r"당\s*반\s*기"), "당분기"),
            (re.compile(r"전\s*반\s*기"), "전분기"),
            (re.compile(r"6\s*개월"), "9개월"),
            # 독립어 '반기'(제 60(당) 반기 / 반기말 / 반기 2026년…) → '3분기'
            # ('반기 총포괄이익'처럼 뒤에 한글 낱말이 이어지면 복합어로 본다)
            (re.compile(r"반\s*기(?=\s*말?(?:$|[\d:(),.\-/]|\s+(?:$|[\d(])))"), "3분기"),
            # 복합어 '반기재무제표·반기순이익·반기검토보고서…' → '분기…'
            (re.compile(r"반\s*기(?=\s*[가-힣])"), "분기"),
        ]
    elif src_q == 3 and dst_q == 1:    # 3분기 → 1분기
        rules += [
            (re.compile(r"3\s*분기"), "1분기"),
            (re.compile(r"9\s*개월"), "3개월"),
        ]
    else:
        raise ValueError(f"지원하지 않는 전환: {QUARTER_WORD[src_q]} → {QUARTER_WORD[dst_q]}")
    return rules


_SPACED_HANGUL_RE = re.compile(r"(?:[가-힣]\s){3,}[가-힣]")


# 치환하면 안 되는 낱말 - '일반기업회계기준'의 '반기', '구분기준'의 '분기' (2026-10-23 3분기 DSD 실측).
# 회사마다 더 있으면 PROTECT_WORDS 에 더하거나 apply_vocab(…, protect=[…]).
PROTECT_WORDS = ["일반기업", "일반기준", "일반적", "일반", "구분기준", "구분기", "기분기"]


def _sub_all(text: str, rules: list[tuple[re.Pattern, str]], protect=None) -> str:
    words = sorted(set(PROTECT_WORDS) | set(protect or []), key=len, reverse=True)
    keep: list[str] = []

    def hide(m):
        keep.append(m.group(0))
        return "\x00%d\x00" % (len(keep) - 1)
    if words:
        text = re.sub("|".join(re.escape(w) for w in words), hide, text)
    for pat, rep in rules:
        text = pat.sub(rep, text)
    return re.sub(r"\x00(\d+)\x00", lambda m: keep[int(m.group(1))], text)


def apply_vocab(text: str, rules: list[tuple[re.Pattern, str]], protect=None) -> str:
    """어휘 규칙 적용. 자간 벌린 구간('반 기 재 무 상 태 표')은 공백을 접어
    규칙을 태운 뒤 같은 자간(한 글자마다 공백)으로 되돌린다. protect = 치환하지 않을 낱말 더하기."""
    if not rules or not text:
        return text
    out: list[str] = []
    pos = 0
    for m in _SPACED_HANGUL_RE.finditer(text):
        out.append(_sub_all(text[pos:m.start()], rules, protect))
        run = m.group(0)
        folded = re.sub(r"\s+", "", run)
        new = _sub_all(folded, rules, protect)
        out.append(run if new == folded else " ".join(new))
        pos = m.end()
    out.append(_sub_all(text[pos:], rules, protect))
    return "".join(out)


# ── 날짜 치환 ────────────────────────────────────────────────────────

_DATE_FORMS = [
    # (정규식, 포매터)  그룹: y, m, d
    (re.compile(r"(?<!\d)(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일"), "kor"),
    (re.compile(r"(?<!\d)(\d{4})\s*[.\-/]\s*(\d{1,2})\s*[.\-/]\s*(\d{1,2})(?![\d])"), "sym"),
]


def date_map(base: Period, target: Period) -> dict[tuple[int, int, int], tuple[int, int, int]]:
    """골격 문서의 기간 날짜 → 목표 문서 날짜.

    같은 해 전환(1Q→2Q, 2Q→3Q): 당기말·전년 동기말만 분기말로 이동, 전기말(12.31) 유지.
    해 넘김(3Q→1Q): 당기말→(y+1).3.31, 전년 동기말→y.3.31, 전기말→(y).12.31,
    당기초 1.1 → (y+1).1.1, 전기초 → y.1.1.
    """
    by, bm, bd = base.end
    ty, tm, td = target.end
    mapping = {
        (by, bm, bd): (ty, tm, td),            # 당기말
        (by - 1, bm, bd): (ty - 1, tm, td),    # 전년 동기말
    }
    if ty != by:
        mapping[(by - 1, 12, 31)] = (ty - 1, 12, 31)   # 전기말
        mapping[(by, 1, 1)] = (ty, 1, 1)               # 당기초
        mapping[(by - 1, 1, 1)] = (ty - 1, 1, 1)       # 전기초
    return mapping


def apply_dates(text: str, mapping: dict) -> str:
    if not text or not mapping:
        return text

    def kor(m: re.Match) -> str:
        key = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if key not in mapping:
            return m.group(0)
        y, mo, d = mapping[key]
        # 자릿수 관례 유지(03월 ↔ 3월)
        ms = f"{mo:02d}" if len(m.group(2)) == 2 else str(mo)
        ds = f"{d:02d}" if len(m.group(3)) == 2 else str(d)
        s = m.group(0)
        s = re.sub(r"^\d{4}", str(y), s)
        s = re.sub(r"(년\s*)\d{1,2}(\s*월)", lambda x: f"{x.group(1)}{ms}{x.group(2)}", s, count=1)
        s = re.sub(r"(월\s*)\d{1,2}(\s*일)", lambda x: f"{x.group(1)}{ds}{x.group(2)}", s, count=1)
        return s

    def sym(m: re.Match) -> str:
        key = (int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if key not in mapping:
            return m.group(0)
        y, mo, d = mapping[key]
        ms = f"{mo:02d}" if len(m.group(2)) == 2 else str(mo)
        ds = f"{d:02d}" if len(m.group(3)) == 2 else str(d)
        s = m.group(0)
        parts = re.split(r"([.\-/])", s)
        # parts: y, sep, m, sep, d  (공백 포함 가능)
        parts[0] = re.sub(r"\d{4}", str(y), parts[0])
        parts[2] = re.sub(r"\d{1,2}", ms, parts[2])
        parts[4] = re.sub(r"\d{1,2}", ds, parts[4])
        return "".join(parts)

    out = text
    for pat, kind in _DATE_FORMS:
        out = pat.sub(kor if kind == "kor" else sym, out)
    return out


def tu_value_map(mapping: dict) -> dict[bytes, bytes]:
    """TU AUNITVALUE(YYYYMMDD) 치환표."""
    return {f"{y}{m:02d}{d:02d}".encode(): f"{ty}{tm:02d}{td:02d}".encode()
            for (y, m, d), (ty, tm, td) in mapping.items()}


def month_word_map(base: Period, target: Period) -> list[tuple[re.Pattern, str]]:
    """'6개월'·'9개월' 등 누적 개월 표기(검토보고서 문단용 — 참고 목록 생성에만 사용)."""
    bm = base.quarter * 3
    tm = target.quarter * 3
    if bm == tm:
        return []
    return [(re.compile(rf"(?<!\d){bm}\s*개월"), f"{tm}개월")]
