"""dsdxlsx check — 규칙 기반 정합성 검증.

C1 AUTOSUM 검산: COLALL/COLPART/ROWALL/ROWPART 셀을 다중 후보로 재계산
C2 그룹 footing: 들여쓰기 계층(Ⅰ.유동자산 = 하위 계정 합) — 손익계산서류는 제외
C3 합계행 footing: 합계/소계/…총계 행의 열 합산 검산
C4 재무제표↔주석 대사: FS 주석참조 열(6,28,31)의 금액이 해당 주석에 존재하는지
C5 주석↔주석 대사: 같은 (계정 라벨, 기간)이 주석 간 다른 금액이면 후보 플래그
C6 표기 검출: 깨진 숫자·공백 낀 숫자·괄호 짝·단위 혼용

검산의 핵심 규칙(샘플 실측): COLALL 합계는 위쪽의 COLPART 소계 값을 더하고
그 소계가 커버하는 구간(직전 섹션 행까지)은 건너뛴다. 소계가 수기(마크 없음)인
표도 같은 구조이므로 합계/소계 라벨 행도 동일 취급하는 '스마트 워크'로 계산한다.

모든 결과는 자문(advisory) — DART 편집기가 본문 합계를 자동 재계산하지 않으므로
(실기 확인) 수정 후에는 반드시 check로 검산한다.
"""
from __future__ import annotations

import re
from contextlib import contextmanager
from dataclasses import dataclass, field
from decimal import Decimal

from . import grid as grid_mod
from .export_xlsx import STATEMENT_KEYWORDS, refined_sections
from .model import Document, Table
from .profile import Profile, default_profile

# 활성 프로파일 — 키워드·정규식·허용오차의 단일 소스 (기본값 = data/_default.toml).
# run_checks(profile=)가 검사 동안 교체한다. CLI 단일 스레드 전제.
_P: Profile = default_profile()


@contextmanager
def use_profile(profile: Profile | None):
    """TableView·검사 헬퍼가 읽는 활성 프로파일을 잠시 교체 (ties 등 외부 모듈용)."""
    global _P
    prev = _P
    _P = profile or default_profile()
    try:
        yield _P
    finally:
        _P = prev


@dataclass
class Finding:
    check: str   # C1~C6
    level: str   # error | warn | info
    where: str
    message: str


@dataclass
class Result:
    """검사한 셀 단위의 구조화 결과 — 엑셀 리포트용 (일치 항목 포함)."""

    check: str
    context: str        # 시트/주석 구역
    table_id: str
    row_id: str         # t014.r004 — 편집용 xlsx 셀 위치 매핑 키
    row_label: str
    col_header: str
    grid_col: int       # 편집용 xlsx 열 = 2 + grid_col
    stated: object      # 기재값
    computed: object    # 계산값 (대사류는 None)
    ok: bool
    note: str = ""
    # 검산 구성 셀 [(row_id, grid_col, 부호)] — 엑셀 수식 생성용 (대사류는 빈 목록)
    terms: list = field(default_factory=list)
    refs: list = field(default_factory=list)  # C4: 참조 주석 번호
    axis: str = "col"  # 'col' 세로(열 합산 — 표 아래 표시) | 'row' 가로(행 합산 — 오른쪽 표시)


@dataclass
class CheckReport:
    findings: list[Finding] = field(default_factory=list)
    results: list[Result] = field(default_factory=list)
    counts: dict = field(default_factory=dict)  # check -> 검사 수
    # 프로파일 suppress 규칙으로 억제된 발견 사항 [(Finding, 사유)] — 숨기지 않고 집계
    suppressed: list[tuple[Finding, str]] = field(default_factory=list)
    profile_name: str = "_default"
    # 표별 검사 커버리지 [{table, context, checks, skipped}] — 침묵 통과 방지
    coverage: list[dict] = field(default_factory=list)

    def add(self, check: str, level: str, where: str, message: str):
        self.findings.append(Finding(check, level, where, message))

    def result(self, check, tv, i, c, stated, computed, ok, note="", terms=None, refs=None,
               axis="col"):
        self.results.append(Result(
            check=check,
            context=tv.context,
            table_id=tv.table.id,
            row_id=tv.table.rows[i].id,
            row_label=tv.row_label(i)[:30],
            col_header=tv.col_header(c)[:30] if c is not None else "",
            grid_col=c if c is not None else 0,
            stated=stated,
            computed=computed,
            ok=ok,
            note=note,
            terms=[(tv.table.rows[j].id, gc, sign) for j, gc, sign in (terms or [])],
            refs=refs or [],
            axis=axis,
        ))

    def tick(self, check: str):
        self.counts[check] = self.counts.get(check, 0) + 1


# ── 값/라벨 유틸 ──────────────────────────────────────────────────────

_NUM_RE = re.compile(r"^\d{1,3}(?:,\d{3})*(?:\.\d+)?$")
# 기간 버킷 보조 정규식 — '당3분기'·'전 1분기'처럼 키워드 목록에 없는 숫자 접두 변형
# (dsdfoot constants.PERIOD_*_RE 와 동일 규칙)
_PERIOD_CUR_RE = re.compile(r"당\d?(반|분)?기|\(당\)")
_PERIOD_PRV_RE = re.compile(r"전\d?(반|분)?기|\(전\)")


def cell_value(text: str):
    """검산용 값: 숫자 → int/Decimal, '-' → 0, 접미 단위('2,978,480주', '62.12%')는
    단위를 뗀 수치, 그 외 None."""
    s = text.strip()
    if s == "-":
        return 0
    from . import numbers as _numbers

    return _numbers.classify(s)


def norm(text: str) -> str:
    return re.sub(r"\s+", "", text)


def fmt(value) -> str:
    return f"({-value:,})" if value < 0 else f"{value:,}"


def is_total_label(label: str) -> bool:
    n = norm(label)
    return bool(_P.total_re.match(n) or _P.grand_total_re.search(n))


class TableView:
    """그리드 기반 값/라벨/헤더 접근."""

    def __init__(self, table: Table, context: str, sec_title: str = ""):
        self.table = table
        self.context = context
        self.sec_title = sec_title
        self.grid = grid_mod.expand(table)
        self.n = len(table.rows)
        self.ncols = table.ncols
        self.body = [i for i, r in enumerate(table.rows) if not r.in_thead]
        self.head = [i for i, r in enumerate(table.rows) if r.in_thead]

    def cell(self, i: int, c: int):
        gc = self.grid[i][c]
        return gc.node if gc is not None and gc.kind == "cell" else None

    def text(self, i: int, c: int) -> str:
        gc = self.grid[i][c]
        return gc.node.text if gc is not None else ""

    def value(self, i: int, c: int):
        node = self.cell(i, c)
        return cell_value(node.text) if node else None

    def row_label(self, i: int) -> str:
        for c in range(self.ncols):
            t = self.text(i, c).strip()
            if t:
                return t
        return ""

    def row_is_total(self, i: int) -> bool:
        """합계류 행 판정 — 첫 열이 '구분'이고 둘째 열에 '소계'가 오는 양식이 있어
        모든 텍스트 셀을 본다."""
        for c in range(self.ncols):
            t = norm(self.text(i, c))
            if t and self.value(i, c) is None and (_P.total_re.match(t)
                                                   or _P.grand_total_re.search(t)):
                return True
        return False

    @property
    def unit(self) -> int:
        """표 단위 배수 (원=1, 천원=1000, …). 헤더에 없으면 1로 간주."""
        if not hasattr(self, "_unit"):
            self._unit = 1
            for i in self.head:
                for c in range(self.ncols):
                    m = _P.unit_re.search(self.text(i, c))
                    if m:
                        u = _P.unit_mult.get(norm(m.group(1)).split(",")[-1])
                        if u:
                            self._unit = u
                        break
        return self._unit

    def label_depth(self, i: int) -> int:
        # K-GAAP 표준 셀(TE)은 ALEVEL 속성이 계층의 정답 — 들여쓰기 추측 불필요
        cell = self.cell(i, 0)
        if cell is not None and "ALEVEL" in cell.attrs:
            try:
                return int(cell.attrs["ALEVEL"])
            except ValueError:
                pass
        raw = self.text(i, 0)
        return len(raw) - len(raw.lstrip(" "))

    def is_section_row(self, i: int) -> bool:
        """값 없는 라벨 행 (자산 / 유동자산 : 등) — 구간 경계."""
        has_value = any(self.value(i, c) is not None for c in range(self.ncols))
        return not has_value and bool(self.row_label(i))

    def col_header(self, c: int) -> str:
        parts = []
        for i in self.head:
            gc = self.grid[i][c]
            if gc is not None:
                t = gc.node.text.strip()
                if t and t not in parts and not t.startswith("(단위"):
                    parts.append(t)
        return "/".join(parts)

    def _gisu_ranks(self) -> dict[int, str]:
        """표 헤더의 기수 집합 → {기수: 버킷}. 최대 기수=당기, 둘째=전기.
        ('제 26 기 1분기말 / 제 25 기말'처럼 (당)/(전) 표기가 없는 회사 대응)"""
        if not hasattr(self, "_gisu"):
            nums = set()
            for c in range(self.ncols):
                m = re.search(r"제(\d+)기", norm(self.col_header(c)))
                if m:
                    nums.add(int(m.group(1)))
            ranked = sorted(nums, reverse=True)
            self._gisu = ({ranked[0]: "당기", ranked[1]: "전기"}
                          if len(ranked) >= 2 else {})
        return self._gisu

    def period_bucket(self, c: int) -> str | None:
        h = norm(self.col_header(c))
        if any(k in h for k in _P.period_current) or _PERIOD_CUR_RE.search(h):
            return "당기"
        if any(k in h for k in _P.period_prior) or _PERIOD_PRV_RE.search(h):
            return "전기"
        m = re.search(r"제(\d+)기", h)
        if m:
            return self._gisu_ranks().get(int(m.group(1)))
        return None

    def where(self, i: int, c: int | None = None) -> str:
        label = self.row_label(i)[:24]
        loc = f"[{self.context}] {self.table.id} '{label}'"
        if c is not None:
            h = self.col_header(c)
            if h:
                loc += f" × '{h[:20]}'"
        return loc


def build_views(doc: Document) -> tuple[list[TableView], dict]:
    """refined_sections 기준으로 각 표에 컨텍스트(시트명/주석번호) 부여."""
    views: list[TableView] = []
    note_tables: dict[int, list[TableView]] = {}
    statement_views: list[TableView] = []
    for sec in refined_sections(doc):
        label = f"주석{sec.note_no} {sec.note_title[:12]}" if sec.note_no else sec.title
        for block in sec.blocks:
            if block.kind != "table":
                continue
            tv = TableView(block.item, label, sec_title=sec.title)
            views.append(tv)
            if sec.note_no is not None:
                note_tables.setdefault(sec.note_no, []).append(tv)
            if sec.title in STATEMENT_KEYWORDS:
                statement_views.append(tv)
    return views, {"notes": note_tables, "statements": statement_views}


# ── 열 합산 워크 (C1·C3 공용) ────────────────────────────────────────
# 후보 = (합, 구성 항 목록). 항 = (행 인덱스, grid 열, 부호) — 엑셀 수식 생성에 쓰인다.

def _is_col_autosum(node) -> bool:
    return node is not None and node.attrs.get("AUTOSUM", "").startswith("COL")


# 흐름표의 결과 행('기말금액' 등)은 소계 역할 — 차감후 금액 = 기말금액 + 차감항목

def _row_is_subtotal(tv: TableView, j: int, c: int) -> bool:
    return (_is_col_autosum(tv.cell(j, c)) or tv.row_is_total(j)
            or bool(_P.flow_end_re.match(norm(tv.row_label(j)))))


def _smart_col_terms(tv: TableView, i: int, c: int) -> list:
    """소계(자동합계 마크 또는 합계류 라벨)를 만나면 그 값을 항으로 넣고,
    소계가 커버하는 구간(직전 섹션 행까지)은 건너뛰는 워크."""
    terms, skipping = [], False
    for j in reversed([b for b in tv.body if b < i]):
        if tv.is_section_row(j) and not tv.row_is_total(j):
            skipping = False
            continue
        node = tv.cell(j, c)
        v = tv.value(j, c)
        subtotal = _row_is_subtotal(tv, j, c)
        if subtotal:
            if v is not None:
                terms.append((v, j))
            skipping = True
            continue
        if v is not None and not skipping:
            terms.append((v, j))
    return terms


def _col_candidates(tv: TableView, i: int, c: int) -> tuple[list, int]:
    """반환: (후보 목록 [(합, [(행,열,부호)])], 비합계 기여 행 수)."""
    contiguous, no_subtotal, totals, same_depth, to_prev = [], [], [], [], []
    depth_i = tv.label_depth(i)
    contiguous_open = True
    autosum_seen = False
    for j in reversed([b for b in tv.body if b < i]):
        v = tv.value(j, c)
        if tv.is_section_row(j):
            contiguous_open = False
            continue
        if v is None:
            continue
        is_autosum = _is_col_autosum(tv.cell(j, c))
        total_row = _row_is_subtotal(tv, j, c)
        if is_autosum:
            autosum_seen = True
        if total_row:
            totals.append((v, j))
            contiguous_open = False
            continue
        if contiguous_open:
            contiguous.append((v, j))
            if tv.label_depth(j) == depth_i:
                same_depth.append((v, j))
        no_subtotal.append((v, j))
        if not autosum_seen:
            to_prev.append((v, j))

    def cand(pairs):
        return (sum(v for v, _ in pairs), [(j, c, 1) for _, j in pairs])

    # 직전 소계까지 구간(소계 + 이후 차감 항목): 차감후 잔액 = 소계 − 유동성 대체
    since_sub = []
    for j in reversed([b for b in tv.body if b < i]):
        if tv.is_section_row(j) and not tv.row_is_total(j):
            break
        v = tv.value(j, c)
        if v is not None:
            since_sub.append((v, j))
        if _row_is_subtotal(tv, j, c):
            break

    cands = [cand(_smart_col_terms(tv, i, c)), cand(contiguous), cand(no_subtotal),
             cand(same_depth), cand(to_prev), cand(_group_block_terms(tv, i, c)),
             cand(_same_key_terms(tv, i, c)), cand(_section_echo_terms(tv, i, c)),
             cand(since_sub)]
    # 합계류 행들의 접두 합 (부채및자본총계 = 자본총계+부채총계 패턴)
    for k in range(1, len(totals) + 1):
        cands.append(cand(totals[:k]))
    return cands, len(no_subtotal)


def _section_echo_terms(tv: TableView, i: int, c: int) -> list:
    """섹션 라벨과 같은 이름의 소계 행('<제품매출원가>' 섹션 끝의 '제품매출원가')이
    섹션 항목들을 대체하는 후보 — 합계 = 에코 소계들의 합."""
    collected: list[tuple] = []
    current: tuple | None = None  # (섹션 코어 라벨, collected 시작 위치)
    for j in tv.body:
        if j >= i:
            break
        if tv.is_section_row(j):
            core = norm(tv.row_label(j)).strip("<>:：")
            current = (core, len(collected))
            continue
        v = tv.value(j, c)
        if v is None:
            continue
        if current and norm(tv.row_label(j)) == current[0]:
            collected = collected[:current[1]]
            collected.append((v, j))
            current = None
        else:
            collected.append((v, j))
    return collected


def _same_key_terms(tv: TableView, i: int, c: int) -> list:
    """'구분' 열 소계: 같은 구분(첫 열 라벨) 블록 안의 위쪽 행만 합산.
    (매입채무 표: '기타금융부채(유동) | 소계' = 같은 구분의 미지급금+미지급비용)"""
    key = norm(tv.text(i, 0))
    if not key:
        return []
    terms = []
    for j in reversed([b for b in tv.body if b < i]):
        if norm(tv.text(j, 0)) != key:
            break
        if tv.row_is_total(j) or _is_col_autosum(tv.cell(j, c)):
            break
        v = tv.value(j, c)
        if v is not None:
            terms.append((v, j))
    return terms


def _group_block_terms(tv: TableView, i: int, c: int) -> list:
    """'구분' 열 블록 단위 합산 후보: 첫 열 라벨이 같은 연속 블록마다, 블록 안에
    소계 행이 있으면 소계만·없으면 항목 전부. (매입채무 표: 합계 = 매입채무 +
    기타금융부채 소계 — 소계가 둘째 열에 있는 양식 대응)"""
    blocks: list[list[int]] = []
    prev_key = None
    for j in tv.body:
        if j >= i:
            break
        if tv.value(j, c) is None:
            continue
        key = norm(tv.text(j, 0))
        if key != prev_key or not blocks:
            blocks.append([])
            prev_key = key
        blocks[-1].append(j)
    terms = []
    for block in blocks:
        subtotal = [j for j in block if tv.row_is_total(j) or _is_col_autosum(tv.cell(j, c))]
        chosen = subtotal[-1:] if subtotal else block
        terms.extend((tv.value(j, c), j) for j in chosen)
    return terms


def _row_candidates(tv: TableView, i: int, c: int) -> list:
    contiguous, all_left = [], []
    stopped = False
    for k in range(c - 1, -1, -1):
        node = tv.cell(i, k)
        if node is not None and node.attrs.get("AUTOSUM", "").startswith("ROW"):
            break
        v = tv.value(i, k)
        if v is None:
            if node is not None and node.text.strip():
                stopped = True  # 라벨/비숫자 셀
            continue
        if not stopped:
            contiguous.append((v, k))
        all_left.append((v, k))

    def cand(pairs):
        return (sum(v for v, _ in pairs), [(i, k, 1) for _, k in pairs])

    return [cand(contiguous), cand(all_left)]


def _pick(cands: list, stated):
    """일치 후보 우선, 없으면 최근접 후보. 반환: (계산값, 항 목록, 일치 여부).
    단수차(기본 ±1)는 일치로 인정 (천원 단위 문서의 원단위 절사 관행)."""
    for total, terms in cands:
        if total == stated:
            return total, terms, True
    total, terms = min(cands, key=lambda x: abs(x[0] - stated))
    if abs(total - stated) <= _P.footing_tol:
        return total, terms, True
    return total, terms, False


# 재무 검산 대상이 아닌 구획 (감사시간 표의 합계는 단순합이 아님, 표지 목차 페이지 등)

def _skip_section(tv: TableView) -> bool:
    return bool(_P.skip_section_re.search(norm(tv.sec_title)))


# ── C1: AUTOSUM 검산 ─────────────────────────────────────────────────

def check_autosum(views: list[TableView], report: CheckReport):
    for tv in views:
        for i in tv.body:
            for c in range(tv.ncols):
                node = tv.cell(i, c)
                if node is None or "AUTOSUM" not in node.attrs:
                    continue
                stated = cell_value(node.text)
                if stated is None:
                    continue
                report.tick("C1")
                kind = node.attrs["AUTOSUM"]
                if kind.startswith("COL"):
                    cands, _ = _col_candidates(tv, i, c)
                else:
                    cands = _row_candidates(tv, i, c)
                computed, terms, ok = _pick(cands, stated)
                report.result("C1", tv, i, c, stated, computed, ok, kind, terms=terms,
                              axis="row" if kind.startswith("ROW") else "col")
                if not ok:
                    report.add(
                        "C1", "warn", tv.where(i, c),
                        f"자동합계({kind}) 기재 {fmt(stated)} ≠ 계산 {fmt(computed)} "
                        f"(차이 {fmt(stated - computed)})",
                    )


# ── C2: 들여쓰기 그룹 footing ─────────────────────────────────────────

def _roman_group_footing(tv: TableView, report: CheckReport):
    """들여쓰기가 없는 표의 폴백: 로마자 접두(Ⅰ. Ⅱ.) 행 = 다음 로마자/합계류 행
    전까지의 값 행 합. ⚠ 현 샘플에는 들여쓰기가 있어 미발동 — 무들여쓰기 샘플로
    캘리브레이션 필요."""
    roman_re = re.compile(rf"^[{_ROMAN}]+\s*\.")
    body = tv.body
    anchors = [i for i in body if roman_re.match(tv.row_label(i))]
    for k, i in enumerate(anchors):
        end = anchors[k + 1] if k + 1 < len(anchors) else None
        children = []
        for j in body:
            if j <= i or (end is not None and j >= end):
                continue
            if is_total_label(tv.row_label(j)):
                break
            children.append(j)
        for c in range(tv.ncols):
            gv = tv.value(i, c)
            if gv is None:
                continue
            terms = [(j, c, 1) for j in children if tv.value(j, c) is not None]
            if len(terms) < 2:
                continue
            report.tick("C2")
            computed = sum(tv.value(j, cc) for j, cc, _ in terms)
            ok = computed == gv
            report.result("C2", tv, i, c, gv, computed, ok,
                          f"로마자 그룹 하위 {len(terms)}개", terms=terms)
            if not ok:
                report.add(
                    "C2", "warn", tv.where(i, c),
                    f"그룹 합계 기재 {fmt(gv)} ≠ 하위 {len(terms)}개 합 "
                    f"{fmt(computed)} (차이 {fmt(gv - computed)})",
                )


# ── C2b: 무들여쓰기 총계-먼저(flat totals-first) footing ──────────────
# 흥구형 IFRS 재무상태표: 들여쓰기·로마자 없이 '유동자산(합계) → 구성항목'
# 순서로 평평하게 나열. 계층이 라벨에 없으므로 **정확한 합 관계 자체**로 트리를
# 재귀 분해한다(12자리 금액의 우연 일치는 사실상 없음 — C1 다중 후보와 동일 논거).
#
# 핵심: 구조 추론과 검증을 분리한다. 편집으로 합이 깨진 열은 분해가 실패하므로,
# 합이 맞는 다른 열(예: 전기)에서 확정한 구조를 그 열에 적용해 차이를 보고한다.
# 어떤 열도 분해되지 않으면 침묵하지 않고 info로 표면화.

_PARSE_BUDGET = 500_000


def _parse_trees(vals: list, i: int, hi: int, target, budget: list, dead: set):
    """vals[i:] 접두를 target으로 합산하는 모든 블록열 분해를 yield: (j, nodes).
    블록 = 잎(행 하나) 또는 부모(행 + 뒤따르는 자식 블록들이 그 값으로 합산).
    첫 해가 부분 소비일 수 있으므로 단일 결과가 아니라 열거한다 — 세그먼트
    파서(호출부)는 전량 소비(j == hi) 해만 채택.

    dead = 해가 하나도 없음이 확정된 (위치, 잔여목표) 상태 — 음수 행 때문에
    합 범위로 가지치기할 수 없으므로 이 메모가 지수 폭주를 막는 핵심.
    budget은 최후 방어선(소진 시 분해 실패 → info 표면화)."""
    if budget[0] <= 0 or (i, target) in dead:
        return
    budget[0] -= 1
    produced = False
    if target == 0:
        produced = True
        yield (i, [])  # 목표 달성 후보 — 0값 행 흡수는 아래 잎 탐색이 계속
    if i < hi:
        _, v = vals[i]
        for j, nodes in _parse_trees(vals, i + 1, hi, target - v, budget, dead):
            produced = True
            yield (j, [(i, [])] + nodes)
        if v != 0:  # 부모: 자식 블록들이 정확히 v로 합산
            for k, children in _parse_trees(vals, i + 1, hi, v, budget, dead):
                if not children:
                    continue
                for j, nodes in _parse_trees(vals, k, hi, target - v, budget, dead):
                    produced = True
                    yield (j, [(i, children)] + nodes)
    if not produced and budget[0] > 0:
        dead.add((i, target))  # 완전 탐색 후 무해 확정 (예산 소진 시엔 미확정)


def _collect_parents(seg: list, nodes: list, checks: list):
    for k, children in nodes:
        if children:
            checks.append((seg[k][0], [seg[c][0] for c, _ in children], "그룹"))
            _collect_parents(seg, children, checks)


def _parse_flat_segment(tv: TableView, c: int, g: int, rows: list[int],
                        prev_totals: list):
    """한 열의 한 세그먼트([직전 총계..총계 g)) 분해.
    반환: [(대상 행, [구성 행], 종류)] | [](총계값 없음) | None(분해 실패)."""
    gval = tv.value(g, c)
    if gval is None:
        return []
    seg = [(i, tv.value(i, c)) for i in rows if tv.value(i, c) is not None]
    if seg:
        budget = [_PARSE_BUDGET]
        gen = _parse_trees(seg, 0, len(seg), gval, budget, set())
        nodes = next((ns for j, ns in gen if j == len(seg)), None)
        if nodes is None:
            return None
        checks = [(g, [seg[k][0] for k, _ in nodes], "총계")]
        _collect_parents(seg, nodes, checks)
        return checks
    # 빈 세그먼트(자본과부채총계류): 직전 총계 행들의 접두 합
    for k in range(1, len(prev_totals) + 1):
        cand = prev_totals[-k:]
        if sum(v for _, v in cand) == gval:
            return [(g, [r for r, _ in cand], "총계 결합")]
    return None


def _shape_suspect(tv: TableView, c: int, ref_col: int, term_rows: list[int]) -> bool:
    """기준열 구조 전이의 이형 신호: 기준열에서 0(또는 없음)이던 항이 이 열에서는
    값이 있고, 그 값이 바로 뒤따르는 형제 항들의 합과 정확히 같으면 — 이 열에서는
    그 항이 부모(뒤 항들이 자식)일 가능성이 높다 → 단정하지 않는다."""
    vals = [tv.value(j, c) for j in term_rows]
    for a, j in enumerate(term_rows):
        rv = tv.value(j, ref_col)
        cv = vals[a]
        if rv in (0, None) and cv not in (0, None):
            acc = 0
            for b in range(a + 1, len(term_rows)):
                if vals[b] is None:
                    break
                acc += vals[b]
                if acc == cv:
                    return True
    return False


def _flat_forward_footing(tv: TableView, report: CheckReport):
    grands = [i for i in tv.body
              if _P.grand_total_re.search(norm(tv.row_label(i)))
              and any(tv.value(i, c) is not None for c in range(tv.ncols))]
    if not grands:
        return
    cols = [c for c in range(tv.ncols)
            if sum(1 for i in tv.body if tv.value(i, c) is not None) >= 6]
    if not cols:
        return
    # 세그먼트 경계(총계 행)는 라벨 기준으로 전 열 공통
    segs: list[tuple[int, list[int]]] = []
    boundary = -1
    for g in grands:
        segs.append((g, [i for i in tv.body if boundary < i < g]))
        boundary = g
    # 열×세그먼트 독립 분해 — 편집으로 깨진 세그먼트만 기준열 구조를 빌린다
    plans: dict[tuple[int, int], list | None] = {}
    for c in cols:
        prev_totals: list = []
        for si, (g, rows) in enumerate(segs):
            plans[(si, c)] = _parse_flat_segment(tv, c, g, rows, prev_totals)
            gv = tv.value(g, c)
            if gv is not None:
                prev_totals.append((g, gv))

    for si, (g, rows) in enumerate(segs):
        ref_col = next((c for c in cols if plans[(si, c)]), None)
        if ref_col is None:
            if any(plans[(si, c)] is None for c in cols):
                report.add("C2", "info", tv.where(g),
                           "무들여쓰기 표의 합산 구조를 해석하지 못함 (전 열) — 수기 확인 권장")
            continue
        for c in cols:
            plan = plans[(si, c)]
            borrowed = False
            if plan is None:
                plan = plans[(si, ref_col)]
                borrowed = True
            for row, term_rows, kind in plan or []:
                stated = tv.value(row, c)
                terms = [(j, c, 1) for j in term_rows if tv.value(j, c) is not None]
                if stated is None or not terms:
                    continue
                computed = sum(tv.value(j, cc) for j, cc, _ in terms)
                ok = abs(computed - stated) <= _P.footing_tol
                if not ok and borrowed and _shape_suspect(tv, c, ref_col, term_rows):
                    # 기수 간 구성 변동(신규 계정 등)으로 트리 모양이 다를 수 있는
                    # 지점 — 오탐 방지를 위해 단정(warn·수식) 대신 참고만
                    report.add(
                        "C2", "info", tv.where(row, c),
                        f"{kind} 기재 {fmt(stated)} vs 기준열 구조 합 {fmt(computed)} — "
                        f"기준열에서 0이던 항목이 있어 구조 이형 가능성 (참고)",
                    )
                    continue
                report.tick("C2")
                note = (f"무들여쓰기 {kind} footing"
                        + (" (기준열 구조 적용)" if borrowed else ""))
                report.result("C2", tv, row, c, stated, computed, ok, note, terms=terms)
                if not ok:
                    report.add(
                        "C2", "warn", tv.where(row, c),
                        f"{kind} 기재 {fmt(stated)} ≠ 구성 {len(terms)}개 합 "
                        f"{fmt(computed)} (차이 {fmt(stated - computed)})",
                    )


def check_group_footing(views: list[TableView], report: CheckReport):
    for tv in views:
        # 손익계산서류는 들여쓰기가 계층이 아님(영업이익 뒤 1칸 들여쓰기,
        # 기타포괄손익 소계가 하위와 같은 깊이) — 실측으로 제외 확정.
        # 주석 안의 포괄손익계산서 표(K-GAAP 주석15 등)도 동일하게 제외.
        # 자본변동표는 C7(기초+변동=기말)이 담당 — 기초 행을 그룹으로 오인 방지.
        if ("손익계산서" in norm(tv.sec_title) or "손익계산서" in norm(tv.context)
                or "자본변동표" in norm(tv.sec_title) or _skip_section(tv)):
            continue
        # 들여쓰기가 없는 회사 대응: 재무제표 표는 로마자 접두 계층으로 폴백,
        # 로마자도 없으면(흥구형 IFRS) 총계-먼저 평면 전개로 footing
        depths = {tv.label_depth(i) for i in tv.body if tv.row_label(i)}
        if len(depths) < 2:
            if tv.sec_title in STATEMENT_KEYWORDS:
                before = len(report.results)
                _roman_group_footing(tv, report)
                if len(report.results) == before:
                    _flat_forward_footing(tv, report)
            continue
        body = tv.body
        for c in range(tv.ncols):
            for gi, i in enumerate(body):
                gv = tv.value(i, c)
                if gv is None:
                    continue
                # '-'(0 플레이스홀더)는 기재된 합계가 아니라 항목 — 그룹 머리 아님
                # (IFRS 연결 CF '종속기업투자주식의 취득', 법인세 '조정사항' 오탐 방지)
                if tv.cell(i, c) is not None and tv.cell(i, c).text.strip() == "-":
                    continue
                # 주당 표(순손실 ÷ 주식수)는 합산 관계가 아님
                if "주당" in norm(tv.row_label(i)):
                    continue
                d = tv.label_depth(i)
                span = []
                for j in body[gi + 1:]:
                    if not tv.row_label(j):
                        continue
                    if tv.label_depth(j) <= d:
                        break
                    span.append(j)
                if not span:
                    continue
                # 직계 자식 = 구간 내 최소 깊이 행만 (중첩 그룹 이중합산 방지)
                min_d = min(tv.label_depth(j) for j in span)
                span_rows = [j for j in span if tv.label_depth(j) == min_d]
                # K-GAAP 2단 금액 열: 그룹 소계는 오른쪽 열, 구성 항목은 왼쪽 열
                # → 같은 열 / 바로 왼쪽 열 두 후보
                cands = []
                for cc in ([c] if c == 0 else [c, c - 1]):
                    ch = [j for j in span_rows if tv.value(j, cc) is not None]
                    if len(ch) >= 2:
                        total = sum(tv.value(j, cc) for j in ch)
                        cands.append((total, [(j, cc, 1) for j in ch]))
                        # K-GAAP 현금흐름표 '…유출액/차감' 그룹: 그룹값은 음수(괄호),
                        # 하위는 양수 나열 → 부호 반전 후보
                        if re.search(r"(유출|차감|감소)", norm(tv.row_label(i))):
                            cands.append((-total, [(j, cc, -1) for j in ch]))
                if not cands:
                    continue
                report.tick("C2")
                computed, terms, ok = _pick(cands, gv)
                report.result("C2", tv, i, c, gv, computed, ok, f"하위 {len(terms)}개",
                              terms=terms)
                if not ok:
                    report.add(
                        "C2", "warn", tv.where(i, c),
                        f"그룹 합계 기재 {fmt(gv)} ≠ 하위 {len(terms)}개 합 "
                        f"{fmt(computed)} (차이 {fmt(gv - computed)})",
                    )


# ── C3: 합계/총계 행 footing ─────────────────────────────────────────

def _depth_diverse(tv: TableView) -> bool:
    """들여쓰기 계층이 실재하는 표인가 (…총계 footing의 전제).

    자본관리처럼 평평한 계산 캐스케이드 표(부채총계−차감=조정부채)에서는
    '…총계' 라벨이 데이터 행이므로 footing 대상이 아니다."""
    depths = {tv.label_depth(i) for i in tv.body if tv.row_label(i)}
    return len(depths) >= 2


def check_total_rows(views: list[TableView], report: CheckReport):
    for tv in views:
        if _skip_section(tv):
            continue
        for i in tv.body:
            # 합계 라벨이 첫 열이 아닐 수 있음(구분|계정과목 양식) — 텍스트 셀 전체 판정
            labels = [norm(tv.text(i, c)) for c in range(tv.ncols)
                      if tv.text(i, c).strip() and tv.value(i, c) is None]
            exact_total = any(_P.total_re.match(t) for t in labels)
            grand_total = (not exact_total
                           and any(_P.grand_total_re.search(t) for t in labels))
            # '차감후 금액'·'…순액'도 합계행 (합계라는 단어 없는 결과 행 — 주석8 퇴직급여)
            deduct_total = (not exact_total and not grand_total
                            and any(_P.deduct_total_re.search(t) for t in labels))
            if not (exact_total or grand_total or deduct_total):
                continue
            if grand_total and not _depth_diverse(tv):
                continue
            for c in range(tv.ncols):
                node = tv.cell(i, c)
                if node is None or "AUTOSUM" in node.attrs:  # C1이 담당
                    continue
                stated = cell_value(node.text)
                if stated is None or node.text.strip() in ("", "-"):
                    continue
                cands, contributing = _col_candidates(tv, i, c)
                if c > 0 and tv.period_bucket(c - 1) == tv.period_bucket(c):
                    # K-GAAP 2단 금액 열(같은 기간의 좌열) — 구성 항목이 왼쪽 열 (자본총계 등)
                    # 당기|전기처럼 기간이 다른 왼쪽 열은 후보로 쓰지 않는다
                    cands_left, contrib_left = _col_candidates(tv, i, c - 1)
                    cands += cands_left
                    contributing = max(contributing, contrib_left)
                if contributing < 2:
                    continue  # 합산 대상이 없는 표(자본관리 등) — 합계행이 아님
                report.tick("C3")
                computed, terms, ok = _pick(cands, stated)
                report.result("C3", tv, i, c, stated, computed, ok, terms=terms)
                if not ok:
                    report.add(
                        "C3", "warn", tv.where(i, c),
                        f"합계행 기재 {fmt(stated)} ≠ 계산 {fmt(computed)} "
                        f"(차이 {fmt(stated - computed)})",
                    )


# ── C3b: 합계 헤더 열 (가로합) ───────────────────────────────────────
# 열 머리글이 '합계/총계/계'인 열은 각 행에서 형제 열들의 합이어야 한다.
# AUTOSUM(ROWPART)을 쓰지 않는 회사에서도 작동하는 구조 규칙.
# 형제 = 그룹 헤더(COLSPAN)를 공유하는 열들, 그룹이 없으면 라벨·주석·합계류 외 전체.

def _leaf_header(tv: TableView, c: int) -> str:
    for i in reversed(tv.head):
        gc = tv.grid[i][c]
        if gc is not None and gc.node.text.strip() and not gc.node.text.startswith("(단위"):
            return gc.node.text.strip()
    return ""


def _header_group_cols(tv: TableView, c: int) -> list[int]:
    """상위 헤더(그룹)가 같은 열들. '(단위 : 원)' 전폭 행은 그룹이 아니므로 건너뛰고,
    의미 있는 그룹(1 < span < 전체 폭)을 위에서부터 찾는다."""
    for i in tv.head:
        gc = tv.grid[i][c]
        if gc is None:
            continue
        owner = gc.node
        if owner.text.strip().startswith("(단위"):
            continue
        cols = [cc for cc in range(tv.ncols)
                if tv.grid[i][cc] is not None and tv.grid[i][cc].node is owner]
        if len(cols) >= tv.ncols:  # 전폭 행 — 그룹 정보 아님
            continue
        if len(cols) > 1:
            return cols
        return []  # span 1 헤더에 도달 — 그룹 없음
    return []


def check_total_columns(views: list[TableView], report: CheckReport):
    for tv in views:
        if not tv.head or _skip_section(tv):
            continue
        total_cols = [c for c in range(tv.ncols)
                      if _P.total_header_re.match(norm(_leaf_header(tv, c)))]
        for c in total_cols:
            group = _header_group_cols(tv, c)
            if group:
                siblings = [cc for cc in group if cc != c
                            and not _P.total_header_re.match(norm(_leaf_header(tv, cc)))]
            else:
                siblings = [cc for cc in range(tv.ncols)
                            if cc != c
                            and not _P.total_header_re.match(norm(_leaf_header(tv, cc)))
                            and "주석" not in norm(_leaf_header(tv, cc))]
            if not siblings:
                continue
            for i in tv.body:
                stated = tv.value(i, c)
                node = tv.cell(i, c)
                if stated is None or node is None:
                    continue
                if "AUTOSUM" in node.attrs and node.attrs["AUTOSUM"].startswith("COL"):
                    continue  # 세로 자동합계 셀은 C1 담당 (가로합이 아님)
                if _P.nonsum_col_re.search(norm(tv.row_label(i))):
                    continue  # 기대손실율 등 비율 행은 가로합 비대상
                # 후보 2개: 형제 전체 / 첫 값 열 제외 — 만기분석표(장부금액|만기…|합계)는
                # 합계가 만기 열들의 합이라 장부금액(첫 값 열)을 빼야 한다
                terms_all = [(i, cc, 1) for cc in siblings if tv.value(i, cc) is not None]
                cands = []
                if len(terms_all) >= 2:
                    cands.append((sum(tv.value(i, cc) for _, cc, _ in terms_all), terms_all))
                if len(terms_all) >= 3:
                    tail = terms_all[1:]
                    cands.append((sum(tv.value(i, cc) for _, cc, _ in tail), tail))
                if not cands:
                    continue
                report.tick("C3")
                computed, terms, ok = _pick(cands, stated)
                report.result("C3", tv, i, c, stated, computed, ok, "합계 열(가로합)",
                              terms=terms, axis="row")
                if not ok:
                    report.add(
                        "C3", "warn", tv.where(i, c),
                        f"합계 열 기재 {fmt(stated)} ≠ 형제 열 합 {fmt(computed)} "
                        f"(차이 {fmt(stated - computed)})",
                    )


# ── C4: 재무제표 ↔ 주석 대사 ─────────────────────────────────────────

def _note_amount_sets(note_views: list[TableView]) -> dict:
    """주석 표 금액 집합 — 표 단위(천원 등)를 원화로 환산해 저장."""
    out = {"당기": set(), "전기": set(), None: set(), "unit": 1}
    for tv in note_views:
        out["unit"] = max(out["unit"], tv.unit)
        for i in tv.body:
            for c in range(tv.ncols):
                v = tv.value(i, c)
                if v is None or v == 0:
                    continue
                out[tv.period_bucket(c)].add(abs(v) * tv.unit)
    return out


def _amount_in_note(target: int, amounts: dict, bucket: str) -> bool:
    """직접 일치 → 두 값의 합/차 조합(유동+비유동, 종속+관계 등) 순으로 탐색."""
    pool = amounts[bucket] | amounts[None]
    if target in pool:
        return True
    for a in pool:
        if (target - a) in pool or (target + a) in pool:
            return True
    return False


# 전역 검색 시 우연 일치 방지: 최소 금액 미만은 참조 주석 안에서만 찾음 (주당이익 등)

def check_fs_to_notes(ctx: dict, report: CheckReport):
    """참조번호 비의존 대사: 금액을 전체 주석에서 전역 검색하고, 참조번호는
    역으로 검증한다 — 참조 주석엔 없는데 다른 주석에서 발견되면 참조번호 오류 의심.
    (주석 번호는 매기 바뀌어 실수가 잦다 — 실무 피드백 반영)"""
    note_amounts = {n: _note_amount_sets(tvs) for n, tvs in ctx["notes"].items()}

    def found_in(n, v, bucket):
        # v는 이미 원화 환산된 값. 주석이 천원 단위면 반올림 차이를 허용
        # (재무제표 55,840,847,426원 ↔ 주석 55,840,847천원)
        pool = note_amounts[n][bucket] | note_amounts[n][None]
        if v in pool:
            return True
        unit = note_amounts[n]["unit"]
        if unit > 1:
            base = (v // unit) * unit
            return base in pool or base + unit in pool
        return False

    for tv in ctx["statements"]:
        note_col = next(
            (c for c in range(tv.ncols) if "주석" in norm(tv.col_header(c))), None
        )
        for i in tv.body:
            refs = ([int(x) for x in re.findall(r"\d+", tv.text(i, note_col))]
                    if note_col is not None else [])
            refs = [n for n in refs if n in note_amounts]
            for c in range(tv.ncols):
                bucket = tv.period_bucket(c)
                v = tv.value(i, c)
                if bucket is None or v is None or v == 0:
                    continue
                won = abs(v) * tv.unit  # 재무제표 표 단위 → 원화
                if not refs and won < _P.global_search_min:
                    continue  # 참조도 없고 소액 — 대사 대상 아님
                report.tick("C4")

                ref_hit = [n for n in refs if found_in(n, won, bucket)]
                global_hit = ([n for n in note_amounts if found_in(n, won, bucket)]
                              if won >= _P.global_search_min else ref_hit)
                display = sorted(set(refs) | set(global_hit))[:6]

                if ref_hit:
                    report.result("C4", tv, i, c, v, None, True,
                                  f"참조 주석 {','.join(map(str, ref_hit))} 확인",
                                  refs=display, axis="row")
                    continue
                if refs and any(_amount_in_note(won, note_amounts[n], bucket)
                                for n in refs):
                    report.result("C4", tv, i, c, v, None, True,
                                  f"참조 주석 {','.join(map(str, refs))} 조합 일치",
                                  refs=display, axis="row")
                    continue
                if global_hit and refs:
                    # 참조 주석엔 없는데 다른 주석에 있음 → 참조번호 오류 의심 (warn)
                    where_found = ",".join(map(str, global_hit[:4]))
                    report.result("C4", tv, i, c, v, None, False,
                                  f"참조 주석 {','.join(map(str, refs))}엔 없고 "
                                  f"주석 {where_found}에서 발견 — 참조번호 확인 필요",
                                  refs=display, axis="row")
                    report.add(
                        "C4", "warn", tv.where(i, c),
                        f"{fmt(v)}: 참조 주석({','.join(map(str, refs))})에 없고 "
                        f"주석 {where_found}에서 발견 — 참조번호 오류 의심",
                    )
                    continue
                if global_hit:
                    # 참조 없이 전역 검색으로 근거 확인
                    report.result("C4", tv, i, c, v, None, True,
                                  f"주석 {','.join(map(str, global_hit[:4]))}에서 확인"
                                  f"{' (참조 표기 없음)' if not refs else ''}",
                                  refs=display, axis="row")
                    continue
                if not refs:
                    continue  # 참조 없고 전역에도 없음 — 주석 미기재 항목일 수 있어 생략
                report.result("C4", tv, i, c, v, None, False,
                              f"참조 주석 {','.join(map(str, refs))} 및 전체 주석에서 미확인",
                              refs=display, axis="row")
                report.add(
                    "C4", "info", tv.where(i, c),
                    f"{fmt(v)} 이(가) 참조 주석 및 전체 주석 표에서 직접·조합으로 "
                    f"확인되지 않음 (세후 표시·부분 참조 가능성)",
                )


# ── C7: 재무제표 계산 라인 검증 ──────────────────────────────────────
# 손익계산서: 계산 라인(매출총이익/영업이익/차감전순이익/순이익/총포괄손익) =
#   직전 계산 라인 + 사이 행들의 부호합 (수익 +, 원가·비용·손실 −).
#   기타영업외비용이 음수(환입)여도 값 자체가 괄호라 부호 규칙이 성립 (실측 확인).
# 라벨 힌트: "…(Ⅰ+Ⅱ+Ⅲ+Ⅳ)"처럼 합산식이 라벨에 명시된 행(현금흐름표)을 그대로 검산.

_ROMAN = "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ"
_HINT_RE = re.compile(rf"\(([{_ROMAN}][+{_ROMAN}\s]*)\)$")


def _pl_sign(label: str) -> int:
    return -1 if _P.pl_minus_re.search(norm(label)) else 1


def _label_core(label: str) -> str:
    """'영업이익(손실)', '기본주당이익(손실) (단위 : 원)' → 괄호 접미 제거."""
    core = norm(label)
    while True:
        stripped = re.sub(r"\([^()]*\)$", "", core)
        if stripped == core:
            return core
        core = stripped


def _check_pl_anchors(tv: TableView, report: CheckReport):
    body = tv.body
    anchors = [
        i for i in body
        if _P.anchor_re.search(_label_core(tv.row_label(i)))
        and "주당" not in norm(tv.row_label(i))
        and any(tv.value(i, c) is not None for c in range(tv.ncols))
    ]
    for ai, i in enumerate(anchors):
        prev = anchors[ai - 1] if ai else None
        between = [j for j in body if (prev if prev is not None else -1) < j < i
                   and tv.row_label(j)]
        if not between:
            continue
        # 사이 행 중 최소 들여쓰기 깊이만 (그룹 하위 항목 이중합산 방지)
        with_val = [j for j in between
                    if any(tv.value(j, c) is not None for c in range(tv.ncols))]
        if not with_val:
            continue
        min_d = min(tv.label_depth(j) for j in with_val)
        comps = [j for j in with_val if tv.label_depth(j) == min_d]
        for c in range(tv.ncols):
            stated = tv.value(i, c)
            if stated is None or not tv.cell(i, c) or tv.cell(i, c).text.strip() in ("", "-"):
                continue
            base = ([(prev, c, 1)] if prev is not None and tv.value(prev, c) is not None
                    else [])
            # 비용 표기 2관행: (a) 양수 기재 + 차감 → 키워드 부호합 (IFRS-A 유형)
            #                (b) 음수 괄호 기재 → 단순합 (IFRS-B 유형)
            # K-GAAP 2단 금액 열: 구성 항목이 바로 왼쪽 열에 올 수 있음
            # K-GAAP 일부 문서는 손실을 양수로 표기 — 직전 앵커('…손실' 라벨, 양수 값)의
            # 실질 부호가 음수인 변형까지 후보에 포함
            bases = [base]
            if (prev is not None and base
                    and "손실" in _label_core(tv.row_label(prev)) and base[0][2] == 1
                    and tv.value(prev, c) is not None and tv.value(prev, c) > 0):
                bases.append([(prev, c, -1)])
            cols = [c] if c == 0 else (
                [c, c - 1] if tv.period_bucket(c - 1) == tv.period_bucket(c) else [c]
            )
            cands = []
            for b in bases:
                for cc in cols:
                    comp_rows = [j for j in comps if tv.value(j, cc) is not None]
                    signed = b + [(j, cc, _pl_sign(tv.row_label(j))) for j in comp_rows]
                    plain = b + [(j, cc, 1) for j in comp_rows]
                    for terms in (signed, plain):
                        if len(terms) >= 2:
                            cands.append(
                                (sum(s * tv.value(j, c2) for j, c2, s in terms), terms)
                            )
            if not cands:
                continue
            report.tick("C7")
            computed, terms, ok = _pick(cands, stated)
            if not ok and "손실" in _label_core(tv.row_label(i)) and stated > 0:
                # 손실 양수 표기: 기재 절대값과 계산 절대값 비교
                computed2, terms2, ok2 = _pick(cands, -stated)
                if ok2:
                    computed, terms, ok = -computed2, terms2, True
            report.result("C7", tv, i, c, stated, computed, ok, "손익 계산 라인", terms=terms)
            if not ok:
                report.add(
                    "C7", "warn", tv.where(i, c),
                    f"계산 라인 기재 {fmt(stated)} ≠ 계산 {fmt(computed)} "
                    f"(차이 {fmt(stated - computed)})",
                )


def _check_hint_lines(tv: TableView, report: CheckReport):
    body = tv.body
    by_roman = {}
    for j in body:
        m = re.match(rf"^([{_ROMAN}]+)\s*\.", tv.row_label(j))
        if m and m.group(1) not in by_roman:
            by_roman[m.group(1)] = j
    for i in body:
        label = norm(tv.row_label(i))
        m = _HINT_RE.search(label)
        if not m:
            continue
        romans = [r.strip() for r in m.group(1).split("+")]
        rows = [by_roman.get(r) for r in romans]
        if any(r is None or r == i for r in rows):
            continue
        for c in range(tv.ncols):
            stated = tv.value(i, c)
            if stated is None:
                continue
            terms = [(j, c, 1) for j in rows if tv.value(j, c) is not None]
            if len(terms) < 2:
                continue
            report.tick("C7")
            computed = sum(tv.value(j, c) for j, _, _ in terms)
            ok = computed == stated
            report.result("C7", tv, i, c, stated, computed, ok,
                          f"라벨 합산식 ({'+'.join(romans)})", terms=terms)
            if not ok:
                report.add(
                    "C7", "warn", tv.where(i, c),
                    f"라벨 합산식({'+'.join(romans)}) 기재 {fmt(stated)} ≠ 계산 "
                    f"{fmt(computed)} (차이 {fmt(stated - computed)})",
                )


# '기초제품/기말상품' 등은 잔액 흐름이 아니라 원가 산식의 재고 항목

def _flow_open(label: str) -> bool:
    return "기초" in label and not _P.inventory_re.match(label)


def _flow_close(label: str) -> bool:
    return "기말" in label and not _P.inventory_re.match(label)


def _check_socie(tv: TableView, report: CheckReport):
    """기초→변동→기말 세로 흐름: 기말 행 = 기초 행 + 사이 변동 행들의 합 (열별).
    자본변동표뿐 아니라 충당부채·운용자산 등 변동 내역 주석 표에도 적용.
    AUTOSUM·들여쓰기에 의존하지 않는 구조 규칙.
    현금흐름표 본표는 제외 — 기말 현금 = 순증감(기초보다 위 행) + 기초라 구조가 다름."""
    if "현금흐름표" in norm(tv.sec_title):
        return
    opening = None
    between: list[int] = []
    for i in tv.body:
        label = norm(tv.row_label(i))
        has_value = any(tv.value(i, c) is not None for c in range(tv.ncols))
        if _flow_open(label) and has_value:
            opening, between = i, []
            continue
        if opening is None:
            continue
        if _flow_close(label) and has_value:
            # 섹션(총포괄손익 : / 주주와의 거래 :)마다 소계 유무가 다르다 —
            # 소계(비들여쓰기 행)가 있는 섹션은 소계만, 없는 섹션은 항목 전부.
            smart_rows: list[int] = []
            section: list[int] = []

            def flush_section():
                valued = [j for j in section
                          if any(tv.value(j, cc) is not None for cc in range(tv.ncols))]
                if not valued:
                    return
                min_d = min(tv.label_depth(j) for j in valued)
                top = [j for j in valued if tv.label_depth(j) == min_d]
                smart_rows.extend(top if len(top) < len(valued) else valued)

            for j in between:
                if tv.is_section_row(j):
                    flush_section()
                    section = []
                else:
                    section.append(j)
            flush_section()

            # 연결 양식: 섹션 라벨('총포괄손익')과 같은 이름의 '… 합계' 재합계 행이
            # 섹션 항목들 뒤에 옴 → 그 행이 섹션 수집분을 대체하는 후보
            replaced: list[int] = []
            stack: list[tuple[str, int]] = []  # (섹션 라벨, replaced 내 시작 위치)
            for j in between:
                if tv.is_section_row(j):
                    stack.append((_label_core(tv.row_label(j)).rstrip(":"), len(replaced)))
                    continue
                if not any(tv.value(j, cc) is not None for cc in range(tv.ncols)):
                    continue
                # '총포괄손익 합계'처럼 라벨이 합계로 끝나는 재합계 행 (정확 일치 아님)
                if tv.row_is_total(j) or re.search(r"(합계|총계)$", norm(tv.row_label(j))):
                    core = re.sub(r"(합\s*계|총\s*계)$", "", norm(tv.row_label(j))).strip()
                    hit = next((k for k in range(len(stack) - 1, -1, -1)
                                if stack[k][0] == core), None)
                    if hit is not None:
                        replaced = replaced[:stack[hit][1]]
                        stack = stack[:hit]
                replaced.append(j)

            candidates_rows = [smart_rows]
            if replaced != smart_rows:
                candidates_rows.append(replaced)
            if between not in candidates_rows:
                candidates_rows.append(between)  # 폴백: 전부 (무들여쓰기 회사)
            for c in range(tv.ncols):
                stated = tv.value(i, c)
                if stated is None or tv.value(opening, c) is None:
                    continue
                if _P.nonsum_col_re.search(norm(tv.col_header(c))):
                    continue  # 평균 가격·비율 열은 흐름 합산 비대상
                cands = []
                for rows in candidates_rows:
                    valued_rows = [j for j in rows if tv.value(j, c) is not None]
                    # 변동 표기 2관행: 괄호 음수(단순합) / 양수 기재+라벨 차감(부호합)
                    plain = [(opening, c, 1)] + [(j, c, 1) for j in valued_rows]
                    signed = [(opening, c, 1)] + [
                        (j, c, -1 if _P.flow_minus_re.search(norm(tv.row_label(j))) else 1)
                        for j in valued_rows
                    ]
                    for terms in (plain, signed):
                        cands.append(
                            (sum(s * tv.value(j, cc) for j, cc, s in terms), terms)
                        )
                if all(len(t) < 2 for _, t in cands):
                    continue
                report.tick("C7")
                computed, terms, ok = _pick(cands, stated)
                report.result("C7", tv, i, c, stated, computed, ok,
                              "자본변동 (기초+변동=기말)", terms=terms)
                if not ok:
                    report.add(
                        "C7", "warn", tv.where(i, c),
                        f"기초+변동 계산 {fmt(computed)} ≠ 기말 기재 {fmt(stated)} "
                        f"(차이 {fmt(stated - computed)})",
                    )
            opening, between = None, []
            continue
        between.append(i)  # 섹션 라벨 행도 포함 — 섹션별 소계 판정의 경계로 쓰임


def _check_flow_columns(tv: TableView, report: CheckReport):
    """가로 흐름: 열이 기초|증감…|기말인 표(유형자산 변동표 등)에서
    각 행의 기말 = 기초 + 사이 열들. 증감이 괄호 음수면 단순합, 양수 표기면
    헤더 키워드(처분·감소·상각 등) 부호합 — 두 후보."""
    headers = [norm(_leaf_header(tv, c)) for c in range(tv.ncols)]
    start = next((c for c, h in enumerate(headers)
                  if re.fullmatch(r"기초(금액|잔액)?", h)), None)
    end = next((c for c in range((start or 0) + 2, tv.ncols)
                if re.fullmatch(r"기말(금액|잔액)?", headers[c])), None)
    if start is None or end is None:
        return
    middle = list(range(start + 1, end))
    for i in tv.body:
        stated = tv.value(i, end)
        if stated is None or tv.value(i, start) is None:
            continue
        base = [(i, start, 1)]
        comps = [c for c in middle if tv.value(i, c) is not None]
        plain = base + [(i, c, 1) for c in comps]
        signed = base + [(i, c, -1 if _P.flow_minus_re.search(headers[c]) else 1)
                         for c in comps]
        cands = [(sum(s * tv.value(i, cc) for _, cc, s in t), t)
                 for t in (plain, signed) if len(t) >= 2]
        if not cands:
            continue
        report.tick("C7")
        computed, terms, ok = _pick(cands, stated)
        report.result("C7", tv, i, end, stated, computed, ok,
                      "가로 흐름 (기초+증감=기말)", terms=terms, axis="row")
        if not ok:
            report.add(
                "C7", "warn", tv.where(i, end),
                f"가로 흐름 기재 {fmt(stated)} ≠ 기초+증감 {fmt(computed)} "
                f"(차이 {fmt(stated - computed)})",
            )


def check_statement_lines(views: list[TableView], report: CheckReport):
    for tv in views:
        if _skip_section(tv):
            continue
        if "손익계산서" in tv.sec_title:
            _check_pl_anchors(tv, report)
        # 기초→기말 세로 흐름: 자본변동표 + 변동 내역 주석 표 전반
        _check_socie(tv, report)
        _check_flow_columns(tv, report)
        _check_hint_lines(tv, report)


# ── C5: 주석 ↔ 주석 대사 ─────────────────────────────────────────────

def check_note_to_note(ctx: dict, report: CheckReport):
    index: dict[tuple, dict] = {}  # (label, bucket) -> {note_no: set(|values|)}
    for n, tvs in ctx["notes"].items():
        for tv in tvs:
            for i in tv.body:
                label = norm(tv.row_label(i))
                if len(label) < 2 or label in _P.generic_labels or is_total_label(label):
                    continue
                for c in range(tv.ncols):
                    bucket = tv.period_bucket(c)
                    v = tv.value(i, c)
                    if bucket is None or v is None or v == 0:
                        continue
                    # 부호는 표 맥락(차감/조정)에 따라 뒤집히므로 절대값, 단위는 원화 환산
                    index.setdefault((label, bucket), {}).setdefault(n, set()).add(
                        abs(v) * tv.unit
                    )
    for (label, bucket), by_note in index.items():
        if len(by_note) < 2:
            continue
        report.tick("C5")
        sets = list(by_note.values())
        if set.intersection(*sets):
            continue
        detail = ", ".join(
            f"주석{n}={'/'.join(fmt(v) for v in sorted(vs))}" for n, vs in sorted(by_note.items())
        )
        report.add("C5", "info", f"'{label}' ({bucket})", f"주석 간 금액 상이 — {detail}")


# ── C6: 표기 검출 ────────────────────────────────────────────────────

def check_format(doc: Document, views: list[TableView], report: CheckReport):
    units = {}
    for tv in views:
        note_cols = {c for c in range(tv.ncols) if "주석" in norm(tv.col_header(c))}
        for i in range(tv.n):
            for c in range(tv.ncols):
                node = tv.cell(i, c)
                if node is None:
                    continue
                t = node.text
                s = t.strip()
                report.tick("C6")
                m = _P.unit_re.search(s)
                if m:
                    units.setdefault(norm(m.group(1)), []).append(tv.where(i))
                if not s:
                    continue
                if s.count("(") != s.count(")"):
                    report.add("C6", "info", tv.where(i, c), f"괄호 짝 불일치: '{s[:30]}'")
                    continue
                if c in note_cols or "%" in s:
                    continue
                # 우측 정렬 + 숫자·콤마 구성인데 숫자 표기가 아님 → 깨진 숫자 의심
                # (목차 점선 '.....'처럼 숫자가 없는 문자열은 제외)
                if (
                    node.align == "RIGHT"
                    and re.fullmatch(r"[\d,.()\s-]+", s)
                    and s != "-"
                    and any(ch.isdigit() for ch in s)
                    and cell_value(t) is None
                ):
                    report.add("C6", "warn", tv.where(i, c), f"깨진 숫자 의심: '{s[:30]}'")
                elif cell_value(t) is not None and t != s and not node.readonly:
                    report.add("C6", "info", tv.where(i, c), f"숫자에 여분 공백: '{t[:30]}'")
    if len(units) > 1:
        detail = ", ".join(f"'{u}' {len(ws)}곳" for u, ws in sorted(units.items()))
        report.add("C6", "info", "문서 전체", f"단위 표기 혼용: {detail}")


# ── 진입점 ───────────────────────────────────────────────────────────

CHECK_NAMES = {
    "C1": "자동합계(AUTOSUM) 검산",
    "C2": "그룹 footing (들여쓰기·로마자 계층)",
    "C3": "합계행·합계열 footing",
    "C7": "재무제표 계산 라인 검증 (손익·자본변동·현금흐름)",
    "C4": "재무제표↔주석 대사 (참조번호 오류 감지)",
    "C5": "주석↔주석 대사",
    "C8": "대사맵 (확정 대사쌍 검증)",
    "C6": "표기 검출",
}


def _apply_suppressions(report: CheckReport):
    """프로파일 suppress 규칙에 걸리는 발견 사항을 억제 목록으로 이동.

    findings(콘솔·요약의 지적 사항)만 억제한다 — annotate가 심는 살아있는
    검산 수식은 계산 결과 그대로 보여준다(수치를 감추지 않는 것이 원칙)."""
    if not _P.suppress:
        return
    kept: list[Finding] = []
    for f in report.findings:
        rule = next((r for r in _P.suppress
                     if r.matches(f.check, f.where, f.message)), None)
        if rule is None:
            kept.append(f)
        else:
            report.suppressed.append((f, rule.reason))
    report.findings = kept


def run_checks(doc: Document, only: set[str] | None = None,
               profile: Profile | None = None) -> CheckReport:
    with use_profile(profile):
        report = CheckReport(profile_name=_P.name)
        views, ctx = build_views(doc)
        steps = {
            "C1": lambda: check_autosum(views, report),
            "C2": lambda: check_group_footing(views, report),
            "C3": lambda: (check_total_rows(views, report),
                           check_total_columns(views, report)),
            "C7": lambda: check_statement_lines(views, report),
            "C4": lambda: check_fs_to_notes(ctx, report),
            "C5": lambda: check_note_to_note(ctx, report),
            "C6": lambda: check_format(doc, views, report),
        }
        for key, fn in steps.items():
            if only is None or key in only:
                fn()
        _apply_suppressions(report)
        # 커버리지: 어느 표에 어떤 재무 검산이 적용됐는지 — "검증 안 된 표"의
        # 침묵 통과를 막는다 (실행한 검사 기준)
        applied: dict[str, set] = {}
        for x in report.results:
            applied.setdefault(x.table_id, set()).add(x.check)
        report.coverage = [
            {"table": tv.table.id, "context": tv.context,
             "checks": sorted(applied.get(tv.table.id, ())),
             "skipped": _skip_section(tv)}
            for tv in views
        ]
        return report
