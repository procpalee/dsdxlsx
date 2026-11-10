"""TableView — 그리드 기반 값/라벨/헤더 접근 (dsdxlsx check.py 포팅, 프로파일 제거).

풋팅·대사 엔진의 공용 추상화. 좌표는 (표 행 인덱스 i, 그리드 열 c) — 0-base.
"""
from __future__ import annotations

import re

from . import constants as C
from .sections import SheetPlan
from .. import grid as grid_mod
from .. import numbers
from ..model import Table


def classify_ext(s: str):
    """numbers.classify 확장: '10,000원' → 10000, '4.35%' → 0.0435(분수 저장).

    퍼센트는 엑셀 표준(분수 값 + % 서식)으로 저장한다 - 유효세율 산식
    (② ÷ ①) 검증이 자연스럽게 성립한다. '천원/백만원' 단위 표기는 숫자로 오인 금지."""
    from decimal import Decimal

    v = numbers.classify(s)
    if v is not None and numbers.suffix_of(s) == "%":
        return Decimal(str(v)) / 100
    if v is None and s.endswith("원") and not s.endswith(("천원", "백만원", "십억원", "만원")):
        v = numbers.classify(s[:-1].strip())
    return v


def cell_value(text: str):
    """검산용 값: 숫자 → int/Decimal, '-' → 0, 접미 단위('1,234,560주', '62.12%',
    '10,000원')는 단위를 뗀 수치, 그 외 None. (전체 일치만 — '1.예적금담보대출'은 텍스트)"""
    s = text.strip()
    if s == "-":
        return 0
    return classify_ext(s)


def is_total_label(label: str) -> bool:
    n = C.norm(label)
    return bool(C.TOTAL_RE.match(n) or C.GRAND_TOTAL_RE.search(n))


class TableView:
    """한 표에 대한 값/라벨/헤더 접근 + 시트 컨텍스트."""

    def __init__(self, table: Table, sheet: str, sec_title: str, note_no: int | None = None):
        self.table = table
        self.sheet = sheet          # 검증이 배치될 시트 이름 ("BS", "1", …)
        self.sec_title = sec_title  # 섹션 제목 (재무제표 키워드 판별용)
        self.note_no = note_no
        self.grid = grid_mod.expand(table)
        self.n = len(table.rows)
        self.ncols = table.ncols
        self.body = [i for i, r in enumerate(table.rows) if not r.in_thead]
        self.head = [i for i, r in enumerate(table.rows) if r.in_thead]
        # 표 밖(직전 문단·캡션)에서 결정되는 값 — build_views가 채운다
        self.unit_fallback: int = 1
        self.caption_bucket: str | None = None
        self.note_title: str = ""       # 주석 제목 (담보·지급보증 표 대사 제외 판별용)

    # ── 셀 접근 ──────────────────────────────────────────────
    def cell(self, i: int, c: int):
        gc = self.grid[i][c]
        return gc.node if gc is not None and gc.kind == "cell" else None

    def owner(self, i: int, c: int):
        """병합 연속 위치(hspan/vspan)면 소유 셀 반환."""
        gc = self.grid[i][c]
        return gc.node if gc is not None else None

    def text(self, i: int, c: int) -> str:
        gc = self.grid[i][c]
        return gc.node.text if gc is not None else ""

    def value(self, i: int, c: int):
        # 주석 참조 열("4,5,32")은 절대 숫자가 아니다 — 검산 항 유입 차단
        # (IFRS BS에서 좌열 후보(cc=c-1)가 주석 열을 집는 사고 방지, 자체감사 실측)
        if c in self.note_ref_cols():
            return None
        node = self.cell(i, c)
        return cell_value(node.text) if node else None

    # ── 행 성격 ──────────────────────────────────────────────
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
            t = C.norm(self.text(i, c))
            if t and self.value(i, c) is None and (C.TOTAL_RE.match(t)
                                                   or C.GRAND_TOTAL_RE.search(t)):
                return True
        return False

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

    # ── 헤더/기간/단위 ───────────────────────────────────────
    def col_header(self, c: int) -> str:
        parts = []
        for i in self.head:
            gc = self.grid[i][c]
            if gc is not None:
                t = gc.node.text.strip()
                if t and t not in parts and not t.startswith("(단위"):
                    parts.append(t)
        return "/".join(parts)

    def leaf_header(self, c: int) -> str:
        for i in reversed(self.head):
            gc = self.grid[i][c]
            if gc is not None and gc.node.text.strip() and not gc.node.text.startswith("(단위"):
                return gc.node.text.strip()
        return ""

    def header_group_cols(self, c: int) -> list[int]:
        """상위 헤더(그룹)가 같은 열들. '(단위 : 원)' 전폭 행은 그룹이 아니므로 건너뛰고,
        의미 있는 그룹(1 < span < 전체 폭)을 위에서부터 찾는다."""
        for i in self.head:
            gc = self.grid[i][c]
            if gc is None:
                continue
            owner = gc.node
            if owner.text.strip().startswith("(단위"):
                continue
            cols = [cc for cc in range(self.ncols)
                    if self.grid[i][cc] is not None and self.grid[i][cc].node is owner]
            if len(cols) >= self.ncols:  # 전폭 행 — 그룹 정보 아님
                continue
            if len(cols) > 1:
                return cols
            return []  # span 1 헤더에 도달 — 그룹 없음
        return []

    def _gisu_ranks(self) -> dict[int, str]:
        """표 헤더의 기수 집합 → {기수: 버킷}. 최대 기수=당기, 둘째=전기."""
        if not hasattr(self, "_gisu"):
            nums = set()
            for c in range(self.ncols):
                m = re.search(r"제(\d+)기", C.norm(self.col_header(c)))
                if m:
                    nums.add(int(m.group(1)))
            ranked = sorted(nums, reverse=True)
            self._gisu = ({ranked[0]: "당기", ranked[1]: "전기"}
                          if len(ranked) >= 2 else {})
        return self._gisu

    def period_bucket(self, c: int) -> str | None:
        h = C.norm(self.col_header(c))
        if any(k in h for k in C.PERIOD_CURRENT) or C.PERIOD_CURRENT_RE.search(h):
            return "당기"
        if any(k in h for k in C.PERIOD_PRIOR) or C.PERIOD_PRIOR_RE.search(h):
            return "전기"
        m = re.search(r"제(\d+)기", h)
        if m:
            return self._gisu_ranks().get(int(m.group(1)))
        return None

    def bucket_of(self, c: int) -> str | None:
        """열 버킷 — 열 헤더 우선, 없으면 표 캡션/직전 문단의 당기·전기."""
        return self.period_bucket(c) or self.caption_bucket

    @property
    def unit(self) -> int:
        """표 단위 배수 (원=1, 천원=1000, …). 헤더에 없으면 직전 문단 폴백."""
        if not hasattr(self, "_unit"):
            self._unit = 0
            for i in range(self.n):
                for c in range(self.ncols):
                    u = C.unit_of(self.text(i, c))
                    if u:
                        self._unit = u
                        break
                if self._unit:
                    break
            if not self._unit:
                self._unit = self.unit_fallback or 1
        return self._unit

    def note_ref_cols(self) -> set[int]:
        """'주석' 참조 열 — 항상 텍스트 유지, 합산 제외."""
        if not hasattr(self, "_note_cols"):
            self._note_cols = {c for c in range(self.ncols)
                               if C.NOTE_COL_HEADER in C.norm(self.leaf_header(c))}
        return self._note_cols


def build_views(plans: list[SheetPlan]) -> tuple[list[TableView], dict]:
    """시트 계획의 각 표에 TableView 부여 + 컨텍스트.

    ctx = {"statements": [tv…], "notes": {원문 주석번호: [tv…]}, "by_table": {id: tv}}
    직전 문단의 (단위:…)와 당기/전기 캡션을 폴백으로 흘려 넣는다 —
    '(단위 : 천원)'이 표 밖 문단인 양식(K-GAAP 주석 등) 대응."""
    views: list[TableView] = []
    ctx: dict = {"statements": [], "notes": {}, "by_table": {}}
    for plan in plans:
        pending_unit = 1
        pending_bucket: str | None = None
        for block in plan.sec.blocks:
            if block.kind in ("p", "label"):
                text = block.item.text
                u = C.unit_of(text)
                if u:
                    pending_unit = u
                n = C.norm(text)
                if n:
                    # 반기·분기 보고서의 '당반기'·'전3분기' 변형도 같은 규칙 (CAPTION_*_RE)
                    if C.CAPTION_CURRENT_RE.match(n) or re.search(r"\(당(반|분)?기\)", n):
                        pending_bucket = "당기"
                    elif C.CAPTION_PRIOR_RE.match(n) or re.search(r"\(전(반|분)?기\)", n):
                        pending_bucket = "전기"
                    # 짧은 캡션 변형: '(1) 당기말'·'<전기>' 등 — 번호·괄호 접두 허용.
                    # 문장('당기 중 …')·양기 병기('당기 및 전기 …')는 길이·상호 배타로 제외
                    elif len(n) <= 12 and C.CAPTION_CURRENT_RE.search(n) \
                            and not C.CAPTION_PRIOR_RE.search(n):
                        pending_bucket = "당기"
                    elif len(n) <= 12 and C.CAPTION_PRIOR_RE.search(n) \
                            and not C.CAPTION_CURRENT_RE.search(n):
                        pending_bucket = "전기"
                    # 제목+캡션 병합 문단('…단기금융상품(1)당기말' — 공단 실측): 말미 판정
                    elif re.search(r"당(반|분)?기말$|\(당(반|분)?기\)$", n):
                        pending_bucket = "당기"
                    elif re.search(r"전(반|분)?기말$|\(전(반|분)?기\)$", n):
                        pending_bucket = "전기"
                continue
            table = block.item
            # 별첨은 정규 주석과 번호가 겹치므로 교차참조 컨텍스트에서는 고유 음수 키를
            # 사용한다. note_no가 None만 아니면 주석 풋팅 규칙은 그대로 적용된다.
            note_key = (-int(plan.name) if plan.kind == "note" and plan.sec.is_appendix
                        else plan.sec.note_no)
            tv = TableView(table, plan.name, plan.sec.title, note_key)
            tv.unit_fallback = pending_unit
            if plan.kind == "note":
                tv.note_title = plan.display_title or ""
            # 표 캡션(본문 첫 셀이 '당 기'/'전 기' — 유형자산 변동표 양식)
            cap = C.norm(table.rows[0].cells[0].text) if table.rows and table.rows[0].cells else ""
            if re.fullmatch(r"당\d?(반|분)?기말?", cap):
                tv.caption_bucket = "당기"
            elif re.fullmatch(r"전\d?(반|분)?기말?", cap):
                tv.caption_bucket = "전기"
            else:
                tv.caption_bucket = pending_bucket
            if tv.n > 2:
                pending_bucket = None   # 실질 표만 캡션 소비 — 사이의 단위 표는 통과
            # 1~2행짜리 '(단위: 천원)' 단위 표 — 다음 표의 단위로 이월 (유형자산 주석 실측)
            if tv.n <= 2:
                for i in range(tv.n):
                    for c in range(tv.ncols):
                        u = C.unit_of(tv.text(i, c))
                        if u:
                            pending_unit = u
                            break
                # 기간 캡션 표('(1) 당기말' 1칸 표 — 공단 실측): 다음 표의 기간으로 이월
                cap_txt = C.norm(" ".join(
                    tv.text(i, c) for i in range(tv.n) for c in range(tv.ncols)))
                if len(cap_txt) <= 12:
                    if C.CAPTION_CURRENT_RE.search(cap_txt) and not C.CAPTION_PRIOR_RE.search(cap_txt):
                        pending_bucket = "당기"
                    elif C.CAPTION_PRIOR_RE.search(cap_txt) and not C.CAPTION_CURRENT_RE.search(cap_txt):
                        pending_bucket = "전기"
            views.append(tv)
            ctx["by_table"][table.id] = tv
            if plan.kind == "note":
                ctx["notes"].setdefault(note_key, []).append(tv)
            else:
                ctx["statements"].append(tv)
    return views, ctx
