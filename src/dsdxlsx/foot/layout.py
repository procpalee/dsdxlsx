"""레이아웃 — 섹션 블록을 시트 행에 배치하고 검증 셀 좌표를 확정한다.

실측 규칙:
- 블록은 인접 행에 연속 배치 (문단=1행, 표=행수). 원본의 빈 문단이 자연 간격.
- 주석 시트: A1=제목(번호 제거), 2~3행 공백, 본문 4행부터. 헤딩 문단은 첫 줄(번호+제목)을
  떼고 나머지만 쓴다 (비어 있으면 행 없음).
- 표 아래 검증 블록: 표 마지막행+2부터 [라벨행, 수식행] 쌍, 그룹(대상 행 단위) 사이 빈행 1,
  블록 뒤 빈행 1 후 다음 블록.
- 오른쪽 검증 슬롯: 시작열 = 시트 내 최대 표 폭 + 2, [라벨/링크, 수식, 공백] 3열 단위로
  행별 왼→오 배치 (풋팅 먼저, 대사 다음). 행을 밀지 않으므로 배치 후 바인딩.
- 원문 시트: 모든 계획 시트를 같은 규칙(검증 없이)으로 이어 붙임, 섹션 간 빈행 2.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal

from . import constants as C
from .checks import CellRef, CheckSet, CrossRef, CrossSpec, FootingCheck, FootingSpec
from .sections import _APPENDIX_RE, _NOTE_LINE_RE, SheetPlan
from .. import grid as grid_mod
from ..model import Table


@dataclass
class PlacedTable:
    table: Table
    sheet: str
    start_row: int  # 표 0행의 xlsx 행 (1-base)

    def row_of(self, i: int) -> int:
        return self.start_row + i


@dataclass
class SheetLayout:
    name: str
    kind: str                                   # statement | note | wonmun
    title: str = ""                             # 헤더 바(2행)에 표시할 시트 제목
    ops: list = field(default_factory=list)     # ("para", row, text, node|None) | ("table", PlacedTable)
    below: dict = field(default_factory=dict)   # (table_id, target_i) → (label_row, formula_row)
    right_start: int = 0
    n_rows: int = 0
    _slots: dict = field(default_factory=dict)  # xlsx row → 사용한 슬롯 수

    def alloc_slot(self, row: int) -> int:
        k = self._slots.get(row, 0)
        self._slots[row] = k + 1
        return self.right_start + C.CHECK_SLOT_WIDTH * k


def _placement(spec: FootingSpec, sheet_name: str) -> str:
    """기준 양식 배치 규칙: 세로축→아래, 가로축→오른쪽. 예외: 손익계산서(PL/CI)는
    단계손익 앵커·라벨힌트만 아래, 나머지(롤업 등)는 오른쪽."""
    if spec.axis == "row":
        return "right"
    if sheet_name in ("PL", C.SECOND_PL_CODE) and spec.kind not in ("pl_anchor", "hint"):
        return "right"
    return "below"


def _heading_line(lines: list[str], min_n: int) -> tuple[int, int] | None:
    """다음 주석 헤딩 줄 위치 → (줄 인덱스, 번호).

    정규 주석은 번호가 증가해야 하지만 ``[별첨N]`` 뒤에서는 1부터 다시 시작한다.
    """
    reset = False
    for idx, ln in enumerate(lines):
        if _APPENDIX_RE.search(ln):
            reset = True
            continue
        m = _NOTE_LINE_RE.match(ln.strip())
        if m and (reset or int(m.group(1)) > min_n):
            return idx, int(m.group(1))
    return None


def compute_render_items(plans: list[SheetPlan]) -> dict[str, list]:
    """플랜별 렌더 항목 [("para", text, node) | ("table", Table)] — 기준 양식 문단 분할 규칙 적용.

    주석 헤딩이 문단 중간·끝에 섞인 경우(혼합 문단), 헤딩 줄 기준 분할:
    앞부분(내용 줄이 있으면) → 이전 주석의 문단 행, 헤딩 줄 → 제거,
    뒷부분(내용이 있으면) → 다음 주석의 첫 문단. (K-GAAP 샘플 주석 경계 실측.
    기준 양식의 공백 행 방출은 비일관적이라 공백뿐인 조각은 행을 만들지 않는다 — 문서화된 편차.)"""
    items: dict[str, list] = {p.name: [] for p in plans}
    # ("para", text, node) — node 는 원본 문단 Node (writeback 셀맵용; 원문 텍스트와
    # 다르게 잘라 쓴 경우(헤딩 제거·분할)는 writeback 이 partial 로 취급한다)
    spill: tuple[str, object] | None = None  # 직전 섹션 혼합 문단의 헤딩 이후 (텍스트, 노드)
    for idx, plan in enumerate(plans):
        out = items[plan.name]
        if plan.kind != "note":
            for b in plan.sec.blocks:
                if b.kind == "table":
                    out.append(("table", b.item))
                elif b.item.tag != "WARNING":  # 편집기 안내문(경고 라벨)은 산출물 제외
                    out.append(("para", b.item.text, b.item))
            spill = None
            continue
        n = plan.sec.note_no
        if spill is not None:
            out.append(("para", spill[0], spill[1]))
            spill = None
        first_para_seen = False
        for b in plan.sec.blocks:
            if b.kind == "table":
                out.append(("table", b.item))
                continue
            if b.kind == "label":
                if b.item.tag != "WARNING":  # 편집기 안내문은 산출물 제외
                    out.append(("para", b.item.text, b.item))
                continue
            text = b.item.text
            lines = text.split("\n")
            if not first_para_seen:
                first_para_seen = True
                # 자기 헤딩('N. 제목')이 첫 비공백 줄이면 제거
                k = next((i for i, ln in enumerate(lines) if ln.strip()), None)
                if k is not None:
                    m = _NOTE_LINE_RE.match(lines[k].strip())
                    if m and int(m.group(1)) == n:
                        rest = "\n".join(lines[k + 1:])
                        if rest.strip():
                            out.append(("para", rest, b.item))
                        continue
            hit = _heading_line(lines, n)
            if hit is not None:  # 문단 중간에 다음 주석 헤딩 — 분할
                cut, _ = hit
                pre = "\n".join(lines[:cut])
                # 기준 양식은 별첨 표지·설명 문구를 별도 주석 시트에 싣지 않고,
                # 별첨 내부 첫 번호 헤딩부터 새 시트를 시작한다.
                if pre.strip() and not _APPENDIX_RE.search(pre):
                    out.append(("para", pre, b.item))
                post = "\n".join(lines[cut + 1:])
                spill = (post, b.item) if post.strip() else None
                continue
            out.append(("para", text, b.item))
    return items


def _render_section(plan: SheetPlan, sheet_name: str, layout: SheetLayout, cursor: int,
                    specs_by_table: dict[str, list[FootingSpec]],
                    items: list) -> int:
    """섹션 렌더 항목을 layout.ops에 추가하고 다음 커서를 반환."""
    if plan.kind == "note" and sheet_name == "원문":
        # 주석 자체 시트에서는 헤더 바(2행)가 제목을 대신한다 — 원문에서만 본문 흐름에 유지
        layout.ops.append(("para", cursor, plan.display_title, None))
        cursor += 3
    for item in items:
        if item[0] == "para":
            layout.ops.append(("para", cursor, item[1], item[2] if len(item) > 2 else None))
            cursor += 1
            continue
        table: Table = item[1]
        grid_mod.expand(table)  # ncols 확정 (irregular 포함)
        placed = PlacedTable(table, sheet_name, cursor)
        layout.ops.append(("table", placed))
        cursor += len(table.rows)
        # 표 아래 검증 블록 행 예약
        below_specs = [s for s in specs_by_table.get(table.id, [])
                       if _placement(s, sheet_name) == "below"]
        if below_specs:
            groups: dict[int, None] = {}
            for s in below_specs:
                groups[s.target[0]] = None
            base = cursor + C.BELOW_GAP_BEFORE
            for gi, ti in enumerate(sorted(groups)):
                label_row = base + gi * (2 + C.BELOW_GAP_BETWEEN)
                layout.below[(table.id, ti)] = (label_row, label_row + 1)
            last_formula = base + (len(groups) - 1) * (2 + C.BELOW_GAP_BETWEEN) + 1
            cursor = last_formula + C.BELOW_GAP_BETWEEN + 1
    return cursor


def build(plans: list[SheetPlan],
          specs_by_table: dict[str, list[FootingSpec]]) -> tuple[list[SheetLayout], dict]:
    """시트별 레이아웃 + 표 배치 지도(placed_map: table_id → 자기 시트의 PlacedTable)."""
    layouts: list[SheetLayout] = []
    placed_map: dict[str, PlacedTable] = {}
    items_by_plan = compute_render_items(plans)

    # 원문 — 검증 없이 전체 이어 붙임
    wonmun = SheetLayout("원문", "wonmun", title="원문")
    cursor = C.BODY_START_ROW
    for plan in plans:
        cursor = _render_section(plan, "원문", wonmun, cursor, {}, items_by_plan[plan.name])
        cursor += C.SECTION_GAP  # 섹션 사이 빈행 2
    wonmun.n_rows = cursor
    layouts.append(wonmun)

    for plan in plans:
        lay = SheetLayout(plan.name, plan.kind,
                          title=plan.display_title or plan.name)
        end = _render_section(plan, plan.name, lay, C.BODY_START_ROW,
                              specs_by_table, items_by_plan[plan.name])
        lay.n_rows = end
        maxw = max((it[1].ncols for it in items_by_plan[plan.name] if it[0] == "table"),
                   default=1)
        # 본문이 B열부터 시작하므로 슬롯 시작도 한 칸 오른쪽
        lay.right_start = maxw + C.RIGHT_START_GAP + (C.CONTENT_START_COL - 1)
        for op in lay.ops:
            if op[0] == "table":
                placed_map[op[1].table.id] = op[1]
        layouts.append(lay)
    return layouts, placed_map


# ── Spec → Check 바인딩 ──────────────────────────────────────────────

def bind(layouts: list[SheetLayout], placed_map: dict,
         footing_specs: list[FootingSpec], cross_specs: list[CrossSpec],
         value_of) -> CheckSet:
    """그리드 좌표 Spec을 셀 좌표 Check로 변환. value_of(table_id, i, c) = 기록될 값.

    오른쪽 슬롯은 여기서 할당한다 — 풋팅 전부 먼저, 대사 다음 (기준 양식 순서)."""
    by_sheet = {lay.name: lay for lay in layouts}
    cs = CheckSet()

    def cellref(table_id: str, i: int, c: int) -> CellRef:
        pt = placed_map[table_id]
        return CellRef(pt.sheet, pt.row_of(i), c + C.CONTENT_START_COL,
                       value_of(table_id, i, c))

    for spec in footing_specs:
        pt = placed_map.get(spec.table_id)
        if pt is None:
            continue
        lay = by_sheet[pt.sheet]
        target = cellref(spec.table_id, *spec.target)
        terms = [(cellref(spec.table_id, j, cc), sign) for j, cc, sign in spec.terms]
        # 방어선: 항은 전부 숫자여야 한다 (텍스트 항 → 엑셀 #VALUE!)
        if not isinstance(target.value, (int, float, Decimal)) or any(
                not isinstance(cell.value, (int, float, Decimal)) for cell, _ in terms):
            continue
        placement = _placement(spec, pt.sheet)
        fc = FootingCheck(target, terms, spec.label, spec.kind, spec.axis, placement,
                          op=spec.op, digits=spec.digits, factor=spec.factor)
        if placement == "below":
            rows = lay.below.get((spec.table_id, spec.target[0]))
            if rows is None:
                continue  # 방어 — 예약 없으면 방출 생략
            fc.label_cell = CellRef(pt.sheet, rows[0], target.col, None)
            fc.formula_cell = CellRef(pt.sheet, rows[1], target.col, None)
        else:
            col = lay.alloc_slot(target.row)
            fc.label_cell = CellRef(pt.sheet, target.row, col, None)
            fc.formula_cell = CellRef(pt.sheet, target.row, col + 1, None)
        cs.footing.append(fc)

    for spec in cross_specs:
        dst_pt = placed_map.get(spec.dst_table)
        src_pt = placed_map.get(spec.src_table)
        if dst_pt is None or src_pt is None:
            continue
        lay = by_sheet[dst_pt.sheet]
        dst = cellref(spec.dst_table, *spec.dst_cell)
        src = [cellref(spec.src_table, i, c) for i, c in spec.src_cells]
        xr = CrossRef(src, dst, spec.label_path, spec.scale, spec.dst_scale, spec.rounder)
        col = lay.alloc_slot(dst.row)
        xr.link_cell = CellRef(dst_pt.sheet, dst.row, col, None)
        xr.formula_cell = CellRef(dst_pt.sheet, dst.row, col + 1, None)
        cs.cross.append(xr)

    return cs
