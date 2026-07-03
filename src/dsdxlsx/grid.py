"""COLGROUP + COLSPAN/ROWSPAN → 그리드 전개, 행 조작 가드 판정."""
from __future__ import annotations

from dataclasses import dataclass

from .model import Node, Table


@dataclass
class GridCell:
    kind: str  # 'cell' | 'hspan' | 'vspan'  (hspan/vspan = 병합 연속 위치)
    node: Node


def expand(table: Table) -> list[list[GridCell | None]]:
    """각 TR을 ncols 폭 그리드로 전개. 결과는 table.grid에 캐시."""
    if table.grid is not None:
        return table.grid
    ncols = table.ncols
    pending: dict[int, list] = {}  # col -> [남은 행 수, owner]
    grid: list[list[GridCell | None]] = []
    for row in table.rows:
        grow: list[GridCell | None] = [None] * ncols
        for c in list(pending):
            remaining, owner = pending[c]
            grow[c] = GridCell("vspan", owner)
            if remaining <= 1:
                del pending[c]
            else:
                pending[c][0] = remaining - 1
        col = 0
        for cell in row.cells:
            while col < ncols and grow[col] is not None:
                col += 1
            span = cell.colspan
            if col + span > ncols:
                # 열 초과 — 그리드 확장 후 이상 테이블로 마킹 (행 조작 금지)
                table.irregular = True
                extra = col + span - ncols
                ncols = col + span
                for g in grid:
                    g.extend([None] * extra)
                grow.extend([None] * extra)
            for cc in range(col, col + span):
                grow[cc] = GridCell("cell" if cc == col else "hspan", cell)
                if cell.rowspan > 1:
                    pending[cc] = [cell.rowspan - 1, cell]
            col += span
        grid.append(grow)
    if ncols != table.ncols:
        table.ncols = ncols
    table.grid = grid
    return grid


def cell_grid_col(table: Table, row_index: int, cell: Node) -> int:
    """해당 셀의 그리드 열(0-base) 반환."""
    grid = expand(table)
    for c, gc in enumerate(grid[row_index]):
        if gc is not None and gc.kind == "cell" and gc.node is cell:
            return c
    raise ValueError(f"그리드에서 셀을 찾지 못함: {cell.id}")


def span_crosses_after(table: Table, row_index: int) -> bool:
    """row_index 행과 다음 행 사이 경계를 가로지르는 ROWSPAN이 있는가 (삽입 가드)."""
    grid = expand(table)
    if row_index + 1 >= len(grid):
        return False
    return any(gc is not None and gc.kind == "vspan" for gc in grid[row_index + 1])


def row_in_span(table: Table, row_index: int) -> bool:
    """해당 행이 ROWSPAN 병합에 관여하는가 (삭제 가드): vspan 연속 위치를 포함하거나
    ROWSPAN>1 셀을 소유."""
    grid = expand(table)
    for gc in grid[row_index]:
        if gc is None:
            continue
        if gc.kind == "vspan":
            return True
        if gc.kind == "cell" and gc.node.rowspan > 1:
            return True
    return False
