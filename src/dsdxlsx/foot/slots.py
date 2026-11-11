"""만든 워크북에 대사 칸을 덧붙일 때 쓰는 도우미 - 빈 슬롯 찾기, 총괄표 오류 카운터에 더하기."""
from __future__ import annotations

from . import constants as C


def slot_for(ws, row: int, right_start: int) -> int:
    """row 의 오른쪽 검증 슬롯 중 비어 있는 첫 칸(링크 칸) 열 번호."""
    k = 0
    while True:
        col = right_start + C.CHECK_SLOT_WIDTH * k
        if ws.cell(row, col).value is None and ws.cell(row, col + 1).value is None:
            return col
        k += 1
        if k > 40:
            return col


def summary_add(ws_sum, sheet: str, addr: str) -> None:
    """총괄표 크로스 오류 COUNTIF 체인에 새 셀 추가 (길이 초과 시 예비 열 M~에 분할)."""
    q = sheet if sheet.isalpha() else f"'{sheet}'"
    term = f'COUNTIF({q}!{addr},"<>TRUE")'
    for r in range(5, ws_sum.max_row + 1):
        if str(ws_sum.cell(r, 7).value) != sheet:
            continue
        cell = ws_sum.cell(r, 10)
        v = cell.value
        if isinstance(v, str) and v.startswith("="):
            if len(v) + len(term) + 1 < C.FORMULA_MAX - 100:
                cell.value = f"{v}+{term}"
                return
            col = 13
            while ws_sum.cell(r, col).value is not None:
                col += 1
            spill = ws_sum.cell(r, col, value=f"={term}")
            cell.value = f"{v}+{spill.coordinate}"
            return
        cell.value = f"={term}"
        return
