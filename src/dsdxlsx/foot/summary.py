"""총괄표 — 시트 목록 + 시트별 푸팅/크로스 오류 COUNTIF 집계 (실측 레이아웃)."""
from __future__ import annotations

from openpyxl.styles import Alignment
from openpyxl.worksheet.worksheet import Worksheet

from . import constants as C
from . import styles as S
from .checks import CheckSet
from .layout import SheetLayout

_HEAD_FILL = S.FILL_DARK      # 테마 primary(진회색) 헤더
_HEAD_FONT = S.FONT_HEAD_INV


def _countif_chain(ws: Worksheet, row: int, col: int, sheet: str, addrs: list[str]):
    """=COUNTIF(...)+... 체인. 수식 길이 한계를 넘으면 오른쪽 예비 열(M~)에 분할."""
    if not addrs:
        ws.cell(row=row, column=col, value=0)
        return
    q = sheet if sheet.isalpha() else f"'{sheet}'"
    parts = [f'COUNTIF({q}!{a},"<>TRUE")' for a in addrs]
    chunks: list[str] = []
    cur = ""
    for p in parts:
        if cur and len(cur) + len(p) + 1 > C.FORMULA_MAX - 100:
            chunks.append(cur)
            cur = p
        else:
            cur = f"{cur}+{p}" if cur else p
    chunks.append(cur)
    if len(chunks) == 1:
        ws.cell(row=row, column=col, value=f"={chunks[0]}")
        return
    # 분할: 본 셀 = 첫 조각 + 예비 열 참조, 예비 열(M부터)에 나머지 조각
    spill_col = 13  # M
    refs = []
    for k, chunk in enumerate(chunks[1:]):
        cell = ws.cell(row=row, column=spill_col + k, value=f"={chunk}")
        cell.font = S.FONT_SMALL_DIM
        refs.append(cell.coordinate)
    ws.cell(row=row, column=col, value="=" + chunks[0] + "".join(f"+{r}" for r in refs))


def write(ws: Worksheet, layouts: list[SheetLayout], checkset: CheckSet,
          display_titles: dict[str, str]) -> None:
    cells = checkset.cells_by_sheet()
    plans = [lay for lay in layouts if lay.kind != "wonmun"]

    ws.cell(row=2, column=2, value="시트정보").font = S.FONT_BOLD
    ws.cell(row=2, column=7, value="푸팅/크로스 에러 사항").font = S.FONT_BOLD
    for col, text in ((2, "시트"), (3, "제목")):
        c = ws.cell(row=4, column=col, value=text)
        c.fill, c.font = _HEAD_FILL, _HEAD_FONT
        c.alignment = Alignment(horizontal="center")
    for col, text in ((7, "시트"), (8, "시트 링크"), (9, "푸팅 오류"),
                      (10, "크로스 오류"), (11, "총합")):
        c = ws.cell(row=4, column=col, value=text)
        c.fill, c.font = _HEAD_FILL, _HEAD_FONT
        c.alignment = Alignment(horizontal="center")

    # 좌측: 원문 + 전체 시트 목록
    left = [("원문", display_titles.get("원문", ""))] + [
        (lay.name, display_titles.get(lay.name, "")) for lay in plans
    ]
    for k, (name, title) in enumerate(left):
        r = 5 + k
        link = ws.cell(row=r, column=2, value=f'=HYPERLINK("#\'{name}\'!A1","{name}")')
        link.font = S.FONT_LINK
        ws.cell(row=r, column=3, value=title)

    # 우측: 시트별 오류 카운터 (원문 제외)
    for k, lay in enumerate(plans):
        r = 5 + k
        ws.cell(row=r, column=7, value=lay.name)
        link = ws.cell(row=r, column=8,
                       value=f'=HYPERLINK("#\'{lay.name}\'!$A$1","시트 바로가기")')
        link.font = S.FONT_LINK
        sheet_cells = cells.get(lay.name, {"footing": [], "cross": []})
        _countif_chain(ws, r, 9, lay.name, sheet_cells["footing"])
        _countif_chain(ws, r, 10, lay.name, sheet_cells["cross"])
        ws.cell(row=r, column=11, value=f"=I{r}+J{r}")
    # 오류 건수 > 0 셀 빨강 음영 (사용자 지정)
    if plans:
        from openpyxl.formatting.rule import CellIsRule
        ws.conditional_formatting.add(
            f"I5:K{4 + len(plans)}",
            CellIsRule(operator="greaterThan", formula=["0"], fill=S.FILL_BAD))

    for col, w in (("B", 10), ("C", 35), ("G", 8), ("H", 15), ("I", 12), ("J", 12), ("K", 10)):
        ws.column_dimensions[col].width = w
