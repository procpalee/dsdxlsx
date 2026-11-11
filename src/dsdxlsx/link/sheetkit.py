"""정산표 시트를 쓰는 작은 도우미 - 머리말, 표, 입력칸 표시.

모든 시트가 같은 자리를 쓴다: A열은 여백, 2~3행 머리말, 5행부터 표(B열 시작).
다른 모듈의 수식이 이 자리(TB!$E:$E 등)에 기대므로 자리를 바꾸지 않는다.
"""
from __future__ import annotations

import re

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ..theme import load_theme

FIRST_COL = 2
HDR_ROW = 5
NUMFMT = '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-'
INPUT_RGB = "FFFBEF"                 # 사람이 넣는 칸(크림) - 왕복 시험이 이 색으로 입력칸을 찾는다
TODO_RGB = "FFFF00"                  # 확인이 필요한 칸(노랑)

_NUM_RE = re.compile(r"(?m)^([ \t]*)(\d{1,2}\.(?!\d)|\(\d{1,2}\)|\d{1,2}\))[ \t]*(?=\S)")


def num_space(text):
    """줄 맨 앞 번호('1.' · '(1)') 뒤 공백을 한 칸으로."""
    return _NUM_RE.sub(lambda m: m.group(1) + m.group(2) + " ", text) if isinstance(text, str) else text


def _style():
    th = load_theme()
    c = th["colors"]
    thin = Side(style="thin", color=c["border_in"])
    return {"font": th["font"], "size": th["size_body"], "box": Border(left=thin, right=thin, top=thin, bottom=thin),
            "head": PatternFill("solid", fgColor=c["header_fill"]), "head_text": c["header_text"]}


def write_header(ws, meta: dict, title: str) -> int:
    """시트 위 머리말(B2:G3) - 회사명 · 문서 번호 · 시트 / 기준일 · 문서 이름 · 작성자. 반환: 표 머리글 행."""
    st = _style()
    ws.column_dimensions["A"].width = 2.0
    grid = [[("회사명", meta.get("company")), ("번호", meta.get("ref")), ("시트", title)],
            [("기준일", meta.get("date")), ("문서", meta.get("title")), ("작성자", meta.get("author"))]]
    for r, line in enumerate(grid, 2):
        for i, (label, value) in enumerate(line):
            lc = ws.cell(r, FIRST_COL + i * 2, label)
            lc.font = Font(name=st["font"], size=st["size"], bold=True)
            vc = ws.cell(r, FIRST_COL + i * 2 + 1, value if value not in (None, "") else None)
            vc.font = Font(name=st["font"], size=st["size"])
            for c in (lc, vc):
                c.border = st["box"]
                c.alignment = Alignment(horizontal="left", vertical="center")
    return HDR_ROW


def write_table(ws, headers, rows, *, header_row: int = HDR_ROW, number_cols: dict | None = None) -> None:
    """headers 와 rows(2차원)를 B열부터 쓴다. number_cols = {열 순번(0부터): 'accounting' 또는 서식 글}."""
    st = _style()
    number_cols = number_cols or {}
    for i, h in enumerate(headers):
        c = ws.cell(header_row, FIRST_COL + i, h)
        c.font = Font(name=st["font"], size=st["size"], bold=True, color=st["head_text"])
        c.fill = st["head"]
        c.border = st["box"]
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for j, row in enumerate(rows, 1):
        for i in range(len(headers)):
            v = row[i] if i < len(row) else None
            c = ws.cell(header_row + j, FIRST_COL + i, v)
            c.font = Font(name=st["font"], size=st["size"])
            c.border = st["box"]
            fmt = number_cols.get(i)
            if fmt:
                c.number_format = NUMFMT if fmt == "accounting" else fmt
            if fmt and fmt != "@" and not isinstance(v, str):
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(v, str) and v.startswith("=") and fmt:
                c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")


def write_section_bar(ws, row: int, text: str, *, first_col: int = FIRST_COL, last_col: int = 9) -> int:
    """구획 제목 줄(주석 번호와 제목). 반환: 다음 표 머리글 행."""
    st = _style()
    for col in range(first_col, last_col + 1):
        c = ws.cell(row, col)
        c.fill = st["head"]
        c.font = Font(name=st["font"], size=st["size"], bold=True, color=st["head_text"])
        c.alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row, first_col, text)
    return row + 2


def _fill(ws, rgb: str, refs) -> None:
    fill = PatternFill("solid", fgColor=rgb)
    for ref in refs:
        cells = ws[ref]
        if not isinstance(cells, tuple):
            cells = ((cells,),)
        for line in cells:
            for c in (line if isinstance(line, tuple) else (line,)):
                c.fill = fill


def mark_cells(ws, *, input=(), todo=()) -> None:
    """입력칸(크림)과 확인할 칸(노랑)을 칠한다."""
    if input:
        _fill(ws, INPUT_RGB, [input] if isinstance(input, str) else input)
    if todo:
        _fill(ws, TODO_RGB, [todo] if isinstance(todo, str) else todo)


def is_input(cell) -> bool:
    """입력칸(크림)으로 칠한 칸인가."""
    fg = getattr(getattr(cell, "fill", None), "fgColor", None)
    return bool(fg is not None and str(getattr(fg, "rgb", "")).upper().endswith(INPUT_RGB))


def col_letter(i: int) -> str:
    return get_column_letter(i)
