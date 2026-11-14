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
NUMFMT = "#,###,##0;[Red](#,###,##0);-"   # 음수는 빨간 괄호, 0 은 '-' (export · 대사 결과와 같은 서식)
INPUT_RGB = "FFFBEF"                 # 사람이 넣는 칸(크림) - 왕복 시험이 이 색으로 입력칸을 찾는다
TODO_RGB = "FFFF00"                  # 확인이 필요한 칸(노랑)
SECTION_RGB = "393939"               # 구획 제목 줄(진회색 바탕에 흰 글씨)
LABELS = ("회사명", "번호", "시트", "기준일", "문서", "작성자")   # 머리말 이름표 - meta["labels"] 로 바꿀 수 있다

_NUM_RE = re.compile(r"(?m)^([ \t]*)(\d{1,2}\.(?!\d)|\(\d{1,2}\)|\d{1,2}\))[ \t]*(?=\S)")


def num_space(text):
    """줄 맨 앞 번호('1.' · '(1)') 뒤 공백을 한 칸으로."""
    return _NUM_RE.sub(lambda m: m.group(1) + m.group(2) + " ", text) if isinstance(text, str) else text


def _style():
    th = load_theme()
    c = th["colors"]
    return {"font": th["font"], "size": th["size_body"], "thin": Side(style="thin", color=c["border_in"]),
            "thick": Side(style="medium", color=c["border_out"]), "head": PatternFill("solid", fgColor=c["header_fill"]),
            "head_text": c["header_text"]}


def write_header(ws, meta: dict, title: str) -> int:
    """시트 위 머리말(B2:G3) - 회사명 · 번호 · 시트 / 기준일 · 문서 · 작성자. 이름표는 meta["labels"](6개)로 바꾼다.
    위아래는 굵은 줄, 가운데는 가는 줄, 양 끝 세로줄은 없다. 반환: 표 머리글 행."""
    st = _style()
    ws.column_dimensions["A"].width = 2.0
    labels = list(meta.get("labels") or LABELS)
    values = [meta.get("company"), meta.get("ref"), title, meta.get("date"), meta.get("title"), meta.get("author")]
    last = FIRST_COL + 5
    for k in range(6):
        r = 2 + k // 3
        for j, (text, bold) in enumerate(((labels[k], True), (values[k], False))):
            col = FIRST_COL + (k % 3) * 2 + j
            c = ws.cell(r, col, text if text not in (None, "") else "")
            c.font = Font(name=st["font"], size=st["size"], bold=bold, color="000000")
            c.border = Border(left=None if col == FIRST_COL else st["thin"], right=None if col == last else st["thin"],
                              top=st["thick"] if r == 2 else st["thin"], bottom=st["thick"] if r == 3 else st["thin"])
            c.alignment = Alignment(horizontal="left", vertical="center")
    return HDR_ROW


def _text_width(text: str) -> float:
    """엑셀 열너비 단위로 어림 - 한글·전각은 2, 그 밖은 1.1."""
    return sum(2.0 if ord(ch) > 0x2E80 else 1.1 for ch in str(text))


def fit_header(wb) -> None:
    """머리말(2~3행) 값 칸이 열보다 넓으면 그 열을 넓힌다 - 열너비를 다 정한 뒤 부른다."""
    for ws in wb.worksheets:
        for r in (2, 3):
            for col in (FIRST_COL + 1, FIRST_COL + 3, FIRST_COL + 5):
                v = ws.cell(r, col).value
                if v in (None, "") or str(v).startswith("="):
                    continue
                letter = get_column_letter(col)
                need = _text_width(v) + 2
                if (ws.column_dimensions[letter].width or 8.43) < need:
                    ws.column_dimensions[letter].width = round(need, 1)


def write_table(ws, headers, rows, *, header_row: int = HDR_ROW, number_cols: dict | None = None) -> None:
    """headers 와 rows(2차원)를 B열부터 쓴다. number_cols = {열 순번(0부터): 'accounting' 또는 서식 글}.
    가로줄만 긋는다 - 머리글 위와 마지막 줄 아래는 굵게, 나머지는 가늘게."""
    st = _style()
    number_cols = number_cols or {}
    for i, h in enumerate(headers):
        c = ws.cell(header_row, FIRST_COL + i, h)
        c.font = Font(name=st["font"], size=st["size"], bold=True, color=st["head_text"])
        c.fill = st["head"]
        c.border = Border(top=st["thick"], bottom=st["thin"])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for j, row in enumerate(rows, 1):
        bottom = st["thick"] if j == len(rows) else st["thin"]
        for i in range(len(headers)):
            v = row[i] if i < len(row) else None
            c = ws.cell(header_row + j, FIRST_COL + i, v)
            c.font = Font(name=st["font"], size=st["size"], color="000000")
            c.border = Border(top=st["thin"], bottom=bottom)
            fmt = number_cols.get(i)
            if fmt:
                c.number_format = NUMFMT if fmt == "accounting" else fmt
            # 정렬은 값으로 정한다 - 숫자는 오른쪽, 글자는 왼쪽. 빈 칸과 수식은 숫자 열이면 오른쪽
            number = isinstance(v, (int, float)) and not isinstance(v, bool)
            text = isinstance(v, str) and not v.startswith("=")
            right = number or (not text and bool(fmt) and fmt != "@")
            c.alignment = Alignment(horizontal="right" if right else "left", vertical="center")


def write_section_bar(ws, row: int, text: str, *, first_col: int = FIRST_COL, last_col: int = 9) -> int:
    """구획 제목 줄(주석 번호와 제목). 반환: 다음 표 머리글 행."""
    st = _style()
    for col in range(first_col, last_col + 1):
        c = ws.cell(row, col)
        c.fill = PatternFill("solid", fgColor=SECTION_RGB)
        c.font = Font(name=st["font"], size=st["size"], bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row, first_col, text)
    return row + 2


def _fill(ws, rgb: str, refs) -> None:
    fill = PatternFill("solid", fgColor=rgb)
    st = _style()
    for ref in refs:
        cells = ws[ref]
        if not isinstance(cells, tuple):
            cells = ((cells,),)
        for line in cells:
            for c in (line if isinstance(line, tuple) else (line,)):
                c.fill = fill
                if c.value is None:                      # 빈 입력칸의 글자색은 자동(값을 넣으면 엑셀 기본색)
                    c.font = Font(name=st["font"], size=st["size"])


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
