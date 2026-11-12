"""왕복 시험 결과를 엑셀 한 장으로 - 사람이 엑셀에서 바로 보는 용도(`link test -o`).

위에는 재무제표 · 주석별로 같음 · 다름 · 빈칸 칸 수, 아래에는 어긋난 칸 목록(다름 = 결과 DSD 의 당기 칸이 전기 금액과 다름,
빈칸 = 당기 칸이 비어 있음). 어긋난 칸이 없으면 목록 자리에 '해당 없음' 한 줄을 둔다.
"""
from __future__ import annotations

from openpyxl import Workbook

from .sheetkit import HDR_ROW, num_space, write_header, write_table

SHEET = "왕복시험"


def rows_of(res: dict, notes: bool) -> list:
    out = []
    kinds = [("재무제표", "다름", res["diff"]), ("재무제표", "빈칸", res["blank"])]
    if notes:
        kinds += [("주석", "다름", res["note_diff"]), ("주석", "빈칸", res["note_blank"])]
    for where, kind, items in kinds:
        for x in items:
            out.append([where, x["section"], kind, num_space(x["label"]), x["built"], x["expected"]])
    return out


def write(out: str, res: dict, *, side: str, notes: bool, company: str = "", author: str | None = None) -> int:
    """결과 통합문서를 쓴다. 반환: 어긋난 칸 수."""
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    write_header(ws, {"company": company, "title": "정산표 왕복 시험(%s)" % side, "author": author}, SHEET)
    summary = [["재무제표", res["same"], len(res["diff"]), len(res["blank"])]]
    if notes:
        summary.append(["주석", res["note_same"], len(res["note_diff"]), len(res["note_blank"])])
    write_table(ws, ["구분", "같음", "다름", "빈칸"], summary, header_row=HDR_ROW, number_cols={1: "#,##0", 2: "#,##0", 3: "#,##0"})
    rows = rows_of(res, notes)
    top = HDR_ROW + len(summary) + 2
    write_table(ws, ["구분", "위치", "결과", "행", "결과 DSD 당기", "전기 금액"],
                rows or [["-", "해당 없음", None, None, None, None]], header_row=top,
                number_cols={4: "accounting", 5: "accounting"})
    for col, w in zip("BCDEFG", (12, 30, 10, 44, 18, 18)):   # C · E 는 머리말의 회사명 · 문서 이름도 담는다
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False
    if author:
        wb.properties.creator = author
    wb.save(out)
    return len(rows)
