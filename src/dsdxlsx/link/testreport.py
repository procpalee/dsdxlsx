"""왕복 시험 결과를 엑셀 한 장으로 - 사람이 엑셀에서 바로 보는 용도(`link test -o`).

위에는 재무제표 · 주석별로 같음 · 다름 · 빈칸 칸 수, 그 아래 어긋난 칸의 원인별 칸 수와 원인의 뜻, 맨 아래 어긋난
칸 목록(다름 = 결과 DSD 의 당기 칸이 전기 금액과 다름, 빈칸 = 당기 칸이 비어 있음). 목록은 '확인 필요'를 맨 위에
둔다 - '연결점검 항목'과 '위 줄을 따라 바뀜'은 정산표를 만들 때 이미 알린 줄과 그 합계이고, '당기 표와 전기 표가
따로'인 주석(변동표 등)은 시험이 같은 자리 칸끼리만 비교하는 표다. 어긋난 칸이 없으면 목록 자리에 '해당 없음' 한 줄.
"""
from __future__ import annotations

from openpyxl import Workbook

from .sheetkit import FIRST_COL, HDR_ROW, fit_header, num_space, write_header, write_table

SHEET = "왕복시험"
ORDER = ("확인 필요", "주석 머리글에 묶음 이름이 빠짐", "채우지 못함", "연결점검 항목", "위 줄을 따라 바뀜",
         "주석 - 당기 표와 전기 표가 따로")
NOTE = {
    "확인 필요": "정산표의 연결(TB 짝, 합계 수식, 주석 칸)이 끊겼을 수 있습니다.",
    "주석 머리글에 묶음 이름이 빠짐": "예전 판으로 만든 정산표입니다. dsdxlsx link fix-headers 로 고치거나 정산표를 다시 만드세요.",
    "채우지 못함": "DSD 를 채울 때 정산표에서 값을 고르지 못했습니다. 이유는 목록 맨 오른쪽 열에 있습니다.",
    "연결점검 항목": "정산표를 만들 때 '연결점검' 시트에 올린 과목입니다(TB 전기 금액이 전기 DSD 와 다름).",
    "위 줄을 따라 바뀜": "위에서 어긋난 과목 줄을 더하거나 이어 받는 줄(합계, 순이익 등)입니다. 과목 줄을 고치면 같이 맞습니다.",
    "주석 - 당기 표와 전기 표가 따로": "변동표처럼 당기 표와 전기 표가 나뉜 주석이라 같은 자리 칸끼리만 비교했습니다.",
}


def rows_of(res: dict, notes: bool) -> list:
    out = []
    kinds = [("재무제표", "다름", res["diff"]), ("재무제표", "빈칸", res["blank"])]
    if notes:
        kinds += [("주석", "다름", res["note_diff"]), ("주석", "빈칸", res["note_blank"])]
    for where, kind, items in kinds:
        for x in items:
            place = x["section"] if where == "재무제표" else "주석 %s" % x["section"]
            out.append([x.get("cause") or "확인 필요", place, kind, num_space(x["label"]), x["built"], x["expected"],
                        x.get("why") or None])
    rank = {k: i for i, k in enumerate(ORDER)}
    out.sort(key=lambda r: rank.get(r[0], len(ORDER)))
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
    top = HDR_ROW + len(summary) + 2
    causes = res.get("causes") or {}
    if causes:
        crow = [[k, causes[k]] for k in ORDER if causes.get(k)]
        write_table(ws, ["원인", "칸 수"], crow, header_row=top, number_cols={1: "#,##0"})
        top += len(crow) + 2
        for k, _n in crow:                         # 원인의 뜻 - 표 밖 설명 줄(오른쪽 빈 칸으로 이어 보인다)
            ws.cell(top, FIRST_COL, "%s: %s" % (k, NOTE[k]))
            top += 1
        top += 1
    rows = rows_of(res, notes)
    write_table(ws, ["원인", "위치", "결과", "행", "결과 DSD 당기", "전기 금액", "못 채운 이유"],
                rows or [["해당 없음", None, None, None, None, None, None]], header_row=top,
                number_cols={4: "accounting", 5: "accounting"})
    # B · C 는 원인 · 위치와 머리말의 회사명, E 는 행 이름과 머리말의 문서 이름을 담는다
    for col, w in zip("BCDEFGH", (30, 30, 10, 44, 18, 18, 48)):
        ws.column_dimensions[col].width = w
    fit_header(wb)
    ws.sheet_view.showGridLines = False
    if author:
        wb.properties.creator = author
    wb.save(out)
    return len(rows)
