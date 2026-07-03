"""check 결과를 편집용 xlsx 안에 직접 기록 — 살아있는 엑셀 수식으로.

- 각 표 오른쪽(마지막 열 + 빈 열 1개)에 검사 대상 열별 '검증' 열 추가.
  검증 셀은 정적 값이 아니라 **엑셀 수식**: `=IF(대상=(구성셀 합),TRUE,대상-(합))`
  → 사용자가 엑셀에서 값을 고치면 즉시 재검증된다 (TRUE ↔ 차이 금액).
  불일치 표시는 조건부 서식(TRUE가 아니면 빨강 음영)이라 이것도 살아있다.
- 재무제표↔주석 대사(C4)·주석↔주석 대사(C5)는 **각 주석 시트 하단 블록**에 표시.
  C4는 `=IF(OR(COUNTIF(본문범위, 재무제표셀), …),TRUE,"미확인")` 수식 — COUNTIF
  범위를 본문(표 영역)까지로 제한해 블록 자신을 세는 자기참조를 차단한다.
- 맨 앞에 '검증요약' 시트 생성(재실행 시 교체): 검사별 집계 + 불일치 목록
  (해당 셀 하이퍼링크) + 표기(C6) 발견 사항.

재변환(import) 안전성: 표 그리드 밖 열은 import가 무시하고, 검증요약·주석 블록의
행은 A열 '!s…'/'!v…' ID라 건너뛴다 — 검증 셀이 있어도 import 결과는 동일하다.
행 좌표는 import 파이프라인의 row_positions(시트 행 ↔ 편집 반영 문서의 TR,
1:1 순서 대응)를 사용 — 행 추가/삭제 후에도 어긋나지 않는다.
"""
from __future__ import annotations

import re
from datetime import datetime
from decimal import Decimal

from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter

from .check import CHECK_NAMES, CheckReport, fmt
from .export_xlsx import (
    FILL_HEAD, FILL_SUBHEAD, FONT, FONT_HEAD, FONT_SUBHEAD, FONT_TITLE, NUMFMT_INT,
    THIN, _C, _T, refined_sections,
)
from .model import Document

SUMMARY_SHEET = "검증요약"
_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
_FILL_BAD = PatternFill("solid", fgColor="FDE9E9")
_FONT_BAD = Font(name=_T["font"], size=_T["size_body"], bold=True, color=_C["accent"])
_FONT_OK = Font(name=_T["font"], size=_T["size_body"], color=_C["primary_dark"])
_NOTE_BLOCK_COLS = 12  # 주석 대사 블록이 쓰는 열 폭 (B~M)


def _num(value):
    return float(value) if isinstance(value, Decimal) else value


def _cell_positions(doc: Document, row_positions: dict) -> dict[str, tuple[str, int]]:
    """편집 반영 문서의 행 ID → (시트, 엑셀 행)."""
    out: dict[str, tuple[str, int]] = {}
    for tid, table in doc.tables.items():
        positions = row_positions.get(tid)
        if positions is None or len(positions) != len(table.rows):
            continue  # 방어: 대응이 안 맞으면 그 표는 주석 생략
        for row, pos in zip(table.rows, positions):
            out[row.id] = pos
    return out


def _note_sheet_map(doc: Document, pos_of: dict) -> dict[int, str]:
    """주석 번호 → 실제 시트명 (그 주석 첫 표의 첫 행 위치로 판별)."""
    out: dict[int, str] = {}
    for sec in refined_sections(doc):
        if sec.note_no is None:
            continue
        for block in sec.blocks:
            if block.kind == "table" and block.item.rows:
                loc = pos_of.get(block.item.rows[0].id)
                if loc:
                    out.setdefault(sec.note_no, loc[0])
                break
    return out


def _find_block_start(ws) -> int:
    """기존 대사 블록('!v' ID 행)의 시작 행. 없으면 내용 끝 + 2."""
    for row in ws.iter_rows(min_col=1, max_col=1):
        v = row[0].value
        if isinstance(v, str) and v.startswith("!v"):
            return row[0].row
    return ws.max_row + 2


def _clear_block(ws, start: int):
    # 이전 블록의 병합부터 해제 (MergedCell은 값 대입 불가)
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= start:
            ws.unmerge_cells(str(rng))
    for r in range(start, ws.max_row + 1):
        for c in range(1, _NOTE_BLOCK_COLS + 2):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.comment = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()


def _sum_formula(x, pos_of, sheet: str, target: str) -> str | None:
    """검산 수식: =IF(대상=(±구성셀…),TRUE,대상-(…)). 구성 셀이 다른 시트면 포기."""
    parts = []
    for rid, gc, sign in x.terms:
        loc = pos_of.get(rid)
        if loc is None or loc[0] != sheet:
            return None
        addr = f"{get_column_letter(2 + gc)}{loc[1]}"
        parts.append(("-" if sign < 0 else "+") + addr)
    if not parts:
        return None
    expr = "".join(parts).lstrip("+")
    formula = f"=IF({target}=({expr}),TRUE,{target}-({expr}))"
    return formula if len(formula) < 8000 else None


def _lookup_formula(x, note_ranges: dict, target: str) -> str | None:
    """C4 대사 수식: 참조 주석 시트 본문 범위에서 값(±)을 COUNTIF로 탐색."""
    conds = []
    for n in x.refs:
        rng = note_ranges.get(n)
        if rng:
            conds += [f"COUNTIF({rng},{target})>0", f"COUNTIF({rng},-{target})>0"]
    if not conds:
        return None
    return f'=IF(OR({",".join(conds)}),TRUE,"미확인")'


def _style(cell, font=FONT, border=True, align=None, numfmt=None):
    cell.font = font
    if border:
        cell.border = _BORDER
    if align:
        cell.alignment = align
    if numfmt:
        cell.number_format = numfmt
    return cell


def annotate(xlsx_path: str, report: CheckReport, doc: Document, row_positions: dict) -> dict:
    wb = load_workbook(xlsx_path)
    pos_of = _cell_positions(doc, row_positions)
    note_sheets = _note_sheet_map(doc, pos_of)

    # 재실행 시 조건부 서식 중복 누적 방지 (이 파일의 조건부 서식은 검증 열 전용)
    for ws in wb.worksheets:
        ws.conditional_formatting = ConditionalFormattingList()

    # 주석 시트별 대사 블록 시작 행(기존 블록 재사용)과 본문 COUNTIF 범위를 먼저 확정
    block_start: dict[str, int] = {}
    note_ranges: dict[int, str] = {}
    for n, sheet in note_sheets.items():
        if sheet not in block_start:
            block_start[sheet] = _find_block_start(wb[sheet])
            _clear_block(wb[sheet], block_start[sheet])
        note_ranges[n] = f"'{sheet}'!$B$1:$Z${max(block_start[sheet] - 2, 1)}"

    # ── 표 검증 배치 ────────────────────────────────────────────
    # 가로 검산(행 합산: ROWPART·합계열·기초→기말 가로 흐름) → 오른쪽 '합계검증' 열
    # 대사(C4) → 오른쪽 '대사검증' 열
    # 세로 검산(열 합산: 자동합계·그룹·합계행·계산라인·세로 흐름) → 표 바로 아래 검증 행
    by_table: dict[str, list] = {}
    for x in report.results:
        by_table.setdefault(x.table_id, []).append(x)

    written = 0
    addr_of: dict[int, tuple[str, str]] = {}  # id(Result) → (시트, 셀주소) — 요약 링크용

    def _cf(ws, rng, anchor):
        ws.conditional_formatting.add(
            rng,
            FormulaRule(formula=[f'AND({anchor}<>"",{anchor}<>TRUE)'],
                        fill=_FILL_BAD, font=_FONT_BAD),
        )

    def _write_check_cell(cell, xs, sheet, target):
        """검사 목록(xs)을 셀 하나에 기록: 수식 우선, 불가 시 정적 값."""
        primary = next((x for x in xs if not x.ok), xs[0])
        formula = None
        if primary.terms:
            formula = _sum_formula(primary, pos_of, sheet, target)
        elif primary.check == "C4" and not (primary.ok and "조합" in primary.note):
            formula = _lookup_formula(primary, note_ranges, target)
        if formula:
            cell.value = formula
            cell.number_format = NUMFMT_INT  # 숫자(차이)일 때만 적용됨
        elif primary.ok:
            cell.value = True
        else:
            cell.value = (_num(primary.stated) - _num(primary.computed)
                          if primary.computed is not None else "미확인")
            if primary.computed is not None:
                cell.number_format = NUMFMT_INT
        cell.comment = None  # 사용자 요청: 검증 셀 메모 없음
        _style(cell, _FONT_OK, align=Alignment(horizontal="center", vertical="center"))

    for tid, results in by_table.items():
        table = doc.tables.get(tid)
        if table is None or not table.rows or table.rows[0].id not in pos_of:
            continue
        head_sheet, head_row = pos_of[table.rows[0].id]
        last_row = max(pos_of[r.id][1] for r in table.rows if r.id in pos_of)
        ws = wb[head_sheet]

        col_axis = [x for x in results if x.axis == "col"]
        row_sums = [x for x in results if x.axis == "row" and x.check != "C4"]
        cross = [x for x in results if x.check == "C4"]

        # 오른쪽 검증 열: 합계검증(세로: 데이터열별) → 합계검증(가로) → 대사검증.
        # 세로검증도 각 합계·소계 행에 개별로 배치한다 (한 셀에 AND로 몰지 않음).
        vcol = 2 + table.ncols + 1
        col_map: dict[tuple, int] = {}  # (헤더, grid_col) → 엑셀 열

        def _add_cols(header, group, with_subhdr):
            nonlocal vcol
            hdr_of: dict[int, str] = {}
            for x in group:
                hdr_of.setdefault(x.grid_col, x.col_header)
            for gc in sorted(hdr_of):
                col_map[(header, gc)] = vcol
                letter = get_column_letter(vcol)
                sub = hdr_of[gc] if with_subhdr else ""
                title = f"{header}\n{sub}" if sub else header
                h = ws.cell(row=head_row, column=vcol, value=title)
                _style(h, FONT_SUBHEAD, align=Alignment(
                    horizontal="center", vertical="center", wrap_text=True))
                h.fill = FILL_SUBHEAD
                ws.column_dimensions[letter].width = 12
                _cf(ws, f"{letter}{head_row + 1}:{letter}{last_row}",
                    f"{letter}{head_row + 1}")
                vcol += 1

        # 세로는 데이터 열마다 별도 열(라벨에 기간 표기), 가로·대사는 단순 헤더
        _add_cols("합계검증(세로)", col_axis, True)
        _add_cols("합계검증(가로)", row_sums, False)
        _add_cols("대사검증", cross, False)

        for header, group in (("합계검증(세로)", col_axis),
                              ("합계검증(가로)", row_sums),
                              ("대사검증", cross)):
            merged: dict[tuple, list] = {}
            for x in group:
                merged.setdefault((x.row_id, x.grid_col), []).append(x)
            for (row_id, gc), xs in merged.items():
                if row_id not in pos_of:
                    continue
                sheet, rnum = pos_of[row_id]
                cell = wb[sheet].cell(row=rnum, column=col_map[(header, gc)])
                _write_check_cell(cell, xs, sheet, f"{get_column_letter(2 + gc)}{rnum}")
                for x in xs:
                    addr_of[id(x)] = (sheet, cell.coordinate)
                written += 1

    # ── 주석 시트 하단 대사 블록 (C4·C5) ─────────────────────────
    written += _write_note_blocks(wb, report, pos_of, note_sheets, note_ranges, block_start)

    _write_summary(wb, report, addr_of)

    if _T.get("modernize"):
        _T["modernize"](wb)
    wb.save(xlsx_path)
    bad = sum(1 for x in report.results if not x.ok)
    return {"cells": written, "bad": bad}


def _write_note_blocks(wb, report, pos_of, note_sheets, note_ranges, block_start) -> int:
    """각 주석 시트 하단: 이 주석을 참조하는 재무제표 대사(C4, 살아있는 수식) +
    이 주석이 연루된 주석 간 대사(C5, 참고)."""
    vid = iter(range(1, 100000))
    written = 0
    c5_by_note: dict[int, list] = {}
    for f in report.findings:
        if f.check != "C5":
            continue
        for n in {int(m) for m in re.findall(r"주석(\d+)=", f.message)}:
            c5_by_note.setdefault(n, []).append(f)

    for n, sheet in sorted(note_sheets.items()):
        ws = wb[sheet]
        r = block_start[sheet]
        c4_items = [x for x in report.results if x.check == "C4" and n in x.refs]
        c5_items = c5_by_note.get(n, [])
        if not c4_items and not c5_items:
            continue

        def put(row, col, value, font=FONT, **kw):
            ws.cell(row=row, column=1, value=f"!v{next(vid):05d}").font = FONT
            cell = ws.cell(row=row, column=col, value=value)
            _style(cell, font, **kw)
            return cell

        if c4_items:
            head = put(r, 2, f"검증: 재무제표 대사 (이 주석을 참조하는 재무제표 행)",
                       FONT_SUBHEAD)
            head.fill = FILL_SUBHEAD
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            r += 1
            for c, htext in enumerate(["재무제표 행", "구분", "기재 금액", "결과"], start=2):
                cell = put(r, c, htext, FONT_SUBHEAD,
                           align=Alignment(horizontal="center"))
                cell.fill = FILL_SUBHEAD
            r += 1
            cf_first = r
            for x in c4_items:
                loc = pos_of.get(x.row_id)
                put(r, 2, f"{x.context} · {x.row_label}")
                put(r, 3, x.col_header, align=Alignment(horizontal="center"))
                amount = put(r, 4, None, align=Alignment(horizontal="right"),
                             numfmt=NUMFMT_INT)
                result = put(r, 5, None, align=Alignment(horizontal="center"),
                             numfmt=NUMFMT_INT)
                if loc:
                    src = f"'{loc[0]}'!{get_column_letter(2 + x.grid_col)}{loc[1]}"
                    amount.value = f"={src}"
                    rng = note_ranges[n]
                    result.value = (
                        f'=IF(OR(COUNTIF({rng},{src})>0,COUNTIF({rng},-{src})>0),'
                        f'TRUE,"미확인")'
                    )
                else:
                    amount.value = _num(x.stated)
                    result.value = True if x.ok else "미확인"
                written += 1
                r += 1
            col_e = get_column_letter(5)
            ws.conditional_formatting.add(
                f"{col_e}{cf_first}:{col_e}{r - 1}",
                FormulaRule(
                    formula=[f'AND({col_e}{cf_first}<>"",{col_e}{cf_first}<>TRUE)'],
                    fill=_FILL_BAD, font=_FONT_BAD,
                ),
            )
            r += 1

        if c5_items:
            head = put(r, 2, "검증: 주석 간 대사 (같은 항목·기간의 금액 상이 — 참고)",
                       FONT_SUBHEAD)
            head.fill = FILL_SUBHEAD
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            r += 1
            for f in c5_items:
                put(r, 2, f.where)
                put(r, 3, f.message, align=Alignment(wrap_text=False))
                ws.merge_cells(start_row=r, start_column=3, end_row=r,
                               end_column=_NOTE_BLOCK_COLS)
                written += 1
                r += 1
    return written


def _write_summary(wb, report, addr_of):
    if SUMMARY_SHEET in wb.sheetnames:
        del wb[SUMMARY_SHEET]
    ws = wb.create_sheet(SUMMARY_SHEET, 0)
    ws.column_dimensions["A"].hidden = True
    for col, w in (("B", 36), ("C", 10), ("D", 8), ("E", 8), ("F", 10),
                   ("G", 30), ("H", 80)):
        ws.column_dimensions[col].width = w
    sid = iter(range(1, 10000))

    def put(r, c, value, font=FONT, **kw):
        ws.cell(row=r, column=1, value=f"!s{next(sid):04d}").font = FONT
        cell = ws.cell(row=r, column=c, value=value)
        cell.font = font
        for k, v in kw.items():
            setattr(cell, k, v)
        return cell

    put(2, 2, "DSD 정합성 검증 요약", FONT_TITLE)
    put(3, 2, f"실행: {datetime.now():%Y-%m-%d %H:%M} — 검증 셀은 수식이라 "
              f"값을 고치면 즉시 재검증됩니다. (행 추가/삭제 후에는 check 재실행)")

    r = 5
    for c, h in enumerate(["검사", "검사 수", "일치", "불일치", "참고(info)"], start=2):
        cell = ws.cell(row=r, column=c, value=h)
        cell.font = FONT_HEAD
        cell.fill = FILL_HEAD
        cell.border = _BORDER
        cell.alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=1, value=f"!s{next(sid):04d}")
    r += 1
    for key, name in CHECK_NAMES.items():
        results = [x for x in report.results if x.check == key]
        infos = [f for f in report.findings if f.check == key and f.level == "info"]
        bad = sum(1 for x in results if not x.ok)
        vals = [f"{key} {name}", report.counts.get(key, 0),
                len(results) - bad if results else "", bad if results else "", len(infos)]
        ws.cell(row=r, column=1, value=f"!s{next(sid):04d}")
        for c, v in enumerate(vals, start=2):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = _FONT_BAD if (c == 5 and isinstance(v, int) and v > 0) else FONT
            cell.border = _BORDER
            if c > 2:
                cell.alignment = Alignment(horizontal="center")
        r += 1

    bad_results = [x for x in report.results if not x.ok]
    if bad_results:
        r += 1
        put(r, 2, "불일치 목록 (클릭하면 해당 셀로 이동)", FONT_TITLE)
        r += 1
        for x in bad_results:
            diff = ("미확인" if x.computed is None
                    else fmt(_num(x.stated) - _num(x.computed)))
            label = f"[{x.check}] {x.context} · {x.row_label} × {x.col_header}: {diff}"
            cell = put(r, 2, label, _FONT_BAD)
            loc = addr_of.get(id(x))
            if loc:
                cell.hyperlink = f"#'{loc[0]}'!{loc[1]}"
                cell.font = Font(name=_T["font"], size=_T["size_body"], bold=True,
                                 color=_C["accent"], underline="single")
            if x.note:
                put(r, 7, x.note)
            r += 1

    others = [f for f in report.findings if f.check == "C6"]
    if others:
        r += 1
        put(r, 2, "표기 발견 사항 (참고)", FONT_TITLE)
        r += 1
        for f in others:
            put(r, 2, f"[{f.check}] ({f.level})")
            put(r, 7, f.where)
            put(r, 8, f.message,
                alignment=Alignment(wrap_text=True, vertical="center"))
            r += 1
