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
    FILL_HEAD, FILL_SUBHEAD, FONT, FONT_HEAD, FONT_LINK, FONT_SUBHEAD, FONT_TITLE,
    NUMFMT_INT, THIN, _C, _T, refined_sections,
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


def annotate(xlsx_path: str, report: CheckReport, doc: Document, row_positions: dict,
             tie_points: list | None = None, tie_suggestions: list | None = None,
             profile=None) -> dict:
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

    # ── 대사맵 시트 (C8: 확정 검증 + 제안) ───────────────────────
    tie_rows = 0
    if tie_points is not None or tie_suggestions is not None:
        tie_rows = _write_tie_sheet(wb, tie_points or [], tie_suggestions or [],
                                    pos_of, profile)

    _write_summary(wb, report, addr_of)

    if _T.get("modernize"):
        _T["modernize"](wb)
    wb.save(xlsx_path)
    bad = sum(1 for x in report.results if not x.ok)
    return {"cells": written, "bad": bad, "ties": tie_rows}


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


def _write_tie_sheet(wb, points: list, suggestions: list, pos_of, profile) -> int:
    """'대사맵' 시트: 확정 대사쌍의 살아있는 검증 + 제안 목록.

    열 배치는 ties.read_tie_sheet와 계약: A=!t행ID, B=상태(사용자 편집),
    C=ID, D=대사A, E=A금액(수식), F=관계, G=대사B, H=B금액(수식), I=기간,
    J=판정(수식 — 확정만 recheck 게이트 대상), K=출처, L=비고(사용자 편집),
    M=_data(JSON, 숨김 — ties save의 복원 원본)."""
    import json as _json

    from . import ties as ties_mod

    if ties_mod.SHEET_NAME in wb.sheetnames:
        del wb[ties_mod.SHEET_NAME]
    ws = wb.create_sheet(ties_mod.SHEET_NAME, 0)
    ws.column_dimensions["A"].hidden = True
    ws.column_dimensions["M"].hidden = True
    for col, w in (("B", 7), ("C", 11), ("D", 34), ("E", 14), ("F", 7),
                   ("G", 34), ("H", 14), ("I", 7), ("J", 12), ("K", 16), ("L", 24)):
        ws.column_dimensions[col].width = w
    tid = iter(range(1, 100000))

    def put(r, c, value, font=FONT, **kw):
        ws.cell(row=r, column=1, value=f"!t{next(tid):05d}").font = FONT
        return _style(ws.cell(row=r, column=c, value=value), font, **kw)

    put(2, 2, "대사맵 — 확정 대사쌍 검증 · 제안", FONT_TITLE, border=False)
    put(3, 2, "상태(확정/제안/제외)·비고 열만 수정하세요 → `dsdxlsx ties save 파일.xlsx` 로 "
              "저장. 확정은 매 check마다 엄격 검증(불일치 빨강), 제외는 다시 제안되지 않음.",
        border=False)

    r = 5
    headers = ["상태", "ID", "대사 A", "A 금액", "관계", "대사 B", "B 금액",
               "기간", "판정", "출처·신뢰도", "비고", "_data"]
    for c, h in enumerate(headers, start=2):
        cell = put(r, c, h, FONT_HEAD, align=Alignment(horizontal="center"))
        cell.fill = FILL_HEAD
    first_data = r + 1
    r += 1

    rel_ko = {"equal": "=", "negate": "= (−)", "abs": "|=|"}

    def amount_cell(row, col, hit, gc):
        cell = put(row, col, None, align=Alignment(horizontal="right"), numfmt=NUMFMT_INT)
        loc = pos_of.get(hit.row_id) if hit else None
        if loc is not None and gc is not None:
            cell.value = f"='{loc[0]}'!{get_column_letter(2 + gc)}{loc[1]}"
        elif hit is not None and gc is not None:
            cell.value = _num(hit.value(gc))
        return cell

    def loc_cell(row, col, sel, hit):
        cell = put(row, col, sel.human())
        loc = pos_of.get(hit.row_id) if hit else None
        if loc is not None:
            cell.hyperlink = f"#'{loc[0]}'!A{loc[1]}"
            cell.font = FONT_LINK
        return cell

    def term(addr, unit, absolute):
        t = f"ABS({addr})" if absolute else addr
        return f"({t}*{unit})" if unit != 1 else t

    def write_row(pt, gating: bool):
        nonlocal r
        tie = pt.tie
        put(r, 2, ties_mod.STATUS_KO.get(tie.status, tie.status),
            align=Alignment(horizontal="center"))
        put(r, 3, tie.id, align=Alignment(horizontal="center"))
        loc_cell(r, 4, tie.a, pt.a)
        amount_cell(r, 5, pt.a, pt.a_col)
        put(r, 6, rel_ko.get(tie.relation, tie.relation),
            align=Alignment(horizontal="center"))
        loc_cell(r, 7, tie.b, pt.b)
        amount_cell(r, 8, pt.b, pt.b_col)
        put(r, 9, pt.bucket, align=Alignment(horizontal="center"))
        verdict = put(r, 10, None, align=Alignment(horizontal="center"),
                      numfmt=NUMFMT_INT)
        if pt.a is None or pt.b is None or pt.a_col is None or pt.b_col is None:
            # 확정인데 해석 불가 → recheck 게이트가 잡도록 마커(,TRUE,) 있는 수식으로
            verdict.value = ('=IF(FALSE,TRUE,"해석불가")' if gating else "해석불가")
            put(r, 11, f"{tie.origin} — {pt.reason}"[:60])
        else:
            ua, ub = pt.a.tv.unit, pt.b.tv.unit
            tol = max(ua, ub) if ua != ub else (profile.footing_tol if profile else 1)
            absolute = tie.relation == "abs"
            ta = term(f"E{r}", ua, absolute)
            tb = term(f"H{r}", ub, absolute)
            if tie.relation == "negate":
                tb = f"(-{tb})"
            good, bad_ = ("TRUE", f"{ta}-{tb}") if gating else ('"일치"', '"상이"')
            verdict.value = f"=IF(ABS({ta}-{tb})<={tol},{good},{bad_})"
            conf = f" ({tie.confidence})" if tie.confidence else ""
            put(r, 11, f"{tie.origin}{conf}")
        put(r, 12, tie.note)
        data = put(r, 13, _json.dumps(tie.to_dict(), ensure_ascii=False), border=False)
        data.font = Font(name=_T["font"], size=6, color="BBBBBB")
        r += 1

    for pt in points:
        write_row(pt, gating=True)
    for pt in suggestions:
        write_row(pt, gating=False)

    if r > first_data:
        col_j = get_column_letter(10)
        ws.conditional_formatting.add(
            f"{col_j}{first_data}:{col_j}{r - 1}",
            FormulaRule(
                formula=[f'AND({col_j}{first_data}<>"",{col_j}{first_data}<>TRUE,'
                         f'{col_j}{first_data}<>"일치")'],
                fill=_FILL_BAD, font=_FONT_BAD,
            ),
        )
    ws.freeze_panes = f"B{first_data}"
    return r - first_data


AI_SHEET = "AI검토"


def annotate_ai(xlsx_path: str, findings: list[dict], doc: Document,
                row_positions: dict) -> dict:
    """AI 검토 소견(findings.json)을 'AI검토' 시트로 기록.

    check가 다시 실행돼도 이 시트는 지우지 않는다(검증요약과 분리) — annotate-ai
    재실행 시에만 교체. 행 ID는 `!a…`라 import가 무시한다. 모든 소견은
    advisory — 반영 여부는 사람이 본문을 직접 고쳐서 결정한다."""
    wb = load_workbook(xlsx_path)
    pos_of = _cell_positions(doc, row_positions)
    if AI_SHEET in wb.sheetnames:
        del wb[AI_SHEET]
    ws = wb.create_sheet(AI_SHEET, 0)
    ws.column_dimensions["A"].hidden = True
    for col, w in (("B", 8), ("C", 8), ("D", 34), ("E", 64), ("F", 44)):
        ws.column_dimensions[col].width = w
    aid = iter(range(1, 100000))

    def put(r, c, value, font=FONT, **kw):
        ws.cell(row=r, column=1, value=f"!a{next(aid):05d}").font = FONT
        return _style(ws.cell(row=r, column=c, value=value), font, **kw)

    put(2, 2, "AI 검토 소견", FONT_TITLE, border=False)
    put(3, 2, f"실행: {datetime.now():%Y-%m-%d %H:%M} — 전 항목 참고용(advisory). "
              "반영은 본문 셀을 직접 수정한 뒤 check 재실행으로 검산하세요.",
        border=False)

    r = 5
    for c, h in enumerate(["심각도", "구분", "위치", "내용", "제안"], start=2):
        cell = put(r, c, h, FONT_HEAD, align=Alignment(horizontal="center"))
        cell.fill = FILL_HEAD
    r += 1
    order = {"warn": 0, "info": 1}
    for f in sorted(findings, key=lambda x: order.get(x.get("severity"), 2)):
        sev = f.get("severity", "info")
        sev_cell = put(r, 2, sev, align=Alignment(horizontal="center"))
        if sev == "warn":
            sev_cell.fill = _FILL_BAD
            sev_cell.font = _FONT_BAD
        put(r, 3, f.get("kind", ""), align=Alignment(horizontal="center"))
        loc_text = f.get("section", "") or ""
        rid = f.get("row_id", "")
        loc = pos_of.get(rid) if rid else None
        if rid:
            loc_text = f"{loc_text} · {rid}" if loc_text else rid
        cell = put(r, 4, loc_text)
        if loc is not None:
            cell.hyperlink = f"#'{loc[0]}'!A{loc[1]}"
            cell.font = FONT_LINK
        put(r, 5, f.get("message", ""),
            align=Alignment(wrap_text=True, vertical="center"))
        put(r, 6, f.get("suggestion", ""),
            align=Alignment(wrap_text=True, vertical="center"))
        r += 1
    wb.save(xlsx_path)
    return {"findings": len(findings)}


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

    if report.coverage:
        covered = sum(1 for c in report.coverage if c["checks"])
        skipped = sum(1 for c in report.coverage if c["skipped"] and not c["checks"])
        uncovered = [c for c in report.coverage
                     if not c["skipped"] and not c["checks"]]
        r += 1
        put(r, 2, f"검사 커버리지 — 표 {len(report.coverage)}개: 검산 적용 {covered} · "
                  f"제외 구획 {skipped} · 미적용 {len(uncovered)}", FONT_TITLE)
        r += 1
        if uncovered:
            put(r, 2, "재무 검산이 하나도 적용되지 않은 표 — 구조 미인식 가능성, 육안 확인 권장:")
            r += 1
            for c in uncovered:
                put(r, 2, f"{c['context']} · {c['table']}")
                r += 1
