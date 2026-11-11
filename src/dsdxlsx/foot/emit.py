"""워크북 작성 — 내용(원문·재무제표·주석) + 검증 블록 + 스타일.

셀 값 규칙 (숫자화 오류 방지):
- 숫자 전체 일치(numbers.classify)만 수치화 — '1.예적금담보대출' 같은 라벨은 텍스트 유지
- '-' 플레이스홀더 → 숫자 0 (서식 3구간이 '-'로 표시, 검증 수식 호환)
- '주석' 참조 열은 항상 텍스트 ("4,5,6,32" → 45632 오변환 방지)
"""
from __future__ import annotations

from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from . import constants as C
from . import mapsheet
from . import styles as S
from .checks import CheckSet
from .layout import PlacedTable, SheetLayout
from .tableview import classify_ext
from .. import grid as grid_mod
from .. import numbers


def cell_content(text: str, is_note_col: bool):
    """셀 텍스트 → (기록 값, 숫자서식|None). 검증 바인딩(value_of)과 공유하는 단일 규칙."""
    s = text.strip()
    if is_note_col:
        return (text if text else None), None
    if s == "-":
        return 0, S.NUMFMT_INT
    value = classify_ext(s)
    if value is None:
        return (text if text else None), None
    if numbers.suffix_of(s) == "%":  # 분수 저장 + 엑셀 % 서식
        dp = numbers.decimal_places(s)
        numfmt = f"0.{'0' * dp}%" if dp else "0%"
    elif numbers.classify(s) is None:  # '원' 접미 확장으로 인정된 경우
        base = s[:-1].strip()
        numfmt = S.number_format(base, numbers.decimal_places(base), "원")
    else:
        numfmt = S.number_format(s, numbers.decimal_places(s), numbers.suffix_of(s))
    if isinstance(value, Decimal):
        return float(value), numfmt
    return value, numfmt


def written_value(tv, i: int, c: int):
    """검산·수식 평가에 쓰는 '기록될 값' — cell_content와 동일 경로."""
    node = tv.cell(i, c)
    if node is None:
        return None
    value, _ = cell_content(node.text, c in tv.note_ref_cols())
    return value


def _write_table(ws: Worksheet, placed: PlacedTable, note_cols: set[int], recorder=None):
    table = placed.table
    g = grid_mod.expand(table)
    bordered = table.attrs.get("BORDER", "1") != "0"  # BORDER=0(제목부·각주 표)는 무테
    nrows = len(table.rows)
    for i, row in enumerate(table.rows):
        r = placed.row_of(i)
        for c in range(table.ncols):
            xc = c + C.CONTENT_START_COL
            cell = ws.cell(row=r, column=xc)
            if bordered:
                # 삼선표(frame) — 상/하 medium, 내부 thin, 좌/우 외곽 없음
                cell.border = S.table_border(i == 0, i == nrows - 1,
                                             c == 0, c == table.ncols - 1)
            if row.in_thead and bordered:
                cell.fill = S.FILL_HEAD
                cell.font = S.FONT_HEAD
            gc = g[i][c]
            if gc is None or gc.kind != "cell":
                continue
            node = gc.node
            value, numfmt = cell_content(node.text, c in note_cols)
            cell.value = value
            if recorder is not None:
                recorder.add(ws.title, r, xc, node, node.text, note_col=c in note_cols)
            if numfmt:
                cell.number_format = numfmt
            # 자동 줄바꿈 없음 — 행 높이가 세로로 늘어나지 않게 (사용자 지정)
            if row.in_thead:
                cell.alignment = Alignment(horizontal="center", vertical="center")
    # 병합 (그리드 좌표 → 시트 좌표)
    for i, row in enumerate(table.rows):
        for c, gc in enumerate(g[i]):
            if gc is None or gc.kind != "cell":
                continue
            node = gc.node
            if node.colspan > 1 or node.rowspan > 1:
                ws.merge_cells(
                    start_row=placed.row_of(i), start_column=c + C.CONTENT_START_COL,
                    end_row=placed.row_of(i) + node.rowspan - 1,
                    end_column=c + node.colspan + C.CONTENT_START_COL - 1,
                )


def _note_cols_of(table) -> set[int]:
    """헤더에 '주석'이 있는 열 (TableView 없이도 판단 — 원문 렌더 공용)."""
    g = grid_mod.expand(table)
    heads = [i for i, r in enumerate(table.rows) if r.in_thead]
    out = set()
    for c in range(table.ncols):
        for i in reversed(heads):
            gc = g[i][c]
            if gc is not None and gc.node.text.strip() and not gc.node.text.startswith("(단위"):
                if C.NOTE_COL_HEADER in C.norm(gc.node.text):
                    out.add(c)
                break
    return out


def write_content(ws: Worksheet, lay: SheetLayout, recorder=None):
    used_cols = 1
    for op in lay.ops:
        if op[0] == "para":
            _, r, text = op[:3]
            node = op[3] if len(op) > 3 else None
            cell = ws.cell(row=r, column=C.CONTENT_START_COL)
            cell.value = text if text and text.strip() else None
            if text and "\n" in text:
                cell.alignment = S.para_alignment(text)
            if recorder is not None and node is not None:
                recorder.add(ws.title, r, C.CONTENT_START_COL, node, text or "",
                             kind="para", partial=(text or "") != node.text)
        else:
            placed: PlacedTable = op[1]
            _write_table(ws, placed, _note_cols_of(placed.table), recorder)
            used_cols = max(used_cols, placed.table.ncols)
    ws.column_dimensions["A"].width = S.GUTTER_WIDTH   # 빈 거터 열
    for c in range(C.CONTENT_START_COL, used_cols + C.CONTENT_START_COL):
        ws.column_dimensions[get_column_letter(c)].width = S.COL_WIDTH
    # 2행 헤더 바 — 시트 제목 (본문 폭만큼, 최소 7칸)
    bar_end = max(used_cols + C.CONTENT_START_COL - 1, C.CONTENT_START_COL + 6)
    for c in range(C.CONTENT_START_COL, bar_end + 1):
        ws.cell(row=C.HEADER_BAR_ROW, column=c).fill = S.FILL_DARK
    tcell = ws.cell(row=C.HEADER_BAR_ROW, column=C.CONTENT_START_COL)
    tcell.value = lay.title or lay.name
    tcell.font = S.FONT_HEAD_INV
    ws.freeze_panes = S.FREEZE
    # 총괄표 바로가기 (B4)
    if lay.name != "총괄표":
        home = ws.cell(row=C.HOME_LINK_ROW, column=C.HOME_LINK_COL)
        home.value = '=HYPERLINK("#총괄표!A1","총괄표")'
        home.font = S.FONT_LINK


COMMENT_AUTHOR = "검증"


def _fmt_amount(v) -> str:
    try:
        d = Decimal(str(v))
    except Exception:
        return str(v)
    if d == d.to_integral_value():
        return f"{int(d):,}"
    return f"{d:,}"


def footing_comment(fc) -> str:
    """풋팅 셀 메모 — 도구가 가정한 관계와 판정 근거를 사람이 읽는 형태로.

    검증 셀마다 판정만 적는 대신 여기서는
    관계식(예: 기말 = 기초 + 취득 − 상각)과 항 값을 적어 검토자가 수식을 풀어보지
    않고도 왜 그 셀이 검증됐는지 알 수 있게 한다."""
    from .checks import KIND_DESC

    desc = KIND_DESC.get(fc.kind, fc.kind)
    sheet = fc.target.sheet
    if fc.op == "+":
        parts = []
        for k, (cell, sign) in enumerate(fc.terms[:10]):
            v = _fmt_amount(cell.value)
            parts.append(("−" if sign < 0 else ("+" if k else "")) + v)
        rel = f"{_fmt_amount(fc.target.value)} = {' '.join(parts)}"
        if len(fc.terms) > 10:
            rel += f" … (항 {len(fc.terms)}개)"
    else:
        a, b = fc.terms[0][0], fc.terms[1][0]
        rel = (f"{_fmt_amount(fc.target.value)} = {_fmt_amount(a.value)} {fc.op} "
               f"{_fmt_amount(b.value)}" + (f" × {fc.factor}" if fc.factor != 1 else ""))
    res = fc.excel_result()
    verdict = "일치" if res is True else f"차이 {_fmt_amount(res)}"
    return (f"[합계검증] {desc}\n{fc.label}\n{fc.d_expr()}\n"
            f"{rel}\n판정: {verdict}")


def cross_comment(xr) -> str:
    """대사 셀 메모 — 어느 재무제표/주석 셀과 어떻게(단위 환산) 대조했는지."""
    src_refs = "+".join(c.ref_abs() for c in xr.src)
    src_raw = sum((Decimal(str(c.value)) for c in xr.src), Decimal(0))
    conv = ""
    if xr.scale > 1:
        conv = f" (원천 ÷{xr.scale:,} {xr.rounder})"
    elif xr.dst_scale > 1:
        conv = f" (대상 ÷{xr.dst_scale:,} {xr.rounder})"
    verdict = "일치" if xr.passed() else f"차이 {_fmt_amount(xr.diff())}"
    return (f"[대사검증] {xr.display_label()}\n"
            f"원천 {src_refs} = {_fmt_amount(src_raw)}\n"
            f"대상 {xr.dst.ref_abs()} = {_fmt_amount(xr.dst.value)}{conv}\n"
            f"판정: {verdict}")


def attach_comment(cell, text: str, width: int = 360, height: int = 110) -> None:
    from openpyxl.comments import Comment

    cm = Comment(text, COMMENT_AUTHOR)
    cm.width, cm.height = width, height
    cell.comment = cm


def write_checks(ws_of, checkset: CheckSet):
    """검증 라벨·수식 셀 기록 + 검증 열너비 맞춤 + 불일치 빨강 음영 조건부서식."""
    label_w: dict[tuple[str, int], int] = {}   # (시트, 열) → 최장 라벨 글자수
    num_cols: set[tuple[str, int]] = set()
    check_addrs: dict[str, list[str]] = {}     # 시트 → 검증 수식 셀 주소들

    def track(sheet: str, col: int, text: str):
        key = (sheet, col)
        label_w[key] = max(label_w.get(key, 0), len(text))

    for fc in checkset.footing:
        ws = ws_of(fc.formula_cell.sheet)
        lab = ws.cell(row=fc.label_cell.row, column=fc.label_cell.col)
        lab.value = fc.label
        lab.font = S.FONT_BOLD
        lab.fill = S.FILL_LABEL
        track(fc.label_cell.sheet, fc.label_cell.col, fc.label)
        cell = ws.cell(row=fc.formula_cell.row, column=fc.formula_cell.col)
        formula = fc.formula()
        assert len(formula) < C.FORMULA_MAX, f"수식 길이 초과: {fc.formula_cell}"
        cell.value = formula
        cell.number_format = S.NUMFMT_CHECK
        cell.border = S.BORDER_CHECK
        attach_comment(cell, footing_comment(fc))
        num_cols.add((fc.formula_cell.sheet, fc.formula_cell.col))
        check_addrs.setdefault(fc.formula_cell.sheet, []).append(cell.coordinate)
    for xr in checkset.cross:
        ws = ws_of(xr.formula_cell.sheet)
        link = ws.cell(row=xr.link_cell.row, column=xr.link_cell.col)
        link.value = xr.hyperlink_formula()
        link.font = S.FONT_BOLD
        link.fill = S.FILL_XREF   # 대사는 세컨더리 헤더색 — 풋팅(F2F2F2)과 구분
        link.alignment = Alignment(horizontal="left")
        track(xr.link_cell.sheet, xr.link_cell.col, xr.display_label())
        cell = ws.cell(row=xr.formula_cell.row, column=xr.formula_cell.col)
        cell.value = xr.compare_formula()
        cell.number_format = S.NUMFMT_CHECK
        cell.border = S.BORDER_CHECK
        attach_comment(cell, cross_comment(xr))
        num_cols.add((xr.formula_cell.sheet, xr.formula_cell.col))
        check_addrs.setdefault(xr.formula_cell.sheet, []).append(cell.coordinate)
    # 검증 라벨 열: 한글 폭(≈2유닛) 기준 맞춤, 수식 열: 고정 폭.
    # 본문 열(아래 배치 라벨)은 이미 COL_WIDTH — max()로 줄어들지 않게만 넓힌다.
    for (sheet, col), n in label_w.items():
        dim = ws_of(sheet).column_dimensions[get_column_letter(col)]
        dim.width = max(dim.width or 0, min(34.0, max(12.0, n * 2.1 + 2)))
    for sheet, col in num_cols:
        if (sheet, col) not in label_w:
            dim = ws_of(sheet).column_dimensions[get_column_letter(col)]
            dim.width = max(dim.width or 0, S.CHECK_NUM_WIDTH)
    # 불일치 빨강 음영: 검증 수식 셀 값이 숫자(=차이)면 채움 — TRUE(논리값)는 제외.
    # 시트당 규칙 1개, 불연속 다중 범위(sqref) — 상대참조가 각 셀에 개별 적용된다.
    from openpyxl.formatting.rule import FormulaRule
    for sheet, addrs in check_addrs.items():
        addrs = sorted(set(addrs), key=lambda a: (len(a), a))
        ws_of(sheet).conditional_formatting.add(
            " ".join(addrs),
            FormulaRule(formula=[f"ISNUMBER({addrs[0]})"], fill=S.FILL_BAD))


def write_workbook(layouts: list[SheetLayout], checkset: CheckSet, out_path: str,
                   summary_writer=None, embed=None) -> None:
    """embed=(dsd_bytes, source_name[, 더 적을 메타]) 이면 셀맵(_map/_cellmap)을 함께 심는다 — writeback 전제."""
    recorder = mapsheet.CellMap() if embed is not None else None
    wb = Workbook()
    # 최신 Office 테마(맑은 고딕 minor) 주입 — 기본 셀 폰트(scheme=minor)가
    # 테마를 따라 맑은 고딕으로 렌더된다 (셀별 폰트 객체 불필요)
    wb.loaded_theme = S.OFFICE_THEME
    summary = wb.active
    summary.title = "총괄표"
    sheets: dict[str, Worksheet] = {"총괄표": summary}
    for lay in layouts:
        ws = wb.create_sheet(lay.name)
        sheets[lay.name] = ws
        write_content(ws, lay, recorder)
    write_checks(lambda name: sheets[name], checkset)
    if summary_writer is not None:
        summary_writer(summary, layouts, checkset)
    if embed is not None:
        mapsheet.write(wb, embed[0], embed[1], recorder, extra=embed[2] if len(embed) > 2 else None)
    from ..theme import set_author
    set_author(wb)
    wb.save(out_path)
