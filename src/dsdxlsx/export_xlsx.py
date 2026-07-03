"""DSD → XLSX 내보내기.

시트: 표지/검토보고서 + 재무제표별(재무상태표·손익계산서·…) + 주석번호별(주석01~).
'주석' 시트에 각 주석 시트로 가는 하이퍼링크 목차를 만든다.
스타일은 excel-theme 스킬의 default 테마(theme.py 경유)를 따른다.
각 행의 A열(숨김)에 안정 ID(p013 / t042.r007 / !003)를 기록 — import의 행 감지 기준.
"""
from __future__ import annotations

import base64
import hashlib
import json
import math
import os
import re
from dataclasses import dataclass, field

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import grid as grid_mod
from . import numbers
from .model import Block, Document, Node, Table
from .theme import load_theme

MAP_SHEET = "_map"
MAP_CHUNK = 32000  # 셀당 32,767자 제한 대비

# ── excel-theme default 테마 기반 스타일 ─────────────────────────────
_T = load_theme()
_C = _T["colors"]
FONT = Font(name=_T["font"], size=_T["size_body"])
FONT_BOLD = Font(name=_T["font"], size=_T["size_body"], bold=True)
FONT_HEAD = Font(name=_T["font"], size=_T["size_header"], bold=True, color=_C["header_text"])
FONT_SUBHEAD = Font(name=_T["font"], size=_T["size_header"], bold=True, color=_C["header2_text"])
FONT_TITLE = Font(name=_T["font"], size=_T["size_title"] + 2, bold=True, color=_C["title_color"])
FONT_LINK = Font(name=_T["font"], size=_T["size_body"], color=_C["hlink"], underline="single")
THIN = Side(style="thin", color=_C["border_in"])
MEDIUM = Side(style="medium", color=_C["border_out"])
FILL_HEAD = PatternFill("solid", fgColor=_C["header_fill"])
FILL_SUBHEAD = PatternFill("solid", fgColor=_C["header2_fill"])
FILL_READONLY = PatternFill("solid", fgColor=_C["linked"])  # TU 자동 필드 — 수정 금지
# default 테마 숫자 서식: 음수 빨강 괄호, 0은 '-' 표시 (DSD 플레이스홀더 관행과 일치)
NUMFMT_INT = "#,###,##0;[Red](#,###,##0);-"
ALIGN_MAP = {"LEFT": "left", "RIGHT": "right", "CENTER": "center"}


def sanitize_sheet_title(text: str, used: set[str]) -> str:
    name = re.sub(r"\s+", " ", text).strip()
    # "분 기 재 무 상 태 표" 같은 자간 벌림 제목은 공백 제거
    tokens = name.split(" ")
    if len(tokens) >= 3 and sum(len(t) == 1 for t in tokens) / len(tokens) > 0.6:
        name = "".join(tokens)
    name = re.sub(r'[\\/*?:\[\]]', "", name) or "Section"
    name = name[:31]
    base, n = name, 2
    while name in used:
        suffix = f" ({n})"
        name = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(name)
    return name


# ---- 시트 세분화: 재무제표별 / 주석번호별 ---------------------------------

@dataclass
class Sec:
    title: str
    blocks: list[Block] = field(default_factory=list)
    note_no: int | None = None
    note_title: str = ""
    notes_intro: bool = False


STATEMENT_KEYWORDS = [
    "이익잉여금처분계산서", "결손금처리계산서", "포괄손익계산서",
    "손익계산서", "재무상태표", "자본변동표", "현금흐름표",
]
# 주석 헤딩: "6. 현금및현금성자산" — 제목 첫 글자가 숫자면(하위 헤딩 "2.1 …") 제외
_NOTE_LINE_RE = re.compile(r"^(\d{1,2})\.\s*([^\d\s][^\n]{0,40})")


def _statement_kw(block: Block) -> str | None:
    """재무제표 제목 테이블 판별: 첫 행 첫 셀이 '분 기 재 무 상 태 표' 류."""
    if block.kind != "table":
        return None
    table = block.item
    if not table.rows or not table.rows[0].cells:
        return None
    first = re.sub(r"\s+", "", table.rows[0].cells[0].text)
    if not first or len(first) > 20:
        return None
    for kw in STATEMENT_KEYWORDS:
        if kw in first:
            return kw
    return None


def _split_statements(title: str, blocks: list[Block]) -> list[Sec]:
    groups = [Sec(title)]
    for block in blocks:
        kw = _statement_kw(block)
        if kw:
            groups.append(Sec(kw, [block]))
        else:
            groups[-1].blocks.append(block)
    return [g for g in groups if g.blocks]


def _note_heading(line: str, current_n: int) -> tuple[int, str] | None:
    """주석 번호는 단조증가만 인정 — 본문 내 번호 목록('1. …')·하위 헤딩 오검출 방지."""
    m = _NOTE_LINE_RE.match(line.strip())
    if m and int(m.group(1)) > current_n:
        return int(m.group(1)), m.group(2).strip()
    return None


def _split_notes(title: str, blocks: list[Block]) -> list[Sec]:
    groups = [Sec(title, notes_intro=True)]
    current_n = 0
    for block in blocks:
        if block.kind == "p":
            lines = [ln for ln in block.item.text.split("\n") if ln.strip()]
            head = _note_heading(lines[0], current_n) if lines else None
            if head:
                current_n, note_title = head
                groups.append(Sec(str(current_n), [block],
                                  note_no=current_n, note_title=note_title))
            else:
                groups[-1].blocks.append(block)
            # 문단 중간에서 새 주석이 시작되면(혼합 P) 이후 블록은 새 그룹으로
            last = None
            for ln in lines[1:] if head else lines:
                h = _note_heading(ln, current_n)
                if h:
                    current_n, t = h
                    last = (current_n, t)
            if last:
                groups.append(Sec(str(last[0]), [],
                                  note_no=last[0], note_title=last[1]))
        else:
            groups[-1].blocks.append(block)
    return [g for g in groups if g.blocks]


def refined_sections(doc: Document) -> list[Sec]:
    """TITLE 구획을 재무제표별/주석번호별로 세분화.

    재무제표가 TITLE 구획 자체인 양식(K-GAAP: '재 무 상 태 표')은 자간 벌림을
    정규화해 표준 명칭 구획으로 인식한다 (IFRS 양식은 '(첨부)재무제표' 내부 분할)."""
    out: list[Sec] = []
    for title, blocks in doc.sections():
        norm = re.sub(r"\s+", "", title)
        kw = next((k for k in STATEMENT_KEYWORDS if k in norm), None)
        if "주석" in norm:
            out.extend(_split_notes(title, blocks))
        elif kw and "보고서" not in norm:
            out.append(Sec(kw, blocks))
        elif "재무제표" in norm and "보고서" not in norm:
            out.extend(_split_statements(title, blocks))
        else:
            out.append(Sec(title, blocks))
    return out


# ---- 셀/블록 쓰기 ----------------------------------------------------------

def _number_format(text: str) -> str:
    dp = numbers.decimal_places(text)
    suffix = numbers.suffix_of(text)
    lit = f'"{suffix}"' if suffix else ""
    if dp:
        frac = "0" * dp
        return f"#,###,##0.{frac}{lit};[Red](#,###,##0.{frac}{lit});-"
    return f"#,###,##0{lit};[Red](#,###,##0{lit});-" if suffix else NUMFMT_INT


def _write_cell(ws, row: int, col: int, node: Node):
    cell = ws.cell(row=row, column=col)
    value = numbers.classify(node.text)
    if value is not None:
        cell.value = float(value) if not isinstance(value, int) else value
        cell.number_format = _number_format(node.text)
    elif node.text.strip() == "-":
        # 플레이스홀더 '-'는 숫자 0으로 (검증 수식 호환) — 서식 3구간이 '-'로 표시
        cell.value = 0
        cell.number_format = NUMFMT_INT
    else:
        cell.value = node.text if node.text else None
    horizontal = ALIGN_MAP.get(node.align, "left")
    if node.tag == "TH":
        cell.font = FONT_HEAD
        cell.fill = FILL_HEAD
        horizontal = "center"
    else:
        cell.font = FONT
    if node.readonly:
        cell.fill = FILL_READONLY
    cell.alignment = Alignment(
        horizontal=horizontal,
        vertical="center",
        wrap_text=("\n" in node.text),
    )
    return cell


def _frame_table(ws, first_row: int, last_row: int, ncols: int, head_rows: int):
    """default 테마 'frame' 룩: 표 상/하 medium, 헤더 아래 구분선."""
    for c in range(2, 2 + ncols):
        top = ws.cell(row=first_row, column=c)
        top.border = Border(
            left=top.border.left, right=top.border.right, top=MEDIUM, bottom=top.border.bottom
        )
        bot = ws.cell(row=last_row, column=c)
        bot.border = Border(
            left=bot.border.left, right=bot.border.right, top=bot.border.top, bottom=MEDIUM
        )
        if head_rows:
            h = ws.cell(row=first_row + head_rows - 1, column=c)
            h.border = Border(
                left=h.border.left, right=h.border.right, top=h.border.top, bottom=MEDIUM
            )


def _write_table(ws, r: int, table: Table) -> int:
    g = grid_mod.expand(table)
    first = r
    for i, row in enumerate(table.rows):
        ws.cell(row=r, column=1, value=row.id).font = FONT
        for c, gc in enumerate(g[i]):
            xc = 2 + c
            ws.cell(row=r, column=xc).border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            if gc is not None and gc.kind == "cell":
                _write_cell(ws, r, xc, gc.node)
        r += 1
    head_rows = sum(1 for row in table.rows if row.in_thead)
    _frame_table(ws, first, r - 1, table.ncols, head_rows)
    # 병합 (그리드 좌표 → 시트 좌표)
    for i, row in enumerate(table.rows):
        for c, gc in enumerate(g[i]):
            if gc is None or gc.kind != "cell":
                continue
            node = gc.node
            if node.colspan > 1 or node.rowspan > 1:
                ws.merge_cells(
                    start_row=first + i,
                    start_column=2 + c,
                    end_row=first + i + node.rowspan - 1,
                    end_column=2 + c + node.colspan - 1,
                )
    return r + 1  # 블록 뒤 빈 행


def _write_para(ws, r: int, node: Node) -> int:
    ws.cell(row=r, column=1, value=node.id).font = FONT
    cell = ws.cell(row=r, column=2, value=node.text if node.text else None)
    if node.id.startswith("!"):
        cell.font = FONT_TITLE
    else:
        cell.font = FONT
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        lines = node.text.count("\n") + 1 + math.ceil(len(node.text) / 90)
        ws.row_dimensions[r].height = min(14 * lines + 4, 380)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
    return r + 2


def _write_note_index(ws, r: int, notes: list[tuple[str, int, str]]) -> None:
    """'주석' 시트 목차: 번호·제목 + 각 주석 시트 하이퍼링크."""
    head = ws.cell(row=r, column=2, value="주석 목차")
    head.font = FONT_TITLE
    ws.cell(row=r, column=1, value="!xh").font = FONT
    r += 2
    for sheet_name, no, title in notes:
        ws.cell(row=r, column=1, value=f"!x{no:02d}").font = FONT
        num = ws.cell(row=r, column=2, value=no)
        num.font = FONT
        num.alignment = Alignment(horizontal="center", vertical="center")
        num.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        link = ws.cell(row=r, column=3, value=title)
        link.hyperlink = f"#'{sheet_name}'!A1"
        link.font = FONT_LINK
        link.alignment = Alignment(vertical="center")
        link.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        r += 1


def exported_sections(doc: Document) -> list[Sec]:
    """엑셀로 내보내는 구획: 재무제표별 시트 + 주석(목차/번호별)만.
    표지·검토(감사)보고서·(첨부) 표지부는 편집 대상이 아니므로 시트를 만들지 않는다
    — 원본 바이트가 마스터라서 재변환 시 그대로 보존된다."""
    return [
        sec for sec in refined_sections(doc)
        if sec.title in STATEMENT_KEYWORDS or sec.note_no is not None or sec.notes_intro
    ]


def export_xlsx(doc: Document, dsd_bytes: bytes, dsd_name: str, out_path: str) -> dict:
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    stats = {"sheets": 0, "tables": 0, "paras": 0}

    intro_ws = None  # 주석 목차를 쓸 시트와 위치
    intro_row = 0
    note_index: list[tuple[str, int, str]] = []
    sections = exported_sections(doc)

    for sec in sections:
        ws = wb.create_sheet(sanitize_sheet_title(sec.title, used))
        ws.column_dimensions["A"].hidden = True
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 42
        for c in range(3, 16):
            ws.column_dimensions[get_column_letter(c)].width = 17
        r = 1
        for block in sec.blocks:
            if block.kind == "table":
                r = _write_table(ws, r, block.item)
                stats["tables"] += 1
            else:
                r = _write_para(ws, r, block.item)
                if block.kind == "p":
                    stats["paras"] += 1
        stats["sheets"] += 1
        if sec.notes_intro:
            intro_ws, intro_row = ws, r + 1
        elif sec.note_no is not None:
            note_index.append((ws.title, sec.note_no, sec.note_title))

    if intro_ws is not None and note_index:
        _write_note_index(intro_ws, intro_row, note_index)

    # _map: 원본 DSD 임베드 → xlsx 단일 파일로 import 가능
    ms = wb.create_sheet(MAP_SHEET)
    ms.sheet_state = "veryHidden"
    exported_tables = [
        block.item.id for sec in sections for block in sec.blocks if block.kind == "table"
    ]
    ms["A1"] = json.dumps(
        {
            "schema": 1,
            "tool": "dsdxlsx",
            "source": os.path.basename(dsd_name),
            "sha256": hashlib.sha256(dsd_bytes).hexdigest(),
            # 내보낸 표 목록 — import가 "시트에 없는 표"를 삭제로 오해하지 않도록
            "exported_tables": exported_tables,
        },
        ensure_ascii=False,
    )
    b64 = base64.b64encode(dsd_bytes).decode("ascii")
    for i in range(0, len(b64), MAP_CHUNK):
        ms.cell(row=i // MAP_CHUNK + 1, column=2, value=b64[i : i + MAP_CHUNK])

    if _T.get("modernize"):
        _T["modernize"](wb)
    wb.save(out_path)
    return stats
