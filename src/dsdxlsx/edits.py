"""apply - JSON 편집 목록을 DSD 에 반영한다(엑셀 없이).

방식: export 가 만드는 시트를 메모리에 그대로 재현하고(가상 시트), 편집을 그 시트에 적용한 뒤
import 와 같은 build_edits → splice 경로로 새 contents.xml 을 만든다. 그래서 쓰기 경로와
가드(고정 표, 병합 행, 자동 필드, 재배열 금지, 표·문단 신설은 주석 구획만 등)가 import 와 하나다.

편집 목록(JSON 배열, 또는 {"edits": [...]}). 행·칸은 ID 나 사람 좌표로 가리킨다.

  {"op": "set", "cell": "t003.r004.c02", "value": 1000}
  {"op": "set", "row": "t003.r004", "col": "당기말", "value": 1000}
  {"op": "set", "section": "4", "row": "보통예금", "col": "당기", "value": 1000}
  {"op": "set_text", "id": "p003", "text": "새 문단 글"}
  {"op": "insert_row", "after": "t003.r004", "values": ["외화예금", 50, 0]}
  {"op": "insert_row", "after": {"section": "4", "row": "보통예금"},
   "values": {"구분": "외화예금", "당기말": 50}}
  {"op": "delete_row", "row": "t003.r004"}
  {"op": "insert_para", "after": "t003", "text": "주석 4 아래 새 문단"}
  {"op": "delete_para", "id": "p005"}
  {"op": "insert_table", "after": "p005", "template": "t005", "header_rows": 1,
   "rows": [["구분", "당기말", "전기말"], ["제품", 10, 20]]}
  {"op": "delete_table", "id": "t005"}

열(col)은 grid 열 번호(정수, inspect 의 [번호]), 머리글 글자(공백 무시), "당기"/"전기" 중 하나다.
칸 ID 끝의 c02 는 행 안 칸 순번이라 열 번호와 다르다.
숫자는 JSON 숫자로 주면 원래 칸의 표기(괄호 음수, 소수 자리, '주' 같은 단위)를 따른다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from . import grid as grid_mod
from . import splice
from .check import norm
from .docview import SectionInfo, cell_value, sections, table_view
from .import_xlsx import ImportError_, Report, SheetEntry, build_edits
from .model import Document, Table
from .parse import parse_contents

ROW_ID_RE = re.compile(r"^t\d{3}\.r\d{3}$")
CELL_ID_RE = re.compile(r"^(t\d{3}\.r\d{3})\.c\d{2}$")
OPS = ("set", "set_text", "insert_row", "delete_row", "insert_para", "delete_para",
       "insert_table", "delete_table")


class EditError(Exception):
    pass


@dataclass
class VLine:
    row_id: str | None
    values: dict[int, object] = field(default_factory=dict)


@dataclass
class VSheet:
    sec: SectionInfo
    lines: list  # VLine | None(빈 행)


def _export_value(text: str):
    """export_xlsx._write_cell 과 같은 값 변환."""
    v = cell_value(text)
    if v is not None:
        return v
    return text if text else None


def _b(value):
    """B열 글자가 '#' 로 시작하면 마커로 읽히지 않게 '##' 로(import 가 한 겹 벗긴다)."""
    if isinstance(value, str) and value.startswith("#"):
        return "#" + value
    return value


class VirtualBook:
    """export 시트 배치를 메모리에 재현. 행 번호는 entries() 에서 매긴다."""

    def __init__(self, doc: Document):
        self.doc = doc
        self.sheets: list[VSheet] = []
        self.exported_tables: set[str] = set()
        self.section_of_table: dict[str, SectionInfo] = {}
        self.section_of_node: dict[str, SectionInfo] = {}
        for sec in sections(doc):
            lines: list = []
            for block in sec.blocks:
                if block.kind == "table":
                    table = block.item
                    self.exported_tables.add(table.id)
                    self.section_of_table[table.id] = sec
                    g = grid_mod.expand(table)
                    for i, row in enumerate(table.rows):
                        vals = {}
                        for c, gc in enumerate(g[i]):
                            if gc is not None and gc.kind == "cell":
                                v = _export_value(gc.node.text)
                                if v is not None:
                                    vals[2 + c] = _b(v) if c == 0 else v
                        lines.append(VLine(row.id, vals))
                    lines.append(None)
                else:
                    node = block.item
                    self.section_of_node[node.id] = sec
                    vals = {2: _b(node.text)} if node.text else {}
                    lines.append(VLine(node.id, vals))
                    lines.append(None)
            self.sheets.append(VSheet(sec, lines))
        self._after: dict[str, VLine] = {}  # 앵커 → 마지막으로 그 뒤에 넣은 줄(순서 유지)

    # ── 찾기 ──
    def locate(self, row_id: str) -> tuple[VSheet, int]:
        for sheet in self.sheets:
            for k, line in enumerate(sheet.lines):
                if line is not None and line.row_id == row_id:
                    return sheet, k
        raise EditError(f"'{row_id}' 없음(지운 행이거나 편집 대상 구획 밖)")

    def block_end(self, block_id: str) -> tuple[VSheet, int]:
        """블록(t### 또는 p###)의 마지막 줄 위치. 그 뒤에 이미 넣은 줄이 있으면 그 뒤."""
        if block_id in self._after:
            anchor = self._after[block_id]
            for sheet in self.sheets:
                for k, line in enumerate(sheet.lines):
                    if line is anchor:
                        return sheet, k
        if re.fullmatch(r"t\d{3}", block_id):
            for sheet in self.sheets:
                ks = [k for k, line in enumerate(sheet.lines)
                      if line is not None and line.row_id
                      and line.row_id.startswith(block_id + ".")]
                if not ks:
                    continue
                k = ks[-1]
                # 표 끝 뒤에 넣은 새 행(ID 없음)까지가 표
                while (k + 1 < len(sheet.lines) and sheet.lines[k + 1] is not None
                       and sheet.lines[k + 1].row_id is None):
                    k += 1
                return sheet, k
            raise EditError(f"표 '{block_id}' 없음(지운 표이거나 편집 대상 구획 밖)")
        return self.locate(block_id)

    def entries(self) -> list[SheetEntry]:
        out = []
        for sheet in self.sheets:
            r = 0
            for line in sheet.lines:
                r += 1
                if line is None or (line.row_id is None and not line.values):
                    continue
                out.append(SheetEntry(sheet.sec.sheet, r, line.row_id, dict(line.values)))
        return out


# ── 좌표 해석 ───────────────────────────────────────────────────────

def _sections_by_key(doc: Document) -> dict[str, SectionInfo]:
    return {s.key: s for s in sections(doc)}


def resolve_row(doc: Document, spec) -> str:
    """행 지정 → 행 ID. 문자열 ID 또는 {"section", "row", ["row_nth"], ["table"]}."""
    if isinstance(spec, str):
        if ROW_ID_RE.match(spec) and spec in doc.rows:
            return spec
        raise EditError(f"행 ID '{spec}' 없음(형식 t003.r004)")
    if not isinstance(spec, dict) or "row" not in spec or "section" not in spec:
        raise EditError(f"행 지정은 ID 문자열이나 section·row 를 가진 객체여야 한다: {spec!r}")
    secs = _sections_by_key(doc)
    key = str(spec["section"])
    sec = secs.get(key)
    if sec is None:
        raise EditError(f"구획 '{key}' 없음 - 있는 구획: {', '.join(secs)}")
    want = norm(str(spec["row"]))
    hits = []
    for block in sec.blocks:
        if block.kind != "table":
            continue
        if spec.get("table") and block.item.id != spec["table"]:
            continue
        tv = table_view(block.item, sec)
        for i in tv.body:
            if norm(tv.row_label(i)) == want:
                hits.append(block.item.rows[i].id)
    if not hits:
        raise EditError(f"구획 '{key}': 행 '{spec['row']}' 없음")
    nth = spec.get("row_nth")
    if nth is None:
        if len(hits) > 1:
            raise EditError(f"구획 '{key}': 행 '{spec['row']}' {len(hits)}개 "
                            f"({', '.join(hits)}) - row_nth 나 행 ID 로 고르라")
        return hits[0]
    if not 1 <= int(nth) <= len(hits):
        raise EditError(f"행 '{spec['row']}' {nth}번째 없음(총 {len(hits)}개)")
    return hits[int(nth) - 1]


def _table_of(doc: Document, row_id: str) -> Table:
    return doc.tables[row_id.split(".")[0]]


def resolve_col(doc: Document, row_id: str, spec, sec: SectionInfo) -> int:
    """열 지정 → grid 열 번호. 정수(inspect 의 [번호], 0부터), 머리글 글자, '당기'/'전기'.

    칸 ID 끝의 c02 는 행 안 칸 순번(1부터)이라 열 번호와 다르다 - 열 지정에는 받지 않는다."""
    table = _table_of(doc, row_id)
    if isinstance(spec, bool):
        raise EditError(f"열 지정이 잘못됐다: {spec!r}")
    if isinstance(spec, str) and re.fullmatch(r"c\d+", spec.strip()):
        raise EditError(f"열 '{spec}' 는 쓸 수 없다 - 열은 번호(0부터, inspect 의 [번호])나 "
                        f"머리글로 주고, 칸 하나를 ID 로 가리키려면 cell 을 쓰라")
    if isinstance(spec, str) and spec.strip().isdigit():
        spec = int(spec.strip())
    if isinstance(spec, int):
        if not 0 <= spec < table.ncols:
            raise EditError(f"표 {table.id} 는 열이 {table.ncols}개다(0~{table.ncols - 1}): {spec}")
        return spec
    if not isinstance(spec, str) or not spec.strip():
        raise EditError(f"열 지정이 잘못됐다: {spec!r}")
    tv = table_view(table, sec)
    want = norm(spec)
    exact = [c for c in range(table.ncols) if norm(tv.col_header(c)) == want]
    if not exact:
        exact = [c for c in range(table.ncols)
                 if want and tv.col_header(c) and want in norm(tv.col_header(c)).split("/")]
    if not exact and spec in ("당기", "전기"):
        exact = [c for c in range(table.ncols) if tv.period_bucket(c) == spec]
    if not exact:
        heads = ", ".join(f"[{c}] {tv.col_header(c)!r}" for c in range(table.ncols))
        raise EditError(f"표 {table.id}: 열 '{spec}' 없음 - 열: {heads}")
    if len(exact) > 1:
        heads = ", ".join(f"[{c}] {tv.col_header(c)!r}" for c in exact)
        raise EditError(f"표 {table.id}: 열 '{spec}' 여러 개({heads}) - 열 번호로 고르라")
    return exact[0]


def _cell_target(doc: Document, book: VirtualBook, op: dict) -> tuple[str, int]:
    if "cell" in op:
        m = CELL_ID_RE.match(str(op["cell"]))
        node = doc.nodes.get(str(op["cell"]))
        if not m or node is None:
            raise EditError(f"칸 ID '{op['cell']}' 없음(형식 t003.r004.c02)")
        row_id = m.group(1)
        table = _table_of(doc, row_id)
        i = next(k for k, r in enumerate(table.rows) if r.id == row_id)
        return row_id, grid_mod.cell_grid_col(table, i, node)
    if "row" not in op or "col" not in op:
        raise EditError("set 은 cell, 또는 row 와 col 이 필요하다")
    row_spec = op["row"]
    if isinstance(row_spec, str) and not ROW_ID_RE.match(row_spec) and "section" in op:
        row_spec = {k: op[k] for k in ("section", "row", "row_nth", "table") if k in op}
    row_id = resolve_row(doc, row_spec)
    sec = book.section_of_table.get(row_id.split(".")[0])
    if sec is None:
        raise EditError(f"행 {row_id} 는 편집 대상 구획(재무제표·주석) 밖이다")
    return row_id, resolve_col(doc, row_id, op["col"], sec)


def _value(v, op: dict):
    if isinstance(v, bool) or isinstance(v, (list, dict)):
        raise EditError(f"값은 숫자나 글자여야 한다: {op}")
    return v


def _row_values(doc: Document, book: VirtualBook, anchor_row: str, values) -> dict[int, object]:
    table = _table_of(doc, anchor_row)
    sec = book.section_of_table[table.id]
    out: dict[int, object] = {}
    if isinstance(values, list):
        if len(values) > table.ncols:
            raise EditError(f"표 {table.id} 는 열이 {table.ncols}개인데 값이 {len(values)}개다")
        items = list(enumerate(values))
    elif isinstance(values, dict):
        items = [(resolve_col(doc, anchor_row, k, sec), v) for k, v in values.items()]
    else:
        raise EditError("insert_row 의 values 는 배열이나 {열: 값} 객체여야 한다")
    for c, v in items:
        if v is None or v == "":
            continue
        out[2 + c] = _b(_value(v, {"values": values})) if c == 0 else _value(v, {"values": values})
    if not out:
        raise EditError("insert_row 의 values 가 비었다")
    return out


# ── 편집 적용 ───────────────────────────────────────────────────────

def _block_id(doc: Document, spec) -> str:
    if isinstance(spec, str) and (re.fullmatch(r"t\d{3}", spec) and spec in doc.tables
                                  or re.fullmatch(r"p\d{3}", spec) and spec in doc.nodes):
        return spec
    if isinstance(spec, str) and ROW_ID_RE.match(spec):
        return spec.split(".")[0]
    raise EditError(f"블록 ID '{spec}' 없음(표 t003 또는 문단 p003)")


def _insert_after(book: VirtualBook, anchor_key: str, sheet: VSheet, k: int, lines: list) -> None:
    sheet.lines[k + 1:k + 1] = lines
    book._after[anchor_key] = next(x for x in reversed(lines) if x is not None)


def apply_op(doc: Document, book: VirtualBook, op: dict) -> None:
    if not isinstance(op, dict) or op.get("op") not in OPS:
        raise EditError(f"알 수 없는 편집: {op!r} - 가능한 op: {', '.join(OPS)}")
    kind = op["op"]
    if kind == "set":
        if "value" not in op:
            raise EditError(f"set 에 value 가 없다: {op}")
        row_id, c = _cell_target(doc, book, op)
        sheet, k = book.locate(row_id)
        line = sheet.lines[k]
        v = op["value"]
        if v is None or v == "":
            line.values.pop(2 + c, None)
        else:
            v = _value(v, op)
            line.values[2 + c] = _b(v) if c == 0 else v
    elif kind == "set_text":
        node_id = str(op.get("id", ""))
        if not re.fullmatch(r"p\d{3}", node_id) or node_id not in doc.nodes:
            raise EditError(f"set_text 는 문단 ID(p003)가 필요하다: {op}")
        sheet, k = book.locate(node_id)
        text = op.get("text")
        if not isinstance(text, str):
            raise EditError(f"set_text 에 text 가 없다: {op}")
        sheet.lines[k].values = {2: _b(text)} if text else {}
    elif kind == "insert_row":
        anchor = resolve_row(doc, op.get("after"))
        vals = _row_values(doc, book, anchor, op.get("values"))
        if anchor in book._after:
            sheet, k = book.block_end(anchor)
        else:
            sheet, k = book.locate(anchor)
        line = VLine(None, vals)
        _insert_after(book, anchor, sheet, k, [line])
    elif kind == "delete_row":
        row_id = resolve_row(doc, op.get("row"))
        sheet, k = book.locate(row_id)
        del sheet.lines[k]
    elif kind == "insert_para":
        anchor = _block_id(doc, op.get("after"))
        text = op.get("text")
        if not isinstance(text, str) or not text.strip():
            raise EditError(f"insert_para 에 text 가 없다: {op}")
        sheet, k = book.block_end(anchor)
        _insert_after(book, anchor, sheet, k, [VLine(None, {2: "#문단 " + text})])
    elif kind == "delete_para":
        node_id = str(op.get("id", ""))
        if not re.fullmatch(r"p\d{3}", node_id) or node_id not in doc.nodes:
            raise EditError(f"delete_para 는 문단 ID(p003)가 필요하다: {op}")
        sheet, k = book.locate(node_id)
        sheet.lines[k].values = {2: "#문단삭제"}
    elif kind == "delete_table":
        tid = str(op.get("id", ""))
        table = doc.tables.get(tid)
        if table is None or not table.rows:
            raise EditError(f"표 '{tid}' 없음")
        sheet, k = book.locate(table.rows[0].id)
        sheet.lines[k].values[2] = "#표삭제"
    elif kind == "insert_table":
        anchor = _block_id(doc, op.get("after"))
        rows = op.get("rows")
        if not isinstance(rows, list) or not rows or not all(isinstance(r, list) for r in rows):
            raise EditError("insert_table 의 rows 는 행 배열의 배열이어야 한다")
        head = int(op.get("header_rows", 1))
        marker = f"#새표 헤더={head}"
        if op.get("template"):
            marker += f" 원본={op['template']}"
        lines = [VLine(None, {2: marker})]
        for r in rows:
            vals = {2 + c: (_b(v) if c == 0 else v) for c, v in enumerate(r)
                    if v is not None and v != ""}
            lines.append(VLine(None, vals or {3: ""}))
        lines.append(None)
        sheet, k = book.block_end(anchor)
        _insert_after(book, anchor, sheet, k, lines)


@dataclass
class ApplyResult:
    contents: bytes
    report: Report

    def changes(self) -> dict:
        r = self.report
        out = {}
        for name in ("cell_edits", "row_inserts", "row_deletes", "table_inserts",
                     "table_deletes", "para_inserts", "para_deletes", "renumbered"):
            v = getattr(r, name)
            if v:
                out[name] = list(v)
        if r.autosum_tables:
            out["autosum_tables"] = sorted(r.autosum_tables)
        return out


def normalize_edits(data) -> list[dict]:
    if isinstance(data, dict) and "edits" in data:
        data = data["edits"]
    if not isinstance(data, list):
        raise EditError('편집 목록은 JSON 배열이거나 {"edits": [...]} 여야 한다')
    return data


_WHERE_RE = re.compile(r"\[([^\[\]!]+)!(\d+)행\]")


def _humanizer(entries: list[SheetEntry], book: VirtualBook):
    """import 메시지의 '[시트!N행]'(가상 시트 위치)을 행 ID 나 '구획 X 새 줄'로 바꾼다."""
    key_of = {s.sec.sheet: s.sec.key for s in book.sheets}
    where = {(e.sheet, e.xlsx_row): e.row_id for e in entries}

    def fix(text: str) -> str:
        def sub(m):
            rid = where.get((m.group(1), int(m.group(2))))
            sec = key_of.get(m.group(1), m.group(1))
            return f"[{rid}]" if rid else f"[구획 {sec} 새 줄]"
        return _WHERE_RE.sub(sub, text)
    return fix


def apply_edits(contents: bytes, edits: list[dict]) -> ApplyResult:
    """contents.xml + 편집 목록 → 새 contents.xml. 실패하면 EditError(무엇이 왜)."""
    doc = parse_contents(contents)
    book = VirtualBook(doc)
    for n, op in enumerate(edits, 1):
        try:
            apply_op(doc, book, op)
        except EditError as e:
            raise EditError(f"편집 {n}번째 {op.get('op') if isinstance(op, dict) else op}: {e}") from None
    entries = book.entries()
    fix = _humanizer(entries, book)
    try:
        ops, report = build_edits(doc, entries, book.exported_tables)
    except ImportError_ as e:
        raise EditError(fix(str(e))) from None
    for name in ("cell_edits", "row_inserts", "row_deletes", "table_inserts", "table_deletes",
                 "para_inserts", "para_deletes", "renumbered", "warnings"):
        setattr(report, name, [fix(s) for s in getattr(report, name)])
    new = splice.apply(contents, ops)
    if report.changed:
        parse_contents(new)  # 형식 재검증
    return ApplyResult(new, report)
