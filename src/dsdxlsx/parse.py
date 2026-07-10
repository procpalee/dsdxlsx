"""expat 기반 오프셋 인식 파서.

핵심: bytes를 그대로 feed하면 CurrentByteIndex가 정확한 UTF-8 바이트 오프셋을 준다.
- StartElementHandler 시점 = 여는 태그 '<' 오프셋
- EndElementHandler 시점 = 닫는 태그 '</' 오프셋
무수정 노드 재직렬화에는 절대 디코드 텍스트를 쓰지 않는다 — 원본 바이트가 마스터.
"""
from __future__ import annotations

import xml.parsers.expat

from .model import Block, Document, Node, Row, Table

# TE = K-GAAP 표준 재무제표 셀 (ACODE 표준 계정코드·ALEVEL 계층·ADELIM 열 위치 보유)
CELL_TAGS = {"TD", "TH", "TU", "TE"}
# 최상위(셀/표 밖) 텍스트 블록으로 내보내는 태그
PARA_TAGS = {"P"}
LABEL_TAGS = {"TITLE", "COVER-TITLE", "WARNING"}


class DsdParseError(Exception):
    pass


def find_tag_end(data: bytes, start: int) -> int:
    """여는 태그 '<' 오프셋부터 따옴표 밖 '>'를 찾아 그 다음 오프셋 반환."""
    i = start
    in_quote = False
    n = len(data)
    while i < n:
        c = data[i]
        if c == 0x22:  # '"'
            in_quote = not in_quote
        elif c == 0x3E and not in_quote:  # '>'
            return i + 1
        i += 1
    raise DsdParseError(f"태그 끝을 찾지 못함 (오프셋 {start})")


def display_text(fragments: list[str]) -> str:
    """디코드된 텍스트 조각 → 표시 텍스트: 원시 개행 제거, 리터럴 &cr; → \\n."""
    text = "".join(fragments)
    text = text.replace("\r\n", "").replace("\n", "").replace("\r", "")
    return text.replace("&cr;", "\n")


class _Builder:
    def __init__(self, data: bytes):
        self.data = data
        self.doc = Document(data=data)
        self.parser = xml.parsers.expat.ParserCreate("utf-8")
        self.parser.buffer_text = True
        self.parser.StartElementHandler = self._start
        self.parser.EndElementHandler = self._end
        self.parser.CharacterDataHandler = self._chardata

        self.table: Table | None = None
        self.row: Row | None = None
        self.cell: Node | None = None
        self.cell_depth = 0  # 셀 내부 자식 요소 중첩 깊이
        self.para: Node | None = None
        self.in_thead = False
        self.in_colgroup = False
        self.text_buf: list[str] = []
        self.n_table = 0
        self.n_para = 0
        self.n_label = 0
        self.group_stack: list[str] = []  # 열려 있는 TABLE-GROUP의 ACLASS
        self.insertion_depth = 0  # INSERTION/LIBRARY 중첩

    def parse(self) -> Document:
        try:
            self.parser.Parse(self.data, True)
        except xml.parsers.expat.ExpatError as e:
            raise DsdParseError(f"XML 파싱 실패: {e}") from e
        return self.doc

    # --- handlers -------------------------------------------------

    def _start(self, tag: str, attrs: dict):
        pos = self.parser.CurrentByteIndex

        if self.cell is not None:
            # 셀 내부 자식 요소(P/SPAN 등) — 혼합 콘텐츠로 마킹, 텍스트는 계속 수집
            self.cell.complex = True
            self.cell_depth += 1
            return
        if self.para is not None:
            self.para.complex = True
            self.cell_depth += 1
            return

        if tag in CELL_TAGS and self.row is not None:
            node = Node(
                id="",  # 행 종료 시 부여
                tag=tag,
                start=pos,
                end=-1,
                inner_start=find_tag_end(self.data, pos),
                inner_end=-1,
                text="",
                attrs=attrs,
                readonly=(tag == "TU"),
            )
            self.cell = node
            self.text_buf = []
            if "AUTOSUM" in attrs and self.table is not None:
                self.table.has_autosum = True
        elif tag == "TR" and self.table is not None:
            self.row = Row(id="", start=pos, end=-1, attrs=attrs, in_thead=self.in_thead)
        elif tag == "TABLE":
            if self.table is not None:
                raise DsdParseError(f"중첩 TABLE은 지원하지 않음 (오프셋 {pos})")
            self.n_table += 1
            self.table = Table(
                id=f"t{self.n_table:03d}", start=pos, end=-1, attrs=attrs,
                group_aclass=self.group_stack[-1] if self.group_stack else None,
                in_insertion=self.insertion_depth > 0,
            )
        elif tag == "TABLE-GROUP":
            self.group_stack.append(attrs.get("ACLASS", ""))
        elif tag in ("INSERTION", "LIBRARY"):
            self.insertion_depth += 1
        elif tag == "THEAD":
            self.in_thead = True
        elif tag == "COLGROUP":
            self.in_colgroup = True
        elif tag == "COL" and self.in_colgroup and self.table is not None:
            self.table.ncols += 1
        elif tag in PARA_TAGS and self.table is None:
            self.n_para += 1
            self.para = Node(
                id=f"p{self.n_para:03d}",
                tag=tag,
                start=pos,
                end=-1,
                inner_start=find_tag_end(self.data, pos),
                inner_end=-1,
                text="",
                attrs=attrs,
            )
            self.text_buf = []
        elif tag in LABEL_TAGS and self.table is None:
            self.n_label += 1
            self.para = Node(
                id=f"!{self.n_label:03d}",
                tag=tag,
                start=pos,
                end=-1,
                inner_start=find_tag_end(self.data, pos),
                inner_end=-1,
                text="",
                attrs=attrs,
                readonly=True,
            )
            self.text_buf = []

    def _end(self, tag: str):
        pos = self.parser.CurrentByteIndex

        if self.cell_depth > 0:
            self.cell_depth -= 1
            return

        if self.cell is not None and tag == self.cell.tag:
            cell = self.cell
            cell.inner_end = pos
            cell.end = self.data.index(b">", pos) + 1
            cell.text = display_text(self.text_buf)
            if b"\r\n" in self.data[cell.inner_start : cell.inner_end]:
                cell.complex = True
            self.row.cells.append(cell)
            self.cell = None
            self.text_buf = []
        elif tag == "TR" and self.row is not None:
            row = self.row
            row.end = self.data.index(b">", pos) + 1
            row.id = f"{self.table.id}.r{len(self.table.rows) + 1:03d}"
            for i, cell in enumerate(row.cells):
                cell.id = f"{row.id}.c{i + 1:02d}"
                self.doc.nodes[cell.id] = cell
            self.table.rows.append(row)
            self.doc.rows[row.id] = row
            self.row = None
        elif tag == "TABLE" and self.table is not None:
            table = self.table
            table.end = self.data.index(b">", pos) + 1
            if table.ncols == 0:
                table.ncols = max((len(r.cells) for r in table.rows), default=1)
            self.doc.tables[table.id] = table
            self.doc.blocks.append(Block(kind="table", item=table))
            self.table = None
        elif tag == "THEAD":
            self.in_thead = False
        elif tag == "COLGROUP":
            self.in_colgroup = False
        elif tag == "TABLE-GROUP":
            if self.group_stack:
                self.group_stack.pop()
        elif tag in ("INSERTION", "LIBRARY"):
            self.insertion_depth = max(0, self.insertion_depth - 1)
        elif self.para is not None and tag == self.para.tag:
            para = self.para
            para.inner_end = pos
            para.end = self.data.index(b">", pos) + 1
            para.text = display_text(self.text_buf)
            if b"\r\n" in self.data[para.inner_start : para.inner_end]:
                para.complex = True
            kind = "label" if para.id.startswith("!") else "p"
            self.doc.nodes[para.id] = para
            self.doc.blocks.append(Block(kind=kind, item=para))
            self.para = None
            self.text_buf = []

    def _chardata(self, text: str):
        if self.cell is not None or self.para is not None:
            self.text_buf.append(text)


def parse_contents(data: bytes) -> Document:
    return _Builder(data).parse()
