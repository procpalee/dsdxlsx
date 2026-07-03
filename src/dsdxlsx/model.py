"""DSD contents.xml 문서 모델.

모든 노드는 원본 bytes 내 바이트 범위(start/end, inner_start/inner_end)를 가진다.
무수정 노드는 원본 바이트를 그대로 보존하고, 수정 시에만 inner 범위를 치환한다.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Node:
    """텍스트를 담는 편집 단위: TD/TH/TU 셀, 문단 P, TITLE 등."""

    id: str
    tag: str
    start: int
    end: int
    inner_start: int
    inner_end: int
    text: str  # 표시 텍스트: 원시 개행 제거, 리터럴 &cr; → \n
    attrs: dict = field(default_factory=dict)
    complex: bool = False  # SPAN/P 자식 또는 원시 CRLF 포함 — 수정 시 평문 강등
    readonly: bool = False  # TU(자동 필드) 등

    @property
    def colspan(self) -> int:
        return int(self.attrs.get("COLSPAN", 1))

    @property
    def rowspan(self) -> int:
        return int(self.attrs.get("ROWSPAN", 1))

    @property
    def align(self) -> str:
        return self.attrs.get("ALIGN", "LEFT").upper()


@dataclass
class Row:
    id: str
    start: int
    end: int
    attrs: dict = field(default_factory=dict)
    in_thead: bool = False
    cells: list[Node] = field(default_factory=list)

    @property
    def acopy(self) -> bool:
        return self.attrs.get("ACOPY", "Y") == "Y"

    @property
    def adelete(self) -> bool:
        return self.attrs.get("ADELETE", "Y") == "Y"


@dataclass
class Table:
    id: str
    start: int
    end: int
    attrs: dict = field(default_factory=dict)
    ncols: int = 0
    rows: list[Row] = field(default_factory=list)
    irregular: bool = False  # 그리드 전개 이상(열 초과 등) — 행 조작 금지
    grid: list | None = None  # grid.expand() 결과 캐시
    has_autosum: bool = False

    @property
    def fixtable(self) -> bool:
        return self.attrs.get("AFIXTABLE", "N") == "Y"


@dataclass
class Block:
    """문서 순서 레이아웃 항목."""

    kind: str  # 'p' | 'table' | 'label'
    item: Node | Table


@dataclass
class Document:
    data: bytes
    blocks: list[Block] = field(default_factory=list)
    tables: dict[str, Table] = field(default_factory=dict)
    rows: dict[str, Row] = field(default_factory=dict)
    nodes: dict[str, Node] = field(default_factory=dict)  # 셀 + 문단 + 라벨

    def sections(self) -> list[tuple[str, list[Block]]]:
        """TITLE(ATOC=Y) 기준으로 시트 구획 분할. 첫 구획은 '표지'."""
        result: list[tuple[str, list[Block]]] = [("표지", [])]
        for block in self.blocks:
            if (
                block.kind == "label"
                and block.item.tag == "TITLE"
                and block.item.attrs.get("ATOC") == "Y"
            ):
                result.append((block.item.text, [block]))
            else:
                result[-1][1].append(block)
        return [(t, bs) for t, bs in result if bs]
