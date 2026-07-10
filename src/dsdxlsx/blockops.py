"""블록(표·문단) 단위 편집 op 계획 + 가드.

바이트 규약(실측·스파이크 확정):
- `<TABLE` 여는 쪽 앞에만 빈 줄(\\r\\n\\r\\n), 그 외 블록 사이는 \\r\\n.
- 블록 삭제 = 블록 범위 + **선행** 구분자 1개 제거 → 남는 이웃이 자연 관례로 수렴.
- 블록 삽입 = 앵커 블록 end 뒤에 [구분자 + 새 블록 바이트].
"""
from __future__ import annotations

from . import splice
from .model import Document, Node, Table

TABLE_SEP = b"\r\n\r\n"
PARA_SEP = b"\r\n"


def is_note_sheet(sheet: str) -> bool:
    """구조 편집(표·문단 신설/삭제)은 주석 시트(숫자 이름)에서만 허용."""
    return sheet.strip().isdigit()


def prev_sep_start(data: bytes, start: int) -> int:
    """블록 앞 실측 구분자 시작 위치 (없으면 start 그대로)."""
    for sep in (TABLE_SEP, PARA_SEP):
        if data[start - len(sep):start] == sep:
            return start - len(sep)
    return start


def table_delete_guard(table: Table) -> str | None:
    if table.group_aclass is not None:
        return f"TABLE-GROUP({table.group_aclass}) 래핑 표는 삭제 불가"
    if table.in_insertion:
        return "INSERTION/LIBRARY 삽입물 내부 표는 삭제 불가"
    if table.fixtable:
        return "고정 표(AFIXTABLE=Y)는 삭제 불가"
    if any(c.tag == "TU" for r in table.rows for c in r.cells):
        return "자동 필드(TU)가 있는 표는 삭제 불가"
    return None


def plan_table_delete(doc: Document, table: Table) -> splice.Op:
    start = prev_sep_start(doc.data, table.start)
    return splice.Op(pos=start, start=start, end=table.end, payload=b"")


def plan_para_delete(doc: Document, node: Node) -> splice.Op:
    start = prev_sep_start(doc.data, node.start)
    return splice.Op(pos=start, start=start, end=node.end, payload=b"")


def insert_after_block(anchor, payload_blocks: list[bytes]) -> splice.Op:
    """앵커(Table|Node) 뒤에 블록들 삽입 — 각 블록 앞 구분자는 종류별 관례."""
    buf = b""
    for block in payload_blocks:
        sep = TABLE_SEP if block.lstrip()[:6] == b"<TABLE" else PARA_SEP
        buf += sep + block
    return splice.Op(pos=anchor.end, start=anchor.end, end=anchor.end, payload=buf)