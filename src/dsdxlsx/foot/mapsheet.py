"""셀맵 임베드 — 워크북 안에 원본 DSD와 '셀 ↔ DSD 노드' 대응표를 숨겨 넣는다.

writeback(엑셀에서 고친 값 → DSD 반영)의 전제. 두 개의 veryHidden 시트:

- ``_map``    : A1 = JSON 메타 {schema, tool, source, sha256, n_cells},
                B열 = 원본 .dsd 파일 전체의 base64 (셀당 32,000자 청크).
                dsdxlsx export_xlsx 의 ``_map`` 규약과 같은 자리·같은 뜻.
- ``_cellmap``: 헤더 1행 + 매핑 셀당 1행
                sheet | row | col | node_id | kind | orig_text | note_col | partial
                kind = cell(표 셀) | para(문단·라벨). partial=1 은 문단을 잘라 쓴 경우
                (주석 헤딩 제거·혼합 문단 분할) — 반영 시 부분 치환으로 처리.

원문 시트와 개별 시트에 같은 노드가 두 번 나타난다 — writeback 이 노드 단위로 묶어
충돌을 검사한다. 검증 수식·AI 슬롯 등 표 밖 셀은 매핑에 없다(반영 대상 아님).
"""
from __future__ import annotations

import base64
import hashlib
import json
from dataclasses import dataclass

from openpyxl import load_workbook

MAP_SHEET = "_map"
CELLMAP_SHEET = "_cellmap"
SCHEMA = 1
TOOL = "dsdxlsx-foot"
TOOLS = (TOOL, "dsdfoot")   # 옛 dsdfoot 으로 만든 워크북도 읽는다
CHUNK = 32000  # 셀당 32,767자 제한 대비
_HEADER = ("sheet", "row", "col", "node_id", "kind", "orig_text", "note_col", "partial")


@dataclass
class MapEntry:
    sheet: str
    row: int
    col: int
    node_id: str
    kind: str          # cell | para
    orig_text: str     # 셀에 기록될 때 쓴 원문 텍스트 (partial 이면 잘린 조각)
    note_col: bool
    partial: bool


class CellMap:
    """emit 이 셀을 쓸 때마다 add — 순서 보존."""

    def __init__(self) -> None:
        self.entries: list[MapEntry] = []

    def add(self, sheet: str, row: int, col: int, node, orig_text: str,
            *, kind: str = "cell", note_col: bool = False, partial: bool = False) -> None:
        if node is None:
            return
        self.entries.append(MapEntry(sheet, row, col, node.id, kind, orig_text,
                                     note_col, partial))

    def __len__(self) -> int:
        return len(self.entries)


class MapError(Exception):
    """워크북에 셀맵이 없거나 스키마가 다름."""


def write(wb, dsd_bytes: bytes, source_name: str, cellmap: CellMap) -> None:
    """워크북(openpyxl Workbook)에 _map/_cellmap 시트를 추가한다. save 전에 호출."""
    for name in (MAP_SHEET, CELLMAP_SHEET):
        if name in wb.sheetnames:
            del wb[name]
    ms = wb.create_sheet(MAP_SHEET)
    ms.sheet_state = "veryHidden"
    ms["A1"] = json.dumps({
        "schema": SCHEMA, "tool": TOOL, "source": source_name,
        "sha256": hashlib.sha256(dsd_bytes).hexdigest(),
        "n_cells": len(cellmap),
    }, ensure_ascii=False)
    b64 = base64.b64encode(dsd_bytes).decode("ascii")
    for i in range(0, len(b64), CHUNK):
        ms.cell(row=i // CHUNK + 1, column=2, value=b64[i:i + CHUNK])

    cs = wb.create_sheet(CELLMAP_SHEET)
    cs.sheet_state = "veryHidden"
    cs.append(list(_HEADER))
    for e in cellmap.entries:
        cs.append([e.sheet, e.row, e.col, e.node_id, e.kind, e.orig_text,
                   1 if e.note_col else 0, 1 if e.partial else 0])


def is_foot_workbook(xlsx_path: str) -> bool:
    """풋팅 워크북(셀맵 시트가 있는 엑셀)인가 - export 로 만든 편집용 엑셀과 가른다."""
    import zipfile

    try:
        with zipfile.ZipFile(xlsx_path) as zf:
            return b'name="' + CELLMAP_SHEET.encode() + b'"' in zf.read("xl/workbook.xml")
    except (OSError, KeyError, zipfile.BadZipFile):
        return False


def read(xlsx_path: str) -> tuple[dict, bytes, list[MapEntry]]:
    """(meta, 원본 dsd bytes, 셀맵) — 없거나 스키마 불일치면 MapError."""
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        if MAP_SHEET not in wb.sheetnames or CELLMAP_SHEET not in wb.sheetnames:
            raise MapError("셀맵(_map/_cellmap)이 없는 워크북 - dsdxlsx foot 으로 만든 워크북만 반영할 수 있습니다")
        ms = wb[MAP_SHEET]
        meta = None
        chunks: list[str] = []
        for row in ms.iter_rows(min_col=1, max_col=2, values_only=True):
            a, b = (row + (None, None))[:2]
            if meta is None and a:
                meta = json.loads(str(a))
            if b:
                chunks.append(str(b))
        if not meta or meta.get("tool") not in TOOLS:
            raise MapError("풋팅 워크북의 셀맵이 아닙니다 (export 로 만든 엑셀은 `dsdxlsx import` 사용)")
        if meta.get("schema") != SCHEMA:
            raise MapError(f"셀맵 스키마 {meta.get('schema')} 미지원 (지원 {SCHEMA})")
        dsd_bytes = base64.b64decode("".join(chunks))
        if hashlib.sha256(dsd_bytes).hexdigest() != meta.get("sha256"):
            raise MapError("임베드된 원본 DSD 무결성 검사 실패")

        cs = wb[CELLMAP_SHEET]
        entries: list[MapEntry] = []
        it = cs.iter_rows(min_col=1, max_col=len(_HEADER), values_only=True)
        next(it, None)  # header
        for r in it:
            if not r or r[0] is None:
                continue
            r = tuple(r) + (None,) * (len(_HEADER) - len(r))
            entries.append(MapEntry(
                sheet=str(r[0]), row=int(r[1]), col=int(r[2]), node_id=str(r[3]),
                kind=str(r[4] or "cell"), orig_text="" if r[5] is None else str(r[5]),
                note_col=bool(r[6]), partial=bool(r[7]),
            ))
        return meta, dsd_bytes, entries
    finally:
        wb.close()
