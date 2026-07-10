"""XLSX → DSD 재변환.

_map 시트의 base64에서 원본 DSD를 복원하고 동일 파서로 모델을 재구성한 뒤(결정론적
ID), 시트를 스캔해 편집 목록(셀 수정/행 추가/행 삭제)을 만들어 바이트 스플라이스한다.

행 감지 규칙:
- A열 ID 있음 → 기존 행. 원본에 있는데 시트에 없음 → 삭제.
- A열 ID 없는데 내용 있음 + 직전 항목이 같은 표의 행 → 신규 행(직전 행 뒤에 삽입).
- 표 내 기존 ID 순서가 원본과 다르면(재배열) 에러 — v1 미지원.
"""
from __future__ import annotations

import base64
import io
import json
import re
from dataclasses import dataclass, field
from decimal import Decimal

from openpyxl import load_workbook

from . import grid as grid_mod
from . import numbers, splice
from .dsd_io import DsdPackage, read_dsd
from .export_xlsx import MAP_SHEET
from .model import Document, Table
from .parse import parse_contents

ROW_ID_RE = re.compile(r"^t\d{3}\.r\d{3}$")


class ImportError_(Exception):
    """가드 위반 등 — 메시지에 위치(시트/행)를 담는다."""


@dataclass
class SheetEntry:
    sheet: str
    xlsx_row: int
    row_id: str | None  # None = 신규 행 후보
    values: dict[int, object]  # excel 열 번호 → 값


@dataclass
class Report:
    cell_edits: list[str] = field(default_factory=list)
    row_inserts: list[str] = field(default_factory=list)
    row_deletes: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    autosum_tables: set[str] = field(default_factory=set)
    # 표별 시트 행 위치 (시트 순서 = 재변환 후 문서의 TR 순서와 1:1) — check 주석용
    row_positions: dict[str, list[tuple[str, int]]] = field(default_factory=dict)

    @property
    def changed(self) -> bool:
        return bool(self.cell_edits or self.row_inserts or self.row_deletes)


def _cell_value_to_text(value, like: str = "") -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float, Decimal)):
        if value == 0:
            return "-"  # 문서 관행: 0은 '-' 표기
        return numbers.to_dsd_text(value, like=like)
    return str(value)


def read_map(xlsx_path: str):
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    if MAP_SHEET not in wb.sheetnames:
        raise ImportError_("이 xlsx에는 _map 시트가 없음 — dsdxlsx export로 만든 파일이 아님")
    ms = wb[MAP_SHEET]
    rows = list(ms.iter_rows(values_only=True))
    meta = json.loads(rows[0][0])
    if meta.get("schema") != 1:
        raise ImportError_(f"지원하지 않는 _map 스키마: {meta.get('schema')}")
    b64 = "".join(r[1] for r in rows if len(r) > 1 and r[1])
    dsd_bytes = base64.b64decode(b64)
    wb.close()
    return meta, dsd_bytes


def scan_sheets(xlsx_path: str) -> list[SheetEntry]:
    # read_only 필수: 일반 모드는 병합 범위가 셀 값을 MergedCell(None)로 가리는데,
    # 행 삽입/삭제 후 병합 범위가 어긋난 파일(openpyxl 편집 등)에서는 실제 값까지
    # 가려져 "빈 값으로 수정됨" 오검출 → 데이터 소실로 이어진다. read_only는
    # 저장된 값을 그대로 스트리밍하므로 안전하다.
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    entries: list[SheetEntry] = []
    for ws in wb.worksheets:
        if ws.title == MAP_SHEET:
            continue
        for row_num, row in enumerate(ws.iter_rows(), start=1):
            row_id = None
            values: dict[int, object] = {}
            for col, cell in enumerate(row, start=1):
                if cell.value is None:
                    continue
                if col == 1:
                    if str(cell.value).strip():
                        row_id = str(cell.value).strip()
                else:
                    values[col] = cell.value
            if row_id is None and not values:
                continue  # 완전 빈 행(구분선)
            entries.append(SheetEntry(ws.title, row_num, row_id, values))
    wb.close()
    return entries


def build_edits(
    doc: Document,
    entries: list[SheetEntry],
    exported_tables: set[str] | None = None,
) -> tuple[list[splice.Op], Report]:
    ops: list[splice.Op] = []
    report = Report()
    errors: list[str] = []

    seen_table_items: dict[str, list] = {}  # table_id -> [(row_id|None, entry)]
    seen_row_ids: set[str] = set()
    prev_table_id: str | None = None
    prev_sheet: str | None = None

    for entry in entries:
        if entry.sheet != prev_sheet:
            prev_table_id = None  # 표가 시트 경계를 넘지 않음
            prev_sheet = entry.sheet
        where = f"[{entry.sheet}!{entry.xlsx_row}행]"
        rid = entry.row_id
        if rid is None:
            # 신규 행 후보 — 직전 항목이 표 행이어야 함
            if prev_table_id is not None:
                seen_table_items.setdefault(prev_table_id, []).append((None, entry))
            else:
                errors.append(f"{where} ID 없는 행 — 표 바로 아래가 아니어서 어디 삽입인지 알 수 없음")
            continue
        if ROW_ID_RE.match(rid):
            table_id = rid.split(".")[0]
            if table_id not in doc.tables or rid not in doc.rows:
                errors.append(f"{where} 알 수 없는 행 ID: {rid}")
                continue
            if rid in seen_row_ids:
                report.warnings.append(f"{where} 중복 ID {rid} — 복사-붙여넣기로 판단, 신규 행으로 처리")
                seen_table_items.setdefault(table_id, []).append((None, entry))
            else:
                seen_row_ids.add(rid)
                seen_table_items.setdefault(table_id, []).append((rid, entry))
            prev_table_id = table_id
            continue
        prev_table_id = None
        if rid.startswith("p"):
            node = doc.nodes.get(rid)
            if node is None:
                errors.append(f"{where} 알 수 없는 문단 ID: {rid}")
                continue
            value = entry.values.get(2)
            if not numbers.values_equal(value, node.text):
                if node.complex:
                    report.warnings.append(
                        f"{where} 문단 {rid}에 인라인 서식(SPAN 등)이 있어 평문으로 대체됨"
                    )
                ops.append(splice.replace_inner(node, _cell_value_to_text(value)))
                report.cell_edits.append(f"{rid} (문단)")
        elif rid.startswith("!"):
            node = doc.nodes.get(rid)
            if node is not None and not numbers.values_equal(entry.values.get(2), node.text):
                report.warnings.append(f"{where} 제목/라벨 {rid}은 수정 불가 — 변경 무시됨")
        else:
            errors.append(f"{where} 해석할 수 없는 ID: {rid}")

    for table_id, items in seen_table_items.items():
        _process_table(doc, doc.tables[table_id], items, ops, report, errors)

    # 시트에서 통째로 사라진 표의 행 → 삭제로 처리하지 않고 에러 (실수 방지)
    # 단, 애초에 내보내지 않은 표(표지·보고서 구획)는 원본 그대로 보존되므로 정상
    for rid, row in doc.rows.items():
        table_id = rid.split(".")[0]
        if table_id in seen_table_items or not doc.tables[table_id].rows:
            continue
        if exported_tables is not None and table_id not in exported_tables:
            continue
        errors.append(
            f"표 {table_id} 전체가 시트에서 사라짐 — 표 단위 삭제는 지원하지 않음"
        )
        break

    if errors:
        raise ImportError_("\n".join(errors))
    return ops, report


def _process_table(doc, table: Table, items, ops, report, errors):
    original_ids = [r.id for r in table.rows]
    index_of = {rid: i for i, rid in enumerate(original_ids)}

    # 1) 순서 검증 (재배열 금지)
    seen_idx = [index_of[rid] for rid, _ in items if rid is not None]
    if seen_idx != sorted(seen_idx):
        errors.append(f"표 {table.id}: 행 순서가 원본과 다름 — 행 재배열은 지원하지 않음")
        return
    report.row_positions[table.id] = [(e.sheet, e.xlsx_row) for _, e in items]

    # 2) 삭제 감지
    present = {rid for rid, _ in items if rid is not None}
    for i, rid in enumerate(original_ids):
        if rid in present:
            continue
        row = doc.rows[rid]
        label = _row_label(table, i)
        if table.fixtable:
            errors.append(f"표 {table.id}: 고정 표(AFIXTABLE=Y)라 행 삭제 불가 — {label}")
        elif row.in_thead:
            errors.append(f"표 {table.id}: 헤더(THEAD) 행은 삭제 불가 — {label}")
        elif not row.adelete:
            errors.append(f"표 {table.id}: ADELETE=N 행은 삭제 불가 — {label}")
        elif grid_mod.row_in_span(table, i):
            errors.append(f"표 {table.id}: 병합(ROWSPAN)에 걸린 행은 삭제 불가(v1) — {label}")
        elif table.irregular:
            errors.append(f"표 {table.id}: 그리드가 불규칙해 행 조작 불가 — {label}")
        else:
            ops.append(splice.delete_row(doc.data, row))
            report.row_deletes.append(f"{rid} {label}")
    deleted = {rid for rid in original_ids if rid not in present}

    # 3) 기존 행 셀 수정
    for rid, entry in items:
        if rid is None:
            continue
        row = doc.rows[rid]
        i = index_of[rid]
        g = grid_mod.expand(table)
        for cell in row.cells:
            col = grid_mod.cell_grid_col(table, i, cell)
            value = entry.values.get(2 + col)
            if numbers.values_equal(value, cell.text):
                continue
            where = f"[{entry.sheet}!{entry.xlsx_row}행] {cell.id}"
            if cell.readonly:
                errors.append(f"{where}: 자동 필드(TU)는 수정 불가 — 원래 값 '{cell.text}'")
                continue
            if cell.complex:
                report.warnings.append(f"{where}: 셀에 서식/중첩 문단이 있어 평문으로 대체됨")
            new_text = _cell_value_to_text(value, like=cell.text)
            # 원본이 접미 단위('2,978,480주')였고 새 값이 숫자면 단위 유지
            suffix = numbers.suffix_of(cell.text)
            if suffix and isinstance(value, (int, float, Decimal)) and new_text != "-":
                new_text += suffix
            ops.append(splice.replace_inner(cell, new_text))
            report.cell_edits.append(cell.id)
            if table.has_autosum:
                report.autosum_tables.add(table.id)

    # 4) 신규 행 삽입 — 연속 신규 행은 하나의 op로 묶어 순서 보존
    pending_new: list = []
    anchor_rid: str | None = None  # 직전 생존 기존 행

    def flush():
        nonlocal pending_new
        if not pending_new:
            return
        _emit_inserts(doc, table, anchor_rid, pending_new, deleted, ops, report, errors)
        pending_new = []

    for rid, entry in items:
        if rid is None:
            pending_new.append(entry)
        else:
            flush()
            if rid not in deleted:
                anchor_rid = rid
    flush()


def _row_label(table: Table, i: int) -> str:
    texts = [c.text for c in table.rows[i].cells if c.text.strip()]
    return f"'{texts[0][:20]}'" if texts else "(빈 행)"


def _emit_inserts(doc, table: Table, anchor_rid, new_entries, deleted, ops, report, errors):
    if table.fixtable:
        errors.append(f"표 {table.id}: 고정 표(AFIXTABLE=Y)라 행 추가 불가")
        return
    if table.irregular:
        errors.append(f"표 {table.id}: 그리드가 불규칙해 행 추가 불가")
        return

    original_ids = [r.id for r in table.rows]
    index_of = {rid: i for i, rid in enumerate(original_ids)}
    surviving = [rid for rid in original_ids if rid not in deleted]
    if not surviving:
        errors.append(f"표 {table.id}: 남는 행이 없어 삽입 기준을 잡을 수 없음")
        return

    # 삽입 위치: anchor 뒤. anchor가 없으면(표 맨 앞) 첫 생존 행 앞.
    if anchor_rid is not None:
        insert_idx = index_of[anchor_rid] + 1
    else:
        insert_idx = index_of[surviving[0]]

    # THEAD 가드: 삽입 지점이 THEAD 내부면 거부
    if insert_idx < len(table.rows) and table.rows[insert_idx].in_thead:
        errors.append(f"표 {table.id}: 헤더(THEAD) 안에는 행을 추가할 수 없음")
        return
    # ROWSPAN 가드: 삽입 경계를 가로지르는 병합 금지
    if insert_idx > 0 and grid_mod.span_crosses_after(table, insert_idx - 1):
        errors.append(
            f"표 {table.id}: 삽입 지점이 병합(ROWSPAN) 범위를 가로지름 — 병합 범위 밖에 삽입하세요"
        )
        return

    template = _pick_template(table, insert_idx)
    if template is None:
        errors.append(f"표 {table.id}: 복제 가능한 템플릿 행(ACOPY=Y, 병합 없음)이 없음")
        return

    tpl_idx = index_of[template.id]
    payload = b""
    for entry in new_entries:
        texts: dict[int, str] = {}
        for ci, cell in enumerate(template.cells):
            col = grid_mod.cell_grid_col(table, tpl_idx, cell)
            texts[ci] = _cell_value_to_text(entry.values.get(2 + col))
        payload += splice.ROW_SEP + splice.clone_row(doc.data, template, texts)
        report.row_inserts.append(
            f"{table.id} {'r%03d' % (insert_idx + 1)} 위치에 신규 행 "
            f"[{entry.sheet}!{entry.xlsx_row}행]"
        )
        if table.has_autosum:
            report.autosum_tables.add(table.id)

    if anchor_rid is not None:
        anchor = doc.rows[anchor_rid]
        ops.append(splice.Op(pos=anchor.end, start=anchor.end, end=anchor.end, payload=payload))
    else:
        first = doc.rows[surviving[0]]
        # 앞 삽입: payload 선두의 ROW_SEP를 뒤로 옮김
        payload = payload[len(splice.ROW_SEP):] + splice.ROW_SEP
        ops.append(splice.Op(pos=first.start, start=first.start, end=first.start, payload=payload))


def _pick_template(table: Table, insert_idx: int):
    """삽입 지점 주변에서 복제 안전한 행 선택: TBODY + ACOPY=Y + 병합 무관 + 전체 폭."""
    g = grid_mod.expand(table)

    def ok(i: int) -> bool:
        row = table.rows[i]
        if row.in_thead or not row.acopy:
            return False
        if any(gc is not None and gc.kind == "vspan" for gc in g[i]):
            return False
        if any(c.rowspan > 1 for c in row.cells):
            return False
        return True

    candidates = list(range(insert_idx - 1, -1, -1)) + list(range(insert_idx, len(table.rows)))
    for i in candidates:
        if ok(i):
            return table.rows[i]
    return None


def import_xlsx(xlsx_path: str) -> tuple[DsdPackage, bytes, Report]:
    """xlsx → (원본 패키지, 새 contents.xml bytes, 리포트)."""
    meta, dsd_bytes = read_map(xlsx_path)
    package = read_dsd(io.BytesIO(dsd_bytes))
    contents = package.contents
    doc = parse_contents(contents)
    entries = scan_sheets(xlsx_path)
    exported = meta.get("exported_tables")
    ops, report = build_edits(
        doc, entries, set(exported) if exported is not None else None
    )
    new_contents = splice.apply(contents, ops)
    if report.changed:
        parse_contents(new_contents)  # well-formedness 재검증
    return package, new_contents, report
