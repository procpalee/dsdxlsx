"""XLSX → DSD 재변환.

_map 시트의 base64에서 원본 DSD를 복원하고 동일 파서로 모델을 재구성한 뒤(결정론적
ID), 시트를 스캔해 편집 목록(셀 수정/행 추가/행 삭제)을 만들어 바이트 스플라이스한다.

행 감지 규칙:
- A열 ID 있음 → 기존 행. 원본에 있는데 시트에 없음 → 삭제.
- A열 ID 없는데 내용 있음 + 직전 항목이 같은 표의 행 → 신규 행(직전 행 뒤에 삽입).
- 표 내 기존 ID 순서가 원본과 다르면(재배열) 에러 — v1 미지원.

구조 편집(마커, B열 — markers.py 문법):
- `#새표 [헤더=k] [원본=t###] [단위=원]` + 이어지는 A열 빈 행들 = 새 표 (주석 시트만)
- `#표삭제`(표 첫 행 B열) / `#문단 텍스트` / `#문단삭제` — clone-only 합성(compose.py),
  가드·바이트 규약은 blockops.py. 표 신설/삭제 후에는 표 ID가 재부여되므로
  **재export로 새 편집본을 받는 것이 표준 워크플로** (row_positions는 내부 재대응).
"""
from __future__ import annotations

import base64
import io
import json
import re
from dataclasses import dataclass, field
from decimal import Decimal

from openpyxl import load_workbook

from . import blockops, compose, markers, renumber
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
    table_inserts: list[str] = field(default_factory=list)
    table_deletes: list[str] = field(default_factory=list)
    para_inserts: list[str] = field(default_factory=list)
    para_deletes: list[str] = field(default_factory=list)
    note_inserts: list[str] = field(default_factory=list)
    note_deletes: list[str] = field(default_factory=list)
    renumbered: list[str] = field(default_factory=list)  # 자동 재번호 치환 내역
    warnings: list[str] = field(default_factory=list)
    autosum_tables: set[str] = field(default_factory=set)
    # 표별 시트 행 위치 (시트 순서 = 재변환 후 문서의 TR 순서와 1:1) — check 주석용
    row_positions: dict[str, list[tuple[str, int]]] = field(default_factory=dict)
    # 구조 편집 내부 기록 — row_positions 재대응용
    _inserted_tables: list[tuple[int, int, list]] = field(default_factory=list)
    _deleted_tables: set[str] = field(default_factory=set)

    @property
    def structural(self) -> bool:
        return bool(self.table_inserts or self.table_deletes
                    or self.para_inserts or self.para_deletes
                    or self.note_inserts or self.note_deletes)

    @property
    def changed(self) -> bool:
        return bool(self.cell_edits or self.row_inserts or self.row_deletes
                    or self.structural)


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


def _texts_of(entry: SheetEntry) -> dict[int, str]:
    """마커 블록 내용 행 → grid 열별 텍스트 (B열 '##' 이스케이프 해제 포함)."""
    out: dict[int, str] = {}
    for col, value in entry.values.items():
        if col < 2:
            continue
        if col == 2:
            value = markers.unescape(value)
        out[col - 2] = _cell_value_to_text(value)
    return out


def _compose_pending_table(doc: Document, p: dict,
                           errors: list[str]) -> tuple[bytes, list] | None:
    """수집된 #새표 블록 → (표 바이트, 시트 행 목록). 실패 시 None."""
    rows = p["entries"]
    n_head = p["n_head"]
    if len(rows) <= n_head:
        errors.append(p["where"] + f"#새표: 헤더 {n_head}행 외에 본문 행이 없음 "
                                   f"(내용 {len(rows)}행)")
        return None
    spec = compose.NewTableSpec(
        header_rows=[_texts_of(e) for e in rows[:n_head]],
        body_rows=[_texts_of(e) for e in rows[n_head:]],
        unit=p["marker"].args.get("단위"),
        where=p["where"],
    )
    spec.sheet_rows = (
        ([(p["sheet"], p["marker_row"])] if spec.unit else [])
        + [(e.sheet, e.xlsx_row) for e in rows]
    )
    try:
        tpl = compose.harvest(doc, p["template"], p["where"])
        block = compose.compose_table(doc.data, tpl, spec)
    except compose.ComposeError as e:
        errors.append(str(e))
        return None
    return block, spec.sheet_rows


def _extract_structures(doc: Document, entries: list[SheetEntry], report: Report,
                        errors: list[str]):
    """마커(구조 편집)를 스트림에서 분리 — 반환: (일반 entry, 삽입 계획, 표/문단
    삭제, 삭제 표 집합, 즉시 op 목록[주석 삽입/삭제/재번호])."""
    filtered: list[SheetEntry] = []
    inserts: list[tuple[object, bytes, list | None]] = []
    table_deletes: list[Table] = []
    para_deletes: list = []
    deleted_tables: set[str] = set()
    extra_ops: list[splice.Op] = []
    skip_ids: set[str] = set()          # 삭제된 주석 범위의 p/! 노드
    note_infos_cache: list | None = None
    new_note: dict | None = None        # {no,title,parts,sheet,marker_row}
    note_deleted = False

    prev_block = None   # 같은 시트의 직전 블록(Table|Node) — 삽입 앵커
    last_table = None   # 같은 시트의 직전 표 — 기본 템플릿
    prev_sheet: str | None = None
    seen_on_sheet = 0   # 이 시트에서 처리한 entry 수 (#새주석은 첫 행이어야)
    pending: dict | None = None  # 수집 중인 #새표 블록

    def notes():
        nonlocal note_infos_cache
        if note_infos_cache is None:
            note_infos_cache = renumber.note_infos(doc)
        return note_infos_cache

    def finalize_pending():
        nonlocal pending
        if pending is None:
            return
        p, pending = pending, None
        made = _compose_pending_table(doc, p, errors)
        if made is None:
            return
        block, sheet_rows = made
        n_body = len(p["entries"]) - p["n_head"]
        if p.get("into_note"):
            new_note["parts"].append((block, sheet_rows))
            report.table_inserts.append(
                f"[{p['sheet']}!{p['marker_row']}행] 새 주석 내 표 — 본문 {n_body}행, "
                f"템플릿 {p['template'].id}")
        else:
            inserts.append((p["anchor"], block, sheet_rows))
            report.table_inserts.append(
                f"[{p['sheet']}!{p['marker_row']}행] 새 표 — 본문 {n_body}행, "
                f"템플릿 {p['template'].id}")

    def finalize_new_note():
        nonlocal new_note
        if new_note is None:
            return
        note, new_note = new_note, None
        infos = notes()
        if not infos:
            errors.append(note["where"] + "#새주석: 문서에서 주석 구획을 찾지 못함")
            return
        existing = sorted(i.no for i in infos)
        no = note["no"]
        if not (1 <= no <= existing[-1] + 1):
            errors.append(note["where"] + f"#새주석 {no}: 번호가 불연속 "
                          f"(현재 1~{existing[-1]}, 허용 최대 {existing[-1] + 1})")
            return
        payload = compose.compose_note(no, note["title"],
                                       [b for b, _ in note["parts"]])
        if no == existing[-1] + 1:  # 말미 append
            anchor = infos[-1].blocks[-1].item
            pos = anchor.end
            extra_ops.append(splice.Op(pos=pos, start=pos, end=pos,
                                       payload=blockops.PARA_SEP + payload))
        else:  # 기존 번호 앞 삽입 (+ 필요 시 재번호 캐스케이드)
            target = next(i for i in infos if i.no >= no)
            pos = target.start
            first = target.blocks[0]
            sep = (blockops.TABLE_SEP if first.kind == "table" else blockops.PARA_SEP)
            extra_ops.append(splice.Op(pos=pos, start=pos, end=pos,
                                       payload=payload + sep))
            if no in existing:
                mapping = {k: k + 1 for k in existing if k >= no}
                try:
                    rn_ops, rn_rep = renumber.renumber_ops(doc, mapping)
                except renumber.RenumberError as e:
                    errors.append(note["where"] + f"#새주석 재번호 불가 — {e}")
                    return
                extra_ops.extend(rn_ops)
                report.renumbered.extend(rn_rep.headings + rn_rep.ref_cells)
                for m in rn_rep.mentions:
                    report.warnings.append(f"재번호 확인 필요(서술 멘션): {m}")
        for _, sheet_rows in note["parts"]:
            if sheet_rows is not None:
                report._inserted_tables.append(
                    (pos, len(report._inserted_tables), sheet_rows))
        report.note_inserts.append(
            f"주석 {no} {note['title'][:16]} — 블록 {len(note['parts']) + 1}개"
            + ("" if no == existing[-1] + 1 else f" (기존 {no}~ → +1 재번호)"))

    for entry in entries:
        if entry.sheet != prev_sheet:
            finalize_pending()
            if new_note is not None and prev_sheet == new_note["sheet"]:
                finalize_new_note()
            prev_block = None
            last_table = None
            prev_sheet = entry.sheet
            seen_on_sheet = 0
        where = f"[{entry.sheet}!{entry.xlsx_row}행] "
        bval = entry.values.get(2)
        in_note_sheet = new_note is not None and entry.sheet == new_note["sheet"]
        first_on_sheet = seen_on_sheet == 0
        seen_on_sheet += 1
        mk = None
        if markers.is_marker(bval):
            try:
                mk = markers.parse_marker(bval)
            except markers.MarkerError as e:
                errors.append(where + str(e))
                continue

        if pending is not None:
            gap = entry.xlsx_row != pending["last_row"] + 1  # 빈 행 = 블록 종료
            if mk is None and entry.row_id is None and not gap:
                pending["entries"].append(entry)
                pending["last_row"] = entry.xlsx_row
                continue
            finalize_pending()
            if mk is not None and mk.kind == "표끝":
                continue
            # 이 entry는 아래에서 정상 처리

        if mk is not None:
            if mk.kind == "새표":
                if not blockops.is_note_sheet(entry.sheet):
                    errors.append(where + "#새표는 주석 시트에서만 지원 — 재무제표 본문 표 신설 금지")
                    continue
                if entry.row_id is not None:
                    errors.append(where + "#새표 행의 A열은 비워야 함")
                    continue
                if not in_note_sheet and prev_block is None:
                    errors.append(where + "#새표 위에 기준 블록(기존 행)이 없어 삽입 위치를 알 수 없음")
                    continue
                if "원본" in mk.args:
                    template = doc.tables.get(mk.args["원본"])
                    if template is None:
                        errors.append(where + f"원본 표 없음: {mk.args['원본']}")
                        continue
                elif in_note_sheet:
                    errors.append(where + "#새주석 시트의 표는 원본=t### 템플릿 지정이 필요")
                    continue
                else:
                    template = last_table
                    if template is None:
                        errors.append(where + "같은 시트에 템플릿 표가 없음 — 원본=t### 지정 필요")
                        continue
                try:
                    n_head = int(mk.args.get("헤더", "1"))
                except ValueError:
                    errors.append(where + f"헤더= 값이 숫자가 아님: {mk.args['헤더']}")
                    continue
                if n_head < 1:
                    errors.append(where + "헤더=1 이상이어야 함")
                    continue
                pending = {"marker": mk, "anchor": prev_block, "template": template,
                           "n_head": n_head, "entries": [], "where": where,
                           "sheet": entry.sheet, "marker_row": entry.xlsx_row,
                           "last_row": entry.xlsx_row, "into_note": in_note_sheet}
                continue
            if mk.kind == "표끝":
                errors.append(where + "#표끝 앞에 #새표가 없음")
                continue
            if mk.kind == "문단":
                if in_note_sheet:
                    new_note["parts"].append(
                        (compose.compose_paragraph(mk.text), None))
                    report.para_inserts.append(
                        f"[{entry.sheet}!{entry.xlsx_row}행] 새 주석 내 문단")
                    continue
                if not blockops.is_note_sheet(entry.sheet):
                    errors.append(where + "#문단은 주석 시트에서만 지원")
                    continue
                if prev_block is None:
                    errors.append(where + "#문단 위에 기준 블록(기존 행)이 없음")
                    continue
                inserts.append((prev_block, compose.compose_paragraph(mk.text), None))
                report.para_inserts.append(f"[{entry.sheet}!{entry.xlsx_row}행] 새 문단")
                continue
            if mk.kind == "새주석":
                if new_note is not None:
                    errors.append(where + "#새주석은 한 번의 변환에 하나만 지원")
                    continue
                if note_deleted:
                    errors.append(where + "#새주석과 #주석삭제는 같은 변환에서 함께 쓸 수 없음 "
                                          "(삭제 → 변환 → 재export 후 추가)")
                    continue
                if not first_on_sheet or entry.row_id is not None:
                    errors.append(where + "#새주석은 새(빈) 시트의 첫 행에 쓰세요")
                    continue
                new_note = {"no": int(mk.args["번호"]), "title": mk.text,
                            "parts": [], "sheet": entry.sheet,
                            "marker_row": entry.xlsx_row, "where": where}
                continue
            if mk.kind == "주석삭제":
                if new_note is not None:
                    errors.append(where + "#새주석과 #주석삭제는 같은 변환에서 함께 쓸 수 없음")
                    continue
                rid = entry.row_id or ""
                node = doc.nodes.get(rid)
                info = next((i for i in notes()
                             if i.own_heading and i.heading.id == rid), None)
                if node is None or info is None:
                    errors.append(where + "#주석삭제는 주석 헤딩 문단 행(p###)의 B열에 쓰세요 "
                                          "(혼합 문단 헤딩은 편집기에서 분리 후)")
                    continue
                reason = renumber.deletable(notes(), info)
                if reason:
                    errors.append(where + reason)
                    continue
                extra_ops.append(renumber.plan_note_delete(doc, info))
                note_deleted = True
                for block in info.blocks:
                    if block.kind == "table":
                        deleted_tables.add(block.item.id)
                    else:
                        skip_ids.add(block.item.id)
                report.note_deletes.append(f"주석 {info.no} {info.title[:16]}")
                _warn_dangling_refs(doc, info.no, report)
                continue
            if mk.kind == "표삭제":
                rid = entry.row_id or ""
                if not ROW_ID_RE.match(rid) or rid not in doc.rows:
                    errors.append(where + "#표삭제는 지우려는 표의 첫 행 B열에 쓰세요 (A열 ID 유지)")
                    continue
                if not blockops.is_note_sheet(entry.sheet):
                    errors.append(where + "재무제표 본문 표는 삭제할 수 없음")
                    continue
                table = doc.tables[rid.split(".")[0]]
                guard = blockops.table_delete_guard(table)
                if guard:
                    errors.append(where + f"표 {table.id}: {guard}")
                    continue
                if table.id not in deleted_tables:
                    deleted_tables.add(table.id)
                    table_deletes.append(table)
                    report.table_deletes.append(f"{table.id} [{entry.sheet}]")
                continue
            if mk.kind == "문단삭제":
                rid = entry.row_id or ""
                node = doc.nodes.get(rid)
                if node is None or not rid.startswith("p"):
                    errors.append(where + "#문단삭제는 기존 문단 행(p###)의 B열에 쓰세요")
                    continue
                para_deletes.append(node)
                report.para_deletes.append(f"{rid} [{entry.sheet}]")
                continue
            errors.append(where + f"#{mk.kind}은(는) 아직 지원하지 않음 — 주석 추가/삭제는 다음 단계")
            continue

        if in_note_sheet:
            errors.append(where + "#새주석 시트에는 #문단/#새표 마커 블록만 올 수 있음")
            continue

        # 일반 행: '##' 이스케이프 해제 + 앵커/템플릿 추적
        if isinstance(bval, str) and bval.startswith("##"):
            entry.values[2] = bval[1:]
        rid = entry.row_id
        if rid and ROW_ID_RE.match(rid):
            tid = rid.split(".")[0]
            if tid in deleted_tables:
                continue  # 삭제 표의 잔여 행은 무시
            table = doc.tables.get(tid)
            if table is not None:
                prev_block = table
                last_table = table
        elif rid and (rid.startswith("p") or rid.startswith("!")):
            if rid in skip_ids:
                continue  # 삭제된 주석 범위의 문단/라벨 행
            node = doc.nodes.get(rid)
            if node is not None:
                prev_block = node
        filtered.append(entry)

    finalize_pending()
    finalize_new_note()
    return (filtered, inserts, table_deletes, para_deletes, deleted_tables,
            extra_ops)


def _warn_dangling_refs(doc: Document, no: int, report: Report):
    """주석 삭제 후 남는 참조(FS 참조 열·서술 멘션) 경고."""
    from .check import build_views, norm, use_profile

    hits: list[str] = []
    with use_profile(None):
        _, ctx = build_views(doc)
        for tv in ctx["statements"]:
            note_col = next((c for c in range(tv.ncols)
                             if "주석" in norm(tv.col_header(c))), None)
            if note_col is None:
                continue
            for i in tv.body:
                nums = [int(x) for x in re.findall(r"\d+", tv.text(i, note_col))]
                if no in nums:
                    hits.append(f"{tv.sec_title} '{tv.row_label(i).strip()[:16]}' 참조 열")
    for node in doc.nodes.values():
        for m in renumber.MENTION_RE.finditer(node.text):
            if no in [int(x) for x in re.findall(r"\d+", m.group())]:
                hits.append(f"{node.id} '{m.group()}'")
    if hits:
        report.warnings.append(
            f"삭제된 주석 {no}을(를) 아직 참조하는 곳 {len(hits)}건 — "
            + "; ".join(hits[:5])
            + (" …" if len(hits) > 5 else "")
            + " — 번호 정리는 `dsdxlsx renumber 파일.dsd --close-gap`")


def build_edits(
    doc: Document,
    entries: list[SheetEntry],
    exported_tables: set[str] | None = None,
) -> tuple[list[splice.Op], Report]:
    ops: list[splice.Op] = []
    report = Report()
    errors: list[str] = []

    # 구조 편집(마커) 분리 — 남은 entries로 기존 셀/행 파이프라인 수행
    (entries, struct_inserts, table_deletes, para_deletes, deleted_tables,
     extra_ops) = _extract_structures(doc, entries, report, errors)
    report._deleted_tables = deleted_tables
    ops.extend(extra_ops)

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
    # 단, 애초에 내보내지 않은 표(표지·보고서 구획)와 #표삭제로 명시 삭제된 표는 예외
    for rid, row in doc.rows.items():
        table_id = rid.split(".")[0]
        if table_id in seen_table_items or not doc.tables[table_id].rows:
            continue
        if table_id in deleted_tables:
            continue
        if exported_tables is not None and table_id not in exported_tables:
            continue
        errors.append(
            f"표 {table_id} 전체가 시트에서 사라짐 — 삭제하려면 표 첫 행 B열에 #표삭제"
        )
        break

    # 구조 편집 op — 같은 앵커의 연속 삽입은 한 op로 묶어 시트 순서 보존
    grouped: dict[int, list[bytes]] = {}
    anchor_of: dict[int, object] = {}
    key_order: list[int] = []
    for anchor, block, sheet_rows in struct_inserts:
        key = id(anchor)
        if key not in grouped:
            grouped[key] = []
            anchor_of[key] = anchor
            key_order.append(key)
        grouped[key].append(block)
        if sheet_rows is not None:  # 표 삽입 — row_positions 재대응용 기록
            report._inserted_tables.append(
                (anchor_of[key].end, len(report._inserted_tables), sheet_rows))
    for key in key_order:
        ops.append(blockops.insert_after_block(anchor_of[key], grouped[key]))
    for table in table_deletes:
        ops.append(blockops.plan_table_delete(doc, table))
    for node in para_deletes:
        ops.append(blockops.plan_para_delete(doc, node))

    if errors:
        raise ImportError_("\n".join(errors))
    return ops, report


def _remap_row_positions(doc: Document, report: Report):
    """표 신설/삭제 후 위치 기반 표 ID(t###)가 밀리므로, 재파싱 문서의 새 ID
    순서에 맞춰 row_positions를 재대응한다 (annotate·check 연속성)."""
    if not report._inserted_tables and not report._deleted_tables:
        return
    events: list[tuple[int, int, list | None]] = []
    for t in doc.tables.values():
        if t.id in report._deleted_tables:
            continue
        events.append((t.start, -1, report.row_positions.get(t.id)))
    for pos, order, sheet_rows in report._inserted_tables:
        events.append((pos, order, sheet_rows))
    events.sort(key=lambda e: (e[0], e[1]))
    report.row_positions = {
        f"t{i:03d}": rows
        for i, (_, _, rows) in enumerate(events, 1)
        if rows is not None
    }


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
    span_tpl = _pick_span_template(table, insert_idx) if template is None else None
    if template is None and span_tpl is None:
        template = _pick_template(table, insert_idx, allow_hspan=True)   # 표 전체가 가로 병합 꼴일 때만
    if template is None and span_tpl is None:
        errors.append(f"표 {table.id}: 복제 가능한 템플릿 행(ACOPY=Y, 병합 없음)이 없음")
        return

    payload = b""
    for entry in new_entries:
        if template is not None:
            tpl_idx = index_of[template.id]
            texts: dict[int, str] = {}
            for ci, cell in enumerate(template.cells):
                col = grid_mod.cell_grid_col(table, tpl_idx, cell)
                texts[ci] = _cell_value_to_text(entry.values.get(2 + col))
            payload += splice.ROW_SEP + splice.clone_row(doc.data, template, texts)
        else:
            texts_by_col = {c: _cell_value_to_text(entry.values.get(2 + c)) for c in range(table.ncols)}
            payload += splice.ROW_SEP + _clone_filling_spans(doc.data, table, span_tpl, texts_by_col)
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


def _pick_template(table: Table, insert_idx: int, allow_hspan: bool = False):
    """삽입 지점 주변에서 복제 안전한 행 선택: TBODY + ACOPY=Y + 병합 무관 + 전체 폭.
    가로 병합(COLSPAN) 행은 빼는 것이 기본 - '합  계'처럼 구분 · 이름 열을 합친 줄을 복제하면
    새 행의 칸이 하나 모자라 값이 사라진다. allow_hspan 은 그런 행밖에 없는 표의 마지막 대안."""
    g = grid_mod.expand(table)

    def ok(i: int) -> bool:
        row = table.rows[i]
        if row.in_thead or not row.acopy:
            return False
        if any(gc is not None and gc.kind == "vspan" for gc in g[i]):
            return False
        if any(c.rowspan > 1 for c in row.cells):
            return False
        if not allow_hspan and any(c.colspan > 1 for c in row.cells):
            return False
        return True

    candidates = list(range(insert_idx - 1, -1, -1)) + list(range(insert_idx, len(table.rows)))
    for i in candidates:
        if ok(i):
            return table.rows[i]
    return None


def _pick_span_template(table: Table, insert_idx: int):
    """보통 행이 없을 때(본문이 모두 세로 병합 묶음): 묶음의 이어지는 행(세로 병합 칸만 빠지고
    자기 칸은 병합 없음)을 고른다. 복제할 때 빠진 열은 묶음 주인 칸을 ROWSPAN 없이 채운다."""
    g = grid_mod.expand(table)

    def ok(i: int) -> bool:
        row = table.rows[i]
        if row.in_thead or not row.acopy or not row.cells:
            return False
        if any(c.rowspan > 1 or c.colspan > 1 for c in row.cells):
            return False
        return any(gc is not None and gc.kind == "vspan" for gc in g[i]) and all(gc is not None for gc in g[i])

    candidates = list(range(insert_idx - 1, -1, -1)) + list(range(insert_idx, len(table.rows)))
    return next((i for i in candidates if ok(i)), None)


def _clone_filling_spans(data: bytes, table: Table, idx: int, texts_by_col: dict[int, str]) -> bytes:
    """idx 행(세로 병합 묶음 안 행)을 복제해 전체 폭 새 행을 만든다: 자기 칸은 그대로, 세로 병합
    자리(vspan)는 묶음 주인 칸을 ROWSPAN 없이 · 이 행 높이로 복제. 칸 사이 공백은 원문 관례."""
    row = table.rows[idx]
    g = grid_mod.expand(table)[idx]
    cells = row.cells
    sep = data[cells[0].end:cells[1].start] if len(cells) > 1 else b"\r\n"
    m = re.search(rb'HEIGHT="(\d+)"', data[cells[0].start:cells[0].inner_start])
    height = m.group(1) if m else None
    parts = []
    for c, gc in enumerate(g):
        if gc.kind == "cell":
            parts.append(splice.cell_with_text(data, gc.node, texts_by_col.get(c, "")))
        elif gc.kind == "vspan" and (c == 0 or g[c - 1].node is not gc.node):
            parts.append(splice.cell_with_text(data, gc.node, texts_by_col.get(c, ""),
                                               drop_rowspan=True, height=height))
    return data[row.start:cells[0].start] + sep.join(parts) + data[cells[-1].end:row.end]


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
    _remap_row_positions(doc, report)
    return package, new_contents, report
