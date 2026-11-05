"""inspect - DSD 를 LLM·스크립트가 읽는 형태(JSON·마크다운)로 펼친다.

엑셀로 내보내는 구획(재무제표별·주석 번호별)만 다룬다 - apply 가 고칠 수 있는 범위와 같다.
칸마다 두 가지 좌표를 준다.
- ID: t003.r004.c02 (문서 순서로 매기므로 같은 파일이면 늘 같다. 구조 편집 뒤에는 바뀐다)
- 사람 좌표: 구획("재무상태표" | 주석 번호 "4") + 행 이름 + 열 머리글(또는 당기·전기)
apply 는 둘 다 받는다.
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass

from . import grid as grid_mod
from . import numbers
from .check import TableView, is_total_label, norm
from .export_xlsx import exported_sections, sanitize_sheet_title
from .model import Document, Node, Table

SCHEMA = "dsdxlsx.inspect/1"


@dataclass
class SectionInfo:
    key: str          # apply 의 section 값: 재무제표 명칭 또는 주석 번호 문자열
    sheet: str        # export 시트 이름(같은 규칙)
    kind: str         # statement | note | notes_intro
    title: str        # 주석 제목(주석이 아니면 key 와 같음)
    blocks: list


def sections(doc: Document) -> list[SectionInfo]:
    used: set[str] = set()
    out = []
    for sec in exported_sections(doc):
        sheet = sanitize_sheet_title(sec.title, used)
        if sec.note_no is not None:
            kind, key, title = "note", str(sec.note_no), sec.note_title
        elif sec.notes_intro:
            kind, key, title = "notes_intro", "주석", sec.title
        else:
            kind, key, title = "statement", sec.title, sec.title
        out.append(SectionInfo(key, sheet, kind, title, sec.blocks))
    return out


def cell_value(text: str):
    """칸 글자 → 숫자(int/float) 또는 None. '-' 는 0."""
    v = numbers.classify(text)
    if v is not None:
        return v if isinstance(v, int) else float(v)
    if text.strip() == "-":
        return 0
    return None


def table_view(table: Table, sec: SectionInfo) -> TableView:
    return TableView(table, sec.key, sec_title=sec.key)


def _unit(table: Table) -> str | None:
    for row in table.rows[:2]:
        for cell in row.cells:
            t = cell.text.strip()
            if t.startswith("(단위"):
                return t.strip("()").split(":", 1)[-1].strip()
    return None


def table_dict(table: Table, sec: SectionInfo) -> dict:
    tv = table_view(table, sec)
    g = grid_mod.expand(table)
    cols = []
    for c in range(table.ncols):
        cols.append({"col": c, "header": tv.col_header(c), "period": tv.period_bucket(c)})
    rows = []
    for i, row in enumerate(table.rows):
        cells = []
        for c, gc in enumerate(g[i]):
            if gc is None or gc.kind != "cell":
                continue
            node: Node = gc.node
            d = {"id": node.id, "col": c, "text": node.text}
            v = cell_value(node.text)
            if v is not None:
                d["value"] = v
            if node.colspan > 1:
                d["colspan"] = node.colspan
            if node.rowspan > 1:
                d["rowspan"] = node.rowspan
            if node.readonly:
                d["readonly"] = True
            cells.append(d)
        r = {"id": row.id, "label": tv.row_label(i), "cells": cells}
        if row.in_thead:
            r["head"] = True
        elif is_total_label(tv.row_label(i)):
            r["total"] = True
        rows.append(r)
    d = {"type": "table", "id": table.id, "ncols": table.ncols, "columns": cols, "rows": rows}
    unit = _unit(table)
    if unit:
        d["unit"] = unit
    locks = []
    if table.fixtable:
        locks.append("고정 표 - 행 추가·삭제 불가")
    if table.irregular:
        locks.append("격자가 불규칙 - 행 추가·삭제 불가")
    if table.has_autosum:
        locks.append("자동합계 칸 있음 - 고친 뒤 편집기에서 합계 확인")
    if locks:
        d["notes"] = locks
    return d


def doc_dict(doc: Document, *, only: list[str] | None = None, source: str = "",
             data: bytes | None = None) -> dict:
    secs = []
    for sec in sections(doc):
        if only and sec.key not in only and sec.sheet not in only:
            continue
        blocks = []
        for block in sec.blocks:
            if block.kind == "table":
                blocks.append(table_dict(block.item, sec))
            elif block.kind == "p":
                blocks.append({"type": "para", "id": block.item.id, "text": block.item.text})
            else:  # 제목 라벨 - 고칠 수 없음
                blocks.append({"type": "label", "id": block.item.id, "text": block.item.text})
        secs.append({"section": sec.key, "kind": sec.kind, "title": sec.title, "blocks": blocks})
    out = {"schema": SCHEMA, "file": source}
    if data is not None:
        out["sha256"] = hashlib.sha256(data).hexdigest()
    out["sections"] = secs
    return out


# ── 마크다운 ─────────────────────────────────────────────────────────

def _md(text: str) -> str:
    """표 칸 글자 → 마크다운 한 칸. 앞 공백(들여쓰기 = 계층)은 줄바꿈 없는 공백으로 남긴다."""
    s = text.replace("|", "\\|").replace("\n", " / ").rstrip()
    body = s.lstrip(" ")
    return "\u00a0" * (len(s) - len(body)) + body


def to_markdown(d: dict) -> str:
    lines = [f"# {d.get('file') or 'DSD'}", "",
             "표의 칸은 행 ID(t003.r004) + 열로 가리킨다. 열은 머리글 앞 [번호](0부터)나 머리글 글자, 당기·전기다. "
             "구획 제목 끝의 [section] 값이 apply 의 section 이다.", ""]
    for sec in d["sections"]:
        head = (f"주석 {sec['section']} {sec['title']}" if sec["kind"] == "note"
                else sec["title"])
        lines.append(f"## {head} [section: {sec['section']}]")
        lines.append("")
        for b in sec["blocks"]:
            if b["type"] in ("para", "label"):
                lines.append(f"[{b['id']}] {_md(b['text'])}")
                lines.append("")
                continue
            meta = [b["id"]]
            if b.get("unit"):
                meta.append(f"단위 {b['unit']}")
            meta.extend(b.get("notes", []))
            lines.append(f"[{' / '.join(meta)}]")
            hdr = ["행 ID"] + [
                f"[{col['col']}] " + _md(col["header"])
                + (f" ({col['period']})" if col["period"] else "") for col in b["columns"]]
            lines.append("| " + " | ".join(hdr) + " |")
            lines.append("|" + "---|" * len(hdr))
            for r in b["rows"]:
                vals = [""] * b["ncols"]
                for cell in r["cells"]:
                    vals[cell["col"]] = _md(cell["text"]) + (" (자동)" if cell.get("readonly") else "")
                tag = " 머리" if r.get("head") else (" 합계" if r.get("total") else "")
                lines.append(f"| {r['id']}{tag} | " + " | ".join(vals) + " |")
            lines.append("")
    return "\n".join(lines)


def dumps(d: dict, pretty: bool = False) -> str:
    if pretty:
        return json.dumps(d, ensure_ascii=False, indent=1)
    return json.dumps(d, ensure_ascii=False, separators=(",", ":"))
