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
import re
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


# ── 목차(--outline) ──────────────────────────────────────────────────

OUTLINE_SCHEMA = "dsdxlsx.outline/1"
FIND_SCHEMA = "dsdxlsx.find/1"


def _one_line(text: str, width: int = 60) -> str:
    s = " ".join(text.split())
    return s if len(s) <= width else s[:width - 1] + "…"


def outline_dict(doc: Document, *, only: list[str] | None = None, source: str = "",
                 data: bytes | None = None) -> dict:
    """구획마다 표·문단 ID 와 크기만 - 큰 문서에서 어디를 읽을지 고르는 용도."""
    secs = []
    for sec in sections(doc):
        if only and sec.key not in only and sec.sheet not in only:
            continue
        blocks = []
        for block in sec.blocks:
            if block.kind == "table":
                t = block.item
                tv = table_view(t, sec)
                d = {"type": "table", "id": t.id, "ncols": t.ncols, "nrows": len(t.rows),
                     "headers": [tv.col_header(c) for c in range(t.ncols)],
                     "labels": [tv.row_label(i).strip() for i in tv.body if tv.row_label(i).strip()][:4]}
                unit = _unit(t)
                if unit:
                    d["unit"] = unit
            else:
                text = block.item.text
                d = {"type": "para" if block.kind == "p" else "label", "id": block.item.id,
                     "chars": len(text), "text": _one_line(text)}
            blocks.append(d)
        secs.append({"section": sec.key, "kind": sec.kind, "title": sec.title,
                     "tables": sum(1 for b in blocks if b["type"] == "table"),
                     "paras": sum(1 for b in blocks if b["type"] == "para"), "blocks": blocks})
    out = {"schema": OUTLINE_SCHEMA, "file": source}
    if data is not None:
        out["sha256"] = hashlib.sha256(data).hexdigest()
    out["sections"] = secs
    return out


def outline_markdown(d: dict) -> str:
    secs = d["sections"]
    lines = [f"# {d.get('file') or 'DSD'} - 목차", "",
             f"구획 {len(secs)}개, 표 {sum(s['tables'] for s in secs)}개, "
             f"문단 {sum(s['paras'] for s in secs)}개. 내용은 inspect --section <값> 으로 본다.", ""]
    for sec in secs:
        head = (f"주석 {sec['section']} {sec['title']}" if sec["kind"] == "note" else sec["title"])
        lines.append(f"## {head} [section: {sec['section']}] - 표 {sec['tables']}, 문단 {sec['paras']}")
        for b in sec["blocks"]:
            if b["type"] == "table":
                unit = f", 단위 {b['unit']}" if b.get("unit") else ""
                heads = " / ".join(_one_line(h, 14) for h in b["headers"] if h)
                labels = ", ".join(_one_line(x, 14) for x in b["labels"])
                lines.append(f"- {b['id']} 표 {b['ncols']}열×{b['nrows']}행{unit}: {heads}"
                             + (f" - {labels}" if labels else ""))
            else:
                kind = "문단" if b["type"] == "para" else "제목"
                lines.append(f"- {b['id']} {kind}({b['chars']}자): {b['text']}")
        lines.append("")
    return "\n".join(lines)


# ── 찾기(--find) ─────────────────────────────────────────────────────

def _snippet(text: str, query: str, width: int = 40) -> str:
    flat = " ".join(text.split())
    k = flat.replace(" ", "").lower().find(query)
    if k < 0 or len(flat) <= 2 * width:
        return _one_line(flat, 2 * width)
    # 공백을 뺀 위치 k → 원문 위치로 대략 환산
    seen = pos = 0
    for pos, ch in enumerate(flat):
        if ch != " ":
            if seen == k:
                break
            seen += 1
    a = max(0, pos - width)
    return ("…" if a else "") + flat[a:pos + width] + ("…" if pos + width < len(flat) else "")


_QNUM_RE = re.compile(r"^(\()?(-)?(\d[\d,]*(?:\.\d+)?)\)?$")


def _query_number(q: str):
    """찾는 글이 숫자면 그 값('1200', '1,200', '(1,200)', '-1200')."""
    m = _QNUM_RE.match(q.replace(" ", ""))
    if not m:
        return None
    v = float(m.group(3).replace(",", ""))
    if m.group(1) or m.group(2):
        v = -v
    return int(v) if v == int(v) else v


def find_dict(doc: Document, queries: list[str], *, only: list[str] | None = None,
              source: str = "", limit: int = 200) -> dict:
    """글(공백·대소문자 무시)이나 숫자 값이 든 문단·행·칸의 위치를 찾는다."""
    wants = [(q, norm(q).lower(), _query_number(q)) for q in queries if q and q.strip()]
    hits: list[dict] = []
    total = 0

    def add(hit: dict) -> None:
        nonlocal total
        total += 1
        if len(hits) < limit:
            hits.append(hit)

    def match(text: str):
        flat = norm(text).lower()
        value = cell_value(text)
        for q, nq, qv in wants:
            if nq and nq in flat:
                return q
            if qv is not None and value is not None and qv == value and qv != 0:
                return q
        return None

    for sec in sections(doc):
        if only and sec.key not in only and sec.sheet not in only:
            continue
        for block in sec.blocks:
            if block.kind != "table":
                q = match(block.item.text)
                if q is not None:
                    add({"section": sec.key, "type": "para" if block.kind == "p" else "label",
                         "id": block.item.id, "query": q,
                         "text": _snippet(block.item.text, norm(q).lower())})
                continue
            table = block.item
            tv = table_view(table, sec)
            g = grid_mod.expand(table)
            for i, row in enumerate(table.rows):
                label = tv.row_label(i)
                for c, gc in enumerate(g[i]):
                    if gc is None or gc.kind != "cell" or not gc.node.text.strip():
                        continue
                    q = match(gc.node.text)
                    if q is None:
                        continue
                    if gc.node.text.strip() == label.strip() and not row.in_thead:
                        add({"section": sec.key, "type": "row", "id": row.id, "table": table.id,
                             "query": q, "label": label.strip()})
                    else:
                        add({"section": sec.key, "type": "cell", "id": gc.node.id, "row": row.id,
                             "table": table.id, "query": q, "label": label.strip(), "col": c,
                             "header": tv.col_header(c), "text": gc.node.text,
                             **({"head": True} if row.in_thead else {})})
    return {"schema": FIND_SCHEMA, "file": source, "queries": [q for q, _, _ in wants],
            "total": total, "truncated": total > len(hits), "hits": hits}


def find_markdown(d: dict) -> str:
    qs = ", ".join(f'"{q}"' for q in d["queries"])
    lines = [f"# {d.get('file') or 'DSD'} - 찾기: {qs}", "",
             f"{d['total']}건" + (f"(앞 {len(d['hits'])}건만 - --limit 로 늘림)" if d["truncated"] else ""),
             ""]
    for h in d["hits"]:
        where = f"[section: {h['section']}]"
        if h["type"] == "row":
            lines.append(f"- {where} 행 {h['id']} '{h['label']}'")
        elif h["type"] == "cell":
            head = " 머리글" if h.get("head") else ""
            lines.append(f"- {where} 칸{head} {h['id']} = {_md(h['text'])} "
                         f"(행 '{h['label']}', 열 [{h['col']}] {_one_line(h['header'], 20)})")
        else:
            kind = "문단" if h["type"] == "para" else "제목"
            lines.append(f"- {where} {kind} {h['id']}: {h['text']}")
    return "\n".join(lines)


def dumps(d: dict, pretty: bool = False) -> str:
    if pretty:
        return json.dumps(d, ensure_ascii=False, indent=1)
    return json.dumps(d, ensure_ascii=False, separators=(",", ":"))
