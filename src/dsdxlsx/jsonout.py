"""리뷰용 구조화 산출 (`check --json`) — AI 검토 스킬(/dsd-review)의 입력.

문서 IR 요약(문단·표·행 라벨·FS 행+주석참조·주석 목록·PGBRK 위치·표 높이)과
검증 결과(findings·커버리지·대사맵)를 하나의 JSON으로 내보낸다. 도구는 결정론
데이터만 내놓고, 오타·주기 적절성·페이지 분할 같은 판단은 스킬(AI)이 한다.

AI 소견의 회신 계약(annotate-ai 입력)은 ai_findings_schema() 참조.
"""
from __future__ import annotations

import re
from decimal import Decimal

from .check import CheckReport, TableView, build_views, norm, use_profile
from .export_xlsx import STATEMENT_KEYWORDS, refined_sections
from .model import Document
from .profile import Profile

SCHEMA = 1


def _jsonable(v):
    if isinstance(v, Decimal):
        return float(v)
    return v


def _sec_label(sec) -> str:
    return f"주석{sec.note_no}" if sec.note_no is not None else sec.title


def _table_height(table) -> int:
    total = 0
    for row in table.rows:
        try:
            total += int(row.attrs.get("HEIGHT", 0))
        except ValueError:
            pass
    return total


def _pgbrks(doc: Document) -> list[dict]:
    """PGBRK 바이트 오프셋 → 직전 블록 ID (페이지 분할 후보 판단 재료)."""
    blocks = []
    for b in doc.blocks:
        item = b.item
        blocks.append((item.start, item.end, item.id))
    out = []
    for m in re.finditer(rb"<PGBRK\b", doc.data):
        off = m.start()
        prev_id = None
        for start, end, bid in blocks:
            if end <= off:
                prev_id = bid
            else:
                break
        out.append({"offset": off, "after_block": prev_id})
    return out


def _fs_rows(ctx) -> list[dict]:
    """재무제표 행: 라벨·주석참조·기간별 금액(표 단위 환산 전 원값 + 단위)."""
    out = []
    for tv in ctx["statements"]:
        note_col = next((c for c in range(tv.ncols)
                         if "주석" in norm(tv.col_header(c))), None)
        for i in tv.body:
            label = tv.row_label(i).strip()
            if not label:
                continue
            refs = ([int(x) for x in re.findall(r"\d+", tv.text(i, note_col))]
                    if note_col is not None else [])
            amounts = {}
            for c in range(tv.ncols):
                b = tv.period_bucket(c)
                v = tv.value(i, c)
                if b and b not in amounts and v is not None:
                    amounts[b] = _jsonable(v)
            if not amounts and not refs:
                continue
            out.append({
                "section": tv.sec_title, "table": tv.table.id,
                "row_id": tv.table.rows[i].id, "label": label,
                "refs": refs, "amounts": amounts, "unit": tv.unit,
            })
    return out


def _tie_point_dict(pt) -> dict:
    d = {"tie": pt.tie.to_dict(), "bucket": pt.bucket}
    if pt.a is not None and pt.a_col is not None:
        d["a_row_id"] = pt.a.row_id
        d["a_value"] = _jsonable(pt.a.value(pt.a_col))
    if pt.b is not None and pt.b_col is not None:
        d["b_row_id"] = pt.b.row_id
        d["b_value"] = _jsonable(pt.b.value(pt.b_col))
    if pt.ok is not None:
        d["ok"] = pt.ok
        if pt.diff is not None:
            d["diff"] = _jsonable(pt.diff)
    if pt.reason:
        d["reason"] = pt.reason
    return d


def build_review_json(doc: Document, report: CheckReport,
                      tie_points: list | None = None,
                      suggestion_points: list | None = None,
                      profile: Profile | None = None,
                      source: str = "") -> dict:
    with use_profile(profile):
        _, ctx = build_views(doc)
        sections, paragraphs, tables, notes = [], [], [], []
        for sec in refined_sections(doc):
            label = _sec_label(sec)
            kind = ("statement" if sec.title in STATEMENT_KEYWORDS
                    else "note" if sec.note_no is not None
                    else "notes_intro" if sec.notes_intro else "other")
            sections.append({"label": label, "kind": kind, "title": sec.title,
                             "note_no": sec.note_no, "note_title": sec.note_title})
            if sec.note_no is not None:
                notes.append({"no": sec.note_no, "title": sec.note_title})
            for block in sec.blocks:
                if block.kind in ("p", "label"):
                    text = block.item.text
                    if text.strip():
                        paragraphs.append({"id": block.item.id, "section": label,
                                           "kind": block.kind, "text": text})
                elif block.kind == "table":
                    tv = TableView(block.item, label, sec_title=sec.title)
                    tables.append({
                        "id": block.item.id, "section": label,
                        "n_rows": len(block.item.rows), "ncols": block.item.ncols,
                        "height": _table_height(block.item),
                        "headers": [tv.col_header(c) for c in range(tv.ncols)],
                        "unit": tv.unit,
                        "rows": [{"id": block.item.rows[i].id,
                                  "label": tv.row_label(i).strip()}
                                 for i in tv.body if tv.row_label(i).strip()],
                    })

        return {
            "schema": SCHEMA, "tool": "dsdxlsx", "source": source,
            "profile": report.profile_name,
            "checks": {
                "counts": report.counts,
                "findings": [{"check": f.check, "level": f.level,
                              "where": f.where, "message": f.message}
                             for f in report.findings],
                "suppressed": [{"check": f.check, "where": f.where,
                                "message": f.message, "reason": reason}
                               for f, reason in report.suppressed],
            },
            "coverage": report.coverage,
            "ties": {
                "confirmed": [_tie_point_dict(p) for p in (tie_points or [])],
                "suggested": [_tie_point_dict(p) for p in (suggestion_points or [])],
            },
            "document": {
                "sections": sections,
                "notes": notes,
                "paragraphs": paragraphs,
                "tables": tables,
                "fs_rows": _fs_rows(ctx),
                "pgbrks": _pgbrks(doc),
            },
        }


def ai_findings_schema() -> dict:
    """annotate-ai 입력 계약 (스킬이 이 형태로 findings.json을 만든다)."""
    return {
        "schema": 1,
        "findings": [{
            "kind": "오타 | 주기 | 페이지 | 대사 | 기타",
            "severity": "warn | info",
            "row_id": "(선택) t014.r004 — 표 행이면 하이퍼링크 생성",
            "section": "(선택) 주석4 / 재무상태표 — 위치 설명",
            "message": "발견 내용 (필수)",
            "suggestion": "(선택) 수정 제안",
        }],
    }
