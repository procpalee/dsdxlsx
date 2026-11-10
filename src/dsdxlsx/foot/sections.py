"""섹션 분리 + 시트 계획.

dsdxlsx export_xlsx.py의 refined_sections를 포팅해 재무제표/주석 섹션으로 나누고,
기준 양식대로 시트 이름(BS/PL/CE/CF, 주석은 "1".."N")을 배정한다.
표지·감사의견·외부감사 실시내용 등 나머지 섹션은 산출물에서 제외한다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .constants import (
    SECOND_PL_CODE,
    STATEMENT_KEYWORDS,
    STATEMENT_SHEET_CODES,
    norm,
)
from ..model import Block, Document

# 주석 헤딩: "6. 현금및현금성자산" — 제목 첫 글자가 숫자면(하위 헤딩 "2.1 …") 제외
_NOTE_LINE_RE = re.compile(r"^(\d{1,2})\.\s*([^\d\s][^\n]{0,40})")
_APPENDIX_RE = re.compile(r"\[별첨\s*\d+\]")


@dataclass
class Sec:
    title: str
    blocks: list[Block] = field(default_factory=list)
    note_no: int | None = None
    note_title: str = ""
    notes_intro: bool = False
    is_appendix: bool = False


# ---- dsdxlsx 포팅: TITLE 구획 → 재무제표별/주석번호별 세분화 ----------------

def _statement_kw(block: Block) -> str | None:
    """재무제표 제목 테이블 판별: 첫 행 첫 셀이 '분 기 재 무 상 태 표' 류."""
    if block.kind != "table":
        return None
    table = block.item
    if not table.rows or not table.rows[0].cells:
        return None
    first = norm(table.rows[0].cells[0].text)
    if not first or len(first) > 20:
        return None
    for kw in STATEMENT_KEYWORDS:
        if kw in first:
            return kw
    return None


def _split_statements(title: str, blocks: list[Block]) -> list[Sec]:
    groups = [Sec(title)]
    for block in blocks:
        kw = _statement_kw(block)
        if kw:
            groups.append(Sec(kw, [block]))
        else:
            groups[-1].blocks.append(block)
    return [g for g in groups if g.blocks]


def _note_heading(line: str, current_n: int) -> tuple[int, str] | None:
    """주석 번호는 단조증가만 인정 — 본문 내 번호 목록·하위 헤딩 오검출 방지."""
    m = _NOTE_LINE_RE.match(line.strip())
    if m and int(m.group(1)) > current_n:
        return int(m.group(1)), m.group(2).strip()
    return None


def _split_notes(title: str, blocks: list[Block]) -> list[Sec]:
    groups = [Sec(title, notes_intro=True)]
    current_n = 0
    appendix = False
    for block in blocks:
        if block.kind == "p":
            lines = [ln for ln in block.item.text.split("\n") if ln.strip()]
            first = lines[0] if lines else ""
            if _APPENDIX_RE.search(first):
                appendix = True
                current_n = 0
            head = _note_heading(first, current_n) if first else None
            if head:
                current_n, note_title = head
                groups.append(Sec(str(current_n), [block],
                                  note_no=current_n, note_title=note_title,
                                  is_appendix=appendix))
            else:
                groups[-1].blocks.append(block)
            # 문단 중간에서 새 주석이 시작되면(혼합 P) 이후 블록은 새 그룹으로
            last = None
            scan = lines[1:] if head else lines
            for ln in scan:
                if _APPENDIX_RE.search(ln):
                    appendix = True
                    current_n = 0
                    continue
                h = _note_heading(ln, current_n)
                if h:
                    current_n, t = h
                    last = (current_n, t, appendix)
            if last:
                groups.append(Sec(str(last[0]), [],
                                  note_no=last[0], note_title=last[1],
                                  is_appendix=last[2]))
        else:
            groups[-1].blocks.append(block)
    # 혼합 문단의 헤딩 이후 내용은 layout.compute_render_items가 spill로 다음 Sec에
    # 넘긴다. 따라서 자체 블록이 아직 없는 경계 Sec도 유지해야 한다.
    return [g for g in groups if g.blocks or g.note_no is not None]


def refined_sections(doc: Document) -> list[Sec]:
    """TITLE 구획을 재무제표별/주석번호별로 세분화.

    재무제표가 TITLE 구획 자체인 양식(K-GAAP: '재 무 상 태 표')은 자간 벌림을
    정규화해 표준 명칭 구획으로 인식한다 (IFRS 양식은 '(첨부)재무제표' 내부 분할)."""
    out: list[Sec] = []
    for title, blocks in doc.sections():
        n = norm(title)
        kw = next((k for k in STATEMENT_KEYWORDS if k in n), None)
        if "주석" in n:
            out.extend(_split_notes(title, blocks))
        elif kw and "보고서" not in n:
            out.append(Sec(kw, blocks))
        elif "재무제표" in n and "보고서" not in n:
            out.extend(_split_statements(title, blocks))
        else:
            out.append(Sec(title, blocks))
    return out


# ---- 시트 계획 --------------------------------------------------------------

@dataclass
class SheetPlan:
    """산출물 시트 하나 — 이름·종류·구성 섹션.

    주석 시트 이름은 실측 규칙대로 **순서 기반 "1".."N"** — 원문 주석 번호가
    아니다 (실측 사례: 원문에 '2. 중요한 회계정책' 대제목이 없어 번호가 1,3..24로
    벌어져도 기준 양식 시트는 1..23). 원문 번호는 sec.note_no에 남아 있으며 재무제표의
    "(주석n)" 참조 → 시트명 매핑은 note_sheet_map()으로 한다.
    """

    name: str          # "BS" | "PL" | "CE" | "CF" | ("CI"/"RE") | "1".."N"(순서)
    kind: str          # "statement" | "note"
    sec: Sec
    display_title: str  # 총괄표 제목 열 텍스트 (재무제표: 원문 표기, 주석: 주석 제목)


def _display_title(sec: Sec) -> str:
    """재무제표 섹션의 원문 표기 제목 ('재 무 상 태 표' — 자간 유지)."""
    for block in sec.blocks:
        if block.kind == "label" and sec.title in norm(block.item.text):
            return block.item.text.strip()
        if block.kind == "table":
            table = block.item
            if table.rows and table.rows[0].cells:
                first = table.rows[0].cells[0].text
                if sec.title in norm(first):
                    return first.strip()
    return sec.title


def build(doc: Document) -> list[SheetPlan]:
    """산출물에 들어갈 시트 목록 (재무제표 → 주석 순, 원문·총괄표 제외)."""
    stmts: list[SheetPlan] = []
    notes: list[SheetPlan] = []
    used_codes: set[str] = set()
    for sec in refined_sections(doc):
        if sec.note_no is not None:
            notes.append(SheetPlan("", "note", sec, sec.note_title or f"주석 {sec.note_no}"))
        elif sec.title in STATEMENT_KEYWORDS:
            code = STATEMENT_SHEET_CODES[sec.title]
            if code in used_codes:
                if code == "PL":
                    code = SECOND_PL_CODE  # 손익+포괄손익 병행
                if code in used_codes:
                    continue  # 같은 계산서 중복 — 첫 것만
            used_codes.add(code)
            stmts.append(SheetPlan(code, "statement", sec, _display_title(sec)))
    # 문서 순서를 그대로 유지한다. 별첨 안에서는 번호가 1부터 다시 시작할 수 있어
    # 번호 정렬을 하면 정규 주석 사이로 별첨이 끼어드는 문제가 생긴다.
    for i, p in enumerate(notes, start=1):
        p.name = str(i)
    return stmts + notes


def note_sheet_map(plans: list[SheetPlan]) -> dict[int, str]:
    """원문 주석 번호 → 시트 이름 ("(주석4)" 참조 해석용)."""
    # 별첨 내부 번호는 정규 주석 참조 번호가 아니며 중복될 수 있다.
    return {p.sec.note_no: p.name for p in plans
            if p.kind == "note" and not p.sec.is_appendix}


def company_name(meta: bytes | None) -> str:
    """meta.xml DOCUMENT-HEADER regname (로그 표시용)."""
    if not meta:
        return ""
    m = re.search(rb'<DOCUMENT-HEADER\b[^>]*?regname="([^"]*)"', meta)
    return m.group(1).decode("utf-8", "replace") if m else ""
