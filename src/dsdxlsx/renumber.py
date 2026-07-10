"""주석 경계·재번호 엔진.

재번호 캐스케이드는 위험도 계층으로 나눈다 (설계 문서·플랜 확정):
    (a) 주석 헤딩 P의 선두 'N.'      → 자동 치환 (파서가 이미 식별)
    (b) 재무제표 주석참조 열 '6,28,31' → 자동 치환 (구분자 보존)
    (c) 서술문 '주석 N (참조)' 멘션    → 확인필요 목록만 (변형 표기 오치환 방지)

안전 원칙:
- 번호가 바뀌는 주석은 **자기 헤딩 P**(분리 문단, 'N.'으로 시작)가 있어야 한다.
  혼합 P 헤딩(앞 문단 꼬리에 붙은 'N. 제목' — 실측 존재)은 안전 치환 불가 → 거부.
- 독립 renumber 명령은 재번호 전/후 **C4 대사 결과의 동형성**(참조 판정이 매핑으로
  정확히 이행)을 게이트로 검사한다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from . import splice
from .blockops import prev_sep_start
from .check import build_views, norm, use_profile
from .export_xlsx import _NOTE_LINE_RE, refined_sections
from .model import Document, Node

# 서술문 멘션: '주석 8', '주석 8, 10' — 목록 형태까지 한 덩어리로
MENTION_RE = re.compile(r"주석\s*\d+(\s*[,·]\s*\d+)*")


class RenumberError(ValueError):  # ValueError 계열 — CLI 최상위에서 '오류:'로 출력
    pass


@dataclass
class NoteInfo:
    no: int
    title: str
    blocks: list                     # Sec.blocks (kind/item)
    heading: Node | None             # 자기 헤딩 P (분리 문단일 때만)

    @property
    def own_heading(self) -> bool:
        return self.heading is not None

    @property
    def start(self) -> int:
        return self.blocks[0].item.start

    @property
    def end(self) -> int:
        return self.blocks[-1].item.end


def note_infos(doc: Document) -> list[NoteInfo]:
    out: list[NoteInfo] = []
    for sec in refined_sections(doc):
        if sec.note_no is None or not sec.blocks:
            continue
        heading = None
        first = sec.blocks[0]
        if first.kind == "p":
            # _split_notes와 같은 기준: 첫 '비어있지 않은' 줄 — 선두 빈 줄(&cr;)은
            # 별도 문단 헤딩으로 취급 (치환 정규식의 \s*가 흡수)
            lines = [ln for ln in first.item.text.split("\n") if ln.strip()]
            m = _NOTE_LINE_RE.match(lines[0].strip()) if lines else None
            if m and int(m.group(1)) == sec.note_no:
                heading = first.item
        out.append(NoteInfo(no=sec.note_no, title=sec.note_title,
                            blocks=sec.blocks, heading=heading))
    return out


def deletable(infos: list[NoteInfo], target: NoteInfo) -> str | None:
    """삭제 가능성 검사 — 사유 반환 (None=가능).

    자기 헤딩이 없으면 경계가 앞 문단에 붙어 있고, 다음 주석이 자기 헤딩이
    없으면 이 주석의 마지막 P 꼬리에 다음 헤딩이 붙어 있다 — 둘 다 거부."""
    if not target.own_heading:
        return (f"주석 {target.no}의 헤딩이 별도 문단이 아님(혼합 문단) — "
                f"DSD 편집기에서 문단을 분리한 뒤 다시 시도")
    idx = infos.index(target)
    if idx + 1 < len(infos) and not infos[idx + 1].own_heading:
        return (f"다음 주석 {infos[idx + 1].no}의 헤딩이 주석 {target.no}의 "
                f"마지막 문단에 붙어 있어 범위를 자를 수 없음")
    for block in target.blocks:
        if block.kind == "label":
            return f"주석 {target.no} 범위에 제목(TITLE) 요소가 있음"
        if block.kind == "table":
            t = block.item
            if t.group_aclass is not None or t.in_insertion:
                return f"주석 {target.no} 범위에 그룹/라이브러리 표({t.id})가 있음"
            if any(c.tag == "TU" for r in t.rows for c in r.cells):
                return f"주석 {target.no} 범위에 자동 필드(TU) 표({t.id})가 있음"
    return None


def plan_note_delete(doc: Document, info: NoteInfo) -> splice.Op:
    start = prev_sep_start(doc.data, info.start)
    return splice.Op(pos=start, start=start, end=info.end, payload=b"")


def build_close_gap_mapping(nums: list[int]) -> dict[int, int]:
    """번호 공백 밀착: [1,2,3,5,6] → {5:4, 6:5}."""
    mapping = {}
    for i, n in enumerate(sorted(nums), 1):
        if n != i:
            mapping[n] = i
    return mapping


@dataclass
class RenumberReport:
    headings: list[str] = field(default_factory=list)   # 자동 치환된 헤딩
    ref_cells: list[str] = field(default_factory=list)  # 자동 치환된 FS 참조 셀
    mentions: list[str] = field(default_factory=list)   # 확인필요 서술 멘션 (위치+원문)


def _sub_ref_cell(text: str, mapping: dict[int, int]) -> str:
    """'6,28,31' — 숫자만 매핑, 구분자·공백 보존."""
    return re.sub(r"\d+",
                  lambda m: str(mapping.get(int(m.group()), int(m.group()))),
                  text)


def renumber_ops(doc: Document, mapping: dict[int, int],
                 profile=None) -> tuple[list[splice.Op], RenumberReport]:
    """매핑에 따른 결정론 치환 op + 확인필요 목록. 안전 위반은 RenumberError."""
    mapping = {a: b for a, b in mapping.items() if a != b}
    if not mapping:
        return [], RenumberReport()
    if len(set(mapping.values())) != len(mapping):
        raise RenumberError(f"매핑 충돌(같은 번호로 두 개 이상): {mapping}")
    infos = note_infos(doc)
    by_no = {i.no: i for i in infos}
    existing = set(by_no)
    for src, dst in mapping.items():
        if src not in existing:
            raise RenumberError(f"주석 {src}이(가) 문서에 없음")
        if dst in existing and dst not in mapping:
            raise RenumberError(f"주석 {src}→{dst}: {dst}번이 이미 있고 이동하지 않음")
        if not by_no[src].own_heading:
            raise RenumberError(
                f"주석 {src}의 헤딩이 별도 문단이 아님(혼합 문단) — 재번호 불가. "
                f"DSD 편집기에서 문단을 분리한 뒤 다시 시도")

    ops: list[splice.Op] = []
    report = RenumberReport()

    # (a) 헤딩 P — 선두 'N.'만 치환
    for src, dst in mapping.items():
        node = by_no[src].heading
        new_text = re.sub(rf"^(\s*){src}\.", rf"\g<1>{dst}.", node.text, count=1)
        ops.append(splice.replace_inner(node, new_text))
        report.headings.append(f"{src}. {by_no[src].title} → {dst}.")

    # (b) 재무제표 주석참조 열 — 구분자 보존 치환
    with use_profile(profile):
        _, ctx = build_views(doc)
        for tv in ctx["statements"]:
            note_col = next((c for c in range(tv.ncols)
                             if "주석" in norm(tv.col_header(c))), None)
            if note_col is None:
                continue
            for i in tv.body:
                cell = tv.cell(i, note_col)
                if cell is None or not cell.text.strip():
                    continue
                nums = [int(x) for x in re.findall(r"\d+", cell.text)]
                if not any(n in mapping for n in nums):
                    continue
                ops.append(splice.replace_inner(cell, _sub_ref_cell(cell.text, mapping)))
                report.ref_cells.append(
                    f"{tv.sec_title} '{tv.row_label(i).strip()[:20]}': "
                    f"{cell.text.strip()} → {_sub_ref_cell(cell.text, mapping)}")

    # (c) 서술문 멘션 — 자동 치환하지 않고 확인필요 목록.
    # 헤딩 문단은 번호 줄(첫 줄)만 제외 — 본문 줄의 '주석 N' 멘션은 검출 대상
    heading_ids = {i.heading.id for i in infos if i.own_heading}
    for node in doc.nodes.values():
        if node.tag == "TU":
            continue
        text = node.text
        if node.id in heading_ids:
            text = text.split("\n", 1)[1] if "\n" in text else ""
        for m in MENTION_RE.finditer(text):
            nums = [int(x) for x in re.findall(r"\d+", m.group())]
            if any(n in mapping for n in nums):
                report.mentions.append(f"{node.id}: '{m.group()}' — 수동 확인 필요")
    return ops, report


def c4_signature(doc: Document, profile=None) -> list[tuple]:
    """C4 동형성 게이트용 서명 — refs는 매핑으로 변하므로 제외하고
    (표, 행, 열, 판정)만. 재번호가 대사 판정을 바꾸면 동형성이 깨진다."""
    from .check import run_checks

    report = run_checks(doc, only={"C4"}, profile=profile)
    return sorted((x.table_id, x.row_id, x.grid_col, x.ok) for x in report.results
                  if x.check == "C4")
