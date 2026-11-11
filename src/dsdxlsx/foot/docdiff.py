"""DSD 두 개 비교 — 정렬 기반 (위치 ID에 의존하지 않는다).

dsdxlsx verify 는 t014.r007 같은 위치 ID를 그대로 맞대어 표가 하나 끼어들면 뒤 전부가
어긋났다. 여기서는 3단 시퀀스 정렬로 '같은 것'을 먼저 찾고 그 안에서 셀을 비교한다:

1. 구획: sections.refined_sections → 키 (주석은 제목, 재무제표는 표준 명칭, 기타는 제목)
2. 블록: 표(헤더 텍스트 + 첫 열 라벨) / 문단(헤딩 접두 뗀 텍스트) / 라벨 시그니처
3. 행·셀: grid 전개 후 행 키(첫 비공백 텍스트, THEAD 여부) 정렬 → 그리드 열별 비교

숫자 셀은 값 차이(delta)까지, 그 외는 텍스트 변경으로 분류한다.
"""
from __future__ import annotations

import difflib
import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from . import sections
from .constants import norm
from .sections import _APPENDIX_RE, _NOTE_LINE_RE, Sec
from .tableview import classify_ext
from .. import grid as grid_mod
from ..dsd_io import read_dsd
from ..model import Block, Node, Table
from ..parse import parse_contents

# ── 결과 모델 ────────────────────────────────────────────────────────

@dataclass
class CellDiff:
    col: int
    kind: str            # value | text
    old: str
    new: str
    delta: Decimal | None = None


@dataclass
class RowDiff:
    status: str          # same | changed | added | deleted
    old_cells: list[str] = field(default_factory=list)   # 그리드 열별 표시 텍스트 (앵커 없는 span 칸은 "")
    new_cells: list[str] = field(default_factory=list)
    cells: list[CellDiff] = field(default_factory=list)   # changed 일 때 열별 차이
    in_thead: bool = False
    old_span: list[tuple[int, int]] = field(default_factory=list)  # (colspan, rowspan) per col, 0 = span 피지배 칸
    new_span: list[tuple[int, int]] = field(default_factory=list)
    old_idx: int | None = None   # 표 안 행 인덱스 (grid 행) — DART 대사 등 후속 매핑용
    new_idx: int | None = None

    @property
    def label(self) -> str:
        for t in (self.new_cells or self.old_cells):
            if t.strip():
                return t.strip()
        return ""


@dataclass
class BlockDiff:
    kind: str            # table | para | label
    status: str          # same | changed | added | deleted
    old: Block | None = None
    new: Block | None = None
    rows: list[RowDiff] = field(default_factory=list)   # table 일 때
    old_text: str = ""   # para/label
    new_text: str = ""
    note: str = ""       # '열 수 5→6' 등

    @property
    def n_changed_cells(self) -> int:
        return sum(len(r.cells) for r in self.rows)


@dataclass
class SectionDiff:
    key: tuple
    title: str
    status: str          # same | changed | added | deleted
    old: Sec | None = None
    new: Sec | None = None
    renumber: tuple[int, int] | None = None
    blocks: list[BlockDiff] = field(default_factory=list)

    @property
    def anchor(self) -> str:
        return "sec-" + re.sub(r"[^0-9A-Za-z가-힣]+", "-", "-".join(map(str, self.key)))


COUNT_KEYS = ("값변경", "텍스트변경", "행추가", "행삭제", "표추가", "표삭제",
              "문단변경", "문단추가", "문단삭제", "구획추가", "구획삭제", "주석재번호")


@dataclass
class DiffReport:
    old_name: str
    new_name: str
    company: str
    sections: list[SectionDiff]
    counts: dict[str, int]

    @property
    def has_changes(self) -> bool:
        return any(v for v in self.counts.values())


# ── 시그니처·키 ─────────────────────────────────────────────────────

def _cell_text(node: Node | None) -> str:
    return node.text if node is not None else ""


def _grid_texts(table: Table) -> list[list[str]]:
    g = grid_mod.expand(table)
    return [[_cell_text(gc.node) if gc is not None and gc.kind == "cell" else ""
             for gc in row] for row in g]


def _grid_spans(table: Table) -> list[list[tuple[int, int]]]:
    g = grid_mod.expand(table)
    return [[(gc.node.colspan, gc.node.rowspan) if gc is not None and gc.kind == "cell" else (0, 0)
             for gc in row] for row in g]


def _strip_heading(text: str) -> str:
    """주석 헤딩('12. 유형자산') 접두를 뗀 본문 — 재번호는 문단 변경으로 세지 않는다."""
    lines = text.split("\n")
    k = next((i for i, ln in enumerate(lines) if ln.strip()), None)
    if k is not None:
        m = _NOTE_LINE_RE.match(lines[k].strip())
        if m:
            lines[k] = m.group(2)
    return "\n".join(lines)


def _table_sig(table: Table) -> tuple:
    texts = _grid_texts(table)
    heads = tuple(norm(t) for i, row in enumerate(texts) if table.rows[i].in_thead for t in row)
    labels = tuple(norm(next((t for t in row if t.strip()), "")) for row in texts)
    return ("t", table.ncols, heads, labels)


def _block_sig(b: Block) -> tuple:
    if b.kind == "table":
        return _table_sig(b.item)
    if b.kind == "label":
        return ("l", norm(b.item.text))
    return ("p", norm(_strip_heading(b.item.text)))


def _block_flat(b: Block) -> str:
    if b.kind == "table":
        return "\n".join("\t".join(r) for r in _grid_texts(b.item))
    return b.item.text


def _sec_key(sec: Sec) -> tuple:
    if sec.note_no is not None:
        return ("note", "별첨" if sec.is_appendix else "", norm(sec.note_title))
    n = norm(sec.title)
    return ("sec", n)


def _row_labels(table: Table) -> list[str]:
    out = []
    for row in _grid_texts(table):
        t = norm(next((x for x in row if x.strip()), ""))
        if t and classify_ext(t) is None:
            out.append(t)
    return out


def _label_containment(ta: Table, tb: Table) -> float:
    """두 표 행 라벨 집합의 포함도 |A∩B| / min(|A|,|B|) (라벨 3개 미만이면 0)."""
    la, lb = set(_row_labels(ta)), set(_row_labels(tb))
    if len(la) < 3 or len(lb) < 3:
        return 0.0
    return len(la & lb) / min(len(la), len(lb))


def _ratio(a: str, b: str) -> float:
    if not a and not b:
        return 1.0
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


# ── 정렬 유틸: SequenceMatcher + replace 구간 재매칭 ────────────────────

def _align(old_keys: list, new_keys: list, old_items: list, new_items: list,
           pair_score, threshold: float) -> list[tuple]:
    """(old_item|None, new_item|None) 쌍 목록. equal 은 1:1, replace 구간은 점수 그리디."""
    sm = difflib.SequenceMatcher(None, old_keys, new_keys, autojunk=False)
    out: list[tuple] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            out.extend(zip(old_items[i1:i2], new_items[j1:j2]))
        elif tag == "delete":
            out.extend((o, None) for o in old_items[i1:i2])
        elif tag == "insert":
            out.extend((None, n) for n in new_items[j1:j2])
        else:  # replace — 유사도 그리디 매칭 (순서 유지)
            olds = old_items[i1:i2]
            news = new_items[j1:j2]
            cands = []
            for a, o in enumerate(olds):
                for b, n in enumerate(news):
                    s = pair_score(o, n)
                    if s >= threshold:
                        cands.append((s, a, b))
            cands.sort(key=lambda x: (-x[0], x[1], x[2]))
            used_a: set[int] = set()
            used_b: set[int] = set()
            pairs: dict[int, int] = {}
            for s, a, b in cands:
                if a in used_a or b in used_b:
                    continue
                # 순서 교차 방지: 이미 짝지은 (a',b') 와 대소 관계가 뒤집히면 건너뜀
                if any((a < a2) != (b < b2) for a2, b2 in pairs.items()):
                    continue
                pairs[a] = b
                used_a.add(a)
                used_b.add(b)
            # 순서대로 방출
            a = b = 0
            while a < len(olds) or b < len(news):
                if a in pairs:
                    tb = pairs[a]
                    while b < tb:
                        out.append((None, news[b]))
                        b += 1
                    out.append((olds[a], news[tb]))
                    a += 1
                    b = tb + 1
                elif a < len(olds):
                    # 이 old 는 짝 없음 — 다만 앞선 new 중 짝 없는 것 먼저 흘려보내지 않고 old 삭제로
                    out.append((olds[a], None))
                    a += 1
                else:
                    out.append((None, news[b]))
                    b += 1
    return out


# ── 표 비교 ──────────────────────────────────────────────────────────

def _classify_cell(old: str, new: str, col: int) -> CellDiff:
    ov, nv = classify_ext(old.strip()) if old.strip() else None, classify_ext(new.strip()) if new.strip() else None
    if old.strip() == "-":
        ov = 0
    if new.strip() == "-":
        nv = 0
    if ov is not None and nv is not None:
        return CellDiff(col, "value", old, new, Decimal(str(nv)) - Decimal(str(ov)))
    return CellDiff(col, "text", old, new)


def _row_key(texts: list[str], in_thead: bool) -> tuple:
    return (norm(next((t for t in texts if t.strip()), "")), in_thead)


def compare_tables(ta: Table, tb: Table) -> tuple[list[RowDiff], str]:
    ga, gb = _grid_texts(ta), _grid_texts(tb)
    sa, sb = _grid_spans(ta), _grid_spans(tb)
    ka = [_row_key(r, ta.rows[i].in_thead) for i, r in enumerate(ga)]
    kb = [_row_key(r, tb.rows[i].in_thead) for i, r in enumerate(gb)]
    ia, ib = list(range(len(ga))), list(range(len(gb)))
    pairs = _align(ka, kb, ia, ib,
                   lambda i, j: _ratio("\t".join(ga[i]), "\t".join(gb[j])), 0.5)
    ncols = min(ta.ncols, tb.ncols)
    note = f"열 수 {ta.ncols}→{tb.ncols}" if ta.ncols != tb.ncols else ""
    rows: list[RowDiff] = []
    for i, j in pairs:
        if i is None:
            rows.append(RowDiff("added", [], gb[j], [], tb.rows[j].in_thead, [], sb[j], None, j))
        elif j is None:
            rows.append(RowDiff("deleted", ga[i], [], [], ta.rows[i].in_thead, sa[i], [], i, None))
        else:
            cells = [_classify_cell(ga[i][c], gb[j][c], c)
                     for c in range(ncols) if norm(ga[i][c]) != norm(gb[j][c])]
            if ta.ncols != tb.ncols:
                extra_old = any(t.strip() for t in ga[i][ncols:])
                extra_new = any(t.strip() for t in gb[j][ncols:])
                if extra_old or extra_new:
                    cells.append(CellDiff(ncols, "text", " | ".join(ga[i][ncols:]), " | ".join(gb[j][ncols:])))
            rows.append(RowDiff("changed" if cells else "same", ga[i], gb[j], cells,
                                tb.rows[j].in_thead, sa[i], sb[j], i, j))
    return rows, note


# ── 구획·블록 비교 ────────────────────────────────────────────────────

def _compare_blocks(old_blocks: list[Block], new_blocks: list[Block],
                    counts: dict[str, int]) -> list[BlockDiff]:
    ok = [_block_sig(b) for b in old_blocks]
    nk = [_block_sig(b) for b in new_blocks]

    def score(a: Block, b: Block) -> float:
        if a.kind != b.kind:
            return 0.0
        r = _ratio(_block_flat(a), _block_flat(b))
        if a.kind == "table":
            # 행 라벨 포함도 — 양식이 달라도(열 수·행 수 차이) 같은 표면 라벨은 대부분 겹친다
            # (연차 재무상태표 ↔ 분기 재무상태표 등). 숫자 셀은 기수가 바뀌면 전부 달라
            # 텍스트 유사도만으로는 0.4 대에 머문다.
            r = max(r, _label_containment(a.item, b.item))
        return r

    pairs = _align(ok, nk, old_blocks, new_blocks, score, 0.5)
    out: list[BlockDiff] = []
    for a, b in pairs:
        if a is None:
            kind = b.kind if b.kind != "p" else "para"
            bd = BlockDiff(kind, "added", None, b, new_text=b.item.text if kind != "table" else "")
            if kind == "table":
                bd.rows = [RowDiff("added", [], r, [], b.item.rows[i].in_thead, [], s)
                           for i, (r, s) in enumerate(zip(_grid_texts(b.item), _grid_spans(b.item)))]
                counts["표추가"] += 1
            else:
                counts["문단추가"] += 1
            out.append(bd)
        elif b is None:
            kind = a.kind if a.kind != "p" else "para"
            bd = BlockDiff(kind, "deleted", a, None, old_text=a.item.text if kind != "table" else "")
            if kind == "table":
                bd.rows = [RowDiff("deleted", r, [], [], a.item.rows[i].in_thead, s, [])
                           for i, (r, s) in enumerate(zip(_grid_texts(a.item), _grid_spans(a.item)))]
                counts["표삭제"] += 1
            else:
                counts["문단삭제"] += 1
            out.append(bd)
        elif a.kind == "table":
            rows, note = compare_tables(a.item, b.item)
            changed = any(r.status != "same" for r in rows) or bool(note)
            bd = BlockDiff("table", "changed" if changed else "same", a, b, rows, note=note)
            for r in rows:
                if r.status == "added":
                    counts["행추가"] += 1
                elif r.status == "deleted":
                    counts["행삭제"] += 1
                for c in r.cells:
                    counts["값변경" if c.kind == "value" else "텍스트변경"] += 1
            out.append(bd)
        else:
            kind = "label" if a.kind == "label" else "para"
            same = norm(_strip_heading(a.item.text)) == norm(_strip_heading(b.item.text)) \
                if kind == "para" else norm(a.item.text) == norm(b.item.text)
            bd = BlockDiff(kind, "same" if same else "changed", a, b,
                           old_text=a.item.text, new_text=b.item.text)
            if not same:
                counts["문단변경"] += 1
            out.append(bd)
    return out


def _sections_of(path: Path):
    package = read_dsd(str(path))
    doc = parse_contents(package.contents)
    return package, doc, _live_sections(doc)


def _live_sections(doc) -> list[Sec]:
    # 주석 구획은 blocks 가 비어 있어도(혼합 헤딩 경계) 유지되지만 비교엔 의미 없음
    return [s for s in sections.refined_sections(doc) if s.blocks]


def compare(old_path: str | Path, new_path: str | Path) -> DiffReport:
    old_path, new_path = Path(old_path), Path(new_path)
    pa, da, sa = _sections_of(old_path)
    pb, db, sb = _sections_of(new_path)
    return compare_docs(da, db, old_path.name, new_path.name, sections.company_name(pb.meta))


def compare_docs(doc_a, doc_b, old_name: str = "old", new_name: str = "new",
                 company: str = "") -> DiffReport:
    """파싱된 두 Document 비교 — DART 원문(전기 보고서) 대사 등이 파일 없이 재사용."""
    sa, sb = _live_sections(doc_a), _live_sections(doc_b)
    counts = {k: 0 for k in COUNT_KEYS}
    ka, kb = [_sec_key(s) for s in sa], [_sec_key(s) for s in sb]

    def sec_score(x: Sec, y: Sec) -> float:
        if (x.note_no is None) != (y.note_no is None):
            return 0.0
        return _ratio("\n".join(_block_flat(b) for b in x.blocks),
                      "\n".join(_block_flat(b) for b in y.blocks))

    pairs = _align(ka, kb, sa, sb, sec_score, 0.6)
    out: list[SectionDiff] = []
    for a, b in pairs:
        if a is None:
            sd = SectionDiff(_sec_key(b), _sec_title(b), "added", None, b)
            sd.blocks = _compare_blocks([], b.blocks, counts)
            # 구획 단위로 세고 블록 단위 카운트는 되돌린다
            _uncount_blocks(sd.blocks, counts)
            counts["구획추가"] += 1
            out.append(sd)
        elif b is None:
            sd = SectionDiff(_sec_key(a), _sec_title(a), "deleted", a, None)
            sd.blocks = _compare_blocks(a.blocks, [], counts)
            _uncount_blocks(sd.blocks, counts)
            counts["구획삭제"] += 1
            out.append(sd)
        else:
            sd = SectionDiff(_sec_key(b), _sec_title(b), "same", a, b)
            if a.note_no is not None and b.note_no is not None and a.note_no != b.note_no:
                sd.renumber = (a.note_no, b.note_no)
                counts["주석재번호"] += 1
            sd.blocks = _compare_blocks(a.blocks, b.blocks, counts)
            if sd.renumber or any(bd.status != "same" for bd in sd.blocks):
                sd.status = "changed"
            out.append(sd)
    return DiffReport(old_name, new_name, company, out, counts)


def _uncount_blocks(blocks: list[BlockDiff], counts: dict[str, int]) -> None:
    for bd in blocks:
        if bd.kind == "table":
            counts["표추가" if bd.status == "added" else "표삭제"] -= 1
        else:
            counts["문단추가" if bd.status == "added" else "문단삭제"] -= 1


def _sec_title(s: Sec) -> str:
    if s.note_no is not None:
        pre = "[별첨] " if s.is_appendix else ""
        return f"{pre}주석 {s.note_no}. {s.note_title}"
    return s.title


def print_summary(r: DiffReport) -> None:
    print(f"\n[DSD 비교] {r.old_name} → {r.new_name}" + (f" ({r.company})" if r.company else ""))
    if not r.has_changes:
        print("  차이 없음")
        return
    parts = [f"{k} {v}" for k, v in r.counts.items() if v]
    print("  " + " · ".join(parts))
    for s in r.sections:
        if s.status == "same":
            continue
        n_cells = sum(bd.n_changed_cells for bd in s.blocks)
        tag = {"added": "추가", "deleted": "삭제", "changed": "변경"}[s.status]
        extra = f" 재번호 {s.renumber[0]}→{s.renumber[1]}" if s.renumber else ""
        print(f"  - [{tag}] {s.title}{extra}" + (f" (셀 {n_cells}건)" if n_cells else ""))
