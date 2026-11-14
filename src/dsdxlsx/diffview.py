"""두 DSD 비교 보고서(diff) - 두 판을 나란히 놓고 문단·표 행 높이를 맞춰 HTML 한 장으로.

용도: 우리 DSD ↔ 감사인이 고쳐 보낸 DSD, 직전 판 ↔ 이번 판처럼 같은 문서의 두 판을 비교한다.
XML 속성·줄바꿈처럼 화면에 안 보이는 차이는 무시하고 '보이는 글'만 비교한다.

- 문서를 (주석 이름, 문단 | 표 행) 목록으로 펴고 순서를 맞춘다(difflib).
- 바뀐 곳마다 종류를 붙인다: 숫자·글·공백·'-'(빈 칸·'-'·0 끼리)·병합(채운 칸은 같고 칸 수만 다름)·추가·삭제.
- 기준 DSD(--base, 예: 직전 분기 최종)를 주면 바뀐 곳마다 기준에 어느 쪽 모양이 있었는지 표시한다
  (A 가 기준과 같음 = 이어받은 것·B 가 기준과 같음 = A 쪽에서 바뀐 것·둘 다 없음 = 이번에 새로 생긴 줄).
- 보고서는 두 화면: '나란히 보기'(두 문서 전체를 왼쪽·오른쪽에, 표는 DSD 모양 그대로 병합 칸까지)와
  '바뀐 곳만'(바뀐 곳 카드 목록).
"""
from __future__ import annotations

import difflib
import html
import re
from dataclasses import dataclass, field
from itertools import zip_longest

from .grid import expand
from .model import Document, Node, Table

NUM = re.compile(r"\(?-?\d[\d,]*(?:\.\d+)?\)?%?")
NOTE_HEAD = re.compile(r"^\s*(\d{1,2})\.\s*([^\d\s].{0,38})$")
STATEMENTS = ("재무상태표", "포괄손익계산서", "손익계산서", "자본변동표", "현금흐름표")
DASHES = {"", "-", "0", "–"}
NUMCELL = re.compile(r"^[\d,.()\-%△▲ ]*\d[\d,.()\-%△▲ ]*$")   # 숫자 칸 - 한 줄로 두고 넘치면 글자만 줄임

KINDS = {  # 종류: (이름, '표시만' 여부)
    "number": ("숫자", False),
    "text": ("글", False),
    "added": ("추가", False),
    "deleted": ("삭제", False),
    "dash": ("'-'·빈 칸", True),
    "space": ("공백", True),
    "layout": ("병합·칸", True),
}


@dataclass
class Item:
    note: str
    kind: str          # 'p' | 'row'
    ref: str           # 문단·행 ID
    cells: list[str]
    empty: bool = False                                                  # 빈 행(나란히 보기에서만 씀)
    node: Node | None = field(default=None, compare=False, repr=False)   # 문단·라벨
    table: Table | None = field(default=None, compare=False, repr=False)
    ri: int = field(default=-1, compare=False, repr=False)               # 표 안 행 번호

    @property
    def text(self) -> str:
        return " | ".join(self.cells) if self.kind == "row" else self.cells[0]


@dataclass
class Change:
    kind: str
    note: str
    a: Item | None
    b: Item | None
    base: str = ""     # 'a' | 'b' | 'both' | 'none' | '' (기준 없음)
    detail: dict = field(default_factory=dict)
    idx: int = field(default=-1, compare=False)


@dataclass
class Entry:
    """나란히 보기 한 줄 - 왼쪽(x)·오른쪽(y) 문단이나 표 행. 한쪽이 None 이면 빈 자리."""
    x: Item | None
    y: Item | None
    c: Change | None = None
    blank: bool = False   # 빈 행끼리 맞춘 줄(차이로 세지 않음)


@dataclass
class Diff:
    changes: list[Change]
    entries: list[Entry]
    a: Document | None = None
    b: Document | None = None


def _norm(t: str) -> str:
    return re.sub(r"\s+", " ", t.replace("\n", " ")).strip()


def _shape(t: str) -> str:
    return NUM.sub("#", _norm(t))


def _period(t: str) -> str:
    """기준(직전 분·반기) 대조용 - 기간 낱말 차이(반기 ↔ 분기)는 같은 것으로."""
    return re.sub(r"반(\s?)기", lambda m: "분" + m.group(1) + "기", t)


def flatten(doc: Document, keep_empty: bool = False) -> list[Item]:
    """문서 순서대로 문단·표 행을 편다. 주석 이름은 'N. 제목' 으로 시작하는 문단에서 잡는다.
    keep_empty: 빈 표 행도 넣는다(empty=True) - 나란히 보기에서 병합 칸 모양을 지키려고."""
    out: list[Item] = []
    for section, blocks in doc.sections():
        note = section
        for b in blocks:
            if b.kind in ("p", "label"):
                t = b.item.text.strip()
                if not t:
                    continue
                m = NOTE_HEAD.match(t.split("\n", 1)[0])
                if m:
                    note = f"{m.group(1)}. {m.group(2).strip()}"
                out.append(Item(note, "p", b.item.id, [t], node=b.item))
            else:
                for ri, r in enumerate(b.item.rows):
                    cells = [c.text.strip() for c in r.cells]
                    head = re.sub(r"^(분기|반기|요약)+", "", re.sub(r"\s", "", cells[0])) if cells else ""
                    if head in STATEMENTS:                 # 재무제표 제목표('재 무 상 태 표')
                        note = head
                    if any(cells) or keep_empty:
                        out.append(Item(note, "row", r.id, cells, empty=not any(cells), table=b.item, ri=ri))
    return out


def _classify(a: Item, b: Item) -> str:
    ta, tb = _norm(a.text), _norm(b.text)
    if re.sub(r"\s", "", ta) == re.sub(r"\s", "", tb):
        return "space"
    if a.kind == b.kind == "row":
        fa = [_norm(c) for c in a.cells if _norm(c)]
        fb = [_norm(c) for c in b.cells if _norm(c)]
        if fa == fb:
            return "layout"
        da = [("-" if _norm(c) in DASHES else _norm(c)) for c in a.cells]
        db = [("-" if _norm(c) in DASHES else _norm(c)) for c in b.cells]
        if da == db:
            return "dash"
        sa = ["#" if _norm(c) in DASHES else _shape(c) for c in a.cells]   # '-' ↔ 숫자도 숫자 차이
        sb = ["#" if _norm(c) in DASHES else _shape(c) for c in b.cells]
        if sa == sb:
            return "number"
    if _shape(ta) == _shape(tb):
        return "number"
    return "text"


def _pair(xs: list[Item], ys: list[Item]) -> list[tuple[Item | None, Item | None]]:
    """바뀐 덩어리 안에서 비슷한 줄끼리 짝짓기(순서 유지). 짝 없는 줄은 추가·삭제."""
    out: list[tuple[Item | None, Item | None]] = []
    j = 0
    for x in xs:
        best, bi = 0.0, -1
        for k in range(j, len(ys)):
            r = difflib.SequenceMatcher(None, _shape(x.text), _shape(ys[k].text), autojunk=False).ratio()
            if r > best:
                best, bi = r, k
        if bi >= 0 and best >= 0.5:
            out += [(None, y) for y in ys[j:bi]]
            out.append((x, ys[bi]))
            j = bi + 1
        else:
            out.append((x, None))
    out += [(None, y) for y in ys[j:]]
    return out


def _align(fa: list[Item], fb: list[Item]) -> tuple[list[Change], list[Entry]]:
    """빈 행을 뺀 줄로 순서를 맞추고(차이 판정), 빈 행은 원래 자리에 되돌려 나란히 보기 줄을 만든다."""
    na = [i for i in fa if not i.empty]
    nb = [i for i in fb if not i.empty]
    sm = difflib.SequenceMatcher(None, [_norm(x.text) for x in na], [_norm(x.text) for x in nb], autojunk=False)
    changes: list[Change] = []
    pairs: list[tuple[Item | None, Item | None, Change | None]] = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            pairs += [(na[i], nb[j], None) for i, j in zip(range(i1, i2), range(j1, j2))]
            continue
        for x, y in _pair(na[i1:i2], nb[j1:j2]):
            if x and y:
                c = Change(_classify(x, y), y.note, x, y)
            elif x:
                c = Change("deleted", x.note, x, None)
            else:
                c = Change("added", y.note, None, y)
            c.idx = len(changes)
            changes.append(c)
            pairs.append((x, y, c))

    entries: list[Entry] = []
    pos_a = {id(it): k for k, it in enumerate(fa)}
    pos_b = {id(it): k for k, it in enumerate(fb)}
    pa = pb = 0

    def flush(ea: list[Item], eb: list[Item]) -> None:
        for e1, e2 in zip_longest(ea, eb):
            entries.append(Entry(e1, e2, None, blank=True))

    for x, y, c in pairs:
        ea: list[Item] = []
        eb: list[Item] = []
        if x is not None:
            k = pos_a[id(x)]
            ea, pa = fa[pa:k], k + 1
        if y is not None:
            k = pos_b[id(y)]
            eb, pb = fb[pb:k], k + 1
        flush(ea, eb)
        if x and y and x.kind != y.kind:           # 문단 ↔ 표 행 짝은 삭제 + 추가 두 줄로
            entries += [Entry(x, None, c), Entry(None, y, c)]
        else:
            entries.append(Entry(x, y, c))
    flush(fa[pa:], fb[pb:])
    return changes, entries


def _mark_base(changes: list[Change], base: Document) -> None:
    items = flatten(base)                  # 같은 주석 안에서만 찾음(짧은 '소 계' 줄이 다른 주석과 겹치지 않게)
    texts = {(_period(i.note), _period(_norm(i.text))) for i in items}
    shapes = {(_period(i.note), _period(_shape(i.text))) for i in items}
    blob = " ".join(_period(_norm(i.text)) for i in items if i.kind == "p")    # 문단은 나뉜 자리가 판마다 달라
    sblob = " ".join(_period(_shape(i.text)) for i in items if i.kind == "p")  # 글 안에 들어 있는지로 봄

    def has(it: Item | None) -> int:   # 2 = 같은 글, 1 = 같은 모양(숫자만 다름), 0 = 없음
        if it is None:
            return 0
        k, tn, ts = _period(it.note), _period(_norm(it.text)), _period(_shape(it.text))
        if it.kind == "p":
            return 2 if tn in blob else (1 if ts in sblob else 0)
        return 2 if (k, tn) in texts else (1 if (k, ts) in shapes else 0)

    for c in changes:
        ha, hb = has(c.a), has(c.b)
        if ha > hb:
            c.base = "a"
        elif hb > ha:
            c.base = "b"
        else:
            c.base = "both" if ha else "none"


def diff(a: Document, b: Document, base: Document | None = None) -> Diff:
    """바뀐 곳 목록 + 나란히 보기 줄 목록."""
    changes, entries = _align(flatten(a, keep_empty=True), flatten(b, keep_empty=True))
    if base is not None:
        _mark_base(changes, base)
    return Diff(changes, entries, a, b)


def compare(a: Document, b: Document, base: Document | None = None) -> list[Change]:
    return diff(a, b, base).changes


def summary(changes: list[Change]) -> dict:
    by_kind: dict[str, int] = {}
    by_note: dict[str, int] = {}
    for c in changes:
        by_kind[c.kind] = by_kind.get(c.kind, 0) + 1
        by_note[c.note] = by_note.get(c.note, 0) + 1
    return {"total": len(changes), "by_kind": by_kind, "by_note": by_note}


CHANGE_LIMIT = 500


def _item_json(it: Item | None) -> dict | None:
    if it is None:
        return None
    if it.kind == "p":
        return {"type": "para", "id": it.ref, "text": it.cells[0]}
    return {"type": "row", "id": it.ref, "label": next((c for c in it.cells if c.strip()), ""),
            "cells": list(it.cells)}


def changes_json(changes: list[Change], limit: int = CHANGE_LIMIT) -> tuple[list[dict], bool]:
    """봉투용 바뀐 곳 목록 - 숫자·글·추가·삭제를 먼저, 표시만 다른 것(공백·'-'·병합)을 뒤에.
    표 행끼리 바뀐 것은 달라진 칸(cols: col·before·after)도 싣는다. limit 를 넘으면 자르고 True."""
    order = sorted(changes, key=lambda c: (KINDS.get(c.kind, ("", False))[1], c.idx))
    out: list[dict] = []
    for c in order[:limit]:
        row = {"kind": KINDS.get(c.kind, (c.kind,))[0], "section": c.note,
               "before": _item_json(c.a), "after": _item_json(c.b)}
        if c.a and c.b and c.a.kind == c.b.kind == "row":
            row["cols"] = [{"col": i, "before": x or "", "after": y or ""}
                           for i, (x, y) in enumerate(zip_longest(c.a.cells, c.b.cells)) if (x or "") != (y or "")]
        if c.base:
            row["base"] = c.base
        out.append(row)
    return out, len(order) > limit


# ── HTML 공통 ─────────────────────────────────────────

def _e(t) -> str:
    return html.escape(str(t)).replace("\n", "<br>")


def _char_diff(a: str, b: str, from_right: bool = False) -> tuple[str, str]:
    """바뀐 글자를 <del>(a)·<ins>(b) 로 감싼다. from_right: 숫자처럼 오른쪽 끝부터 자리를 맞춘다."""
    ra, rb = (a[::-1], b[::-1]) if from_right else (a, b)
    ops = difflib.SequenceMatcher(None, ra, rb, autojunk=False).get_opcodes()
    sa: list[tuple[int, int, bool]] = []   # (시작, 끝, 바뀜) - 원래 글 기준
    sb: list[tuple[int, int, bool]] = []
    for op, i1, i2, j1, j2 in ops:
        if from_right:
            i1, i2, j1, j2 = len(a) - i2, len(a) - i1, len(b) - j2, len(b) - j1
        if i2 > i1:
            sa.append((i1, i2, op != "equal"))
        if j2 > j1:
            sb.append((j1, j2, op != "equal"))

    def build(t: str, segs: list, tag: str) -> str:
        return "".join(f"<{tag}>{_e(t[i:j])}</{tag}>" if ch else _e(t[i:j]) for i, j, ch in sorted(segs))
    return build(a, sa, "del"), build(b, sb, "ins")


def _mark(t: str, o: str | None, side: str) -> str:
    """이쪽 글 t 를 반대쪽 o 와 견줘 바뀐 글자만 칠한다(거의 다 다르면 통째로)."""
    if not t:
        return ""
    if o is None or not o or _norm(t) == _norm(o):
        return _e(t)
    a, b = (t, o) if side == "a" else (o, t)
    if difflib.SequenceMatcher(None, a, b, autojunk=False).ratio() < 0.4:
        return f"<del>{_e(t)}</del>" if side == "a" else f"<ins>{_e(t)}</ins>"
    xa, xb = _char_diff(a, b, from_right=bool(NUMCELL.match(a) and NUMCELL.match(b)))
    return xa if side == "a" else xb


BASE_TXT = {
    "a": "기준과 같은 쪽: {A} (기준을 이어받음)",
    "b": "기준과 같은 쪽: {B} ({A}에서 달라짐)",
    "both": "기준에 두 모양 모두 있음",
    "none": "기준에 없음 (이번에 생긴 줄)",
}
BASE_SHORT = {"a": "{A}가 이어받음", "b": "{A}에서 달라짐", "both": "기준에 둘 다", "none": "새로 생긴 줄"}


# ── 나란히 보기 ───────────────────────────────────────

def _style(node: Node, cell: bool = True) -> str:
    """DSD 셀·문단 속성 중 화면에 보이는 것(정렬·굵게·글자 크기)만 옮긴다."""
    st = []
    al = node.attrs.get("ALIGN", "").upper()
    if al in ("RIGHT", "CENTER", "LEFT", "JUSTIFY"):
        st.append(f"text-align:{al.lower()}")
    va = node.attrs.get("VALIGN", "").upper()
    if cell and va in ("TOP", "BOTTOM"):
        st.append(f"vertical-align:{va.lower()}")
    um = node.attrs.get("USERMARK", "").split()
    if any(u == "B" or u.startswith("F-B") for u in um):
        st.append("font-weight:700")
    for u in um:
        m = re.match(r"F-B?T?(\d+)$", u)
        if m:
            st.append(f"font-size:{max(0.75, min(1.6, int(m.group(1)) / 10)):.2f}em")
    return f' style="{";".join(st)}"' if st else ""


def _col_widths(doc: Document | None, t: Table, n: int) -> list[float]:
    ws: list[int] = []
    if doc is not None:
        seg = doc.data[t.start:t.end]
        m = re.search(rb"<COLGROUP.*?</COLGROUP>", seg, re.S)
        if m:
            ws = [int(x) for x in re.findall(rb'<COL\b[^>]*?WIDTH="(\d+)"', m.group(0))]
    if len(ws) != n or not sum(ws):
        ws = [1] * n
    return [w / sum(ws) for w in ws]


HALF = 49.25   # 양쪽 폭(%) - 가운데 1.5% 는 바뀐 줄 표시 띠


def _changed_cells(e: Entry) -> set[int]:
    if not (e.c and e.x and e.y and e.x.kind == e.y.kind == "row") or e.c.kind == "layout":
        return set()
    cx, cy = e.x.cells, e.y.cells
    return {k for k in range(max(len(cx), len(cy)))
            if _norm(cx[k] if k < len(cx) else "") != _norm(cy[k] if k < len(cy) else "")}


def _side_table(entries: list[Entry], side: str, t: Table | None, doc: Document | None
                ) -> tuple[list[str], list[str]]:
    """한쪽 표를 줄 목록(entries)에 맞춰 그린다 - 반대쪽에만 있는 줄은 빈 자리(빗금)로,
    ROWSPAN 은 빈 자리까지 덮도록 늘리고, 이 덩어리 밖에서 시작한 병합은 빈 칸으로 잇는다.
    반환: (줄마다 td html, col html)."""
    n = len(entries)
    items = [(e.x if side == "a" else e.y) for e in entries]
    if t is None:                                      # 이쪽엔 표가 없음 - 빗금 한 칸으로
        out = [""] * n
        out[0] = f'<td class="{side} ph" rowspan="{n}"></td>' if n > 1 else f'<td class="{side} ph"></td>'
        return out, [f'<col style="width:{HALF}%">']
    grid = expand(t)
    ncols = max([len(g) for g in grid] + [t.ncols, 1])
    bd = "" if t.attrs.get("BORDER", "1") == "0" else " bd"
    pos = {it.ri: q for q, it in enumerate(items) if it is not None}
    cover = [[False] * ncols for _ in range(n)]
    cells: list[list[tuple[int, str]]] = [[] for _ in range(n)]
    for q, it in enumerate(items):
        if it is None:
            continue
        e = entries[q]
        row = t.rows[it.ri]
        kidx = {id(nd): k for k, nd in enumerate(row.cells)}
        chg = _changed_cells(e)
        other = (e.y if side == "a" else e.x)
        ocells = other.cells if other is not None and other.kind == "row" else None
        g = grid[it.ri]
        for c, gc in enumerate(g):
            if gc is None or gc.kind != "cell":
                continue
            nd = gc.node
            cs = max(1, min(nd.colspan, ncols - c))
            span_rows = [r for r in range(it.ri, min(it.ri + nd.rowspan, len(grid))) if r in pos]
            qe = pos[max(span_rows)] if span_rows else q
            for qq in range(q, qe + 1):
                for cc in range(c, c + cs):
                    cover[qq][cc] = True
            k = kidx.get(id(nd), -1)
            txt = nd.text.strip()
            cls = side + bd
            if nd.tag == "TH" or row.in_thead:
                cls += " h"
            if NUMCELL.match(txt):
                cls += " n"
            if k in chg:
                cls += " x"
                body = _mark(txt, (ocells[k] if ocells is not None and k < len(ocells) else ""), side)
            else:
                body = _e(txt)
            attrs = (f' colspan="{cs}"' if cs > 1 else "") + (f' rowspan="{qe - q + 1}"' if qe > q else "")
            cells[q].append((c, f'<td class="{cls}"{attrs}{_style(nd)}>{body}</td>'))
    for q in range(n):
        ph = items[q] is None
        c = 0
        while c < ncols:
            if cover[q][c]:
                c += 1
                continue
            c2 = c
            while c2 < ncols and not cover[q][c2]:
                c2 += 1
            cls = f"{side} ph" if ph else f"{side}{bd} fl"
            cells[q].append((c, f'<td class="{cls}"' + (f' colspan="{c2 - c}"' if c2 - c > 1 else "") + "></td>"))
            c = c2
    rows = ["".join(h for _, h in sorted(cs)) for cs in cells]
    cols = [f'<col style="width:{HALF * w:.3f}%">' for w in _col_widths(doc, t, ncols)]
    return rows, cols


def _row_attrs(e: Entry, A: str, B: str) -> str:
    if e.c is None:
        return ' class="eq"'
    c = e.c
    tip = f"{KINDS[c.kind][0]}·{c.note}·{(c.b or c.a).ref}"
    if c.base:
        tip += "·" + BASE_TXT[c.base].format(A=A, B=B)
    return (f' class="d k-{c.kind}{" b-" + c.base if c.base else ""}" data-c="{c.idx}"'
            f' title="{html.escape(tip, quote=True)}"')


def _p_block(e: Entry, A: str, B: str) -> str:
    def half(it: Item | None, other: Item | None, side: str) -> str:
        if it is None:
            return f'<div class="{side} ph"></div>'
        st = _style(it.node, cell=False) if it.node is not None else ""
        body = _mark(it.text, other.text if other is not None and e.c else None, side) if e.c else _e(it.text)
        return f'<div class="{side}"{st}>{body}</div>'
    it = e.x or e.y
    ttl = " ttl" if it.node is not None and it.node.tag == "TITLE" else ""
    a = _row_attrs(e, A, B).replace(' class="', f' class="pp{ttl} ', 1)
    return f'<div{a}>{half(e.x, e.y, "a")}<div class="g"></div>{half(e.y, e.x, "b")}</div>'


def _blocks(entries: list[Entry]) -> list[dict]:
    """줄 목록을 문단 하나·표 하나(양쪽 표가 같은 동안) 덩어리로 묶는다."""
    out: list[dict] = []
    for e in entries:
        it = e.x or e.y
        if it.kind == "p":
            out.append({"kind": "p", "entries": [e], "note": it.note})
            continue
        tx = e.x.table if e.x else None
        ty = e.y.table if e.y else None
        cur = out[-1] if out and out[-1]["kind"] == "t" else None
        if cur and (tx is None or cur["ta"] in (None, tx)) and (ty is None or cur["tb"] in (None, ty)):
            cur["entries"].append(e)
            cur["ta"] = cur["ta"] or tx
            cur["tb"] = cur["tb"] or ty
        else:
            out.append({"kind": "t", "entries": [e], "ta": tx, "tb": ty, "note": it.note})
    return out


def _est(blk: dict) -> int:
    """덩어리 높이 어림(px) - 화면 밖 덩어리는 그리지 않고(content-visibility) 이 높이로 자리만 잡는다."""
    if blk["kind"] == "t":
        return 20 * len(blk["entries"]) + 2
    lines = max(sum(len(ln) // 52 + 1 for ln in it.text.split("\n"))
                for it in (blk["entries"][0].x, blk["entries"][0].y) if it is not None)
    return 21 * lines + 6


def _side_html(d: Diff, A: str, B: str) -> tuple[str, list[str], dict[str, int]]:
    """나란히 보기 본문 + 주석 목록(문서 순서) + 주석별 바뀐 곳 수."""
    blocks = _blocks(d.entries)
    notes = list(dict.fromkeys(b["note"] for b in blocks))
    nidx = {n: i for i, n in enumerate(notes)}
    seen: set[str] = set()
    out: list[str] = []
    for blk in blocks:
        ni = nidx[blk["note"]]
        anchor = ""
        if blk["note"] not in seen:
            seen.add(blk["note"])
            anchor = f' id="sn{ni}"'
        out.append(f'<div class="blk"{anchor} data-note="{ni}" style="contain-intrinsic-size:auto {_est(blk)}px">')
        if blk["kind"] == "p":
            out.append(_p_block(blk["entries"][0], A, B))
        else:
            ents = blk["entries"]
            ra, ca = _side_table(ents, "a", blk["ta"], d.a)
            rb, cb = _side_table(ents, "b", blk["tb"], d.b)
            out.append('<table class="pair"><colgroup>' + "".join(ca) + '<col class="g">' + "".join(cb)
                       + "</colgroup>")
            for e, ha, hb in zip(ents, ra, rb):
                out.append(f'<tr{_row_attrs(e, A, B)}>{ha}<td class="g"></td>{hb}</tr>')
            out.append("</table>")
        out.append("</div>")
    counts: dict[str, int] = {}
    for c in d.changes:
        counts[c.note] = counts.get(c.note, 0) + 1
    return "".join(out), notes, counts


# ── 바뀐 곳만(카드) ───────────────────────────────────

def _row_html(cells: list[str], other: list[str] | None, side: str) -> str:
    tds = []
    for k, c in enumerate(cells):
        o = other[k] if other is not None and k < len(other) else None
        cls = ""
        if other is not None and _norm(c) != _norm(o or ""):
            cls = ' class="chg-a"' if side == "a" else ' class="chg-b"'
        if not c and not cls:
            cls = ' class="empty"'
        tds.append(f"<td{cls}>{_e(c) if c else '&nbsp;'}</td>")
    return "<tr>" + "".join(tds) + "</tr>"


def _cards_html(changes: list[Change], A: str, B: str, side_link: bool) -> str:
    s = summary(changes)
    kinds = [k for k in KINDS if s["by_kind"].get(k)]
    notes = list(dict.fromkeys(c.note for c in changes))
    meaning = {"number": "숫자만 다름", "text": "글이 다름", "added": f"{B}에만 있는 줄", "deleted": f"{A}에만 있는 줄",
               "dash": "빈 칸·'-'·0 끼리만 다름", "space": "띄어쓰기만 다름", "layout": "채운 칸은 같고 칸 수(병합)만 다름"}
    out = [f"<p>바뀐 곳은 모두 <b>{s['total']}</b>건입니다. XML 속성·줄바꿈처럼 화면에 안 보이는 차이는 뺐습니다. "
           f"표 행은 {_e(A)} 칸을 빨강, {_e(B)} 칸을 초록으로, 문단은 지운 글자를 빨강·더한 글자를 초록으로 칠했습니다.</p>",
           "<table class='sum'><tr><th>종류</th><th class='num'>건수</th><th>뜻</th></tr>"]
    for k in kinds:
        out.append(f"<tr><td><i class='sw k-{k}'></i>{KINDS[k][0]}</td><td class='num'>{s['by_kind'][k]}</td>"
                   f"<td>{_e(meaning[k])}</td></tr>")
    out.append("</table>")
    if not changes:
        out.append("<p>보이는 글의 차이가 없습니다.</p>")
        return "".join(out)
    out.append("<h2>주석별</h2><div class='toc'>")
    for i, n in enumerate(notes):
        out.append(f"<div><a href='#n{i}'>{_e(n)}</a> <span class='none'>{s['by_note'][n]}</span></div>")
    out.append("</div><h2>바뀐 곳</h2>")
    seen: set[str] = set()
    for c in changes:
        anchor = ""
        if c.note not in seen:
            anchor = f" id='n{notes.index(c.note)}'"
            seen.add(c.note)
        out.append(f"<section class='chg k-{c.kind}{' b-' + c.base if c.base else ''}' data-c='{c.idx}'{anchor}>"
                   f"<div class='meta'><span class='tag'><i class='sw k-{c.kind}'></i>{KINDS[c.kind][0]}</span>"
                   f"<b>{_e(c.note)}</b> <span class='none'>{_e((c.b or c.a).ref)}</span>")
        if c.base:
            out.append(f"·{_e(BASE_TXT[c.base].format(A=A, B=B))}")
        if side_link:
            out.append(f" <a href='#' class='goside' data-c='{c.idx}'>나란히 보기에서 보기</a>")
        out.append("</div><div class='side'>")
        if c.a and c.b and c.a.kind == c.b.kind == "p":
            xa, xb = _char_diff(c.a.text, c.b.text)
            out.append(f"<div class='who'>{_e(A)}</div><div class='para'>{xa}</div>"
                       f"<div class='who'>{_e(B)}</div><div class='para'>{xb}</div>")
        else:
            for who, it, other, side in ((A, c.a, c.b, "a"), (B, c.b, c.a, "b")):
                out.append(f"<div class='who'>{_e(who)}</div>")
                if it is None:
                    out.append("<div class='none'>(없음)</div>")
                elif it.kind == "p":
                    out.append(f"<div class='para'>{_e(it.text)}</div>")
                else:
                    oc = other.cells if other is not None and other.kind == "row" and c.kind != "layout" else None
                    out.append(f"<div class='scroll'><table class='row'>{_row_html(it.cells, oc, side)}</table></div>")
        out.append("</div></section>")
    return "".join(out)


# ── 화면 틀 ──────────────────────────────────────────

CSS = """
:root{--bg:#f4f4f4;--paper:#fff;--ink:#1f1f1f;--sub:#666;--dim:#8c8c8c;--line:#d0d0d0;--head:#f2f2f2;
--neg:#b42318;--negbg:#fde8e6;--pos:#1a7f37;--posbg:#e6f4ea;--hl:#ffe58a;--ph:#e4e4e4;
--k-number:#d97706;--k-text:#7c3aed;--k-added:#16a34a;--k-deleted:#dc2626;--k-cos:#94a3b8;
--cols:49.25% 1.5% 49.25%;--top:96px;--nav:250px}
@media (prefers-color-scheme:dark){:root{--bg:#161616;--paper:#1f1f1f;--ink:#e8e8e8;--sub:#a0a0a0;--dim:#7a7a7a;
--line:#3a3a3a;--head:#2a2a2a;--neg:#ff8a80;--negbg:#3b1f1d;--pos:#7ee2a0;--posbg:#1c3324;--hl:#5c4a12;--ph:#2c2c2c;
--k-number:#f59e0b;--k-text:#a78bfa;--k-added:#4ade80;--k-deleted:#f87171;--k-cos:#64748b}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.55 "맑은 고딕","Malgun Gothic","Noto Sans KR",sans-serif}
button{font:inherit;color:inherit;background:var(--paper);border:1px solid var(--line);border-radius:4px;padding:2px 10px;cursor:pointer}
button:hover{background:var(--head)}
#top{position:sticky;top:0;z-index:6;background:var(--paper);border-bottom:1px solid var(--line);padding:8px 16px 6px}
#top .r1{display:flex;flex-wrap:wrap;gap:6px 14px;align-items:center}
#top h1{font-size:17px;margin:0}
.files{color:var(--sub);font-size:12.5px}.files b{color:var(--ink)}
.tabs{display:flex;gap:0}.tabs button{border-radius:0}.tabs button:first-child{border-radius:4px 0 0 4px}
.tabs button:last-child{border-radius:0 4px 4px 0;border-left:0}
.tabs button.on{background:var(--ink);color:var(--paper);border-color:var(--ink)}
.navb{display:flex;gap:6px;align-items:center;margin-left:auto}#pos{min-width:64px;text-align:center;color:var(--sub);font-variant-numeric:tabular-nums}
kbd{font:11px monospace;border:1px solid var(--line);border-bottom-width:2px;border-radius:3px;padding:0 4px;color:var(--sub)}
.filters{margin-top:6px;font-size:13px;line-height:1.9}.filters .tot{margin-right:14px}.filters label{margin-right:12px;white-space:nowrap;cursor:pointer}
.sw{display:inline-block;width:10px;height:10px;border-radius:2px;margin:0 4px 0 1px;vertical-align:-1px;background:var(--k-cos)}
.sw.k-number{background:var(--k-number)}.sw.k-text{background:var(--k-text)}.sw.k-added{background:var(--k-added)}.sw.k-deleted{background:var(--k-deleted)}
.none{color:var(--sub)}
body[data-view=side] #cards,body[data-view=cards] #sidev{display:none}
/* 나란히 보기 */
#sidev{display:grid;grid-template-columns:var(--nav) minmax(0,1fr)}
#nav{position:sticky;top:var(--top);height:calc(100vh - var(--top));display:grid;grid-template-columns:18px 1fr;
border-right:1px solid var(--line);background:var(--paper)}
#map{position:relative;background:var(--head);cursor:pointer;border-right:1px solid var(--line);overflow:hidden}
#map .tk{position:absolute;left:2px;right:2px;height:3px;border-radius:1px;background:var(--k-cos)}
#map .tk.k-number{background:var(--k-number)}#map .tk.k-text{background:var(--k-text)}
#map .tk.k-added{background:var(--k-added)}#map .tk.k-deleted{background:var(--k-deleted)}
#vp{position:absolute;left:0;right:0;border:2px solid var(--ink);opacity:.35;pointer-events:none;border-radius:2px}
#notes{overflow:auto;padding:6px 6px 24px;font-size:12.5px}
#notes .hd{color:var(--sub);font-size:11.5px;margin:2px 4px 4px}
#notes a{display:flex;gap:6px;justify-content:space-between;color:var(--ink);text-decoration:none;padding:1px 6px;border-radius:3px}
#notes a:hover{background:var(--head)}#notes a.act{background:var(--head);font-weight:700}
#notes a.zero{color:var(--dim)}#notes .cnt{color:var(--neg);font-variant-numeric:tabular-nums;font-weight:600}
main#doc{min-width:0;background:var(--paper);padding:0 14px 60vh}
.colhead{position:sticky;top:var(--top);z-index:4;display:grid;grid-template-columns:var(--cols);background:var(--paper);
border-bottom:2px solid var(--ink);padding:6px 0 4px;font-size:13px}
.colhead .a,.colhead .b{padding:0 6px}.colhead small{color:var(--sub);font-weight:400}
.blk{margin:0 0 8px;content-visibility:auto;contain-intrinsic-size:auto 200px}
table.pair{width:100%;table-layout:fixed;border-collapse:collapse;font-size:12px;line-height:1.45;font-variant-numeric:tabular-nums}
table.pair td{padding:1px 4px;vertical-align:middle;overflow-wrap:anywhere;word-break:keep-all}
table.pair td.n{white-space:nowrap;overflow:hidden;padding:1px 3px}table.pair td.bd{border:1px solid var(--line)}table.pair td.h{background:var(--head)}
td.g,.pp>.g{padding:0!important;border:0}
.eq td.a,.eq td.b,.pp.eq>.a,.pp.eq>.b{color:var(--dim)}
.d td.g,.d>.g{background:var(--k-cos)}
.d.k-number td.g,.d.k-number>.g{background:var(--k-number)}.d.k-text td.g,.d.k-text>.g{background:var(--k-text)}
.d.k-added td.g,.d.k-added>.g{background:var(--k-added)}.d.k-deleted td.g,.d.k-deleted>.g{background:var(--k-deleted)}
td.x.a{background:var(--negbg);color:var(--neg)}td.x.b{background:var(--posbg);color:var(--pos);font-weight:600}
.k-deleted td.a:not(.ph),.pp.k-deleted>.a{background:var(--negbg)}
.k-added td.b:not(.ph),.pp.k-added>.b{background:var(--posbg)}
.ph{background:repeating-linear-gradient(135deg,var(--ph) 0 5px,transparent 5px 10px)!important}
.pp{display:grid;grid-template-columns:var(--cols);font-size:13px;line-height:1.6}
.pp>.a,.pp>.b{padding:2px 6px;word-break:keep-all;overflow-wrap:anywhere;min-height:1.6em}
.pp.ttl>.a,.pp.ttl>.b{font-weight:700;font-size:15px;padding-top:10px}
del{background:var(--negbg);color:var(--neg);text-decoration:line-through}
ins{background:var(--posbg);color:var(--pos);text-decoration:underline;font-weight:600}
td.x del,td.x ins{background:transparent}
.d.off td.g,.d.off>.g{background:transparent}
.d.off td.a,.d.off td.b,.d.off>.a,.d.off>.b{background:none;color:var(--dim);font-weight:inherit}
.d.off del,.d.off ins{background:none;color:inherit;text-decoration:none;font-weight:inherit}
.cur td.g,.cur>.g{box-shadow:inset 0 0 0 2px var(--ink)}
.flash td.a,.flash td.b,.flash>.a,.flash>.b{animation:fl 1.4s ease-out}
@keyframes fl{0%,35%{background-color:var(--hl)}}
/* 바뀐 곳만 */
#cards{max-width:1180px;margin:16px auto;background:var(--paper);padding:20px 32px 48px;border:1px solid var(--line)}
h2{font-size:16px;margin:26px 0 8px;border-bottom:1px solid var(--ink);padding-bottom:4px}
table.sum{border-collapse:collapse;margin:6px 0 10px}table.sum td,table.sum th{border-top:1px solid var(--line);
border-bottom:1px solid var(--line);padding:4px 12px;text-align:left}table.sum th{background:var(--head)}
.num{text-align:right;font-variant-numeric:tabular-nums}
.toc{columns:3;column-gap:24px;margin:6px 0}.toc a,.meta a{color:var(--ink)}.toc a{text-decoration:none}.toc a:hover{text-decoration:underline}
.chg{border-top:1px solid var(--line);padding:10px 0;scroll-margin-top:calc(var(--top) + 8px)}.chg.hide{display:none}
.meta{color:var(--sub);font-size:12.5px;margin-bottom:6px}.meta b{color:var(--ink)}.meta a.goside{font-size:12px;margin-left:6px}
.tag{display:inline-block;border:1px solid var(--line);padding:0 6px;margin-right:6px;font-size:12px}
.side{display:grid;grid-template-columns:72px 1fr;gap:4px 10px;align-items:start}
.side .who{color:var(--sub);font-size:12.5px;padding-top:3px}
.scroll{overflow-x:auto}table.row{border-collapse:collapse;font-variant-numeric:tabular-nums}
table.row td{border:1px solid var(--line);padding:2px 8px;white-space:nowrap}
td.empty{border-style:dashed;min-width:24px}td.chg-a{background:var(--negbg);color:var(--neg)}td.chg-b{background:var(--posbg);color:var(--pos);font-weight:600}
.para{white-space:normal;word-break:keep-all}
.chg.cur{box-shadow:-4px 0 0 var(--ink)}
@media (max-width:900px){#sidev{grid-template-columns:minmax(0,1fr)}#nav{display:none}main#doc{padding:0 4px 60vh}
#cards{padding:16px}.toc{columns:1}.navb{margin-left:0}}
"""

JS = r"""
(function(){
var R=document.documentElement,B=document.body,top=document.getElementById('top');
var side=document.getElementById('doc'),cards=document.getElementById('cards'),colhead=document.querySelector('.colhead');
var map=document.getElementById('map'),vp=document.getElementById('vp'),posEl=document.getElementById('pos');
var inputs=[].slice.call(document.querySelectorAll('input[data-k],input[data-b]'));
var links=[].slice.call(document.querySelectorAll('#notes a'));
var anchors=links.map(function(a){return document.getElementById('sn'+a.dataset.ni)});
var targets=[],cur=-1,saved={side:0,cards:0},OFF=90;
var io=('IntersectionObserver' in window)?new IntersectionObserver(function(es){es.forEach(function(e){
  if(e.isIntersecting){io.unobserve(e.target);fitIn(e.target)}})},{rootMargin:'1200px 0px'}):null;
function fit(){if(!side||B.dataset.view!=='side')return;if(!io){fitIn(side);return}
  side.querySelectorAll('.blk').forEach(function(b){if(b.querySelector('td.n')){io.unobserve(b);io.observe(b)}})}
function fitIn(root){var cs=root.querySelectorAll('td.n'),big=[];
  cs.forEach(function(c){if(c.dataset.fs===undefined)c.dataset.fs=c.style.fontSize;else c.style.fontSize=c.dataset.fs});
  cs.forEach(function(c){if(c.scrollWidth>c.clientWidth+1)big.push([c,parseFloat(getComputedStyle(c).fontSize)*c.clientWidth/c.scrollWidth])});
  big.forEach(function(b){b[0].style.fontSize=Math.max(7,b[1]*.97)+'px'})}
function setTop(){R.style.setProperty('--top',top.offsetHeight+'px')}
function hdr(){return top.offsetHeight+(B.dataset.view==='side'&&colhead?colhead.offsetHeight:0)}
function isOff(el,off){for(var i=0;i<el.classList.length;i++)if(off[el.classList[i]])return true;return false}
function kindOf(el){for(var i=0;i<el.classList.length;i++)if(el.classList[i].indexOf('k-')===0)return el.classList[i];return ''}
function apply(){
  var off={};inputs.forEach(function(b){if(!b.checked)off[b.dataset.k?'k-'+b.dataset.k:'b-'+b.dataset.b]=1});
  if(cards)cards.querySelectorAll('.chg').forEach(function(s){s.classList.toggle('hide',isOff(s,off))});
  var cnt={};
  if(side)side.querySelectorAll('.d').forEach(function(el){var o=isOff(el,off);el.classList.toggle('off',o);
    if(!o){var n=el.closest('.blk').dataset.note;(cnt[n]=cnt[n]||{})[el.dataset.c]=1}});
  links.forEach(function(a){var c=cnt[a.dataset.ni],k=c?Object.keys(c).length:0;
    a.querySelector('.cnt').textContent=k||'';a.classList.toggle('zero',!k)});
  collect();drawMap();
}
function collect(){
  var list=B.dataset.view==='cards'?(cards?cards.querySelectorAll('.chg:not(.hide)'):[]):(side?side.querySelectorAll('.d:not(.off)'):[]);
  targets=[].slice.call(list);setCur(-1);
}
function setCur(i){
  if(cur>=0&&targets[cur])targets[cur].classList.remove('cur');
  cur=i;if(i>=0&&targets[i])targets[i].classList.add('cur');
  posEl.textContent=(i>=0?i+1:'-')+' / '+targets.length;
}
function jump(i){
  if(i<0||i>=targets.length)return;var t=targets[i];setCur(i);
  var y=function(){return Math.max(0,t.getBoundingClientRect().top+scrollY-hdr()-OFF)};scrollTo(0,y());
  var fix=2;(function again(){requestAnimationFrame(function(){var z=y();if(Math.abs(z-scrollY)>2)scrollTo(0,z);if(--fix>0)again()})})();
  t.classList.remove('flash');void t.offsetWidth;t.classList.add('flash');
}
function go(dir){
  if(!targets.length)return;var line=hdr()+OFF,i=-1,k;
  if(dir>0){for(k=0;k<targets.length;k++)if(targets[k].getBoundingClientRect().top>line+2){i=k;break}if(i<0)i=targets.length-1}
  else{for(k=targets.length-1;k>=0;k--)if(targets[k].getBoundingClientRect().top<line-2){i=k;break}if(i<0)i=0}
  jump(i);
}
function drawMap(){
  if(!map)return;map.querySelectorAll('.tk').forEach(function(t){t.remove()});
  if(B.dataset.view!=='side')return;
  var H=R.scrollHeight,mh=map.clientHeight,f=document.createDocumentFragment();
  targets.forEach(function(t,i){var d=document.createElement('div');d.className='tk '+kindOf(t);
    d.style.top=Math.min(mh-3,(t.getBoundingClientRect().top+scrollY)/H*mh)+'px';d.dataset.i=i;
    d.title=t.getAttribute('title')||'';f.appendChild(d)});
  map.appendChild(f);updVp();
}
function updVp(){if(!map||B.dataset.view!=='side')return;var H=R.scrollHeight,mh=map.clientHeight;
  vp.style.top=(scrollY/H*mh)+'px';vp.style.height=Math.max(10,innerHeight/H*mh)+'px'}
function activeNote(){
  if(B.dataset.view!=='side'||!anchors.length)return;var line=hdr()+6,a=0;
  for(var i=0;i<anchors.length;i++){if(anchors[i]&&anchors[i].getBoundingClientRect().top<=line)a=i;else if(anchors[i])break}
  links.forEach(function(l,i){l.classList.toggle('act',i===a)});
  var box=document.getElementById('notes'),l=links[a];
  if(l&&(l.offsetTop<box.scrollTop||l.offsetTop>box.scrollTop+box.clientHeight-20))box.scrollTop=l.offsetTop-box.clientHeight/3;
}
function show(v){
  saved[B.dataset.view]=scrollY;B.dataset.view=v;
  document.querySelectorAll('.tabs button').forEach(function(b){b.classList.toggle('on',b.dataset.view===v)});
  setTop();fit();collect();scrollTo(0,saved[v]||0);drawMap();activeNote();
}
inputs.forEach(function(b){b.addEventListener('change',apply)});
var cosm=document.getElementById('cosm');
if(cosm)cosm.addEventListener('change',function(e){
  document.querySelectorAll('input[data-cosm]').forEach(function(b){b.checked=!e.target.checked});apply()});
document.querySelectorAll('.tabs button').forEach(function(b){b.addEventListener('click',function(){show(b.dataset.view)})});
document.getElementById('next').addEventListener('click',function(){go(1)});
document.getElementById('prev').addEventListener('click',function(){go(-1)});
links.forEach(function(a){a.addEventListener('click',function(e){e.preventDefault();var t=anchors[links.indexOf(a)];
  if(t)scrollTo(0,t.getBoundingClientRect().top+scrollY-hdr()+1)})});
if(map)map.addEventListener('click',function(e){
  if(e.target.dataset&&e.target.dataset.i){jump(+e.target.dataset.i);return}
  var r=map.getBoundingClientRect();scrollTo(0,(e.clientY-r.top)/r.height*R.scrollHeight-innerHeight/2)});
document.querySelectorAll('a.goside').forEach(function(a){a.addEventListener('click',function(e){
  e.preventDefault();show('side');var t=side.querySelector('.d[data-c="'+a.dataset.c+'"]');
  if(t){var i=targets.indexOf(t);if(i<0){targets.push(t);i=targets.length-1}jump(i)}})});
addEventListener('keydown',function(e){
  if(e.ctrlKey||e.metaKey||e.altKey||(e.target.closest&&e.target.closest('input,textarea,select')))return;
  if(e.code==='KeyN'||e.code==='KeyJ'){go(1);e.preventDefault()}
  else if(e.code==='KeyP'||e.code==='KeyK'){go(-1);e.preventDefault()}});
var tick=false,mt;addEventListener('scroll',function(){clearTimeout(mt);mt=setTimeout(drawMap,300);if(tick)return;tick=true;
  requestAnimationFrame(function(){tick=false;updVp();activeNote()})},{passive:true});
var rz;addEventListener('resize',function(){clearTimeout(rz);rz=setTimeout(function(){setTop();fit();drawMap()},150)});
addEventListener('load',function(){setTop();fit();drawMap()});
setTop();fit();apply();show(B.dataset.view);
})();
"""


def render_html(changes: list[Change] | Diff, labels: tuple[str, str, str | None],
                files: tuple[str, str, str | None], view: str = "side") -> str:
    """보고서 HTML 한 장. changes 에 Diff 를 주면 '나란히 보기' 화면도 만든다(바뀐 곳 목록만 주면 카드 화면만)."""
    d = changes if isinstance(changes, Diff) else None
    changes = d.changes if d is not None else changes
    A, B, C = labels
    s = summary(changes)
    kinds = [k for k in KINDS if s["by_kind"].get(k)]
    base_counts = {b: sum(1 for c in changes if c.base == b) for b in ("a", "b", "none", "both")}
    base_counts = {b: n for b, n in base_counts.items() if n}
    if d is None:
        view = "cards"
    out = ["<!doctype html><html lang='ko'><head><meta charset='utf-8'>",
           "<meta name='viewport' content='width=device-width,initial-scale=1'>",
           f"<title>DSD 비교 - {_e(A)}·{_e(B)}</title><style>{CSS}</style></head>",
           f"<body data-view='{view}'><header id='top'><div class='r1'><h1>DSD 비교</h1>",
           f"<span class='files'><b>{_e(A)}</b> ({_e(files[0])}) → <b>{_e(B)}</b> ({_e(files[1])})"
           + (f", 기준 <b>{_e(C)}</b> ({_e(files[2])})" if C else "") + "</span>"]
    if d is not None:
        out.append("<span class='tabs'><button data-view='side'>나란히 보기</button>"
                   "<button data-view='cards'>바뀐 곳만</button></span>")
    out.append("<span class='navb'><button id='prev' title='이전 차이 (p)'>이전 <kbd>p</kbd></button>"
               "<span id='pos'></span><button id='next' title='다음 차이 (n)'>다음 <kbd>n</kbd></button></span></div>")
    out.append(f"<div class='filters'><span class='tot'>바뀐 곳 <b>{s['total']}</b>건</span>"
               "<label><input type='checkbox' id='cosm'> 표시만 바뀐 것 숨기기(공백·'-'·병합)</label>"
               + "".join(f"<label><input type='checkbox' checked data-k='{k}'"
                         f"{' data-cosm' if KINDS[k][1] else ''}><i class='sw k-{k}'></i>{KINDS[k][0]} "
                         f"{s['by_kind'][k]}</label>" for k in kinds)
               + (("<br><span class='none'>기준 " + _e(C) + ":</span> " + "".join(
                   f"<label><input type='checkbox' checked data-b='{b}'> {_e(BASE_SHORT[b].format(A=A, B=B))} {n}</label>"
                   for b, n in base_counts.items())) if C else "")
               + "</div></header>")
    if d is not None:
        body, notes, counts = _side_html(d, A, B)
        out.append("<div id='sidev'><nav id='nav'><div id='map' title='바뀐 곳 위치 - 누르면 그 자리로'><div id='vp'></div></div>"
                   "<div id='notes'><div class='hd'>주석·바뀐 곳 수</div>")
        for i, n in enumerate(notes):
            k = counts.get(n, 0)
            out.append(f"<a href='#sn{i}' data-ni='{i}'{'' if k else ' class=zero'}><span>{_e(n)}</span>"
                       f"<span class='cnt'>{k or ''}</span></a>")
        out.append("</div></nav><main id='doc'>"
                   f"<div class='colhead'><div class='a'><b>{_e(A)}</b> <small>{_e(files[0])}</small></div><div></div>"
                   f"<div class='b'><b>{_e(B)}</b> <small>{_e(files[1])}</small></div></div>")
        out.append(body if d.entries else "<p>문서가 비어 있습니다.</p>")
        out.append("</main></div>")
    out.append("<div id='cards'>" + _cards_html(changes, A, B, d is not None) + "</div>")
    out.append(f"<script>{JS}</script></body></html>")
    return "".join(out)
