"""두 DSD 비교 보고서(diff) - 문단 · 표 행 단위로 맞춰 바뀐 곳만 HTML 한 장으로.

용도: 우리 DSD ↔ 감사인이 고쳐 보낸 DSD, 직전 판 ↔ 이번 판처럼 같은 문서의 두 판을 비교한다.
XML 속성 · 줄바꿈처럼 화면에 안 보이는 차이는 무시하고 '보이는 글'만 비교한다.

- 문서를 (주석 이름, 문단 | 표 행) 목록으로 펴고 순서를 맞춘다(difflib).
- 바뀐 곳마다 종류를 붙인다: 숫자 · 글 · 공백 · '-'(빈 칸 · '-' · 0 끼리) · 병합(채운 칸은 같고 칸 수만 다름) · 추가 · 삭제.
- 기준 DSD(--base, 예: 직전 분기 최종)를 주면 바뀐 곳마다 기준에 어느 쪽 모양이 있었는지 표시한다
  (A 가 기준과 같음 = 이어받은 것 · B 가 기준과 같음 = A 쪽에서 바뀐 것 · 둘 다 없음 = 이번에 새로 생긴 줄).
"""
from __future__ import annotations

import difflib
import html
import re
from dataclasses import dataclass, field

from .model import Document

NUM = re.compile(r"\(?-?\d[\d,]*(?:\.\d+)?\)?%?")
NOTE_HEAD = re.compile(r"^\s*(\d{1,2})\.\s*([^\d\s].{0,38})$")
STATEMENTS = ("재무상태표", "포괄손익계산서", "손익계산서", "자본변동표", "현금흐름표")
DASHES = {"", "-", "0", "–"}

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
    ref: str           # 문단 · 행 ID
    cells: list[str]

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


def _norm(t: str) -> str:
    return re.sub(r"\s+", " ", t.replace("\n", " ")).strip()


def _shape(t: str) -> str:
    return NUM.sub("#", _norm(t))


def _period(t: str) -> str:
    """기준(직전 분 · 반기) 대조용 - 기간 낱말 차이(반기 ↔ 분기)는 같은 것으로."""
    return re.sub(r"반(\s?)기", lambda m: "분" + m.group(1) + "기", t)


def flatten(doc: Document) -> list[Item]:
    """문서 순서대로 문단 · 표 행을 편다. 주석 이름은 'N. 제목' 으로 시작하는 문단에서 잡는다."""
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
                out.append(Item(note, "p", b.item.id, [t]))
            else:
                for r in b.item.rows:
                    cells = [c.text.strip() for c in r.cells]
                    head = re.sub(r"^(분기|반기|요약)+", "", re.sub(r"\s", "", cells[0])) if cells else ""
                    if head in STATEMENTS:                 # 재무제표 제목표('재 무 상 태 표')
                        note = head
                    if any(cells):
                        out.append(Item(note, "row", r.id, cells))
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
    """바뀐 덩어리 안에서 비슷한 줄끼리 짝짓기(순서 유지). 짝 없는 줄은 추가 · 삭제."""
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


def compare(a: Document, b: Document, base: Document | None = None) -> list[Change]:
    ia, ib = flatten(a), flatten(b)
    sm = difflib.SequenceMatcher(None, [_norm(x.text) for x in ia], [_norm(x.text) for x in ib], autojunk=False)
    out: list[Change] = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue
        for x, y in _pair(ia[i1:i2], ib[j1:j2]):
            if x and y:
                out.append(Change(_classify(x, y), y.note, x, y))
            elif x:
                out.append(Change("deleted", x.note, x, None))
            else:
                out.append(Change("added", y.note, None, y))
    if base is not None:
        items = flatten(base)                  # 같은 주석 안에서만 찾음(짧은 '소 계' 줄이 다른 주석과 겹치지 않게)
        texts = {(_period(i.note), _period(_norm(i.text))) for i in items}
        shapes = {(_period(i.note), _period(_shape(i.text))) for i in items}
        blob = " ".join(_period(_norm(i.text)) for i in items if i.kind == "p")    # 문단은 나뉜 자리가 판마다 달라
        sblob = " ".join(_period(_shape(i.text)) for i in items if i.kind == "p")  # 글 안에 들어 있는지로 봄
        for c in out:
            def has(it: Item | None) -> int:   # 2 = 같은 글, 1 = 같은 모양(숫자만 다름), 0 = 없음
                if it is None:
                    return 0
                k, tn, ts = _period(it.note), _period(_norm(it.text)), _period(_shape(it.text))
                if it.kind == "p":
                    return 2 if tn in blob else (1 if ts in sblob else 0)
                return 2 if (k, tn) in texts else (1 if (k, ts) in shapes else 0)
            ha, hb = has(c.a), has(c.b)
            if ha > hb:
                c.base = "a"
            elif hb > ha:
                c.base = "b"
            else:
                c.base = "both" if ha else "none"
    return out


def summary(changes: list[Change]) -> dict:
    by_kind: dict[str, int] = {}
    by_note: dict[str, int] = {}
    for c in changes:
        by_kind[c.kind] = by_kind.get(c.kind, 0) + 1
        by_note[c.note] = by_note.get(c.note, 0) + 1
    return {"total": len(changes), "by_kind": by_kind, "by_note": by_note}


# ── HTML ─────────────────────────────────────────────

def _e(t) -> str:
    return html.escape(str(t)).replace("\n", "<br>")


def _char_diff(a: str, b: str) -> tuple[str, str]:
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    xa, xb = [], []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            xa.append(_e(a[i1:i2]))
            xb.append(_e(b[j1:j2]))
            continue
        if i2 > i1:
            xa.append(f'<del>{_e(a[i1:i2])}</del>')
        if j2 > j1:
            xb.append(f'<ins>{_e(b[j1:j2])}</ins>')
    return "".join(xa), "".join(xb)


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


BASE_TXT = {
    "a": "기준과 같은 쪽: {A} (기준을 이어받음)",
    "b": "기준과 같은 쪽: {B} ({A}에서 달라짐)",
    "both": "기준에 두 모양 모두 있음",
    "none": "기준에 없음 (이번에 생긴 줄)",
}
BASE_SHORT = {"a": "{A}가 이어받음", "b": "{A}에서 달라짐", "both": "기준에 둘 다", "none": "새로 생긴 줄"}

CSS = """
:root{--bg:#f4f4f4;--paper:#fff;--ink:#1f1f1f;--sub:#666;--line:#d0d0d0;--head:#f2f2f2;
--neg:#b42318;--negbg:#fde8e6;--pos:#1a7f37;--posbg:#e6f4ea;--hl:#fff4c2}
@media (prefers-color-scheme:dark){:root{--bg:#161616;--paper:#202020;--ink:#e8e8e8;--sub:#a0a0a0;
--line:#3a3a3a;--head:#2a2a2a;--neg:#ff8a80;--negbg:#3b1f1d;--pos:#7ee2a0;--posbg:#1c3324;--hl:#4a4020}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);
font:14px/1.6 "맑은 고딕","Malgun Gothic","Noto Sans KR",sans-serif}
.doc{max-width:1180px;margin:24px auto;background:var(--paper);padding:28px 32px 48px;border:1px solid var(--line)}
h1{font-size:20px;margin:0 0 4px}.sub{color:var(--sub);margin:0 0 14px;border-bottom:3px solid var(--ink);padding-bottom:10px}
h2{font-size:16px;margin:28px 0 8px;border-bottom:1px solid var(--ink);padding-bottom:4px}
table.sum{border-collapse:collapse;margin:6px 0 10px}table.sum td,table.sum th{border-top:1px solid var(--line);
border-bottom:1px solid var(--line);padding:4px 12px;text-align:left}table.sum th{background:var(--head)}
.num{text-align:right;font-variant-numeric:tabular-nums}
.filters{position:sticky;top:0;background:var(--paper);padding:8px 0;border-bottom:1px solid var(--line);z-index:2}
.filters label{margin-right:14px;white-space:nowrap;cursor:pointer}
.toc{columns:3;column-gap:24px;margin:6px 0}.toc a{color:var(--ink);text-decoration:none}.toc a:hover{text-decoration:underline}
.chg{border-top:1px solid var(--line);padding:10px 0}.chg.hide{display:none}
.meta{color:var(--sub);font-size:12.5px;margin-bottom:6px}.meta b{color:var(--ink)}
.tag{display:inline-block;border:1px solid var(--line);padding:0 6px;margin-right:6px;font-size:12px}
.side{display:grid;grid-template-columns:72px 1fr;gap:4px 10px;align-items:start}
.side .who{color:var(--sub);font-size:12.5px;padding-top:3px}
.scroll{overflow-x:auto}table.row{border-collapse:collapse;font-variant-numeric:tabular-nums}
table.row td{border:1px solid var(--line);padding:2px 8px;white-space:nowrap}
td.empty{border-style:dashed;min-width:24px}td.chg-a{background:var(--negbg);color:var(--neg)}td.chg-b{background:var(--posbg);color:var(--pos);font-weight:600}
.para{white-space:normal;word-break:keep-all}
del{background:var(--negbg);color:var(--neg)}ins{background:var(--posbg);color:var(--pos);text-decoration:underline;font-weight:600}
.none{color:var(--sub)}
@media (max-width:640px){.doc{padding:16px}.toc{columns:1}}
"""

JS = """
function apply(){var off={};document.querySelectorAll('input[data-k],input[data-b]').forEach(function(b){
if(!b.checked)off[(b.dataset.k?'k-'+b.dataset.k:'b-'+b.dataset.b)]=1});
document.querySelectorAll('.chg').forEach(function(s){var h=false;s.classList.forEach(function(c){if(off[c])h=true});
s.classList.toggle('hide',h)})}
document.querySelectorAll('input[data-k],input[data-b]').forEach(function(b){b.addEventListener('change',apply)});
document.getElementById('cosm').addEventListener('change',function(e){
document.querySelectorAll('input[data-cosm]').forEach(function(b){b.checked=!e.target.checked});apply()});
"""


def render_html(changes: list[Change], labels: tuple[str, str, str | None], files: tuple[str, str, str | None]) -> str:
    A, B, C = labels
    s = summary(changes)
    kinds = [k for k in KINDS if s["by_kind"].get(k)]
    notes = list(dict.fromkeys(c.note for c in changes))
    base_counts = {b: sum(1 for c in changes if c.base == b) for b in ("a", "b", "none", "both")}
    base_counts = {b: n for b, n in base_counts.items() if n}
    out = ["<!doctype html><html lang='ko'><head><meta charset='utf-8'>",
           "<meta name='viewport' content='width=device-width,initial-scale=1'>",
           f"<title>DSD 비교 - {_e(A)} · {_e(B)}</title><style>{CSS}</style></head><body><div class='doc'>",
           "<h1>DSD 비교</h1>",
           f"<p class='sub'>{_e(A)} <span class='none'>({_e(files[0])})</span> → {_e(B)} "
           f"<span class='none'>({_e(files[1])})</span>"
           + (f" · 기준 {_e(C)} <span class='none'>({_e(files[2])})</span>" if C else "") + "</p>",
           f"<p>바뀐 곳은 모두 <b>{s['total']}</b>건입니다. XML 속성 · 줄바꿈처럼 화면에 안 보이는 차이는 뺐습니다. "
           f"표 행은 {_e(A)} 칸을 빨강, {_e(B)} 칸을 초록으로, 문단은 지운 글자를 빨강 · 더한 글자를 초록으로 칠했습니다.</p>",
           "<table class='sum'><tr><th>종류</th><th class='num'>건수</th><th>뜻</th></tr>"]
    meaning = {"number": "숫자만 다름", "text": "글이 다름", "added": f"{B}에만 있는 줄", "deleted": f"{A}에만 있는 줄",
               "dash": "빈 칸 · '-' · 0 끼리만 다름", "space": "띄어쓰기만 다름", "layout": "채운 칸은 같고 칸 수(병합)만 다름"}
    for k in kinds:
        out.append(f"<tr><td>{KINDS[k][0]}</td><td class='num'>{s['by_kind'][k]}</td><td>{_e(meaning[k])}</td></tr>")
    out.append("</table>")
    out.append("<div class='filters'><label><input type='checkbox' id='cosm'> 표시만 바뀐 것 숨기기(공백 · '-' · 병합)</label>"
               + "".join(f"<label><input type='checkbox' checked data-k='{k}'"
                         f"{' data-cosm' if KINDS[k][1] else ''}> {KINDS[k][0]} {s['by_kind'][k]}</label>" for k in kinds)
               + (("<br><span class='none'>기준 " + _e(C) + ":</span> " + "".join(
                   f"<label><input type='checkbox' checked data-b='{b}'> {_e(BASE_SHORT[b].format(A=A, B=B))} {n}</label>"
                   for b, n in base_counts.items())) if C else "")
               + "</div>")
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
        out.append(f"<section class='chg k-{c.kind}{' b-' + c.base if c.base else ''}'{anchor}><div class='meta'><span class='tag'>{KINDS[c.kind][0]}</span>"
                   f"<b>{_e(c.note)}</b> <span class='none'>{_e((c.b or c.a).ref)}</span>")
        if c.base:
            out.append(f" · {_e(BASE_TXT[c.base].format(A=A, B=B))}")
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
    if not changes:
        out.append("<p>보이는 글의 차이가 없습니다.</p>")
    out.append(f"</div><script>{JS}</script></body></html>")
    return "".join(out)
