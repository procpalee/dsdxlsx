"""정산표 → DSD 다시 만들기(build) - 뼈대 DSD + 정산표 + 건별 레시피로 매번 같은 DSD.

세 층으로 나눈다. 뼈대 DSD = 구조(행 · 표 · 문구), 정산표 = 숫자, 레시피(JSON) = 둘을 잇는 규칙.
DSD 를 손으로 고치지 않고 셋 중 하나를 고친 뒤 다시 build 한다.

순서: fill(행 이름 짝짓기, recon 설정) → 내보낸 엑셀에 레시피 적용
(labels → tables → cells → subtotals → paragraphs → keep → unit_suffix) → import → text → 저장.
fill 이 못 맞추는 다열 · 그룹 · 중복 이름 표를 레시피 tables 가 정산표 칸으로 덮어쓴다.

레시피 키(모두 선택):
  reference   직전 최종 DSD(.dsd 또는 export 한 .xlsx) - keep · unit_suffix 의 원천. 레시피 파일 기준 상대경로 가능
  aliases     {정규화 이름: 정규화 이름} - 행 이름 짝짓기 별칭
  labels      [{sheet, table, row_label, src}]              DSD 행 이름을 정산표 칸 글로
  tables      [{sheet, table, match: label|order, ...}]     표 덮어쓰기(아래)
  cells       [{sheet, table, row_label | row, col, src, template}] 한 칸을 정산표 칸 글로('{D12*100:.2f}' · '{E46+E47}' · '{-F124}')
  subtotals   [{sheet, table, row_label, cols, prev}]       바로 위 prev 줄 합
  paragraphs  [{sheet, para, src, mode: numbers|sentence|line, contains, set}] - line: contains 로 고른 줄의 숫자만 바꿈(set {줄 숫자 순번: 정산표 글 숫자 순번 | 칸 틀})
  keep        [{sheet, table, skip_header, cols, num_cols}] 같은 행 ID · 이름이면 reference 값 유지
  unit_suffix true 면 reference 같은 표 · 행 · 열 서식의 '주' · '%' 를 글로 붙임
  text        [{start, end, replace: [[옛, 새]], protect: [낱말]}]  contents.xml 구간 글 바꾸기
  output      (xlwork 리본이 읽음) 출력 이름 틀 '{dir}/회사_재무제표_v{version}.dsd' - build 자체는 -o 를 쓴다

tables 항목:
  공통  sheet(DSD export 시트) · table(t###) · src(정산표 시트) · rows('34-53' · [157, 158] · '7-11,54')
        cols {DSD 열: 정산표 열} · formats {DSD 열: rate2 | pct2(×100) | raw | {prefix_by_row: {정산표 행: 'USD'}}}
  label  dsd_label(['B']) · src_label(['B']) · skip_header(1) · range(['r002', 'r009'])
         total_pattern + total_cols - 구분 합계 줄(현금흐름표 Ⅰ~Ⅶ)은 다른 열 대응
         keep_currency - DSD 칸이 'USD 1,800,000' 꼴이면 그 통화를 붙여 씀 / none_as_zero - 빈 칸은 0
  order  first(r### | 행 이름 글귀 '(당기초)') - 이 행부터 rows 순서대로(rows 의 null = 그 DSD 행 건너뜀)
"""
from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string as CI

from .dsd_io import read_dsd, write_dsd
from .parse import parse_contents

ROW_ID = re.compile(r"t\d+\.r\d+")
NUM = r"\d[\d,]*(?:\.\d+)?"
SUFFIX = re.compile(r'"(주|%)"')


@dataclass
class BuildReport:
    filled: int = 0
    unfilled: int = 0
    steps: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def load_recipe(path: str | None) -> dict:
    if not path:
        return {}
    with open(path, encoding="utf-8") as f:
        recipe = json.load(f)
    base = Path(path).resolve().parent
    ref = recipe.get("reference")
    if ref and not os.path.isabs(ref):
        recipe["reference"] = str(base / ref)
    return recipe


# ── 공통 ─────────────────────────────────────────────

def _norm_fn(aliases: dict):
    def norm(s):
        s = re.sub(r"\(\*\d?(,\d)*\)", "", str(s or ""))
        s = re.sub(r"[\s　\xa0/:·]", "", s).upper()
        return aliases.get(s, s)
    return norm


def num(v):
    """정산표 계산값 → DSD 칸 값: 정수에 가깝거나 1,000 이상이면 정수로."""
    if isinstance(v, bool) or v is None:
        return v
    if isinstance(v, (int, float)):
        return int(round(v)) if abs(v - round(v)) < 1e-6 or abs(v) >= 1000 else v
    return v


def _rows_spec(spec) -> list[int]:
    if isinstance(spec, int):
        return [spec]
    if isinstance(spec, str):
        out = []
        for part in spec.split(","):
            part = part.strip()
            if "-" in part:
                a, b = part.split("-")
                out += list(range(int(a), int(b) + 1))
            elif part:
                out.append(int(part))
        return out
    out = []
    for x in spec:
        out += _rows_spec(x)
    return out


def _rows_of(ws, tid):
    return [r for r in ws.iter_rows() if isinstance(r[0].value, str) and r[0].value.startswith(tid + ".")]


def _label(cells, cols):
    for c in cols:
        v = cells[CI(c) - 1].value
        if v not in (None, "") and str(v).strip():
            return v
    return None


def _fmt(v, how, src_row):
    if how == "raw":
        return v
    if how == "rate2":
        return f"{v:.2f}" if isinstance(v, (int, float)) else v
    if how == "pct2":                                   # 정산표 비율(0.0459) → DSD % 숫자(4.59)
        return round(v * 100, 2) if isinstance(v, (int, float)) and not isinstance(v, bool) else v
    if isinstance(how, dict) and "prefix_by_row" in how:
        cur = how["prefix_by_row"].get(str(src_row))
        return f"{cur} {v:,.0f}" if cur else num(v)
    return num(v)


_TPL = re.compile(r"\{(-?[A-Z]+\d+(?:[+-][A-Z]+\d+)*)(?:\*([\d.]+))?(?::([^}]*))?\}")
_TERM = re.compile(r"([+-]?)([A-Z]+\d+)")


def _template(tpl: str, js):
    """'{D64:,.0f}\\nUSD {D47:,.0f}' · '{D12*100:.2f}' · '{E46+E47}' · '{-F124}' - 칸 값이 숫자가 아니면 None
    (더하기 · 빼기 식은 빈 칸을 0 으로)."""
    ok = True

    def rep(m):
        nonlocal ok
        terms = _TERM.findall(m.group(1))
        if len(terms) == 1 and not terms[0][0]:
            v = js[terms[0][1]].value
        else:
            v = 0
            for sign, ref in terms:
                x = js[ref].value
                if x is None:
                    continue
                if not isinstance(x, (int, float)) or isinstance(x, bool):
                    ok = False
                    return ""
                v += -x if sign == "-" else x
        if not isinstance(v, (int, float)) or isinstance(v, bool):
            ok = False
            return ""
        if m.group(2):
            v = v * float(m.group(2))
        vals.append(v)
        return format(v, m.group(3) or "")
    vals: list = []
    out = _TPL.sub(rep, tpl)
    if not ok:
        return None
    whole = _TPL.fullmatch(tpl)
    if whole and not whole.group(3):                   # 식 하나뿐(서식 없음) → 숫자로 넣어 DSD 숫자 서식을 따름
        return num(vals[0])
    return out


def _src_cell(jw, ref: str):
    sheet, cell = ref.rsplit("!", 1)
    return jw[sheet][cell]


# ── 레시피 단계 ─────────────────────────────────────

class _Ctx:
    def __init__(self, dw, jw, recipe, rep: BuildReport):
        self.dw, self.jw, self.recipe, self.rep = dw, jw, recipe, rep
        self.norm = _norm_fn(recipe.get("aliases") or {})
        self._ref = None

    def log(self, msg):
        self.rep.warnings.append(msg)

    def ref(self):
        """reference 엑셀(직전 최종 DSD export) - .dsd 면 임시로 내보낸다."""
        if self._ref is None:
            path = self.recipe.get("reference")
            if not path:
                raise ValueError("레시피 reference(직전 최종 DSD)가 없음 - keep · unit_suffix 에 필요")
            if path.lower().endswith(".dsd"):
                from .export_xlsx import export_xlsx
                tmp = os.path.join(tempfile.mkdtemp(prefix="dsdbuild_"), "reference.xlsx")
                with open(path, "rb") as f:
                    data = f.read()
                export_xlsx(parse_contents(read_dsd(path).contents), data, path, tmp)
                path = tmp
            self._ref = openpyxl.load_workbook(path)
        return self._ref


def _step_labels(c: _Ctx):
    for it in c.recipe.get("labels", []):
        want = c.norm(it["row_label"])
        for r in _rows_of(c.dw[it["sheet"]], it["table"]):
            if c.norm(r[1].value) == want:
                r[1].value = _src_cell(c.jw, it["src"]).value


def _slice(drows, tid, rng, norm):
    """range = [처음, 끝] - 'r002' 꼴은 행 ID, 그 밖은 행 이름에 들어 있는 글(뼈대에 행을 넣어도 안 밀림)."""
    def pos(key, start=0):
        for i in range(start, len(drows)):
            r = drows[i]
            if re.fullmatch(r"r\d+", key) and r[0].value == f"{tid}.{key}":
                return i
            if not re.fullmatch(r"r\d+", key) and norm(key) in norm(r[1].value):
                return i
        raise ValueError(f"{tid} 범위 '{key}' 행 없음")
    a = pos(rng[0])
    return drows[a:pos(rng[1], a) + 1]


def _table_label(c: _Ctx, it):
    ws, js = c.dw[it["sheet"]], c.jw[it["src"]]
    tid, norm = it["table"], c.norm
    src_label = it.get("src_label", ["B"])
    jrows: dict[str, list[int]] = {}
    for i in _rows_spec(it["rows"]):
        lab = _label([js.cell(i, k) for k in range(1, 15)], src_label)
        if lab is not None:
            jrows.setdefault(norm(lab), []).append(i)
    drows = _rows_of(ws, tid)[it.get("skip_header", 2):]
    if it.get("range"):
        drows = _slice(drows, tid, it["range"], norm)
    tot = re.compile(it["total_pattern"]) if it.get("total_pattern") else None
    formats = it.get("formats", {})
    used: dict[str, int] = {}
    n_set = 0
    for r in drows:
        lab = _label(r, it.get("dsd_label", ["B"]))
        if lab is None:
            continue
        k = norm(lab)
        cand = jrows.get(k) or next((v for kk, v in jrows.items() if kk.startswith(k) or k.startswith(kk)), None)
        if not cand:
            c.log(f"{tid} '{str(lab)[:30]}' 정산표에 없음")
            continue
        n = used.get(k, 0)
        used[k] = n + 1
        if n >= len(cand):
            c.log(f"{tid} '{str(lab)[:30]}' {n + 1}번째 짝 없음")
            continue
        ji = cand[n]
        pairs = it.get("cols", {})
        if tot and tot.match(str(lab).strip()):
            pairs = it.get("total_cols", pairs)
        cur = None
        if it.get("keep_currency"):
            cur = next((str(x.value).split()[0] for x in r[2:] if isinstance(x.value, str)
                        and re.match(r"[A-Z]{3} ", x.value)), None)
        for dc, jc in pairs.items():
            raw = js[f"{jc}{ji}"].value
            if it.get("none_as_zero") and raw is None:
                raw = 0
            if cur:
                v = f"{cur} {raw:,.0f}" if raw else "0"
            else:
                v = _fmt(raw, formats.get(dc), ji)
            if v == "-":
                v = 0
            cell = r[CI(dc) - 1]
            if cell.value != v:
                cell.value = v
                n_set += 1
    return n_set


def _table_order(c: _Ctx, it):
    ws, js, tid = c.dw[it["sheet"]], c.jw[it["src"]], it["table"]
    rs = _rows_of(ws, tid)
    first = it["first"]
    if re.fullmatch(r"r\d+", first):
        idx = [i for i, r in enumerate(rs) if r[0].value == f"{tid}.{first}"][0]
    else:                                       # 행 이름 글귀('(당기초)') - 밀기로 앞 행 수가 바뀌어도 안 밀림
        idx = next(i for i, r in enumerate(rs) if c.norm(first) in c.norm(r[1].value))
    formats = it.get("formats", {})
    seq = [None if x is None else x for x in it["rows"]] if isinstance(it["rows"], list) and None in it["rows"] \
        else _rows_spec(it["rows"])
    for k, ji in enumerate(seq):
        if ji is None:                          # null = 이 DSD 행은 건너뜀(구분 머리글 · 소계 줄)
            continue
        if idx + k >= len(rs):
            c.log(f"{tid} 행 순서 대응: 정산표 {ji}행에 맞는 DSD 행 없음")
            break
        r = rs[idx + k]
        for dc, jc in it["cols"].items():
            if isinstance(r[CI(dc) - 1], MergedCell):     # 병합된 칸(구분 합계 줄 등)은 건너뜀
                continue
            r[CI(dc) - 1].value = _fmt(js[f"{jc}{ji}"].value, formats.get(dc), ji)
    return len(seq)


def _step_tables(c: _Ctx):
    for it in c.recipe.get("tables", []):
        try:
            (_table_order if it.get("match") == "order" else _table_label)(c, it)
        except (KeyError, IndexError, ValueError) as e:
            c.log(f"표 대응 {it.get('sheet')}/{it.get('table')} 건너뜀: {e}")


def _step_cells(c: _Ctx):
    for it in c.recipe.get("cells", []):
        js = c.jw[it["src"]]
        want = c.norm(it.get("row_label"))
        rid = f"{it['table']}.{it['row']}" if it.get("row") else None     # 행 이름이 빈 줄(병합된 차입처 등)은 행 ID
        hit = False
        for r in _rows_of(c.dw[it["sheet"]], it["table"]):
            if (r[0].value == rid) if rid else \
                    (want in c.norm(r[1].value) if it.get("contains") else c.norm(r[1].value) == want):
                v = _template(it["template"], js)
                if v is not None:
                    r[CI(it["col"]) - 1].value = v
                hit = True
        if not hit:
            c.log(f"칸 {it['sheet']}/{it['table']} '{it.get('row') or it.get('row_label')}' 행 없음")


def _step_subtotals(c: _Ctx):
    for it in c.recipe.get("subtotals", []):
        rs = _rows_of(c.dw[it["sheet"]], it["table"])
        want, prev = c.norm(it.get("row_label", "소계")), it.get("prev", 3)
        for i, r in enumerate(rs):
            if c.norm(r[1].value) == want:
                for col in it["cols"]:
                    vals = [rs[j][CI(col) - 1].value for j in range(i - prev, i)]
                    r[CI(col) - 1].value = sum(num(v) or 0 for v in vals if isinstance(v, (int, float)))


def _line_numbers(c: _Ctx, it, ln: str) -> str:
    """mode line: contains 로 고른 줄의 숫자만 바꾼다. set {DSD 줄 숫자 순번: 정산표 글 숫자 순번(int) | 칸 틀(str)},
    set 이 없으면 정산표 글(src = '시트!칸') 숫자를 순서대로(개수가 같을 때). 칸 틀의 시트는 src 의 시트."""
    dn = re.findall(NUM, ln)
    sheet = it["src"].rsplit("!", 1)[0]
    jt = str(_src_cell(c.jw, it["src"]).value) if "!" in it["src"] else ""
    jn = re.findall(NUM, jt)
    pick = it.get("set") or ({str(i): i for i in range(len(dn))} if len(dn) == len(jn) else None)
    if pick is None:
        c.log(f"{it['para']} 숫자 개수 다름 {dn} / {jn}")
        return ln
    new = {}
    for k, src in pick.items():
        if isinstance(src, int):
            new[int(k)] = jn[src] if src < len(jn) else None
        else:
            v = _template(src, c.jw[sheet])
            new[int(k)] = f"{v:,.0f}" if isinstance(v, (int, float)) else v
    if any(v is None for v in new.values()):
        c.log(f"{it['para']} 정산표 숫자 없음 {pick}")
        return ln
    pos = iter(range(len(dn)))
    return re.sub(NUM, lambda m: (lambda i: new.get(i, m.group(0)))(next(pos)), ln)


def _step_paragraphs(c: _Ctx):
    for it in c.recipe.get("paragraphs", []):
        jt = str(_src_cell(c.jw, it["src"]).value) if "!" in it["src"] else ""
        for r in c.dw[it["sheet"]].iter_rows():
            if r[0].value != it["para"]:
                continue
            lines = str(r[1].value).split("\n")
            for k, ln in enumerate(lines):
                if it.get("mode") == "line":
                    if all(m in ln for m in it.get("contains", [])):
                        lines[k] = _line_numbers(c, it, ln)
                elif it.get("mode") == "sentence":
                    marks = it.get("contains", [])
                    if marks and all(m in ln for m in marks):
                        lines[k] = ln[:ln.index(marks[0])] + jt
                elif ln.strip()[:12] and re.sub(NUM, "#", ln.strip())[:15] == re.sub(NUM, "#", jt.strip())[:15]:
                    dn, jn = re.findall(NUM, ln), re.findall(NUM, jt)
                    if len(dn) == len(jn):
                        nums = iter(jn)
                        lines[k] = re.sub(NUM, lambda m: next(nums), ln)
                    else:
                        c.log(f"{it['para']} 숫자 개수 다름 {dn} / {jn}")
            new = "\n".join(lines)
            if new == r[1].value:
                c.log(f"{it['para']} 문단 바뀐 것 없음")
            r[1].value = new


def _step_keep(c: _Ctx):
    for it in c.recipe.get("keep", []):
        ref = c.ref()
        tid = it["table"]
        arows = {r[0]: r for r in ref[it["sheet"]].iter_rows(values_only=True)
                 if r[0] and str(r[0]).startswith(tid + ".")}
        num_cols = set(it.get("num_cols", it["cols"]))
        for r in _rows_of(c.dw[it["sheet"]], tid)[it.get("skip_header", 2):]:
            a = arows.get(r[0].value)
            if a and c.norm(a[1]) == c.norm(r[1].value):
                for col in it["cols"]:
                    v = a[CI(col) - 1]
                    r[CI(col) - 1].value = num(v) if col in num_cols else v


def _step_unit_suffix(c: _Ctx):
    if not c.recipe.get("unit_suffix"):
        return
    afmt: dict = {}
    for ws in c.ref().worksheets:
        for r in ws.iter_rows():
            if isinstance(r[0].value, str) and ROW_ID.match(r[0].value):
                afmt.setdefault((ws.title, r[0].value.split(".")[0], c.norm(r[1].value)),
                                {x.column: x.number_format for x in r[2:]
                                 if isinstance(x.value, (int, float)) and SUFFIX.search(x.number_format)})
    fixed = 0
    for ws in c.dw.worksheets:
        for r in ws.iter_rows():
            if not (isinstance(r[0].value, str) and ROW_ID.match(r[0].value)):
                continue
            fm = afmt.get((ws.title, r[0].value.split(".")[0], c.norm(r[1].value)), {})
            for x in r[2:]:
                if x.column in fm and isinstance(x.value, (int, float)) and not isinstance(x.value, bool) and x.value != 0:
                    f = fm[x.column]
                    m = re.search(r"\.(0+)", f)
                    t = f"{abs(x.value):,.{len(m.group(1)) if m else 0}f}{SUFFIX.search(f).group(1)}"
                    x.value = f"({t})" if x.value < 0 else t
                    fixed += 1
    c.rep.steps.append(f"단위 기호 {fixed}칸")


def _restore_paras(orig_path: str, wb) -> int:
    """레시피가 건드리지 않은 문단은 글자 그대로(앞뒤 공백 · 줄바꿈만 달라진 것 되돌림)."""
    a = openpyxl.load_workbook(orig_path)
    n = 0
    for s in a.sheetnames:
        if s == "_map" or s not in wb.sheetnames:
            continue
        orig = {r[0].value: r for r in a[s].iter_rows() if isinstance(r[0].value, str) and r[0].value.startswith("p")}
        for r in wb[s].iter_rows():
            if r[0].value in orig:
                for c0, c1 in zip(orig[r[0].value], r):
                    if c0.value != c1.value and (c0.value or "").strip() == (c1.value or "").strip():
                        c1.value = c0.value
                        n += 1
    return n


def apply_text(contents: bytes, rules: list[dict], rep: BuildReport) -> bytes:
    if not rules:
        return contents
    x = contents.decode("utf-8")
    for rule in rules:
        a = x.find(rule["start"])
        b = x.find(rule["end"], a) if a >= 0 else -1
        if a < 0 or b < 0:
            rep.warnings.append(f"글 바꾸기 구간 못 찾음: {rule['start'][:20]} ~ {rule['end'][:20]}")
            continue
        seg = x[a:b]
        keep = {f"\0{i}\0": w for i, w in enumerate(rule.get("protect", []))}
        for k, w in keep.items():
            seg = seg.replace(w, k)
        for old, new in rule.get("replace", []):
            seg = seg.replace(old, new)
        for k, w in keep.items():
            seg = seg.replace(k, w)
        x = x[:a] + seg + x[b:]
    rep.steps.append(f"글 바꾸기 {len(rules)}구간")
    return x.encode("utf-8")


# ── 빌드 ────────────────────────────────────────────

RECIPE_STEPS = (_step_labels, _step_tables, _step_cells, _step_subtotals,
                _step_paragraphs, _step_keep, _step_unit_suffix)


def build(skeleton: str, workpaper: str, out: str, *, recipe: dict, fill_conf: dict, side_sheets,
          workdir: str | None = None) -> BuildReport:
    """뼈대 + 정산표 + 레시피 → out DSD. fill_conf · side_sheets 는 cli._recon_inputs 결과."""
    from . import fill as F
    from . import recon as R
    from . import splice
    from .export_xlsx import export_xlsx
    from .import_xlsx import import_xlsx

    rep = BuildReport()
    work = workdir or tempfile.mkdtemp(prefix="dsdbuild_")
    os.makedirs(work, exist_ok=True)

    package = read_dsd(skeleton)
    doc = parse_contents(package.contents)
    ops, frep = F.fill_ops(doc, side_sheets, R.ReconConfig.from_dict(fill_conf))
    filled = splice.apply(package.contents, ops)
    rep.filled, rep.unfilled = frep.count("채움"), frep.count("못 채움")
    rep.steps.append(f"채움 {rep.filled}칸 · 못 채움 {rep.unfilled}칸")
    f_dsd, f_xlsx, r_xlsx = (os.path.join(work, n) for n in ("1_채움.dsd", "1_채움.xlsx", "2_레시피.xlsx"))
    write_dsd(f_dsd, package, filled)
    with open(f_dsd, "rb") as f:
        data = f.read()
    export_xlsx(parse_contents(filled), data, f_dsd, f_xlsx)

    shutil.copy2(f_xlsx, r_xlsx)
    dw = openpyxl.load_workbook(r_xlsx)
    jw = openpyxl.load_workbook(workpaper, data_only=True)
    ctx = _Ctx(dw, jw, recipe, rep)
    for step in RECIPE_STEPS:
        step(ctx)
    n = _restore_paras(f_xlsx, dw)
    if n:
        rep.steps.append(f"문단 공백 되돌림 {n}칸")
    dw.save(r_xlsx)

    pkg2, contents, irep = import_xlsx(r_xlsx)
    rep.steps.append(f"레시피 반영 {len(irep.cell_edits)}칸")
    rep.warnings += [f"import: {w}" for w in irep.warnings]
    contents = apply_text(contents, recipe.get("text", []), rep)
    parse_contents(contents)
    write_dsd(out, pkg2, contents)
    return rep


# ── 빌드 밖 수정 감지 ───────────────────────────────

def last_build_path(out: str) -> str:
    p = Path(out)
    return str(p.parent / "_build" / (p.stem + ".last.dsd"))


def outside_edits(out: str) -> list[str] | None:
    """out 이 지난 build 결과와 다르면 다른 곳 목록(사람이 DSD 를 직접 고친 흔적). 비교 기준이 없으면 None."""
    last = last_build_path(out)
    if not (os.path.exists(out) and os.path.exists(last)):
        return None
    if read_dsd(out).contents == read_dsd(last).contents:
        return []
    from .verify import diff_dsd
    lines = diff_dsd(last, out)
    return [ln for ln in lines if ln != "구조·텍스트 차이 없음"] or ["(바이트 차이 - 표 · 문단 글은 같음)"]


def save_last_build(out: str) -> None:
    last = last_build_path(out)
    os.makedirs(os.path.dirname(last), exist_ok=True)
    shutil.copy2(out, last)
