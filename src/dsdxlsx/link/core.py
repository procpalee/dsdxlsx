"""뼈대 DSD 의 행 그대로 정산표를 만든다.

재무상태표 · 손익계산서 · 자본변동표 · 현금흐름표 · 주석을 DSD 행 순서대로 시트에 옮기고,
TB 의 정산 과목을 DSD 행에 짝지어 수식으로 끌어온다. 이렇게 만든 정산표는 fill · build 가 그대로 읽는다.

짝짓기는 전기 금액으로 확인한다: 뼈대 DSD 의 전기 칸과 TB 전기 잔액이 같아야 한 짝이다.
회사 이름과 금액은 코드에 넣지 않는다.
"""
from __future__ import annotations

import re
from decimal import Decimal

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

from ..guard import InputError
from .sheetkit import HDR_ROW, mark_cells, num_space, write_header, write_section_bar, write_table

ACCT = "accounting"
LINK_SHEET = "_link"

CONTRA = ("대손충당금", "감가상각누계액", "손상차손누계액", "정부보조금", "국고보조금", "현재가치할인차금", "사채할인발행차금",
          "퇴직연금운용자산", "국민연금전환금", "상각누계액", "전환권조정", "재평가차손누계액")
ALIAS = {"법인세등": ("법인세비용", "법인세등"), "법인세비용": ("법인세비용", "법인세등"),
         "복리후생비": ("복리후생비",), "보증금": ("보증금", "임차보증금"), "임차보증금": ("임차보증금", "보증금"),
         "접대비": ("접대비", "기업업무추진비"), "기업업무추진비": ("기업업무추진비", "접대비"),
         "잡손실": ("잡손실", "기타의영업외비용"), "잡이익": ("잡이익", "기타의영업외수익"),
         "미처분이익잉여금": ("미처분이익잉여금", "이월이익잉여금"), "미처리결손금": ("미처리결손금", "이월결손금", "미처분이익잉여금", "이월이익잉여금"),
         "이월결손금": ("이월결손금", "미처리결손금", "미처분이익잉여금", "이월이익잉여금"),
         "이익잉여금(결손금)": ("이익잉여금", "미처분이익잉여금", "이월이익잉여금"),
         "결손금": ("결손금", "미처리결손금", "이월결손금", "미처분이익잉여금", "이월이익잉여금")}
RETAINED = re.compile(r"^(미처분이익잉여금|미처리결손금|이월결손금|이익잉여금\(결손금\)|결손금|이월이익잉여금|이익잉여금$)")
ROMAN = "IVXLⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ"
SP_RE = re.compile(r"[\s\xa0\u3000]+")


def squeeze(s) -> str:
    return SP_RE.sub("", str(s or ""))


def tidy(s) -> str:
    """DSD 글 - 칸 안 공백을 하나로, 앞뒤 공백 제거."""
    return SP_RE.sub(" ", str(s or "")).strip()


def key_of(label: str) -> str:
    s = squeeze(label)
    s = re.sub(r"\(주\s*석?[\d,및\s~\-]*\)$", "", s)
    s = re.sub(r"\(주석[^)]*\)", "", s)
    s = re.sub(r"\(\*\d*\)", "", s)
    s = re.sub(r"^[%s]+\.?(?=[가-힣(])" % ROMAN, "", s)
    s = re.sub(r"^\(\d+\)", "", s)
    s = re.sub(r"^\d+\.", "", s)
    return s


T_NAMES = {"유동자산", "비유동자산", "유동부채", "비유동부채", "자본금", "자본잉여금", "자본조정", "기타포괄손익누계액", "이익잉여금", "결손금", "이익잉여금(결손금)"}
M_NAMES = {"당좌자산", "재고자산", "투자자산", "유형자산", "무형자산", "기타비유동자산"}
IS_T_NAMES = {"매출총이익", "매출총손실", "영업이익", "영업손실", "영업이익(손실)", "법인세비용차감전순이익", "법인세비용차감전순손실", "법인세차감전순이익",
              "법인세차감전순손실", "법인세비용차감전순이익(손실)", "당기순이익", "당기순손실", "당기순이익(손실)"}


def level_of(label: str, ctx: str = "") -> str:
    s = squeeze(label)
    if re.match(r"^[%s]+\.?(?=[가-힣(])" % ROMAN, s):
        return "T"
    if re.match(r"^\(\d+\)", s) or (ctx == "CF" and re.match(r"^\d+\.", s)):
        return "M"
    if ctx == "CF":
        return "L"
    k = key_of(label)
    if k in T_NAMES or k in IS_T_NAMES:
        return "T"
    if k in M_NAMES:
        return "M"
    return "L"


# ── DSD 표 읽기 ───────────────────────────────────────

def main_table(sec):
    tabs = [b for b in sec["blocks"] if b["type"] == "table" and any(c.get("period") for c in b["columns"])]
    return max(tabs, key=lambda b: len(b["rows"])) if tabs else None


def row_vals(row):
    return {c["col"]: c for c in row["cells"]}


def num_of(c):
    v = c.get("value")
    return Decimal(str(v)) if v is not None else None


class Line:
    __slots__ = ("label", "key", "lv", "prior", "sign", "tb", "kind", "kids", "note", "flip", "ref", "cur_only", "neg", "extra", "parent")

    def __init__(self, label, prior, ctx=""):
        self.label, self.key, self.lv, self.prior = tidy(label), key_of(label), level_of(label, ctx), prior
        self.sign, self.tb, self.kind, self.kids, self.note, self.flip, self.ref = 1, None, "", [], "", {}, None
        self.cur_only, self.neg, self.extra, self.parent = False, False, [], None


def lines_of(table, ctx="") -> tuple[list, str, str]:
    cols = table["columns"]
    cur_cols, pri_cols = [c["col"] for c in cols if c.get("period") == "당기"], [c["col"] for c in cols if c.get("period") == "전기"]
    h_c = next((tidy(c["header"]) for c in cols if c.get("period") == "당기"), "당기")
    h_p = next((tidy(c["header"]) for c in cols if c.get("period") == "전기"), "전기")
    out = []
    for r in table["rows"]:
        if r.get("head"):
            continue
        vals = row_vals(r)
        p = next((num_of(vals[c]) for c in pri_cols if c in vals and num_of(vals[c]) is not None), None)
        dash = any(tidy(vals[c].get("text")) == "-" for c in pri_cols if c in vals)
        line = Line(r["label"], p if p is not None else (Decimal(0) if dash else None), ctx)
        if line.label:
            out.append(line)
    return out, h_c, h_p


# ── TB 짝짓기 ─────────────────────────────────────────

class Matcher:
    def __init__(self, tb):
        self.settles: dict = {}
        self.accts: list = []
        for r in tb:
            s = squeeze(r["settle"])
            if s:
                self.settles.setdefault(s, Decimal(0))
                self.settles[s] += r["pri"]
                self.accts.append({"code": r["code"], "name": squeeze(r["name"]), "settle": s, "pri": r["pri"]})
        self.used: dict = {}
        self.acct_used: dict = {}             # 코드 → 행(계정 단위로 잡힌 것)
        self.subs: dict = {}
        for r in tb:
            sub = squeeze(r.get("sub") or "")
            if sub and squeeze(r["settle"]) != sub:
                self.subs.setdefault(sub, []).append(squeeze(r["settle"]))

    def take_sub(self, line):
        """정산 과목(공시용 묶음)이 DSD 행 이름과 같으면 그 아래 계정을 모두 짝."""
        names = [n for n in self.subs.get(line.key, []) if n not in self.used and n in self.settles]
        if not names:
            return []
        for n in names:
            self.used[n] = line
        return names

    def amount(self, key):
        if key in self.settles:
            return self.settles[key]
        return sum((a["pri"] for a in self.accts if a["code"] in self.acct_used and a["name"] == key), Decimal(0))

    def take(self, key, line):
        self.used[key] = line.key_full if hasattr(line, "key_full") else line.key
        return key

    def find(self, line, prev_key, exact=False):
        k = line.key
        cands = []
        if k.startswith(CONTRA):
            base = k
            cands += ["%s(%s)" % (base, prev_key)] if prev_key else []
        cands += [k]
        for c in cands:
            if c in self.settles and c not in self.used and k.startswith(CONTRA):
                self.used[c] = line
                return c
        if k not in self.settles and line.prior is not None and not k.startswith(CONTRA):
            hit = [a for a in self.accts if a["name"] == k and a["code"] not in self.acct_used and a["pri"] * line.sign == line.prior and a["pri"] != 0]
            if len(hit) == 1:                     # 정산 과목 하나에 DSD 행이 둘인 경우 - 계정 단위로 짝
                a = hit[0]
                self.acct_used[a["code"]] = line
                self.settles[a["settle"]] -= a["pri"]
                if self.settles[a["settle"]] == 0:
                    self.used[a["settle"]] = line
                return k
        cands += list(ALIAS.get(k, ()))
        for c in cands:
            if c in self.settles and c not in self.used:
                if exact and line.prior is not None and self.settles[c] * line.sign != line.prior:
                    continue
                self.used[c] = line
                return c
        return None

    def by_amount(self, line, sign):
        """이름으로 못 찾은 행 - 전기 금액이 같은 미사용 정산 과목."""
        if line.prior is None or line.prior == 0:
            return None
        target = line.prior * sign
        hits = [s for s, a in self.settles.items() if s not in self.used and a == target]
        if len(hits) == 1:
            self.used[hits[0]] = line
            return hits[0]
        return None

    def leftovers(self):
        return {s: a for s, a in self.settles.items() if s not in self.used}


def absorb(lines, m: Matcher, sign_of):  # noqa: C901
    """금액이 어긋난 행에 미사용 정산 과목(최대 3개 조합)을 묶어 차이를 없앤다."""
    from itertools import combinations
    left = m.leftovers()
    pool = [s for s, a in left.items() if a != 0]
    done = []
    for ln in lines:
        if ln.kind != "leaf" or ln.prior is None or (not ln.tb and ln.prior == 0):
            continue
        cur = sum(m.amount(x) for x in ln.tb) * sign_of(ln) if ln.tb else Decimal(0)
        gap = ln.prior - cur
        if gap == 0:
            continue
        for n in (1, 2, 3, 4):
            hit = next((c for c in combinations(pool, n) if sum(left[x] for x in c) * sign_of(ln) == gap), None)
            if hit:
                for x in hit:
                    pool.remove(x)
                    m.used[x] = ln
                ln.tb.extend(hit)
                done.append((ln.key, hit))
                break
    return done


# ── 합계 구조 ─────────────────────────────────────────

def sum_ok(parent, kids, flip=None):
    if parent.prior is None or any(k.prior is None for k in kids):
        return None
    s = sum(k.prior * ((flip or {}).get(id(k), 1)) for k in kids)
    return s == parent.prior


def set_sum(p, kids, signed=("차감", "유출", "기말", "감소", "결손")):
    """합계행 p = kids 합. 전기로 검증되지 않으면 차감 낱말 있는 자식의 부호를 뒤집어 본다. 그래도 아니면 입력칸."""
    if not kids:
        return False
    p.kids = kids
    p.kind = "sum"
    if sum_ok(p, kids) in (True, None):
        return True
    flips = [k for k in kids if any(w in k.key for w in signed)]
    from itertools import product
    for bits in product((1, -1), repeat=min(len(flips), 6)):
        f = {id(k): b for k, b in zip(flips, bits)}
        if sum_ok(p, kids, f):
            p.flip = f
            return True
    p.kind, p.kids = "input", []
    p.note = "전기 합계와 구성이 맞지 않아 입력칸으로 둠"
    return False


IFRS_T = {"유동자산", "비유동자산", "유동부채", "비유동부채"}


def build_bs(lines, m):
    if not any(x.key == "당좌자산" for x in lines):     # 분류 없는 재무상태표(K-IFRS 등) - 이름만으로 합계 줄을 만들지 않음
        for x in lines:
            if x.lv in ("T", "M") and (x.key in M_NAMES or (x.key in T_NAMES and x.key not in IFRS_T)) and not re.match(r"^[%s(]" % ROMAN, squeeze(x.label)):
                x.lv = "L"
    sect, prev_key, tops = None, None, {"자산": [], "부채": [], "자본": []}
    loose = {"자산": [], "부채": [], "자본": []}       # 소계 줄 아래에 있지 않은 과목(분류 없는 자본 등) - 총계에 직접 들어감
    totals = {}
    cur_t = cur_m = None
    for ln in lines:
        k = ln.key
        if k in ("자산", "부채", "자본"):
            sect, ln.kind, cur_t, cur_m = k, "head", None, None
            continue
        if "총계" in k:
            ln.kind, cur_t, cur_m = "total", None, None
            totals[k] = ln
            continue
        ln.sign = 1 if sect == "자산" else -1
        if ln.lv == "T":
            ln.kind = "T"
            cur_t, cur_m = ln, None
            tops[sect].append(ln)
            ln.kids = []
            continue
        if ln.lv == "M":
            ln.kind = "M"
            cur_m = ln
            ln.kids = []
            if cur_t:
                cur_t.kids.append(ln)
            continue
        ln.kind = "leaf"
        if cur_m:
            cur_m.kids.append(ln)
        elif cur_t:
            cur_t.kids.append(ln)
        elif sect in loose and not k.startswith("지배기업"):
            loose[sect].append(ln)
        ln.ref = prev_key                  # 앞 자산 이름(차감 계정 짝)
        ln.tb = []
        if not k.startswith(CONTRA):
            prev_key = k
    for s_ in tops:
        tops[s_] = tops[s_] + loose[s_]
    for exact in (True, False):         # 이름 + 금액이 같은 줄 먼저(유동 · 비유동에 같은 이름)
        for ln in lines:
            if ln.kind == "leaf" and not ln.tb:
                t = m.find(ln, ln.ref, exact)
                ln.tb = [t] if t else []
    for ln in lines:
        if ln.kind == "leaf":
            ln.ref = None
            if not ln.tb:
                ln.tb = m.take_sub(ln)
    for ln in lines:                    # 금액으로 다시
        if ln.kind == "leaf" and not ln.tb:
            t = m.by_amount(ln, ln.sign)
            ln.tb = [t] if t else []
    absorb(lines, m, lambda l: l.sign)
    for ln in lines:
        if ln.kind in ("M", "T") and not ln.kids:   # 합계가 아니라 잎인 줄(이름만 합계형)
            ln.kind = "leaf"
            t = m.find(ln, None) or m.by_amount(ln, ln.sign)
            ln.tb = [t] if t else []
    for ln in lines:
        if ln.kind == "leaf" and ln.tb and ln.prior:
            tot = sum(m.amount(x) for x in ln.tb) * ln.sign
            if tot == -ln.prior:                    # 표시 부호 반대(결손금을 양수로 쓰는 표 등)
                ln.sign *= -1
        if ln.kind == "leaf" and RETAINED.match(ln.key) and ln.prior and ln.prior > 0 and "결손" in ln.key:
            ln.sign = 1
    for ln in lines:
        if ln.kind == "M":
            set_sum(ln, ln.kids)
        elif ln.kind == "T":
            mids = [x for x in ln.kids if x.kind == "M"]
            if mids and all(x.kind == "M" for x in ln.kids):
                set_sum(ln, ln.kids)
            else:
                set_sum(ln, ln.kids)
    sects = {"자산총계": "자산", "부채총계": "부채", "자본총계": "자본"}
    for k, ln in totals.items():
        for name, s in sects.items():
            if k == name:
                set_sum(ln, tops[s])
        if k.startswith("부채") and "자본" in k:
            set_sum(ln, [totals[x] for x in totals if x in ("부채총계", "자본총계")])
    for ln in lines:
        if ln.kind == "leaf" and RETAINED.match(ln.key):
            ln.note = "ni"
        if ln.key.startswith("지배기업") and ln.kind == "leaf" and not ln.tb:   # 연결: 지배기업 소유주지분 = 자본 구성 합
            if not set_sum(ln, tops["자본"]):
                ln.kids, ln.kind, ln.flip = tops["자본"], "sum", {}
                ln.note = "전기 DSD 소계가 구성 합과 다름"
    return lines


REV_RE = re.compile(r"^(매출액|수익\(매출액\)|영업수익|영업외수익|금융수익|기타수익|기타영업외수익)")
COST_RE = re.compile(r"^(매출원가|영업비용|판매비와관리비|영업외비용|금융비용|기타비용|기타영업외비용|법인세등|법인세비용$)")
RESULT_RE = re.compile(r"(이익|손실|손익)(\(손실\)|\(이익\))?$")
MEMO_RE = re.compile(r"^(기초|당기|기말|타계정|재고자산평가)")
INV_RE = re.compile(r"(상품|제품|재공품|원재료|부재료|저장품|반제품|미착품|재고)")


def build_is(lines, m, bs_lines):
    sign, cur, cur_m, last = 1, None, None, None
    blocks = []
    for ln in lines:
        k = ln.key
        top = ln.lv == "T" or bool(REV_RE.match(k) or COST_RE.match(k))
        if top:
            ln.kind, ln.kids, cur_m = "T", [], None
            cur = ln
            if RESULT_RE.search(k) and not REV_RE.match(k) and not COST_RE.match(k):
                ln.kind, cur = "result", None
            else:
                sign = -1 if REV_RE.match(k) else 1
                ln.sign = sign
            blocks.append(ln)
            last = None
            continue
        if cur is None:
            ln.kind = "plain"
            continue
        ln.sign = sign
        if ln.lv == "M":
            ln.kind, ln.kids = "M", []
            cur.kids.append(ln)
            cur_m, last = ln, ln
            continue
        if (MEMO_RE.match(k) and INV_RE.search(k)) or (MEMO_RE.match(k) and ("제조원가" in k or "매입" in k)):
            ln.kind, ln.parent = "memo", last
            continue
        ln.kind = "leaf"
        ln.tb = []
        if cur_m is not None:
            cur_m.kids.append(ln)
        else:
            cur.kids.append(ln)
            last = ln
    for exact in (True, False):
        for ln in lines:
            if ln.kind == "leaf" and not ln.tb:
                t = m.find(ln, None, exact)
                ln.tb = [t] if t else []
    for ln in lines:
        if ln.kind == "leaf" and not ln.tb:
            ln.tb = m.take_sub(ln)
    memo_of = {}
    for ln in lines:
        if ln.kind == "memo" and ln.parent is not None:
            memo_of.setdefault(id(ln.parent), []).append(ln)
    for ln in lines:
        if ln.kind == "M" and id(ln) in memo_of:   # 재고 구성을 품은 줄 = 잎 + TB
            ln.extra, ln.kids = ln.kids, []
            ln.kind = "leaf"
            t = m.find(ln, None) or m.by_amount(ln, ln.sign)
            ln.tb = [t] if t else []
    for ln in lines:
        if ln.kind == "leaf" and not ln.tb:
            t = m.by_amount(ln, ln.sign)
            ln.tb = [t] if t else []
    absorb(lines, m, lambda l: l.sign)
    for ln in lines:
        if ln.kind == "leaf" and ln.tb and ln.prior:
            if sum(m.amount(x) for x in ln.tb) * ln.sign == -ln.prior:
                ln.sign *= -1
    for ln in lines:
        if ln.kind == "M":
            set_sum(ln, ln.kids)
    running, block_seq = None, []
    for ln in blocks:
        if ln.kind == "T":
            if ln.kids:
                set_sum(ln, ln.kids)
            else:
                ln.kind = "leaf"
                t = m.find(ln, None) or m.by_amount(ln, ln.sign)
                ln.tb = [t] if t else []
            block_seq.append((ln, ln.sign))
        elif ln.kind == "result":
            kids = ([running] if running is not None else []) + [b for b, _s in block_seq]
            flip = {id(b): (1 if s == -1 else -1) for b, s in block_seq}
            if running is not None and running.neg:
                flip[id(running)] = -1
            ln.kids, ln.flip, ln.kind = kids, flip, "sum"
            if sum_ok(ln, kids, flip) is False:
                tmp = ln.prior
                ln.prior = -tmp if tmp is not None else None
                ok = sum_ok(ln, kids, flip)
                ln.prior = tmp
                if ok and "손실" in ln.key:
                    ln.neg = True
                else:
                    ln.kind, ln.kids, ln.flip = "input", [], {}
                    ln.note = "전기 합계와 구성이 맞지 않아 입력칸으로 둠"
            running, block_seq = ln, []
    inv = {re.sub(r"\s", "", x.key): x for x in bs_lines if x.kind == "leaf"}
    for ln in lines:
        if ln.kind != "memo":
            continue
        name = next((n for n in sorted(inv, key=len, reverse=True) if n in ln.key), None)
        if name and ln.key.startswith("기초"):
            ln.ref = ("bs_pri", inv[name])
        elif name and ln.key.startswith("기말"):
            ln.ref = ("bs_cur", inv[name])
        elif ln.parent is not None:
            ln.ref = ("plug", ln.parent)
        else:
            ln.kind = "memo_input"
    return lines


# ── 쓰기 ─────────────────────────────────────────────

def ni_line(is_lines):
    hits = [x for x in is_lines if x.kind == "sum" and re.search(r"당기순(이익|손실|손익)", x.key)]
    return hits[-1] if hits else None


def ni_expr(ni, col, idx):
    return ("-" if ni.neg else "") + "IS!%s%d" % (col, idx[id(ni)])


def fs_sheet(wb, name, title, meta, lines, h_c, h_p, *, is_lines=None):
    ws = wb.create_sheet(name)
    write_header(ws, meta, title)
    rows = []
    idx = {id(ln): HDR_ROW + 1 + i for i, ln in enumerate(lines)}
    ni_row = idx.get(id(ni_line(is_lines))) if is_lines and name == "IS" else None
    for i, ln in enumerate(lines):
        r = HDR_ROW + 1 + i
        c, e = None, None
        if ln.kind in ("leaf",):
            key = ln.tb[0] if ln.tb else ln.key
            sgn = "" if ln.sign == 1 else "-"
            if ln.tb:
                c = '=%sSUMIF(TB!$E:$E,$G%d,TB!$H:$H)' % (sgn, r)
                e = '=%sSUMIF(TB!$E:$E,$G%d,TB!$I:$I)' % (sgn, r)
                for x in ln.extra:              # 재고 구성 밖 잎(관세환급금 등)은 아래 줄 금액을 더함
                    c += "+C%d" % idx[id(x)]
                    e += "+E%d" % idx[id(x)]
            ln.ref = key
        elif ln.kind == "sum" or ln.kind in ("T", "M", "total"):
            if ln.kids:
                parts = ["%s%%s%d" % ("-" if ln.flip.get(id(k), 1) == -1 else "+", idx[id(k)]) for k in ln.kids]
                expr = "".join(parts).lstrip("+")
                c = "=" + ("-(%s)" % expr if ln.neg else expr).replace("%s", "C")
                e = "=" + ("-(%s)" % expr if ln.neg else expr).replace("%s", "E")
        rows.append([num_space(ln.label), c, nv(ln.prior), e, ("=D%d-E%d" % (r, r)) if e else None,
                     ln.tb[0] if ln.tb else None])
    write_table(ws, ["과목", h_c, h_p, "TB 금액", "차이", "TB 연결 과목"], rows, header_row=HDR_ROW,
                number_cols={1: ACCT, 2: ACCT, 3: ACCT, 4: ACCT})
    sub, tot = [], []
    inputs = []
    for i, ln in enumerate(lines):
        r = HDR_ROW + 1 + i
        if ln.kind in ("total",) or (ln.kind == "sum" and name == "IS") or "총계" in ln.key:
            tot.append(r)
        if ln.kind in ("input", "memo_input") or (ln.kind == "leaf" and not ln.tb) or (ln.kind == "plain" and ln.prior is not None):
            inputs.append("C%d" % r)            # plain = 당기순이익 아래 수기 줄(기타포괄손익 · 주당손익 · 귀속)
    if inputs:
        mark_cells(ws, input=inputs)
    for col, w in zip("BCDEFG", (34, 18, 18, 18, 14, 26)):
        ws.column_dimensions[col].width = w
    return ws, idx


def nv(v):
    return float(v) if isinstance(v, Decimal) else v


def link_retained(ws, bs_lines, idx, is_ws_name, is_idx, is_lines):
    ni = ni_line(is_lines)
    if ni is None:
        return
    for ln in bs_lines:
        if ln.note == "ni":
            r = idx[id(ln)]
            for col, tbc in (("C", "H"), ("E", "I")):
                n = ni_expr(ni, col, is_idx)
                f = "SUMIF(TB!$E:$E,$G%d,TB!$%s:$%s)" % (r, tbc, tbc)
                ws["%s%d" % (col, r)] = ("=-%s+%s" % (f, n)) if ln.sign == -1 else ("=%s-(%s)" % (f, n))
            ln.note = ""
            break


def tb_sheet(wb, meta, tb, m, info):
    ws = wb.create_sheet("TB")
    write_header(ws, meta, "TB")
    rows = []
    for i, r in enumerate(tb):
        s = squeeze(r["settle"])
        line = m.acct_used.get(r["code"]) or m.used.get(s)
        dsd = line.tb[0] if line is not None and line.tb else None
        rows.append([r["code"], r["name"], r["settle"], dsd, float(r["cur"]), float(r["aje"]), "=F%d+G%d" % (HDR_ROW + 1 + i, HDR_ROW + 1 + i),
                     float(r["pri"]), float(r["plan"]) if r["plan"] is not None else None])
    write_table(ws, ["코드", "계정과목", "정산 과목", "DSD 행", info.get("cur_label", "회사제시"), info.get("aje_label", "수정사항"), info["cur"], info["pri"], info["plan"] or "참고"], rows, header_row=HDR_ROW, number_cols={0: "@", 4: ACCT, 5: ACCT, 6: ACCT, 7: ACCT, 8: ACCT})
    for col, w in zip("BCDEFGHIJ", (22, 28, 28, 28, 17, 15, 17, 17, 17)):
        ws.column_dimensions[col].width = w
    empties = ["E%d" % (HDR_ROW + 1 + i) for i, r in enumerate(tb) if not (m.acct_used.get(r["code"]) or m.used.get(squeeze(r["settle"])) or MFG_RE.search(squeeze(r["settle"])))]
    if empties:
        mark_cells(ws, todo=empties)
    adj = ["G%d" % (HDR_ROW + 1 + i) for i, r in enumerate(tb) if r.get("adj")]
    if adj:
        mark_cells(ws, input=adj)
    return ws


def table_grid(t):
    """주석 · 자본변동표 표 → 행렬(열 수 ncols). 병합 칸은 왼쪽에만."""
    grid = []
    for r in t["rows"]:
        row = [None] * t["ncols"]
        for c in r["cells"]:
            v = c.get("value")
            txt = tidy(c.get("text"))
            row[c["col"]] = float(v) if v is not None else (txt if txt else None)
        grid.append((row, bool(r.get("head"))))
    return grid


CAPTION_RE = re.compile(r"(?:^|[.\-:\s])(당\s*기\s*말?|전\s*기\s*말?)\s*[:.]?\s*$")
UNIT_RE = re.compile(r"\(?\s*단위\s*[:：]\s*[^)\s]+\s*\)?")


def notes_sheet(wb, meta, sections):
    ws = wb.create_sheet("주석")
    write_header(ws, meta, "주석")
    r = HDR_ROW
    total, pairs = 0, []
    for sec in sections:
        if sec["kind"] != "note":
            continue
        tabs = [b for b in sec["blocks"] if b["type"] == "table" and b["ncols"] >= 2]
        if not tabs:
            continue
        no, title = sec["section"], tidy(sec.get("title") or "")
        t0 = re.split(r"\s*\(1\)|/", re.sub(r"^\s*\d+\s*\.\s*", "", title))[0].strip()
        t0 = re.split(r"(?:당기|보고기간|당사|회사의|재무제표의 작성|당기말과)", t0)[0].strip() or t0
        head_txt = "%s. %s" % (no, t0[:16].strip())
        write_section_bar(ws, r, head_txt, first_col=2, last_col=9)
        ws.cell(r, 2).value = head_txt
        r += 2
        unit, caption, pending = None, None, None
        for b in sec["blocks"]:
            if b["type"] == "para":
                txt = tidy(b.get("text"))
                if not txt:
                    continue
                u = UNIT_RE.search(txt)
                if u:
                    unit = "(단위: %s)" % re.sub(r".*[:：]\s*", "", u.group(0)).strip(" )")
                m = CAPTION_RE.search(txt)
                caption = squeeze(m.group(1)) if m else None
                continue
            if b["type"] != "table":
                continue
            if b["ncols"] < 2:
                txt = " ".join(tidy(c.get("text")) for row in b["rows"] for c in row["cells"])
                u = UNIT_RE.search(txt)
                if u:
                    unit = "(단위: %s)" % re.sub(r".*[:：]\s*", "", u.group(0)).strip(" )")
                continue
            grid = table_grid(b)
            if not any(isinstance(x, float) for row, h in grid if not h for x in row) and not any(c.get("period") for c in b["columns"]):
                continue                       # 숫자 없는 글 표는 정산표에 두지 않음
            head = [row for row, h in grid if h]
            body = [row for row, h in grid if not h]
            h0 = squeeze(head[0][0]) if head and head[0][0] else ""
            mark = caption
            if re.fullmatch(r"당(분|반)?기말?", h0):
                mark = "당기"
            elif re.fullmatch(r"전(분|반)?기말?", h0):
                mark = "전기"
            if unit:
                ws.cell(r, 2, unit)
                r += 1
            if caption:
                ws.cell(r, 2, caption)
                r += 1
            headers = [str(x) if x is not None else "" for x in (head[0] if head else [None] * b["ncols"])]
            for extra in head[1:]:
                headers = [("%s %s" % (a, x)).strip() if x else a for a, x in zip(headers, extra)]
            # 여러 줄 머리글은 열마다 위 묶음까지 붙인다('당기말' 아래 '소유주식수 · 지분율' → 두 열 모두 '당기말 …').
            # 합친 칸의 둘째 열부터 묶음 이름이 빠지면 fill 이 그 열의 기간을 몰라 옆 열 값을 넣는다.
            full = {c["col"]: " ".join(c["header"].split("/")) for c in b["columns"] if c.get("header")}
            if len(head) > 1:
                headers = [full.get(i) or h for i, h in enumerate(headers)]
            body = [[num_space(x) if isinstance(x, str) else x for x in row] for row in body]
            write_table(ws, headers, body or [[None] * b["ncols"]], header_row=r,
                        number_cols={i: ACCT for i in range(1, b["ncols"])})
            for rr in range(r, r + len(body) + 1):
                for cc in range(2, 2 + b["ncols"]):
                    c_ = ws.cell(rr, cc)
                    if isinstance(c_.value, str) and len(c_.value) > 10:
                        c_.alignment = Alignment(wrap_text=True, vertical="center", horizontal=c_.alignment.horizontal)
            periods = [c.get("period") for c in b["columns"]]
            cur = [i for i, p in enumerate(periods) if p == "당기"]
            pri = [i for i, p in enumerate(periods) if p == "전기"]
            cream = []
            first = r + 1
            if cur and len(cur) == len(pri):
                for i, row in enumerate(body):
                    for c0, p0 in zip(cur, pri):
                        if row[c0] is None and isinstance(row[p0], float):
                            cell = ws.cell(first + i, 2 + c0).coordinate
                            cream.append(cell)
                            pairs.append([cell, ws.cell(first + i, 2 + p0).coordinate])
            elif mark and mark.startswith("당"):
                for i, row in enumerate(body):
                    for c0 in range(1, b["ncols"]):
                        if row[c0] is None:
                            cream.append(ws.cell(first + i, 2 + c0).coordinate)
                pending = (first, b["ncols"], len(body))
            elif mark and mark.startswith("전") and pending and pending[1] == b["ncols"] and pending[2] == len(body):
                for i, row in enumerate(body):
                    for c0 in range(1, b["ncols"]):
                        if isinstance(row[c0], float):
                            pairs.append([ws.cell(pending[0] + i, 2 + c0).coordinate, ws.cell(first + i, 2 + c0).coordinate])
                pending = None
            if cream:
                mark_cells(ws, input=cream)
            r += len(body) + 2
            total += 1
            caption = None
    ws.column_dimensions["B"].width = 44
    for col in "CDEFGHIJKLMNOPQR":
        ws.column_dimensions[col].width = 18
    return ws, total, pairs


def sce_sheet(wb, meta, sec, is_lines, is_idx):
    t = main_table(sec) or next(b for b in sec["blocks"] if b["type"] == "table" and b["ncols"] >= 3)
    ws = wb.create_sheet("SCE")
    write_header(ws, meta, "SCE")
    grid = table_grid(t)
    head = [row for row, h in grid if h]
    body = [row for row, h in grid if not h]
    headers = [str(x or "") for x in (head[0] if head else body[0])]
    for extra in head[1:]:
        headers = [("%s %s" % (a, x or "")).strip() for a, x in zip(headers, extra)]
    n = t["ncols"]
    tot_col = n - 1
    ni = ni_line(is_lines)
    rows, cream = [], []
    start = None
    for i, row in enumerate(body):
        out = list(row)
        r = HDR_ROW + 1 + i
        lab = squeeze(row[0] or "")
        is_open = bool(re.search(r"기초|\.0?1\.0?1", lab))
        is_end = bool(re.search(r"기말|12\.31", lab))
        cur_block = start is not None
        if is_open and any(isinstance(x, float) for x in row[1:]) and i > 0 and not all(x is None for x in row[1:]) and not cur_block:
            prior_blank = all(body[j][k] is None for j in range(i + 1, min(i + 3, len(body))) for k in range(1, n))
            if prior_blank:
                start = i
        elif cur_block and is_end:
            for c in range(1, n):
                L = _col(c + 2)
                out[c] = "=SUM(%s%d:%s%d)" % (L, HDR_ROW + 1 + start, L, r - 1)
            start = None
        elif cur_block:
            for c in range(1, n):
                if out[c] is None and c != tot_col:
                    cream.append("%s%d" % (_col(c + 2), r))
            if "당기순" in lab and ni is not None:
                retained = next((c for c in range(1, n) if "이익잉여금" in squeeze(headers[c]) or "결손금" in squeeze(headers[c])), None)
                if retained:
                    out[retained] = "=" + ni_expr(ni, "C", is_idx)
                    cream[:] = [x for x in cream if x != "%s%d" % (_col(retained + 2), r)]
            if out[tot_col] is None:
                out[tot_col] = "=SUM(%s%d:%s%d)" % (_col(3), r, _col(n), r)
        rows.append(out)
    rows = [[num_space(x) if isinstance(x, str) and not str(x).startswith("=") else x for x in row] for row in rows]
    write_table(ws, headers, rows, header_row=HDR_ROW, number_cols={i: ACCT for i in range(1, n)})
    if cream:
        mark_cells(ws, input=cream)
    ws.column_dimensions["B"].width = 34
    for c in range(1, n):
        ws.column_dimensions[_col(c + 2)].width = 18
    return ws


def _col(i):
    from openpyxl.utils import get_column_letter
    return get_column_letter(i)


def build_cf(lines, bs_lines, bs_idx, is_lines, is_idx):
    cur_t = cur_m = None
    tops = []
    for ln in lines:
        k = ln.key
        ln.kind = "input"
        if ln.lv == "T":
            cur_t, cur_m = ln, None
            ln.kind, ln.kids = "T", []
            tops.append(ln)
        elif ln.lv == "M":
            cur_m = ln
            ln.kind, ln.kids = "M", []
            if cur_t:
                cur_t.kids.append(ln)
        else:
            (cur_m or cur_t).kids.append(ln) if (cur_m or cur_t) else None
    for ln in lines:
        if ln.kind == "M":
            if ln.kids:
                set_sum(ln, ln.kids)
            else:
                ln.kind = "input"
    act = []
    cash = next((x for x in bs_lines if x.key.startswith("현금및현금성자산") and x.kind == "leaf"), None)
    for ln in tops:
        k = ln.key
        if re.search(r"활동", k) and ln.kids:
            set_sum(ln, ln.kids)
            act.append(ln)
        elif re.search(r"환율", k):
            act.append(ln)
            ln.kind = "input"
        elif re.search(r"증가|감소", k) and "현금" in k and act:
            set_sum(ln, act)
        elif "기초" in k and cash is not None:
            ln.kind, ln.ref = "cash_open", cash
        elif "기말" in k and cash is not None:
            ln.kind, ln.ref = "cash_end", cash
    ni = ni_line(is_lines)
    for ln in lines:
        if ln.kind == "input" and re.search(r"당기순(이익|손실|손익)", ln.key) and ni is not None:
            ln.kind, ln.ref = "ni", ni
    return lines


def cf_sheet(wb, meta, lines, h_c, h_p, bs_idx, is_idx):
    ws = wb.create_sheet("CF")
    write_header(ws, meta, "CF")
    idx = {id(ln): HDR_ROW + 1 + i for i, ln in enumerate(lines)}
    rows, cream = [], []
    for i, ln in enumerate(lines):
        r = HDR_ROW + 1 + i
        c = None
        if ln.kids and ln.kind in ("T", "M", "sum"):
            parts = ["%s%%s%d" % ("-" if ln.flip.get(id(k), 1) == -1 else "+", idx[id(k)]) for k in ln.kids]
            c = "=" + "".join(parts).lstrip("+").replace("%s", "C")
        elif ln.kind == "cash_open":
            c = "=BS!D%d" % bs_idx[id(ln.ref)]
        elif ln.kind == "cash_end":
            c = "=BS!C%d" % bs_idx[id(ln.ref)]
        elif ln.kind == "ni":
            c = "=" + ni_expr(ln.ref, "C", is_idx)
        elif ln.kind == "input":
            cream.append("C%d" % r)
        rows.append([num_space(ln.label), c, nv(ln.prior)])
    write_table(ws, ["과목", h_c, h_p], rows, header_row=HDR_ROW, number_cols={1: ACCT, 2: ACCT})
    if cream:
        mark_cells(ws, input=cream)
    for col, w in zip("BCD", (42, 18, 18)):
        ws.column_dimensions[col].width = w
    return ws


def link_is_memo(ws, lines, idx, bs_idx):
    for ln in lines:
        if ln.kind != "memo" or not ln.ref:
            continue
        r = idx[id(ln)]
        kind, target = ln.ref
        if kind == "plug":
            others = [x for x in lines if x.kind == "memo" and x.parent is target and x is not ln]
            ex = "".join("-C%d" % idx[id(x)] for x in others + list(target.extra))
            ws["C%d" % r] = "=C%d%s" % (idx[id(target)], ex)
        else:
            ws["C%d" % r] = ("=BS!D%d" % bs_idx[id(target)]) if kind == "bs_pri" else ("=-BS!C%d" % bs_idx[id(target)])


MFG_RE = re.compile(r"\((제|도|분|공)\)$|제조비용대체")


def check_sheet(wb, meta, bs, is_, leftovers, m):
    ws = wb.create_sheet("연결점검")
    write_header(ws, meta, "연결점검")
    rows = []
    for name, lines in (("BS", bs), ("IS", is_)):
        for ln in lines:
            if ln.kind == "leaf" and ln.tb and ln.prior is not None and not ln.extra and not RETAINED.match(ln.key):
                gap = ln.prior - sum(m.amount(x) for x in ln.tb) * ln.sign
                if gap:
                    rows.append([name, ln.label, "전기 금액 차이(DSD - TB)", float(gap)])
            if ln.kind == "leaf" and not ln.tb and ln.prior:
                rows.append([name, num_space(ln.label), "TB에 짝 없음", nv(ln.prior)])
            elif ln.kind in ("input", "sum") and ln.note:
                rows.append([name, num_space(ln.label), ln.note, nv(ln.prior)])
    for s, a in leftovers.items():
        if a != 0 and not MFG_RE.search(s):
            rows.append(["TB", s, "DSD 행에 짝 없음(전기 잔액 있음)", float(a)])
    write_table(ws, ["구분", "과목", "내용", "전기 금액"], rows or [["-", "해당 없음", None, None]], header_row=HDR_ROW,
                number_cols={3: ACCT})
    for col, w in zip("BCDEFG", (8, 40, 44, 18, 10, 10)):
        ws.column_dimensions[col].width = w
    return ws, rows


# ── 연결 기록(숨은 시트) ──────────────────────────────

LINK_SCHEMA = "dsdxlsx.link/1"
_PAIR_ROW = 8


def link_sheet(wb, *, side: str, source: str, sha256: str, note_pairs: list) -> None:
    """정산표가 어느 DSD 에서 만들어졌는지와 주석의 당기 · 전기 칸 짝을 숨은 시트에 적는다."""
    ws = wb.create_sheet(LINK_SHEET)
    for i, (k, v) in enumerate((("schema", LINK_SCHEMA), ("side", side), ("source", source), ("sha256", sha256),
                                ("note_pairs", len(note_pairs))), 1):
        ws.cell(i, 1, k)
        ws.cell(i, 2, v)
    ws.cell(_PAIR_ROW - 1, 1, "cur")
    ws.cell(_PAIR_ROW - 1, 2, "pri")
    for i, (cur, pri) in enumerate(note_pairs):
        ws.cell(_PAIR_ROW + i, 1, cur)
        ws.cell(_PAIR_ROW + i, 2, pri)
    ws.sheet_state = "hidden"


def read_link(wb) -> dict | None:
    """정산표의 연결 기록. 없으면 None."""
    if LINK_SHEET not in wb.sheetnames:
        return None
    ws = wb[LINK_SHEET]
    meta = {str(ws.cell(r, 1).value): ws.cell(r, 2).value for r in range(1, _PAIR_ROW - 1) if ws.cell(r, 1).value}
    if meta.get("schema") != LINK_SCHEMA:
        return None
    pairs = []
    for r in range(_PAIR_ROW, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if a and b:
            pairs.append([str(a), str(b)])
    meta["note_pairs"] = pairs
    return meta


def is_link_workbook(path: str) -> bool:
    """dsdxlsx 가 만든 정산표인가(숨은 시트로 가린다)."""
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            return ('name="%s"' % LINK_SHEET).encode() in z.read("xl/workbook.xml")
    except (OSError, KeyError, zipfile.BadZipFile):
        return False


# ── 만들기 · 다시 넣기 ────────────────────────────────

class LinkError(InputError):
    """정산표를 만들 수 없는 입력(재무제표 구획이 없는 DSD 등)."""


def recon_config() -> dict:
    """이 정산표를 recon · fill · build 가 읽을 때 쓰는 시트 지정."""
    return {"재무제표": {"BS": ["BS"], "PL": ["IS"], "SCE": ["SCE"], "CF": ["CF"]}, "주석": ["주석"],
            "빼는 시트": [], "빼는 칸": [], "이름 다른 행": {}}


def make(doc: dict, tb: list, info: dict, *, company: str = "", title: str | None = None, ref: str = "",
         author: str | None = None, labels=None, consolidated: bool = False, source: str = "", sha256: str = ""):
    """뼈대 DSD(inspect 사전)와 TB 로 정산표 통합문서를 만든다. 반환: (Workbook, 요약 사전)."""
    sec = {s["section"]: s for s in doc["sections"]}
    bs_s = sec.get("재무상태표")
    is_s = sec.get("손익계산서") or sec.get("포괄손익계산서")
    bs_t = main_table(bs_s) if bs_s else None
    is_t = main_table(is_s) if is_s else None
    if bs_t is None or is_t is None:
        raise LinkError("뼈대 DSD 에서 당기 · 전기 열이 있는 재무상태표와 손익계산서를 찾지 못했습니다")
    m = Matcher(tb)
    bs, h_c, h_p = lines_of(bs_t)
    bs = build_bs(bs, m)
    is_, ic, ip = lines_of(is_t)
    is_ = build_is(is_, m, bs)
    meta = {"company": company, "ref": ref, "date": info["cur"], "title": title or ("연결정산표" if consolidated else "정산표"),
            "author": author, "labels": labels}
    wb = Workbook()
    wb.remove(wb.active)
    tb_sheet(wb, meta, tb, m, info)
    ws_bs, bs_idx = fs_sheet(wb, "BS", "재무상태표", meta, bs, h_c, h_p)
    ws_is, is_idx = fs_sheet(wb, "IS", "손익계산서", meta, is_, ic, ip, is_lines=is_)
    link_retained(ws_bs, bs, bs_idx, "IS", is_idx, is_)
    link_is_memo(ws_is, is_, is_idx, bs_idx)
    sheets = ["TB", "BS", "IS"]
    sce_s = sec.get("자본변동표")
    if sce_s and any(b["type"] == "table" and b["ncols"] >= 3 for b in sce_s["blocks"]):
        sce_sheet(wb, meta, sce_s, is_, is_idx)
        sheets.append("SCE")
    cf_t = main_table(sec["현금흐름표"]) if sec.get("현금흐름표") else None
    if cf_t is not None:
        cf, cc, cp = lines_of(cf_t, "CF")
        cf = build_cf(cf, bs, bs_idx, is_, is_idx)
        cf_sheet(wb, meta, cf, cc, cp, bs_idx, is_idx)
        sheets.append("CF")
    _, n_tab, note_pairs = notes_sheet(wb, meta, doc["sections"])
    left = m.leftovers()
    _, chk = check_sheet(wb, meta, bs, is_, left, m)
    side = "연결" if consolidated else "별도"
    link_sheet(wb, side=side, source=source, sha256=sha256, note_pairs=note_pairs)
    unlinked = [s for s, a in left.items() if a != 0 and not MFG_RE.search(s)]
    summary = {"side": side, "sheets": sheets + ["주석", "연결점검"], "accounts": len(tb), "note_tables": n_tab,
               "to_check": [{"sheet": r[0], "label": r[1], "issue": r[2], "prior": r[3]} for r in chk],
               "unlinked_tb": unlinked}
    return wb, summary


def refresh(wb, tb: list) -> dict:
    """정산표 TB 시트의 금액만 다시 넣는다. 사람이 고친 'DSD 행'은 그대로, 새 계정은 맨 아래(DSD 행 비움, 노란 칸)."""
    if "TB" not in wb.sheetnames:
        raise LinkError("정산표에 TB 시트가 없습니다")
    ws = wb["TB"]
    pos = {str(ws.cell(r, 2).value): r for r in range(HDR_ROW + 1, ws.max_row + 1) if ws.cell(r, 2).value is not None}
    last = max(pos.values()) if pos else HDR_ROW
    added = []
    for r0 in tb:
        r = pos.get(str(r0["code"]))
        if r is None:
            last += 1
            r = last
            added.append(r0["code"])
            ws.cell(r, 2).value, ws.cell(r, 3).value, ws.cell(r, 4).value = r0["code"], r0["name"], r0["settle"]
            ws.cell(r, 8).value = "=F%d+G%d" % (r, r)
            for c in range(2, 11):
                ws.cell(r, c)._style = ws.cell(r - 1, c)._style
            mark_cells(ws, todo=["E%d" % r])
        ws.cell(r, 6).value, ws.cell(r, 7).value, ws.cell(r, 9).value = float(r0["cur"]), float(r0["aje"]), float(r0["pri"])
        if r0["plan"] is not None:
            ws.cell(r, 10).value = float(r0["plan"])
    return {"accounts": len(tb), "added": added}
