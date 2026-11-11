"""계정 분류 규칙 - 재무제표 안 위치로 계정을 분류하고 회사 분개장의 계정코드에 짝짓는다.

계정코드는 회사마다 다르므로 표준 계정표를 쓰지 않는다. 재무상태표 · 손익계산서 · 제조원가명세서의 줄이
곧 TB 의 계정이고, 분류(대분류 · 소분류 · 공시 · 정산 과목)는 그 줄이 놓인 자리로 정한다.
  build()   TB 행 [[코드, 이름, 당기, 전기]] + 분류 {코드: [대분류, 소분류, 공시, 정산]}
"""
import re
from collections import Counter, OrderedDict, defaultdict

from .douzone import _ROMAN, _load_rows, norm_name, parse_cost, to_num

TOL = 1.0
TAG = re.compile(r"\((판|제|도|분|공|운|용)\)$")                      # 더존 계정 이름 끝의 원가 구분 표시
COST_TAGS = ("제", "도", "분", "공", "운", "용")
CONTRA_ASSET = re.compile(r"감가상각누계액|손상차손누계액|대손충당금|국고보조금|정부보조금|현재가치할인차금|평가충당금|감모충당금|\(국고\)$")   # 기계장치(국고) = 세무사랑 국고보조금
CONTRA_LIAB = re.compile(r"퇴직연금운용자산|퇴직보험예치금|국민연금전환금|현재가치할인차금|사채할인발행차금|전환권조정|신주인수권조정")
RETAINED = ("미처분이익잉여금", "이월이익잉여금", "미처리결손금", "이월결손금")
CLOSING = ("손익", "미처분이익잉여금", "미처리결손금")                  # 이 계정이 든 전표 = 마감 대체(손익 · 이익잉여금 대체)
PL_GROUPS = ("매출액", "매출원가", "판매비와관리비", "영업외수익", "영업외비용", "법인세등", "법인세비용")
SALES_DEDUCT = re.compile(r"환입|에누리|매출할인")


def base(name):
    return TAG.sub("", norm_name(name))


def tag(name):
    m = TAG.search(norm_name(name))
    return m.group(1) if m else ""


def ckey(c):
    s = str(c)
    return (0, int(s), s) if s.isdigit() else (1, 0, s)


def heading(key):
    """'Ⅰ.유동자산' → '유동자산', '(1)당좌자산' → '당좌자산'."""
    return re.sub(r"^\(\d+\)", "", key.lstrip(_ROMAN).lstrip("."))


# ───── 재무제표 줄 ─────
def bs_lines(path):
    """재무상태표 계정 줄 → [{key, cur, pri, section, sub, group, contra}] (금액은 표시 그대로 - 부호는 build 에서)."""
    out, section, sub, group = [], "", "", ""
    for r in _load_rows(path):
        name = str(r[0] or "").strip()
        if not name:
            continue
        key = norm_name(name)
        c = [to_num(r[i]) if len(r) > i else 0.0 for i in range(5)]
        if "총계" in key:
            continue
        if key in ("자산", "부채", "자본") and not any(c[1:]):
            section, sub, group = key, "", ""
            continue
        if name[0] in _ROMAN:
            sub, group = heading(key), ""
            if "결손" in sub or ("이익" in sub and "잉여금" in sub):
                sub = "이익잉여금"
            continue
        if re.match(r"^\(\d+\)", key):
            group = heading(key)
            continue
        if name.startswith("(") and not any(c[1:]):
            continue                                                     # (당기순이익 …) 같은 설명 줄 - (청도)입회보증금처럼 금액이 있으면 계정
        if key.startswith(("당기", "전기")):
            continue
        contra = bool((CONTRA_ASSET if section == "자산" else CONTRA_LIAB if section == "부채" else re.compile("$^")).search(key))
        if contra and not (c[1] or c[3]) and (c[2] or c[4]):               # 세무사랑: 차감 계정을 바깥 칸 음수로(건물국고보조금 · 퇴직연금운용자산) - 보인 부호 그대로
            contra = False                                               # 바깥 칸 양수(현재가치할인차금 차변 9,000,000 - 합계에 더해짐)도 같은 규칙
        cur = c[1] if contra or (c[1] and not c[2]) else c[2]
        pri = c[3] if contra or (c[3] and not c[4]) else c[4]
        if cur or pri or contra:
            out.append({"key": key, "cur": cur, "pri": pri, "section": section, "sub": sub, "group": group or sub, "contra": contra})
    return out


def is_lines(path):
    """손익계산서 → (줄[{group, key, cur, pri, sign}], 매출원가 세부[(매출원가 키, 키, 당기, 전기)], 요약{키: (당기, 전기)}).
    매출원가 그룹은 바깥 칸 금액 줄(상품 · 제품매출원가 · 운송원가 …)이 계정이고, 그 아래 안쪽 칸 줄 중 기초 · 기말 · 당기매입 · 당기제조원가가
    아닌 것(관세환급금 · 타계정대체 …)은 세부. 그 밖 그룹은 안쪽 칸 줄이 계정이고, 줄이 없는 그룹(법인세등)은 그룹 줄이 계정.
    매출액의 환입 · 에누리 · 매출할인은 차감(차변)."""
    rows = _load_rows(path)
    lines, detail, summary = [], [], {}
    group, gline, parent, n_in = "", None, "", 0

    def close():
        if gline and group in PL_GROUPS and not n_in and (gline[1] or gline[2]):
            lines.append({"group": group, "key": group, "cur": gline[1], "pri": gline[2], "sign": -1 if group in ("매출액", "영업외수익") else 1})
    for r in rows:
        name = str(r[0] or "").strip()
        if not name:
            continue
        key = norm_name(name)
        c = [to_num(r[i]) if len(r) > i else 0.0 for i in range(5)]
        if name[0] in _ROMAN:
            close()
            group = heading(key)
            summary[group] = (c[2], c[4])
            gline, parent, n_in = (key, c[2], c[4]), "", 0
            continue
        if re.match(r"^\(\d+\)", key) or (name.startswith("(") and not any(c[1:])):
            continue
        if group not in PL_GROUPS:
            continue
        if group == "매출원가":
            if (c[2] or c[4]) and not (c[1] or c[3]):
                lines.append({"group": group, "key": key, "cur": c[2], "pri": c[4], "sign": 1})
                parent, n_in = key, n_in + 1
            elif c[1] or c[3]:
                if key.startswith(("기초", "기말")) or (key.startswith("당기") and ("매입" in key or "원가" in key)):
                    continue
                detail.append((parent, key, c[1], c[3]))
            continue
        if c[1] or c[3]:
            s = -1 if group in ("매출액", "영업외수익") else 1
            if group == "매출액" and SALES_DEDUCT.search(key):
                s = 1
            lines.append({"group": group, "key": key, "cur": c[1], "pri": c[3], "sign": s})
            n_in += 1
    close()
    return lines, detail, summary


# ───── TB ─────
def build(bs_p, is_p, journal, cost_p=None, book=None, info=None):
    """반환 (rows[[코드, 이름, 당기, 전기]], classes{코드: [대분류, 소분류, 공시, 정산]}, warnings[], errors[]).
    info(dict)를 주면 이월이익잉여금 값을 채운다 - info['ret'] = {code, op(마감 전 기초 당기 · 전기), ni(당기순이익), closed(재무상태표 잔액, 차변 +)}."""
    warns, errors = [], []
    names, nets, nets_open = OrderedDict(), defaultdict(float), defaultdict(float)
    by_v = defaultdict(list)
    for r in journal:
        c = str(r["계정코드"])
        names.setdefault(c, r["계정과목"])
        by_v[r["전표키"]].append(r)
    for v, rs in by_v.items():
        shut = any(base(x["계정과목"]) in CLOSING for x in rs)
        for x in rs:
            amt = (x["차변"] or 0.0) - (x["대변"] or 0.0)
            nets[str(x["계정코드"])] += amt
            if not shut:
                nets_open[str(x["계정코드"])] += amt
    known = OrderedDict((str(k), str(v)) for k, v in (book or {}).items() if not str(k).startswith("*"))
    known.update(names)                                                  # 분개장 이름이 우선
    bidx = defaultdict(list)
    for c in sorted(known, key=ckey):
        bidx[base(known[c])].append(c)
    rows, classes, used = [], {}, set()

    def add(code, name, cur, pri, cls):
        rows.append([code, name, round(cur, 2), round(pri, 2)])
        classes.setdefault(code, cls)
        used.add(code)

    # 1) 재무상태표
    bl = bs_lines(bs_p)
    ret = next((x for x in reversed(bl) if x["section"] == "자본" and x["key"] in RETAINED), None)
    prev_code, prev_key = None, ""
    last_code, last_sec = None, ""
    for x in bl:
        if x is ret:
            continue
        settle = "%s(%s)" % (x["key"], prev_key) if x["contra"] and prev_key else x["key"]
        if x["contra"] and prev_key and x["key"].endswith("(국고)"):        # 세무사랑 기계장치(국고) → 정산 과목 국고보조금(기계장치)
            settle = "국고보조금(%s)" % prev_key
        s = 1 if x["section"] == "자산" else -1
        if x["contra"]:
            s = -s
        cands = [c for c in bidx.get(base(x["key"]), []) if c not in used]
        delta = s * (x["cur"] - x["pri"])                                # 재무상태표 증감 = 그 해 분개장 순액(같은 이름 후보 가리기)
        fit = [c for c in cands if abs(nets.get(c, 0.0) - delta) <= TOL]
        if fit and (len(cands) > 1 or not delta):
            cands = fit
        elif not delta and cands and all(abs(nets.get(c, 0.0)) > TOL for c in cands):
            cands = []                                                   # 증감 없는 줄에 거래가 있는 코드는 붙이지 않음(다른 줄의 코드)
        if x["contra"] and x["section"] == "자산" and last_code and last_sec == "자산":   # 자산 차감 계정은 앞서 정해진 코드보다 뒤 코드(부채 차감 퇴직연금운용자산은 자산 번호라 제외)
            cands = [c for c in cands if ckey(c) > ckey(last_code)]
        if x["contra"] and not prev_code and len(bidx.get(base(x["key"]), [])) > 1:   # 바로 위 자산이 이름 키(코드 모름)면 다른 자산 바로 다음 코드(212 비품 → 213)는 그 자산 몫
            cands = [c for c in cands if not (c.isdigit() and str(int(c) - 1) in known and base(known[str(int(c) - 1)]) != base(x["key"]))]
        if len(cands) > 1 and prev_code:
            after = [c for c in cands if ckey(c) > ckey(prev_code)]
            cands = after or cands
        code = cands[0] if cands else "*" + settle
        if code in used:
            errors.append("재무상태표 같은 줄이 둘: %s" % settle)
            continue
        if not code.startswith("*"):
            last_code, last_sec = code, x["section"]
        disc = x["group"] if x["section"] == "자산" else x["sub"]
        add(code, known.get(code, x["key"]), s * x["cur"], s * x["pri"], [x["section"], x["sub"], disc, settle])
        if not x["contra"]:
            prev_code, prev_key = (None if code.startswith("*") else code), x["key"]

    # 2) 손익계산서 · 제조원가명세서
    lines, detail, summary = is_lines(is_p)
    cost_ps = [p for p in (cost_p if isinstance(cost_p, (list, tuple)) else [cost_p]) if p]   # 원가명세서 여럿(세무사랑 공사원가 · 분양(용역)원가)
    costs = [parse_cost(p)["leaves"] for p in cost_ps]
    pl_used = set()
    cost_tag = ""

    def fits(c, kind, group):                                            # 원가 구분 표시가 이 줄과 맞는 계정
        t_ = tag(known[c])
        if kind == "제조":
            return t_ != "판"
        if group == "판매비와관리비":
            return t_ in ("판", "")
        return t_ not in COST_TAGS and t_ != "판"

    def pick(key, kind, group, target):
        cs = [c for c in bidx.get(base(key), []) if c not in used and c not in pl_used and fits(c, kind, group)]
        if not cs and target:                                            # 재무제표 출력 이름이 계정 이름과 다른 회사(접대비 ↔ 접대비(기업업무추진비)(판)): 금액이 같은 계정이 하나뿐이면
            amt = [c for c in names if c not in used and c not in pl_used and fits(c, kind, group) and abs(nets_open.get(c, 0.0) - target) <= TOL]
            if len(amt) == 1:
                warns.append("%s: 이름이 같은 계정이 없어 금액이 같은 %s %s 로 넣음" % (key, amt[0], names[amt[0]]))
                return [(amt[0], target)]
        if not cs:
            return []
        exact = [c for c in cs if abs(nets_open.get(c, 0.0) - target) <= TOL]
        live = [c for c in cs if abs(nets_open.get(c, 0.0)) > TOL]
        if target and len(exact) == 1:
            return [(exact[0], target)]
        if target and len(live) > 1 and abs(sum(nets_open[c] for c in live) - target) <= TOL:
            return [(c, nets_open[c]) for c in live]
        if len(cs) > 1 and kind == "제조" and cost_tag:                   # 원가명세서에서 하나로 정해진 계정들의 표시((제) · (분) …)를 따름
            cs = [c for c in cs if tag(known[c]) == cost_tag] or cs
        if len(cs) > 1 and target:
            warns.append("%s: 같은 이름 계정 %s 중 분개장 금액이 맞는 것이 없어 %s 로 넣음" % (key, ", ".join(cs), cs[0]))
        elif target and live and abs(nets_open[cs[0]] - target) > TOL and group != "매출원가":   # 매출원가는 세부(관세환급금 등)를 뗀 뒤에 본다
            warns.append("%s(%s): 재무제표 %s · 분개장 순액 %s" % (key, cs[0], "{:,.0f}".format(target), "{:,.0f}".format(nets_open[cs[0]])))
        return [(cs[0], target)]

    def place(key, kind, group, cur, pri, cls):
        hit = pick(key, kind, group, cur)
        if not hit:                                                      # 제조는 '(제)'를 붙여 판관비 같은 이름(퇴직급여 등)과 키가 겹치지 않게
            settle = cls[3] + ("(제)" if kind == "제조" else "")
            add("*" + settle, key, cur, pri, cls[:3] + [settle])
            return ["*" + settle]
        for i, (c, v) in enumerate(hit):                                 # 제조의 정산 과목 = 회사 계정 이름 …)
            add(c, known[c], v, pri if i == 0 else 0.0, cls[:3] + [known[c]] if kind == "제조" else cls)
            pl_used.add(c)
        return [c for c, _ in hit]

    mfg = [0.0, 0.0]
    for cost in costs:
        one = [bidx.get(base(k), []) for _, k, _, _ in cost]
        tags = Counter(tag(known[cs[0]]) for cs in (x for x in ([c for c in cs if fits(c, "제조", "")] for cs in one) if len(x) == 1))
        cost_tag = tags.most_common(1)[0][0] if tags else ""           # 명세서마다 원가 표시((도) 공사 · (분) 분양 …)
        for grp, key, cur, pri in cost:
            place(key, "제조", grp, cur, pri, ["제조", grp, grp, key])
            mfg[0] += cur
            mfg[1] += pri
    # 결산 전(중간) 세무사랑 손익계산서는 매출원가를 대체하지 않아 0 으로 나온다(2026.09) - 그 기간은 원가명세서 금액이 그대로 비용이므로
    # 대체 줄을 넣지 않고 이월이익잉여금 계산의 당기순이익에서 뺀다(TB 부호합은 그대로 0).
    unclosed = [mfg[i] if abs(mfg[i]) > TOL and abs((summary.get("매출원가") or (0.0, 0.0))[i]) <= TOL else 0.0 for i in (0, 1)]
    if any(costs):
        add("*당기총제조비용 대체(재공품)", "당기총제조비용 대체(재공품)", -(mfg[0] - unclosed[0]), -(mfg[1] - unclosed[1]),
            ["제조", "대체", "대체", "당기총제조비용 대체(재공품)"])
    for i, label in ((0, "회사제시"), (1, "전기")):
        if unclosed[i]:
            warns.append("%s 손익계산서 매출원가가 0 이라(결산 전) 원가명세서 %s 을 대체 없이 비용으로 둠" % (label, "{:,.0f}".format(unclosed[i])))
    for x in lines:
        g = x["group"]
        cls = ["손익", g, g, base(x["key"])]
        codes = place(x["key"], "손익", g, x["sign"] * x["cur"], x["sign"] * x["pri"], cls)
        if g != "매출원가":
            continue
        row = next(r for r in rows if r[0] == codes[0])
        for parent, dkey, dcur, _ in detail:                            # 관세환급금 등: 분개장에 따로 있으면 떼어 내고 합계에서 뺌
            if parent != x["key"]:
                continue
            hit = [c for c in bidx.get(base(dkey), []) if c not in used and abs(abs(nets_open.get(c, 0.0)) - abs(dcur)) <= TOL and nets_open.get(c)]
            if len(hit) == 1:
                add(hit[0], known[hit[0]], nets_open[hit[0]], 0.0, ["손익", "매출원가", "매출원가", base(dkey)])
                row[2] = round(row[2] - nets_open[hit[0]], 2)
        if codes[0].startswith("*") or abs(nets_open.get(codes[0], 0.0) - row[2]) > TOL:
            # 원가명세서 없이 매출원가 한 줄(운송원가 등)만 있는 회사: 원가 표시 (도)(분)… 계정 순액 합이 그 금액이면 그 계정들로
            rest = [c for c in names if c not in used and tag(names[c]) in COST_TAGS and abs(nets_open.get(c, 0.0)) > TOL]
            if rest and abs(sum(nets_open[c] for c in rest) - row[2]) <= TOL:
                for c in rest:
                    add(c, names[c], nets_open[c], 0.0, ["제조", x["key"], x["key"], names[c]])
                row[2] = 0.0
            elif not codes[0].startswith("*") and abs(nets_open.get(codes[0], 0.0)) > TOL:
                warns.append("%s(%s): 재무제표 %s · 분개장 순액 %s" % (x["key"], codes[0], "{:,.0f}".format(row[2]), "{:,.0f}".format(nets_open[codes[0]])))

    # 3) 이월이익잉여금(마감 전) · 전기오류수정손익
    if ret:
        sgn = -1 if "결손" in ret["key"] else 1
        unap = (sgn * ret["cur"], sgn * ret["pri"])
        if "당기순손실" in summary:
            ni = tuple(-v for v in summary["당기순손실"])
        else:
            ni = summary.get("당기순이익") or summary.get("당기순손익") or (0.0, 0.0)
        ni = (ni[0] - unclosed[0], ni[1] - unclosed[1])                  # 결산 전 손익계산서에 빠진 원가(위 2) 참고)
        adj, adj_code = 0.0, None
        for v, rs in by_v.items():
            if any(base(x["계정과목"]) in CLOSING for x in rs):
                continue
            for x in rs:
                if base(x["계정과목"]) in ("전기오류수정손실", "전기오류수정이익"):
                    adj += (x["차변"] or 0.0) - (x["대변"] or 0.0)
                    adj_code = adj_code or str(x["계정코드"])
        op = (unap[0] - ni[0] + adj, unap[1] - ni[1])
        cls = ["자본", ret["sub"], ret["sub"], ret["key"]]
        cands = [c for nm in ("이월이익잉여금", "이월결손금", ret["key"]) for c in bidx.get(nm, []) if c not in used]
        code = cands[0] if cands else "*" + ret["key"]
        add(code, known.get(code, "이월이익잉여금"), -op[0], -op[1], cls)
        if abs(adj) > TOL:
            add(adj_code, known[adj_code], adj, 0.0, cls)
        if info is not None:
            info["ret"] = {"code": code, "op": op, "ni": ni, "closed": (-unap[0], -unap[1]), "key": ret["key"]}

    for i, label in ((2, "회사제시"), (3, "전기")):
        tot = sum(r[i] for r in rows)
        if abs(tot) > TOL:
            errors.append("TB %s 부호합 %s ≠ 0" % (label, "{:,.0f}".format(tot)))
    left = [c for c in names if c not in used and abs(nets_open.get(c, 0.0)) > TOL and c not in {r[0] for r in rows}]
    if left:
        warns.append("TB 에 안 들어간 분개장 계정(순액 있음, 재무상태표 계정이면 정상): " + ", ".join("%s %s" % (c, names[c]) for c in left[:15]))
    return rows, classes, warns, errors
