"""풋팅(합계검증) 엔진 — dsdxlsx check.py의 검산 휴리스틱 포팅 + 수식 방출 게이트.

검사 종류(kind)와 방출 정책:
- 확실형(총계 라벨·AUTOSUM·단계손익·기초/기말 흐름): 후보가 있으면 항상 방출 —
  불일치면 최근접 후보를 수식화해 엑셀에 빨간 차이를 노출한다.
- 추측형(group/rollup — 들여쓰기 계층 추측): |차이| ≤ 허용오차일 때만 방출 —
  계층 추측이 빗나간 표에 노이즈 수식을 만들지 않는다.

핵심 규칙(다중 후보): 대상 셀에 대해 그럴듯한 합 여러 개를 만들어 Python(Decimal)으로
전부 평가하고, 정확히 일치하는 첫 후보를 수식으로 확정한다. 12자리 금액의 우연 일치는
사실상 없어 후보 다중화가 오탐을 만들지 않는다 (dsdxlsx 8개사 캘리브레이션 근거).
"""
from __future__ import annotations

import re
from decimal import Decimal

from . import constants as C
from .checks import FootingSpec
from .tableview import TableView, is_total_label

_ROMAN = "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ"
_HINT_RE = re.compile(rf"\(([{_ROMAN}][+{_ROMAN}\s]*)\)$")
_PREFIX_RE = re.compile(rf"^([{_ROMAN}]+|[0-9]+|[가-힣])\s*[.)]\s*|^\(\d+\)\s*")

# maturity: IFRS 만기분석은 할인 전 계약상 현금흐름이라 장부금액과 다른 게 정상 —
# 일치할 때만 방출(불일치는 오류가 아니라 표시 관행)
# socie_sub: 섹션 항목 나열이 계층(소계 포함)일 수 있어 정확 일치만 방출
# fv_equal: 원가측정·지분법 종목은 장부가액≠공정가치가 정상 — 일치할 때만
# summary_fin: 피투자회사 요약재무 자본=자산−부채 — NCI 등 예외 대비 일치만
# eq_carry: 무변동 이월(기초=기말) — 값이 같을 때만 (구현 후보 ⑤)
# net_pair: 이연법인세 순액 = 자산 계 − 부채 계 (라벨 짝 추측) — 일치만
GUESS_KINDS = {"group", "rollup", "combo", "cascade", "maturity", "socie_sub",
               "fv_equal", "summary_fin", "net_pair", "eq_carry"}
# 동일 대상 셀 중복 시 우선순위 (일치 우선 → 의미 순)
KIND_PRIORITY = ["pl_anchor", "cf_anchor", "hint", "socie", "socie_sub", "flow_h",
                 "asset_cost", "maturity", "fv_equal", "summary_fin", "net_pair",
                 "group", "rollup", "total_row", "total_col", "autosum", "cascade",
                 "combo", "ratio", "eq_carry"]
# 같은 대상 셀·같은 축이라도 별개 관계라 공존하는 kind (유형자산 기말: 흐름+취득원가)
COEXIST_KINDS = {"flow_h", "asset_cost", "fv_equal"}

# 고재현율 검출 프로필(--detection-profile exhaustive) — analyze(relaxed=True)가 켠다
RELAXED = False


def _mirror_pass(d: Decimal, digits: int = 1) -> bool:
    """방출될 수식의 판정 미러: IF(ABS(IF(d=TRUNC(d),d,ROUND(d,n)))<1e-6,TRUE,…)."""
    q = Decimal(1).scaleb(-digits)
    d2 = d if d == d.to_integral_value() else d.quantize(q)
    return abs(d2) < Decimal("0.000001")


def _label_core(label: str) -> str:
    """'영업이익(손실)', '기본주당이익(손실) (단위 : 원)' → 괄호 접미 제거."""
    core = C.norm(label)
    while True:
        stripped = re.sub(r"\([^()]*\)$", "", core)
        if stripped == core:
            return core
        core = stripped


def clean_label(label: str) -> str:
    """검증 라벨용 정리: 로마자/번호 접두·괄호 접미·공백 제거."""
    s = label.strip()
    s = _PREFIX_RE.sub("", s)
    return _label_core(s)


def _despace(text: str) -> str:
    return C.norm(text)


def _clean_header(text: str) -> str:
    """헤더 정규화: 공백 제거 + 각주 표시 제거 ('상각누계액(*1)' → '상각누계액').

    각주 마커((*)·(*1)·(주1))는 표시용이라 구조 판별에서 빼야 한다 — 안 그러면
    '누계액$' 같은 접미 규칙이 통째로 빗나간다 (IFRS 주석 실측)."""
    return re.sub(r"\((?:\*|주)[^)]*\)", "", C.norm(text))


def _headers_of(tv: TableView) -> list[str]:
    hs = [_clean_header(tv.leaf_header(c)) for c in range(tv.ncols)]
    if any(hs):
        return hs
    # THEAD 없는 수기 표 폴백 (공단 실측): 값이 전혀 없는 선두 본문 행(≤3)을
    # 헤더로 간주 — 아래 행 텍스트 우선(leaf 의미). 해당 행들은 값이 없어
    # 헤더 기반 검사(flow·asset_cost 등)의 대상/항이 되지 않는다.
    header_rows = []
    for i in tv.body:
        if any(tv.value(i, c) is not None for c in range(tv.ncols)):
            break
        header_rows.append(i)
        if len(header_rows) >= 3:
            break
    if header_rows:
        texts = ["" for _ in range(tv.ncols)]
        for i in header_rows:
            for c in range(tv.ncols):
                t = _clean_header(tv.text(i, c))
                if t:
                    texts[c] = t
        if sum(1 for t in texts if t) >= 2:
            return texts
    return hs


# ── 수집기 ───────────────────────────────────────────────────────────

class Collector:
    def __init__(self):
        self.raw: list[dict] = []

    def add(self, tv: TableView, kind: str, axis: str, i: int, c: int,
            cands: list, stated, label: str | None = None, op: str = "+",
            force_certain: bool = False, digits: int = 1, factor: int = 1):
        """후보 평가 → 게이트 통과 시 수집. cands = [(합, [(j,cc,sign)…]), …]."""
        cands = [(t, terms) for t, terms in cands if terms]
        if stated is None or not cands:
            return
        picked = next(((t, terms, True) for t, terms in cands if t == stated), None)
        if picked is None:
            t, terms = min(cands, key=lambda x: abs(Decimal(str(x[0])) - Decimal(str(stated))))
            picked = (t, terms, abs(Decimal(str(t)) - Decimal(str(stated))) <= C.FOOTING_TOL)
        computed, terms, ok = picked
        diff = Decimal(str(stated)) - Decimal(str(computed))
        if op != "+":
            ok = _mirror_pass(diff, digits)  # 수식 ROUND(d,n)이 흡수하는 만큼 인정
            if not ok:
                if op == "/" and not force_certain:
                    return           # 나눗셈 추측(유효세율)은 일치할 때만
                if op in ("-", "*") and abs(diff) > C.FOOTING_TOL:
                    return           # 합·차·곱 조합은 절사(±1)까지 방출 (원단위 절사 실측)
        elif (kind in GUESS_KINDS and not force_certain
              and abs(diff) > C.FOOTING_TOL):
            return  # 추측형은 빗나간 계층 추측을 수식화하지 않는다
        # 항에 대상 자신이 들어가는 퇴화 후보 방지
        if any(j == i and cc == c for j, cc, _ in terms):
            return
        self.raw.append(dict(tv=tv, kind=kind, axis=axis, i=i, c=c,
                             terms=terms, ok=ok, diff=diff, label=label, op=op,
                             digits=digits, factor=factor))

    def has(self, tv: TableView, i: int, c: int) -> bool:
        return any(r["tv"] is tv and r["i"] == i and r["c"] == c for r in self.raw)

    def finish(self) -> list[FootingSpec]:
        """대상 셀·축 단위 중복 제거(일치 우선 → kind 우선순위) 후 FootingSpec 변환.

        세로(아래)와 가로(오른쪽) 검증은 같은 셀에 공존한다 (유형자산 기말 합계 등).
        가로축 안에서도 흐름·취득원가 검증은 별개 관계라 공존(COEXIST_KINDS)."""
        best: dict[tuple, dict] = {}
        for r in self.raw:
            kind_slot = r["kind"] if r["kind"] in COEXIST_KINDS else ""
            key = (r["tv"].table.id, r["i"], r["c"], r["axis"], kind_slot)
            rank = (not r["ok"], KIND_PRIORITY.index(r["kind"]))
            cur = best.get(key)
            if cur is None or rank < (not cur["ok"], KIND_PRIORITY.index(cur["kind"])):
                best[key] = r
        specs = []
        for r in best.values():
            tv: TableView = r["tv"]
            specs.append(FootingSpec(
                table_id=tv.table.id,
                target=(r["i"], r["c"]),
                terms=r["terms"],
                label=r["label"] or _auto_label(tv, r),
                kind=r["kind"],
                axis=r["axis"],
                ok=r["ok"],
                diff=r["diff"],
                op=r["op"],
                digits=r.get("digits", 1),
                factor=r.get("factor", 1),
            ))
        return specs


def _auto_label(tv: TableView, r: dict) -> str:
    """검증 라벨: 재무제표 아래=정리된 행 라벨, 주석 세로=열머리글+행라벨,
    가로=행라벨+열머리글, 롤업=행라벨(+' 합계')."""
    i, c, kind = r["i"], r["c"], r["kind"]
    row = clean_label(tv.row_label(i))
    leaf = _despace(tv.leaf_header(c))
    if kind == "rollup":
        return f"{row} 합계" if len(r["terms"]) > 1 else row
    if r["axis"] == "row":
        return f"{row} {leaf}".strip()
    if (tv.sec_title in C.STATEMENT_KEYWORDS
            and not C.SOCIE_TITLE_RE.search(tv.sec_title)):
        return row
    return f"{leaf} {row}".strip() if leaf else row


# ── 열 합산 후보 (check.py 포팅) ─────────────────────────────────────

def _is_col_autosum(node) -> bool:
    return node is not None and node.attrs.get("AUTOSUM", "").startswith("COL")


def _row_is_subtotal(tv: TableView, j: int, c: int) -> bool:
    return (_is_col_autosum(tv.cell(j, c)) or tv.row_is_total(j)
            or bool(C.FLOW_END_RE.match(C.norm(tv.row_label(j)))))


# '채권 합계'처럼 접두 붙은 명명 소계 — TOTAL_RE(정확 일치)·GRAND_TOTAL_RE(총계$)에
# 안 걸려 위쪽 합산 워크가 지나치며 이중 합산된다 (NPP연결 주석10 실측)
_NAMED_TOTAL_RE = re.compile(r"(합계|총계|소계)$")


def _smart_col_terms(tv: TableView, i: int, c: int) -> list:
    """소계를 만나면 그 값을 항으로 넣고, 소계가 커버하는 구간은 건너뛰는 워크."""
    terms, skipping = [], False
    for j in reversed([b for b in tv.body if b < i]):
        if tv.is_section_row(j) and not tv.row_is_total(j):
            skipping = False
            continue
        v = tv.value(j, c)
        subtotal = _row_is_subtotal(tv, j, c)
        if subtotal:
            if v is not None:
                terms.append((v, j))
            skipping = True
            continue
        if v is not None and not skipping:
            terms.append((v, j))
    return terms


def _section_echo_terms(tv: TableView, i: int, c: int) -> list:
    """섹션 라벨과 같은 이름의 소계 행이 섹션 항목들을 대체하는 후보."""
    collected: list[tuple] = []
    current: tuple | None = None
    for j in tv.body:
        if j >= i:
            break
        if tv.is_section_row(j):
            core = C.norm(tv.row_label(j)).strip("<>:：")
            current = (core, len(collected))
            continue
        v = tv.value(j, c)
        if v is None:
            continue
        if current and C.norm(tv.row_label(j)) == current[0]:
            collected = collected[:current[1]]
            collected.append((v, j))
            current = None
        else:
            collected.append((v, j))
    return collected


def _same_key_terms(tv: TableView, i: int, c: int) -> list:
    """'구분' 열 소계: 같은 구분(첫 열 라벨) 블록 안의 위쪽 행만 합산."""
    key = C.norm(tv.text(i, 0))
    if not key:
        return []
    terms = []
    for j in reversed([b for b in tv.body if b < i]):
        if C.norm(tv.text(j, 0)) != key:
            break
        if tv.row_is_total(j) or _is_col_autosum(tv.cell(j, c)):
            break
        v = tv.value(j, c)
        if v is not None:
            terms.append((v, j))
    return terms


def _group_block_terms(tv: TableView, i: int, c: int) -> list:
    """'구분' 열 블록 단위: 블록마다 소계 행이 있으면 소계만, 없으면 항목 전부."""
    blocks: list[list[int]] = []
    prev_key = None
    for j in tv.body:
        if j >= i:
            break
        if tv.value(j, c) is None:
            continue
        key = C.norm(tv.text(j, 0))
        if key != prev_key or not blocks:
            blocks.append([])
            prev_key = key
        blocks[-1].append(j)
    terms = []
    for block in blocks:
        subtotal = [j for j in block if tv.row_is_total(j) or _is_col_autosum(tv.cell(j, c))]
        chosen = subtotal[-1:] if subtotal else block
        terms.extend((tv.value(j, c), j) for j in chosen)
    return terms


def _col_candidates(tv: TableView, i: int, c: int) -> tuple[list, int]:
    """반환: (후보 목록 [(합, [(행,열,부호)])], 비합계 기여 행 수)."""
    contiguous, no_subtotal, totals, same_depth, to_prev = [], [], [], [], []
    depth_i = tv.label_depth(i)
    contiguous_open = True
    autosum_seen = False
    for j in reversed([b for b in tv.body if b < i]):
        v = tv.value(j, c)
        if tv.is_section_row(j):
            contiguous_open = False
            continue
        if v is None:
            continue
        is_autosum = _is_col_autosum(tv.cell(j, c))
        total_row = _row_is_subtotal(tv, j, c)
        if is_autosum:
            autosum_seen = True
        if total_row:
            totals.append((v, j))
            contiguous_open = False
            continue
        if contiguous_open:
            contiguous.append((v, j))
            if tv.label_depth(j) == depth_i:
                same_depth.append((v, j))
        no_subtotal.append((v, j))
        if not autosum_seen:
            to_prev.append((v, j))

    def cand(pairs):
        return (sum(v for v, _ in pairs), [(j, c, 1) for _, j in pairs])

    # 직전 소계까지 구간(소계 + 이후 차감 항목): 차감후 잔액 = 소계 − 유동성 대체
    since_sub = []
    for j in reversed([b for b in tv.body if b < i]):
        if tv.is_section_row(j) and not tv.row_is_total(j):
            break
        v = tv.value(j, c)
        if v is not None:
            since_sub.append((v, j))
        if _row_is_subtotal(tv, j, c):
            break

    # 직전 '…합계/…총계/…소계' 명명 소계까지 (그 행 포함): 합계 = 명명 소계 + 이후 행
    since_named, named_hit = [], False
    for j in reversed([b for b in tv.body if b < i]):
        if tv.is_section_row(j) and not tv.row_is_total(j):
            break
        v = tv.value(j, c)
        if v is not None:
            since_named.append((v, j))
        if _row_is_subtotal(tv, j, c):
            break  # since_sub와 같은 종착 — 중복 후보라 버린다
        if _NAMED_TOTAL_RE.search(C.norm(tv.row_label(j))):
            named_hit = True
            break

    cands = [cand(_smart_col_terms(tv, i, c)), cand(contiguous), cand(no_subtotal),
             cand(same_depth), cand(to_prev), cand(_group_block_terms(tv, i, c)),
             cand(_same_key_terms(tv, i, c)), cand(_section_echo_terms(tv, i, c)),
             cand(since_sub)]
    if named_hit:
        cands.append(cand(since_named))
    # 합계류 행들의 접두 합 (부채및자본총계 = 자본총계+부채총계 패턴)
    for k in range(1, len(totals) + 1):
        cands.append(cand(totals[:k]))
    return cands, len(no_subtotal)


def _row_candidates(tv: TableView, i: int, c: int) -> list:
    contiguous, all_left = [], []
    stopped = False
    for k in range(c - 1, -1, -1):
        node = tv.cell(i, k)
        if node is not None and node.attrs.get("AUTOSUM", "").startswith("ROW"):
            break
        v = tv.value(i, k)
        if v is None:
            if node is not None and node.text.strip():
                stopped = True  # 라벨/비숫자 셀
            continue
        if not stopped:
            contiguous.append((v, k))
        all_left.append((v, k))

    def cand(pairs):
        return (sum(v for v, _ in pairs), [(i, k, 1) for _, k in pairs])

    return [cand(contiguous), cand(all_left)]


def _skip_section(tv: TableView) -> bool:
    return bool(C.SKIP_SECTION_RE.search(C.norm(tv.sec_title)))


# ── 개별 검사 (check.py C1·C2·C3·C7 포팅) ────────────────────────────

def check_autosum(views: list[TableView], col: Collector):
    for tv in views:
        if _skip_section(tv):
            continue
        for i in tv.body:
            for c in range(tv.ncols):
                node = tv.cell(i, c)
                if node is None or "AUTOSUM" not in node.attrs:
                    continue
                stated = tv.value(i, c)
                if stated is None:
                    continue
                kind = node.attrs["AUTOSUM"]
                if kind.startswith("COL"):
                    cands, _ = _col_candidates(tv, i, c)
                    col.add(tv, "autosum", "col", i, c, cands, stated)
                else:
                    col.add(tv, "autosum", "row", i, c, _row_candidates(tv, i, c), stated)


def _flat_outline_footing(tv: TableView, col: Collector) -> int:
    """들여쓰기 없는 K-GAAP 표의 로마자/괄호번호 계층 합계.

    ``Ⅰ. 비유동자산`` 아래 ``(1) 투자자산`` 같은 소그룹은 소계 열(C),
    소그룹 아래 세부 계정은 인접 금액 열(B)에 놓이는 2단 구조를 함께 시도한다.
    정확 일치(또는 기존 ±1 절사 게이트)일 때만 방출한다.
    """
    added = 0
    roman_re = re.compile(rf"^[{_ROMAN}]+\s*\.")
    sub_re = re.compile(r"^\(\d+\)")
    body = tv.body
    anchors = [(i, 0 if roman_re.match(tv.row_label(i)) else 1)
               for i in body if roman_re.match(tv.row_label(i))
               or sub_re.match(tv.row_label(i))]
    for k, (i, level) in enumerate(anchors):
        end = next((j for j, lv in anchors[k + 1:] if lv <= level), None)
        span = [j for j in body if j > i and (end is None or j < end)]
        if not span:
            continue
        direct = [j for j, lv in anchors[k + 1:]
                  if j in span and lv == level + 1]
        leaf = [j for j in span if not (roman_re.match(tv.row_label(j))
                                        or sub_re.match(tv.row_label(j)))]
        row_sets = [direct] if direct else []
        if level == 1 or not direct:
            row_sets.append(leaf)
        for c in range(tv.ncols):
            gv = tv.value(i, c)
            if gv is None:
                continue
            cands = []
            for rows in row_sets:
                for cc in ([c] if c == 0 else [c, c - 1]):
                    terms = [(j, cc, 1) for j in rows if tv.value(j, cc) is not None]
                    if len(terms) >= 2:
                        cands.append((sum(tv.value(j, cc) for j, cc, _ in terms), terms))
            if cands:
                before = len(col.raw)
                col.add(tv, "group", "col", i, c, cands, gv)
                added += len(col.raw) - before
    return added


def check_group_footing(views: list[TableView], col: Collector):
    """들여쓰기 계층 그룹 footing. 손익계산서는 'rollup'으로 방출(오른쪽 배치),
    자본변동표는 socie가 담당하므로 제외 (check.py 실측 규칙)."""
    for tv in views:
        if _skip_section(tv):
            continue
        if C.SOCIE_TITLE_RE.search(C.norm(tv.sec_title)):
            continue
        is_pl = ("손익계산서" in C.norm(tv.sec_title)
                 or (tv.note_no is not None and "손익계산서" in C.norm(tv.sec_title)))
        kind = "rollup" if is_pl else "group"
        depths = {tv.label_depth(i) for i in tv.body if tv.row_label(i)}
        if len(depths) < 2:
            _flat_outline_footing(tv, col)
            continue
        body = tv.body
        for c in range(tv.ncols):
            for gi, i in enumerate(body):
                gv = tv.value(i, c)
                if gv is None:
                    continue
                # '-'(0) 대상도 검증한다 (실측: BS 비유동부채 0도 방출) —
                # 정확 일치 게이트가 있어 빗나간 그룹 추측은 어차피 걸러진다
                d = tv.label_depth(i)
                span = []
                for j in body[gi + 1:]:
                    if not tv.row_label(j):
                        continue
                    if tv.label_depth(j) <= d:
                        break
                    span.append(j)
                if not span:
                    continue
                # 직계 자식 = 구간 내 최소 깊이 행만 (중첩 그룹 이중합산 방지)
                min_d = min(tv.label_depth(j) for j in span)
                span_rows = [j for j in span if tv.label_depth(j) == min_d]
                # 차감계정 전용 자식(대손충당금·연금운용자산 등)뿐인 그룹 머리 —
                # '총액 = 충당금' 오인 방출 방지 (농협 BS 3행 차감 표시 실측).
                # 부모가 유출/차감류(부호 반전 후보 있음)면 정상 그룹이므로 제외
                if (span_rows
                        and all(_contra_child(tv, j, gv > 0) for j in span_rows)
                        and not C.CONTRA_RE.search(C.norm(tv.row_label(i)))
                        and not re.search(r"(유출|차감|감소|결손)",
                                          C.norm(tv.row_label(i)))):
                    continue
                # K-GAAP 2단 금액 열: 그룹 소계는 오른쪽 열, 구성 항목은 왼쪽 열
                cands = []
                reversed_group = bool(
                    re.search(r"(유출|차감|감소|결손)", C.norm(tv.row_label(i))))
                for cc in ([c] if c == 0 else [c, c - 1]):
                    ch = [j for j in span_rows if tv.value(j, cc) is not None]
                    if not ch:
                        continue
                    total = sum(tv.value(j, cc) for j in ch)
                    cands.append((total, [(j, cc, 1) for j in ch]))
                    # 현금흐름표 '…유출액/차감' 그룹: 그룹값 음수(괄호), 하위 양수 나열.
                    # '결손금'은 반대(그룹 양수 표기, 하위 괄호 음수) — 부호 반전 후보
                    if reversed_group:
                        cands.append((-total, [(j, cc, -1) for j in ch]))
                    # 깊이 무시 연속 평면합 — 차감(음수) 자식 행 포함 후보.
                    # 총액/(충당금) 차감 표시는 충당금이 한 단계 더 깊어 직계 합산에서
                    # 빠진다. 기준 양식은 구간 전 행 평면합 (농협/은행 BS 실측)
                    ch_all = [j for j in span if tv.value(j, cc) is not None]
                    if ch_all != ch:
                        total_all = sum(tv.value(j, cc) for j in ch_all)
                        cands.append((total_all, [(j, cc, 1) for j in ch_all]))
                        if reversed_group:
                            cands.append((-total_all, [(j, cc, -1) for j in ch_all]))
                # 재무상태표 본문의 들여쓰기 계층(ALEVEL·표준 들여쓰기)은 신뢰할 수 있어
                # 불일치를 빨간 수식으로 노출한다 — 주석 표는 정확 일치만(오탐 방지).
                # 예외 (실측 오탐 대응):
                #  - '-'(0 플레이스홀더) 대상: 그룹 머리가 아닐 수 있음(IFRS CF
                #    '종속기업투자주식의 취득') → 합계가 실제 0일 때만
                #  - 현금흐름표: 원본 ALEVEL 오지정(leaf가 depth 0)이 실존 →
                #    정확 일치만 (오류는 라벨힌트 Ⅳ=Ⅰ+Ⅱ+Ⅲ 등 확실형이 잡는다)
                is_dash = (tv.cell(i, c) is not None
                           and tv.cell(i, c).text.strip() == "-")
                col.add(tv, kind, "col", i, c, cands, gv,
                        force_certain=(tv.sec_title in C.STATEMENT_KEYWORDS
                                       and kind == "group" and not is_dash
                                       and "현금흐름표" not in C.norm(tv.sec_title)))


def _contra_child(tv: TableView, j: int, parent_positive: bool) -> bool:
    """차감계정 자식 판정 — 라벨(충당금·누계액 등) 또는 [부모가 양수일 때] 행 전체 음수.

    값 기준이 필요한 이유: '퇴직연급운용자산'(오타 포함 실측)처럼 라벨 사전에 없는
    차감 표시 행이 있다 — 총액(+) 부모 밑의 전량 음수 자식은 소계 관계일 수 없다.
    부모가 음수면 음수 자식과 정당한 소계일 수 있어(주당손익=주당순손실) 값 기준 제외."""
    if C.CONTRA_RE.search(C.norm(tv.row_label(j))):
        return True
    if not parent_positive:
        return False
    vals = [tv.value(j, c) for c in range(tv.ncols)]
    vals = [v for v in vals if v is not None]
    return bool(vals) and all(v < 0 for v in vals)


def _rate_cols(tv: TableView) -> set[int]:
    """세로 합산이 무의미한 요율 열(이자율 등) — AI검토가 발견한 오탐 대응."""
    hs = _headers_of(tv)
    return {c for c, h in enumerate(hs) if C.RATE_COL_RE.search(h)}


def _depth_diverse(tv: TableView) -> bool:
    depths = {tv.label_depth(i) for i in tv.body if tv.row_label(i)}
    return len(depths) >= 2


def _downward_terms(tv: TableView, i: int, cc: int, target_c: int) -> list:
    """총계-먼저형: 대상 행 '아래쪽'의 연속 세부 행들 (구현 후보 ③).

    소계가 위에 오고 '- 세부항목'이 뒤따르는 양식(포괄손익 명세 등).
    다른 소계·섹션 행을 만나거나, 세부 열 순회 중 대상 열에 값이 다시 나타나면
    (다음 블록 시작) 멈춘다."""
    terms = []
    for j in tv.body:
        if j <= i:
            continue
        if tv.is_section_row(j) or tv.row_is_total(j):
            break
        if cc != target_c and tv.value(j, target_c) is not None:
            break  # 대상 열에 값 등장 — 다음 소계 블록
        v = tv.value(j, cc)
        if v is None:
            if terms:
                break  # 연속 구간 종료
            continue
        terms.append((v, j))
    return terms


def check_total_rows(views: list[TableView], col: Collector):
    for tv in views:
        if _skip_section(tv):
            continue
        rate_cols = _rate_cols(tv)
        for i in tv.body:
            labels = [C.norm(tv.text(i, c)) for c in range(tv.ncols)
                      if tv.text(i, c).strip() and tv.value(i, c) is None]
            exact_total = any(C.TOTAL_RE.match(t) for t in labels)
            grand_total = (not exact_total
                           and any(C.GRAND_TOTAL_RE.search(t) for t in labels))
            deduct_total = (not exact_total and not grand_total
                            and any(C.DEDUCT_TOTAL_RE.search(t) for t in labels))
            if not (exact_total or grand_total or deduct_total):
                continue
            if grand_total and not _depth_diverse(tv):
                continue
            for c in range(tv.ncols):
                node = tv.cell(i, c)
                if node is None or "AUTOSUM" in node.attrs:  # autosum 검사가 담당
                    continue
                stated = tv.value(i, c)
                if stated is None or not node.text.strip() or c in rate_cols:
                    continue
                cands, contributing = _col_candidates(tv, i, c)
                cols_here = [c] if c == 0 else (
                    [c, c - 1] if tv.period_bucket(c - 1) == tv.period_bucket(c) else [c])
                if len(cols_here) > 1:
                    cands_left, contrib_left = _col_candidates(tv, i, c - 1)
                    cands += cands_left
                    contributing = max(contributing, contrib_left)
                # 아래 방향 후보 (총계-먼저형 — 구현 후보 ③)
                down = []
                for cc in cols_here:
                    pairs = _downward_terms(tv, i, cc, c)
                    if len(pairs) >= 2:
                        down.append((sum(v for v, _ in pairs),
                                     [(j, cc, 1) for _, j in pairs]))
                cands += down
                if contributing < 2:
                    # 위쪽 합산 대상이 부족한 표 — '정확히' 맞을 때만 방출:
                    # 아래 방향 합 또는 단일 구성항 소계(구현 후보 ⑤ — 소계=유일 항)
                    exact = next((cd for cd in down if cd[0] == stated), None)
                    if exact is None and contributing == 1:
                        exact = next((cd for cd in cands
                                      if cd[0] == stated and len(cd[1]) == 1), None)
                        # 차감계정 단독항(총액=충당금 우연 일치)은 소계 관계가 아니다
                        if exact is not None:
                            j0 = exact[1][0][0]
                            if (C.CONTRA_RE.search(C.norm(tv.row_label(j0)))
                                    and not C.CONTRA_RE.search(C.norm(tv.row_label(i)))):
                                exact = None
                    if exact is None:
                        continue
                    col.add(tv, "total_row", "col", i, c, [exact], stated)
                    continue
                col.add(tv, "total_row", "col", i, c, cands, stated)


# 최종 합계 열 (연결 CE '총자본' — 소계 열들 위의 상위 합계, 기준 양식은 총자본=소계+비지배)
_GRAND_COL_RE = re.compile(r"^(총자본|자본총계)$")
# 상위 헤더 그룹 안의 '소계' 열 (연결 CE '지배기업 소유주지분 > 소계' — 삼보 반기연결 실측).
# 소계 = 그룹 구성 열의 합, 상위 '합계' 열 = 소계 열들 + 그룹에 안 덮인 열(비지배지분).
_SUBTOTAL_HEADER_RE = re.compile(r"^소\s*계$")


def check_total_columns(views: list[TableView], col: Collector):
    for tv in views:
        if _skip_section(tv):
            continue
        # 수기 표(THEAD 없음)도 _headers_of 폴백으로 합계 열을 찾는다 (의료재단 실측)
        hs = _headers_of(tv)
        if not any(hs):
            continue
        total_cols = [c for c in range(tv.ncols)
                      if C.TOTAL_HEADER_RE.match(hs[c])]
        grand_cols = [c for c in range(tv.ncols)
                      if c not in total_cols and _GRAND_COL_RE.match(hs[c])]
        # 그룹 소계 열: '소계' 머리글이 상위 그룹(span>1) 안에 있을 때만 (단독 '소계' 열은
        # 합계 열과 같은 취급을 받을 이유가 없어 제외 — 기존 동작 유지)
        sub_cols = [c for c in range(tv.ncols)
                    if c not in total_cols and c not in grand_cols
                    and _SUBTOTAL_HEADER_RE.match(hs[c])
                    and len(tv.header_group_cols(c)) > 1]
        # 중첩 열 집합: [소계 열들] + [소계 열 그룹에 안 덮인 열] — 총자본=지배소계+비지배
        covered = set()
        for t2 in total_cols + sub_cols:
            covered |= set(tv.header_group_cols(t2) or [])
        for c in total_cols + grand_cols + sub_cols:
            group = tv.header_group_cols(c)
            if group and c not in grand_cols:
                siblings = [cc for cc in group if cc != c
                            and not C.TOTAL_HEADER_RE.match(hs[cc])
                            and cc not in sub_cols]
            else:
                siblings = [cc for cc in range(tv.ncols)
                            if cc != c and cc not in grand_cols
                            and not C.TOTAL_HEADER_RE.match(hs[cc])
                            and cc not in sub_cols
                            and "주석" not in hs[cc]]
            if c in sub_cols:
                nested_cols = []  # 그룹 소계는 자기 그룹 구성 열의 평면합만
            else:
                nested_cols = ([t2 for t2 in total_cols if t2 != c]
                               + [t2 for t2 in sub_cols if t2 != c]
                               + [cc for cc in range(tv.ncols)
                                  if cc != c and cc not in covered
                                  and cc not in total_cols and cc not in grand_cols
                                  and cc not in sub_cols
                                  and "주석" not in hs[cc]])
            if not siblings and not nested_cols:
                continue
            # 그룹 소계 열이 있는 표의 상위 합계 열은 중첩(소계+비지배)이 정식 관계
            prefer_nested = c in grand_cols or (bool(sub_cols) and c not in sub_cols)
            for i in tv.body:
                stated = tv.value(i, c)
                node = tv.cell(i, c)
                if stated is None or node is None:
                    continue
                if "AUTOSUM" in node.attrs and node.attrs["AUTOSUM"].startswith("COL"):
                    continue
                if C.NONSUM_COL_RE.search(C.norm(tv.row_label(i))):
                    continue  # 이자율 등 비율 행은 가로합 비대상
                terms_all = [(i, cc, 1) for cc in siblings if tv.value(i, cc) is not None]
                terms_nested = [(i, cc, 1) for cc in nested_cols
                                if tv.value(i, cc) is not None]
                flat = []
                if len(terms_all) >= 2:
                    flat.append((sum(tv.value(i, cc) for _, cc, _ in terms_all), terms_all))
                if len(terms_all) >= 3:
                    # 만기분석표(장부금액|만기…|합계): 첫 값 열 제외 후보
                    tail = terms_all[1:]
                    flat.append((sum(tv.value(i, cc) for _, cc, _ in tail), tail))
                nested = []
                if len(terms_nested) >= 2 and terms_nested != terms_all:
                    nested.append((sum(tv.value(i, cc) for _, cc, _ in terms_nested),
                                   terms_nested))
                # 그랜드 열은 중첩(소계+비지배)이 정식 관계라 우선, 합계 열은 기존 우선
                cands = (nested + flat) if prefer_nested else (flat + nested)
                col.add(tv, "total_col", "row", i, c, cands, stated)


# ── 재무제표 계산 라인 (C7 포팅) ─────────────────────────────────────

def _pl_sign(label: str) -> int:
    return -1 if C.PL_MINUS_RE.search(C.norm(label)) else 1


_ROMAN_PREFIX_ONLY = re.compile(rf"^[{_ROMAN}]+\s*[.)]")
# 로마자 섹션의 실질 부호 — 어미 고정이 핵심: '법인세비용차감전순이익'은 순이익$ → +,
# '계속사업손익법인세비용'은 비용$ → − (_pl_sign의 anywhere-search로는 구분 불가)
_ROMAN_MINUS_RE = re.compile(r"(비용|원가|손실|판매비|법인세등?|전입액?)$")


def _roman_sign(label: str) -> int:
    return -1 if _ROMAN_MINUS_RE.search(_label_core(label)) else 1


def _check_pl_anchors(tv: TableView, col: Collector):
    body = tv.body
    anchors = [
        i for i in body
        if C.ANCHOR_RE.search(_label_core(tv.row_label(i)))
        and "주당" not in C.norm(tv.row_label(i))
        and any(tv.value(i, c) is not None for c in range(tv.ncols))
    ]
    for ai, i in enumerate(anchors):
        prev = anchors[ai - 1] if ai else None
        between = [j for j in body if (prev if prev is not None else -1) < j < i
                   and tv.row_label(j)]
        with_val = [j for j in between
                    if any(tv.value(j, c) is not None for c in range(tv.ncols))]
        # 기타포괄손익 항목이 없는 문서: '총포괄이익 = 순이익' 동액 이월 (반기검토보고서
        # 실측 — 기준 양식도 C19-(C18) 단일항으로 방출). 직전 앵커가 순이익류일 때만.
        carry_only = (not with_val and prev is not None
                      and "총포괄" in _label_core(tv.row_label(i))
                      and re.search(r"순(이익|손실|손익)$", _label_core(tv.row_label(prev))))
        if not with_val and not carry_only:
            continue
        if carry_only:
            for c in range(tv.ncols):
                stated = tv.value(i, c)
                pv = tv.value(prev, c)
                if stated is None or pv is None:
                    continue
                if not tv.cell(i, c) or not tv.cell(i, c).text.strip():
                    continue
                col.add(tv, "pl_anchor", "col", i, c, [(pv, [(prev, c, 1)])], stated,
                        label=clean_label(tv.row_label(i)))
            continue
        min_d = min(tv.label_depth(j) for j in with_val)
        comps = [j for j in with_val if tv.label_depth(j) == min_d]
        for c in range(tv.ncols):
            stated = tv.value(i, c)
            if stated is None or not tv.cell(i, c) or not tv.cell(i, c).text.strip():
                continue
            base = ([(prev, c, 1)] if prev is not None and tv.value(prev, c) is not None
                    else [])
            bases = [base]
            # 손실 양수 표기 문서: 직전 앵커의 실질 부호가 음수인 변형
            if (prev is not None and base
                    and "손실" in _label_core(tv.row_label(prev)) and base[0][2] == 1
                    and tv.value(prev, c) is not None and tv.value(prev, c) > 0):
                bases.append([(prev, c, -1)])
            cols = [c] if c == 0 else (
                [c, c - 1] if tv.period_bucket(c - 1) == tv.period_bucket(c) else [c]
            )
            cands = []
            for b in bases:
                for cc in cols:
                    comp_rows = [j for j in comps if tv.value(j, cc) is not None]
                    signed = b + [(j, cc, _pl_sign(tv.row_label(j))) for j in comp_rows]
                    plain = b + [(j, cc, 1) for j in comp_rows]
                    for terms in (signed, plain):
                        if len(terms) >= 2:
                            cands.append(
                                (sum(s * tv.value(j, c2) for j, c2, s in terms), terms)
                            )
            # 로마자 섹션 체인 후보 (후미 append — 첫 정확 일치 규칙이 기존 후보 우선):
            # 표준 앵커 체인이 비표준 단계 구조에서 빗나갈 때 대비. 의료 PL
            # 'Ⅵ=Ⅲ+Ⅳ−Ⅴ'(Ⅱ.의료비용에 합계 금액이 없는 양식), 조합 PL 'ⅩⅡ=ⅩⅠ'
            # (동액 이월·ⅩⅠ이 앵커 라벨이 아닌 양식) — 신규 샘플 실측.
            roman_above = [j for j in body if j < i
                           and _ROMAN_PREFIX_ONLY.match(tv.row_label(j).strip())]
            for cc in cols:
                rv = [j for j in roman_above if tv.value(j, cc) is not None]
                for k in range(1, min(len(rv), 5) + 1):
                    tail = rv[-k:]
                    signed = [(j, cc, _roman_sign(tv.row_label(j))) for j in tail]
                    plain = [(j, cc, 1) for j in tail]
                    for terms in (signed, plain):
                        cands.append(
                            (sum(s * tv.value(j, c2) for j, c2, s in terms), terms)
                        )
            if not cands:
                continue
            # 손실 양수 표기: 기재 절대값 비교 변형 — 부호 반전 후보 추가
            if "손실" in _label_core(tv.row_label(i)) and stated > 0:
                cands = cands + [(-t, [(j, c2, -s) for j, c2, s in terms])
                                 for t, terms in cands]
            col.add(tv, "pl_anchor", "col", i, c, cands, stated,
                    label=clean_label(tv.row_label(i)))


# 한글 순서 기호 힌트: '마. 법인세비용 (가+나+다+라)' (구현 후보 ②)
_HINT_KR_RE = re.compile(r"\(([가-하]\s*(?:\+\s*[가-하]\s*)+)\)$")
_KR_PREFIX_RE = re.compile(r"^([가-하])\s*[.)]")


def _check_hint_lines(tv: TableView, col: Collector):
    """라벨에 합산식이 명시된 행 검산 — 로마자 '(Ⅰ+Ⅱ+Ⅲ)'와 한글 '(가+나+다+라)'."""
    body = tv.body
    by_key = {}
    for j in body:
        m = re.match(rf"^([{_ROMAN}]+)\s*\.", tv.row_label(j))
        if m and m.group(1) not in by_key:
            by_key[m.group(1)] = j
        mk = _KR_PREFIX_RE.match(tv.row_label(j).strip())
        if mk and mk.group(1) not in by_key:
            by_key[mk.group(1)] = j
    for i in body:
        label = C.norm(tv.row_label(i))
        m = _HINT_RE.search(label) or _HINT_KR_RE.search(label)
        if not m:
            continue
        keys = [r.strip() for r in m.group(1).split("+")]
        rows = [by_key.get(r) for r in keys]
        # 참조 행이 하나라도 없으면 검산하지 않는다 (엄격 — 확실형이므로).
        # 행은 있되 값이 빈 경우는 아래 terms 필터가 자연히 제외한다.
        if any(r is None or r == i for r in rows):
            continue
        for c in range(tv.ncols):
            stated = tv.value(i, c)
            if stated is None:
                continue
            terms = [(j, c, 1) for j in rows if tv.value(j, c) is not None]
            if len(terms) < 2:
                continue
            total = sum(tv.value(j, cc) for j, cc, _ in terms)
            col.add(tv, "hint", "col", i, c, [(total, terms)], stated)


def _flow_open(label: str) -> bool:
    return "기초" in label and not C.INVENTORY_RE.match(label)


def _flow_close(label: str) -> bool:
    return "기말" in label and not C.INVENTORY_RE.match(label)


def _socie_subtotal_variants(tv: TableView, between: list[int]) -> list[list[int]]:
    """기초~기말 사이 '소계' 행이 있으면 [소계로 구성 행 대체, 소계 제외] 두 행 목록.
    소계가 대체하는 구간은 직전 섹션 행 또는 직전 소계 다음부터."""
    def is_sub(j: int) -> bool:
        return (any(tv.value(j, cc) is not None for cc in range(tv.ncols))
                and (tv.row_is_total(j) or re.search(r"소계$", C.norm(tv.row_label(j))) is not None))

    if not any(is_sub(j) for j in between):
        return []
    collapsed: list[int] = []
    seg_start = 0
    for j in between:
        if tv.is_section_row(j):
            seg_start = len(collapsed)
        elif is_sub(j):
            collapsed = collapsed[:seg_start] + [j]
            seg_start = len(collapsed)
            continue
        collapsed.append(j)
    return [collapsed, [j for j in between if not is_sub(j)]]


def _check_socie(tv: TableView, col: Collector):
    """기초→변동→기말 세로 흐름: 기말 행 = 기초 행 + 사이 변동 행들의 합 (열별).
    현금흐름표 본표는 제외 — 기말 현금 = 순증감+기초라 구조가 다름."""
    if "현금흐름표" in C.norm(tv.sec_title):
        return
    opening = None
    between: list[int] = []
    for i in tv.body:
        label = C.norm(tv.row_label(i))
        has_value = any(tv.value(i, c) is not None for c in range(tv.ncols))
        if _flow_open(label) and has_value:
            opening, between = i, []
            continue
        if opening is None:
            continue
        if _flow_close(label) and has_value:
            # 섹션마다 소계 유무가 다르다 — 소계가 있는 섹션은 소계만, 없으면 전부
            smart_rows: list[int] = []
            section: list[int] = []

            def flush_section():
                valued = [j for j in section
                          if any(tv.value(j, cc) is not None for cc in range(tv.ncols))]
                if not valued:
                    return
                min_d = min(tv.label_depth(j) for j in valued)
                top = [j for j in valued if tv.label_depth(j) == min_d]
                smart_rows.extend(top if len(top) < len(valued) else valued)

            for j in between:
                if tv.is_section_row(j):
                    flush_section()
                    section = []
                else:
                    section.append(j)
            flush_section()

            # 연결 양식: 섹션 라벨과 같은 이름의 '… 합계' 재합계 행이 대체하는 후보
            replaced: list[int] = []
            stack: list[tuple[str, int]] = []
            for j in between:
                if tv.is_section_row(j):
                    stack.append((_label_core(tv.row_label(j)).rstrip(":"), len(replaced)))
                    continue
                if not any(tv.value(j, cc) is not None for cc in range(tv.ncols)):
                    continue
                if tv.row_is_total(j) or re.search(r"(합계|총계)$", C.norm(tv.row_label(j))):
                    core = re.sub(r"(합\s*계|총\s*계)$", "", C.norm(tv.row_label(j))).strip()
                    hit = next((k for k in range(len(stack) - 1, -1, -1)
                                if stack[k][0] == core), None)
                    if hit is not None:
                        replaced = replaced[:stack[hit][1]]
                        stack = stack[:hit]
                replaced.append(j)

            candidates_rows = [smart_rows]
            if replaced != smart_rows:
                candidates_rows.append(replaced)
            # 같은 깊이의 '소계' 행(총포괄손익 소계) — 구성 행과 소계를 둘 다 더하면
            # 차이 = −소계로 빗나간다(분기 DSD 실측). 소계만 / 구성 행만 두 후보를 앞에 둔다
            # (소계가 0인 열에서도 이중 합산 식이 먼저 뽑히지 않게).
            for rows in reversed(_socie_subtotal_variants(tv, between)):
                if rows in candidates_rows:
                    candidates_rows.remove(rows)
                candidates_rows.insert(0, rows)
            if between not in candidates_rows:
                candidates_rows.append(between)
            all_valued = [j for j in between if any(
                tv.value(j, cc) is not None for cc in range(tv.ncols))]
            for c in range(tv.ncols):
                stated = tv.value(i, c)
                if stated is None or tv.value(opening, c) is None:
                    continue
                if C.NONSUM_COL_RE.search(C.norm(tv.col_header(c))):
                    continue
                cands = []
                item_cands = []
                for k, rows in enumerate(candidates_rows):
                    valued_rows = [j for j in rows if tv.value(j, c) is not None]
                    plain = [(opening, c, 1)] + [(j, c, 1) for j in valued_rows]
                    signed = [(opening, c, 1)] + [
                        (j, c, -1 if C.FLOW_MINUS_RE.search(C.norm(tv.row_label(j))) else 1)
                        for j in valued_rows
                    ]
                    for terms in (plain, signed):
                        cands.append(
                            (sum(s * tv.value(j, cc) for j, cc, s in terms), terms)
                        )
                        if rows is between:
                            item_cands.append(cands[-1])
                if all(len(t) < 2 for _, t in cands):
                    continue
                col.add(tv, "socie", "col", i, c, cands, stated)
                # 잔액행 이중 변형 (실측): 아래에는 소계 대체본, 오른쪽에는 항목
                # 전체본 — 두 변형의 항이 실제로 다르고 항목본도 정확 일치할 때만
                exact_item = next((cd for cd in item_cands if cd[0] == stated), None)
                if exact_item is not None and set(all_valued) != set(smart_rows):
                    col.add(tv, "socie", "row", i, c, [exact_item], stated)
            opening, between = None, []
            continue
        between.append(i)


def _check_section_subtotal(tv: TableView, col: Collector):
    """섹션 소계행: '<섹션라벨> 합계' 행 = 섹션 라벨 행 ~ 소계행 사이 항목들의 합.

    IFRS 자본변동표('총 포괄손익 합계', '소유주와의 거래 합계') 실측 — 기준 양식은
    오른쪽·아래에 이중 방출한다."""
    section_at: int | None = None
    section_core = ""
    items: list[int] = []
    for i in tv.body:
        label = C.norm(tv.row_label(i))
        if tv.is_section_row(i):
            section_at, section_core, items = i, _label_core(tv.row_label(i)).rstrip(":"), []
            continue
        if section_at is not None and re.search(r"(합계|총계)$", label):
            core = re.sub(r"(합\s*계|총\s*계)$", "", label).strip()
            if core and core.replace(" ", "") == section_core.replace(" ", ""):
                for c in range(tv.ncols):
                    stated = tv.value(i, c)
                    if stated is None:
                        continue
                    terms = [(j, c, 1) for j in items if tv.value(j, c) is not None]
                    if not terms:
                        continue
                    total = sum(tv.value(j, cc) for j, cc, _ in terms)
                    col.add(tv, "socie_sub", "col", i, c, [(total, terms)], stated)
                    col.add(tv, "socie_sub", "row", i, c, [(total, terms)], stated)
                section_at, items = None, []
                continue
        if section_at is not None:
            items.append(i)


def _check_eq_carry(tv: TableView, col: Collector):
    """무변동 이월 (구현 후보 ⑤): 인접한 기초 행과 기말 행이 같은 금액이면 이월 확인.

    자본금·주식수처럼 변동이 없는 항목의 '전기초=전기말' 확인 — 값이 같을 때만
    방출한다 (사이에 변동 행이 있으면 socie가 담당하므로 여기서는 인접 쌍만)."""
    pairs = {"기초": "기말", "당기초": "당기말", "전기초": "전기말"}
    body = tv.body
    for k in range(len(body) - 1):
        j, i = body[k], body[k + 1]
        lj = _label_core(tv.row_label(j))
        li = _label_core(tv.row_label(i))
        if C.INVENTORY_RE.match(C.norm(tv.row_label(j))):
            continue
        matched = next((o for o, cl in pairs.items()
                        if lj.endswith(o) and li.endswith(cl)
                        and lj[:-len(o)] == li[:-len(cl)]), None)
        if matched is None:
            continue
        for c in range(tv.ncols):
            a, b = tv.value(j, c), tv.value(i, c)
            if a is None or b is None or a != b:
                continue
            col.add(tv, "eq_carry", "col", i, c, [(a, [(j, c, 1)])], b)


_BULLET_RE = re.compile(r"^[-ㆍ·•]")


def _check_bullet_subtotal(tv: TableView, col: Collector):
    """총계-먼저형 소계 (구현 후보 ③): '- 세부항목' 불릿 행들이 뒤따르는 행 =
    그 불릿 행들의 합. 합계 라벨이 없어 기존 트리거가 못 잡는 양식(포괄손익 명세 등).

    구조 추정이므로 정확 일치할 때만 방출한다."""
    body = tv.body
    for k, i in enumerate(body):
        if k + 1 >= len(body):
            break
        nxt = body[k + 1]
        if not _BULLET_RE.match(tv.row_label(nxt).strip()):
            continue
        # 불릿 연속 구간
        run = []
        for j in body[k + 1:]:
            if _BULLET_RE.match(tv.row_label(j).strip()):
                run.append(j)
            else:
                break
        if len(run) < 2:
            continue
        for c in range(tv.ncols):
            stated = tv.value(i, c)
            if stated is None:
                continue
            cols_here = [c] if c == 0 else (
                [c, c - 1] if tv.period_bucket(c - 1) == tv.period_bucket(c) else [c])
            cands = []
            for cc in cols_here:
                terms = [(j, cc, 1) for j in run if tv.value(j, cc) is not None]
                if len(terms) >= 2:
                    cands.append((sum(tv.value(j, c2) for j, c2, _ in terms), terms))
            exact = next((cd for cd in cands if cd[0] == stated), None)
            if exact is not None:
                col.add(tv, "total_row", "col", i, c, [exact], stated)


def _check_summary_fin(tv: TableView, col: Collector):
    """피투자회사 요약재무: 자본 = 자산 − 부채 (구현 후보 ④).

    열 머리글에 자산·부채·자본이 모두 있을 때만, 행별 가로 산술.
    비지배지분 표시 등 예외 대비 정확 일치만 방출."""
    if tv.note_no is None:
        return
    headers = _headers_of(tv)
    asset_c = next((c for c, h in enumerate(headers)
                    if re.fullmatch(r"자산(총계|총액)?", h)), None)
    liab_c = next((c for c, h in enumerate(headers)
                   if re.fullmatch(r"부채(총계|총액)?", h)), None)
    eq_c = next((c for c, h in enumerate(headers)
                 if re.fullmatch(r"자본(총계|총액)?|순자산", h)), None)
    if asset_c is None or liab_c is None or eq_c is None:
        return
    for i in tv.body:
        stated = tv.value(i, eq_c)
        a, l = tv.value(i, asset_c), tv.value(i, liab_c)
        if stated is None or a is None or l is None:
            continue
        terms = [(i, asset_c, 1), (i, liab_c, -1)]
        col.add(tv, "summary_fin", "row", i, eq_c, [(a - l, terms)], stated)


def _check_net_pair(tv: TableView, col: Collector):
    """이연법인세자산(부채)순액 = [자산 계행] − [부채 계행] (비인접 행 짝).

    소계 워크(위쪽 합산 후보)로는 못 잡는 라벨 짝 관계 — 정확 일치만 방출.
    일치 시 finish의 일치-우선 랭크가 같은 셀의 total_row 최근접 오답을 대체한다."""
    if tv.note_no is None:
        return
    for i in tv.body:
        n = C.norm(tv.row_label(i))
        if not n.endswith("순액") or "자산" not in n or "부채" not in n:
            continue
        above = [j for j in tv.body if j < i]
        ja = next((j for j in reversed(above)
                   if "이연법인세자산" in C.norm(tv.row_label(j))
                   and not C.norm(tv.row_label(j)).endswith("순액")), None)
        jl = next((j for j in reversed(above)
                   if "이연법인세부채" in C.norm(tv.row_label(j))
                   and not C.norm(tv.row_label(j)).endswith("순액")), None)
        if ja is None or jl is None or ja == jl:
            continue
        for c in range(tv.ncols):
            stated = tv.value(i, c)
            a, l = tv.value(ja, c), tv.value(jl, c)
            if stated is None or a is None or l is None:
                continue
            cands = [(a - l, [(ja, c, 1), (jl, c, -1)]),   # 부채 양수 표기
                     (a + l, [(ja, c, 1), (jl, c, 1)])]    # 부채 음수 저장
            col.add(tv, "net_pair", "col", i, c, cands, stated)


_CF_ACT_RE = re.compile(r"(영업|투자|재무)활동으?로?\s*인한\s*현금흐름")
_CF_NET_RE = re.compile(r"현금(및현금성자산)?의?\s*순?(증가|감소)")
_CF_FX_RE = re.compile(r"환율변동")
_CF_SUB_RE = re.compile(r"^\(\d+\)")


def _check_cf_anchors(tv: TableView, col: Collector):
    """현금흐름표 앵커 캐스케이드 — 로마자 힌트가 없는 양식(IFRS) 대응.

    - 순증가(감소) = 영업+투자+재무 활동 현금흐름 (+환율변동효과)
    - 활동별 합계 = 앵커 사이 구성 행들의 합 (전체 행 / '(n)' 소계 행 두 후보)
    앵커 라벨이 명확해 확실형 — 불일치는 빨간 차이로 노출한다."""
    if "현금흐름표" not in C.norm(tv.sec_title):
        return
    body = tv.body
    acts, nets, fx = [], [], []
    for i in body:
        core = _label_core(tv.row_label(i))
        if _CF_ACT_RE.search(core):
            acts.append(i)
        elif _CF_NET_RE.search(core) and "기초" not in core and "기말" not in core:
            nets.append(i)
        elif _CF_FX_RE.search(core):
            fx.append(i)
    if not acts:
        return
    boundaries = sorted(acts + nets)
    for c in range(tv.ncols):
        # K-GAAP 2단 금액 열: 합계는 오른쪽 열, 구성 소계는 바로 왼쪽 열 (구현 후보 ①)
        cols = [c] if c == 0 else (
            [c, c - 1] if tv.period_bucket(c - 1) == tv.period_bucket(c) else [c])
        # 활동별 합계 = 다음 앵커 전까지의 구성 행 합
        for k, i in enumerate(acts):
            stated = tv.value(i, c)
            if stated is None:
                continue
            nxt = next((b for b in boundaries if b > i), None)
            span = [j for j in body if i < j < (nxt if nxt is not None else 10 ** 9)]
            cands = []
            for cc in cols:
                valued = [j for j in span if tv.value(j, cc) is not None]
                if not valued:
                    continue
                subs = [j for j in valued
                        if _CF_SUB_RE.match(tv.row_label(j).strip()) or tv.row_is_total(j)]
                plain = [j for j in valued if j not in subs]
                # 후보: '(n)' 소계 행만 / 소계 제외 전체 / 전체 — 정확 일치를 _pick이 고른다
                for rows in (subs, plain, valued):
                    if not rows:
                        continue
                    terms = [(j, cc, 1) for j in rows]
                    cands.append((sum(tv.value(j, cc) for j in rows), terms))
            if cands:
                col.add(tv, "cf_anchor", "col", i, c, cands, stated,
                        label=clean_label(tv.row_label(i)))
        # 순증가(감소) = 활동 앵커 합 (+환율변동)
        for i in nets:
            stated = tv.value(i, c)
            if stated is None:
                continue
            base = [j for j in acts if j < i]
            if len(base) < 2:
                continue
            extra = [j for j in fx if j < i]
            cands = []
            for cc in cols:
                for rows in (base, base + extra):
                    terms = [(j, cc, 1) for j in rows if tv.value(j, cc) is not None]
                    if len(terms) >= 2:
                        cands.append((sum(tv.value(j, c2) for j, c2, _ in terms), terms))
            if cands:
                col.add(tv, "cf_anchor", "col", i, c, cands, stated,
                        label=clean_label(tv.row_label(i)))


_BV_RE = re.compile(r"장부(금액|가액)$")
_FV_RE = re.compile(r"공정가(치|액)$")  # K-GAAP 표기 '공정가액' 실측


def _check_fv_equal(tv: TableView, col: Collector):
    """장부가액 = 공정가치 등가 확인 (일치할 때만 — 원가측정 종목은 다른 게 정상).

    가로: 같은 헤더 그룹의 (공정가치, 장부가액) 열 짝 → 행별 등가
    세로: 인접한 (장부금액 행, 공정가치 행) 짝 → 열별 등가 (공정가치 주석 합계표)"""
    headers = _headers_of(tv)
    bv_cols = [c for c, h in enumerate(headers) if _BV_RE.search(h)]
    for bv in bv_cols:
        group = tv.header_group_cols(bv)
        scope = group if (group and bv in group) else list(range(tv.ncols))
        fv = next((c for c in scope if _FV_RE.search(headers[c]) and c != bv), None)
        if fv is None:
            continue
        for i in tv.body:
            stated = tv.value(i, bv)
            v = tv.value(i, fv)
            if stated is None or v is None or (stated == 0 and v == 0):
                continue
            col.add(tv, "fv_equal", "row", i, bv, [(v, [(i, fv, 1)])], stated)
    # 세로 짝: 행 라벨 (장부금액|장부가액) 바로 아래/위 (공정가치)
    for i in tv.body:
        li = C.norm(tv.row_label(i))
        if not _FV_RE.search(_label_core(tv.row_label(i)) or li):
            continue
        j = next((j for j in (i - 1, i + 1) if j in tv.body
                  and re.search(r"장부(금액|가액)|취득원가", C.norm(tv.row_label(j)))
                  and not _FV_RE.search(C.norm(tv.row_label(j)))), None)
        if j is None:
            continue
        for c in range(tv.ncols):
            stated = tv.value(i, c)
            v = tv.value(j, c)
            if stated is None or v is None or (stated == 0 and v == 0):
                continue
            col.add(tv, "fv_equal", "col", i, c, [(v, [(j, c, 1)])], stated)


def _check_flow_columns(tv: TableView, col: Collector):
    """가로 흐름: 열이 기초|증감…|기말인 표에서 각 행의 기말 = 기초 + 사이 열들.

    헤더 변형 실측: '당기초|…|당기말'(연결 이연법인세 명세),
    '기초미상각잔액|…|기말미상각잔액'(공단 유형자산) — 접두·수식어를 허용한다."""
    headers = _headers_of(tv)
    start = next((c for c, h in enumerate(headers)
                  if re.fullmatch(r"(당|전)?기초(장부|미상각)?(금액|가액|잔액)?", h)), None)
    end = next((c for c in range((start or 0) + 2, tv.ncols)
                if re.fullmatch(r"(당|전)?기말(장부|미상각)?(금액|가액|잔액)?",
                                headers[c])), None)
    if start is None or end is None:
        return
    middle = list(range(start + 1, end))
    for i in tv.body:
        stated = tv.value(i, end)
        if stated is None or tv.value(i, start) is None:
            continue
        base = [(i, start, 1)]
        comps = [c for c in middle if tv.value(i, c) is not None]
        plain = base + [(i, c, 1) for c in comps]
        signed = base + [(i, c, -1 if C.FLOW_MINUS_RE.search(headers[c]) else 1)
                         for c in comps]
        cands = [(sum(s * tv.value(i, cc) for _, cc, s in t), t)
                 for t in (plain, signed) if len(t) >= 2]
        col.add(tv, "flow_h", "row", i, end, cands, stated)


# 잔액 차감 열만 (누계액·충당금). '손상차손·처분' 등은 흐름 변동 열이라 넣으면
# 흐름표(기초|취득|손상차손|기말)를 가로채 오탐이 된다 — IFRS 주석 실측.
# '보조금' 열은 이중적이다: 잔액 차감(취득원가 우측 '국고보조금')일 수도, 흐름 변동
# ('기초|취득|정부보조금|처분|기말' — 금오ADS 실측)일 수도 있어 무조건 넣으면 흐름표를
# 가로챈다 — 명명된 취득원가 열이 있고 그보다 오른쪽일 때만 차감 열로 인정한다.
_ACCUM_RE = re.compile(r"(누계액|충당금)$")
_SUBSIDY_RE = re.compile(r"보조금$")


def _check_asset_cost(tv: TableView, col: Collector):
    """장부금액 산식(가로): 장부금액 열 = 총액 열 + 차감 열(누계액·충당금, 음수 기재).

    총액 열은 이름이 회사·계정마다 다르다(취득원가·취득가액·채권액·총장부금액…).
    이름으로 못 찾으면 **차감 열 바로 앞의 숫자 열**을 총액으로 본다 — 단 그 차감
    열이 실제로 차감 성격(전부 ≤ 0)일 때만. 당기말|전기말이 나란한 표는 상위 헤더
    그룹 안에서만 짝을 찾는다 (기간이 섞이면 오탐)."""
    headers = _headers_of(tv)
    targets = [c for c, h in enumerate(headers)
               if re.fullmatch(r"기말(장부)?(금액|가액|잔액)?|순?장부(금액|가액)", h)]
    for target_c in targets:
        group = tv.header_group_cols(target_c)
        scope = group if (group and target_c in group) else list(range(tv.ncols))
        def is_contra(cx: int) -> bool:
            # 차감 성격: 값이 하나라도 양수면 그 열은 차감이 아니다
            return not any(v is not None and v > 0
                           for v in (tv.value(i, cx) for i in tv.body))

        # 총액 열 이름: 취득원가류 + 기말 미상각잔액 (공단 '기말 미상각잔액 + 취득보조금
        # = 기말 장부가액' 실측). '기초미상각잔액'은 흐름 시작 열이라 제외
        def is_gross_header(h: str) -> bool:
            if re.search(r"취득(원가|가액)", h):
                return True
            return h.endswith("미상각잔액") and not re.match(r"(당|전)?기초", h)

        named_cost = next((c for c in scope
                           if is_gross_header(headers[c]) and c != target_c), None)
        accum = [c for c in scope
                 if c != target_c and c != named_cost and is_contra(c)
                 and (_ACCUM_RE.search(headers[c])
                      or (named_cost is not None and c > named_cost
                          and _SUBSIDY_RE.search(headers[c])))]
        if not accum:
            continue
        cost_c = named_cost
        if cost_c is None:
            # 이름 미상 총액 열: 차감 열 바로 앞의 숫자 열
            before = [c for c in scope
                      if c < min(accum) and c != target_c
                      and any(tv.value(i, c) is not None for i in tv.body)]
            cost_c = before[-1] if before else None
        if cost_c is None:
            continue
        # 이름 미상 차감열 안전망: 취득원가 '이름이 확인된' 경우에 한해, 그 오른쪽의
        # 음수 열 전부 포함하는 확장 후보 (할인차금 등 — 기준 양식은 우측 열 전체를 합산).
        # 폴백 총액 열에서는 흐름 열(처분·감가상각)을 끌어들일 수 있어 쓰지 않는다.
        right_all = [] if named_cost is None else [
            cx for cx in scope
            if cx > cost_c and cx != target_c and cx not in accum
            and any(tv.value(i, cx) is not None for i in tv.body)
            and is_contra(cx)]
        for i in tv.body:
            stated = tv.value(i, target_c)
            if stated is None or tv.value(i, cost_c) is None:
                continue
            t1 = [(i, cost_c, 1)] + [(i, c, 1) for c in accum
                                     if tv.value(i, c) is not None]
            cands = []
            if len(t1) >= 2:
                cands.append((sum(tv.value(j, cc) for j, cc, _ in t1), t1))
            extra = [c for c in right_all if tv.value(i, c) is not None]
            if extra:
                t2 = t1 + [(i, c, 1) for c in extra]
                if len(t2) >= 2:
                    cands.append((sum(tv.value(j, cc) for j, cc, _ in t2), t2))
            if not cands:
                continue
            col.add(tv, "asset_cost", "row", i, target_c, cands, stated)


_MATURITY_RE = re.compile(r"(\d+\s*년|\d+\s*개월|만기|이하|초과|이내)")


def _check_maturity(tv: TableView, col: Collector):
    """만기분석 가로합: '장부금액' 열 = 만기 구간 열들의 합 (합계 열 제외).

    금융부채 만기분석 표(K-GAAP·IFRS 샘플 실측) 구조 규칙."""
    headers = _headers_of(tv)
    target_c = next((c for c, h in enumerate(headers)
                     if re.fullmatch(r"장부(금액|가액)", h)), None)
    if target_c is None:
        return
    terms_c = [c for c, h in enumerate(headers)
               if c != target_c and _MATURITY_RE.search(h)
               and not C.TOTAL_HEADER_RE.match(h)]
    if len(terms_c) < 2:
        return
    for i in tv.body:
        stated = tv.value(i, target_c)
        if stated is None:
            continue
        terms = [(i, c, 1) for c in terms_c if tv.value(i, c) is not None]
        if len(terms) < 2:
            continue
        total = sum(tv.value(j, cc) for j, cc, _ in terms)
        col.add(tv, "maturity", "row", i, target_c, [(total, terms)], stated)


def _check_cascade(tv: TableView, col: Collector):
    """캐스케이드 합(주석 소형 표, 세로): 대상 행 = 바로 위 연속 숫자 행들의 합.

    '총 법인세효과 = 부담액+조정+변동액', '법인세비용 = 적용세율+조정사항들' 등
    합계 라벨 없는 계산 행 대응 (K-GAAP 법인세 주석 실측). IFRS 현금흐름표의
    '영업활동현금흐름 = 창출현금+이자수취−법인세납부' 캐스케이드도 같은 구조.
    정확 일치만 방출."""
    if len(tv.body) > 60 or (tv.note_no is None
                             and tv.sec_title not in C.STATEMENT_KEYWORDS):
        return
    if tv.note_no is not None and len(tv.body) > 20:
        return
    rate_cols = _rate_cols(tv)
    # 0 행으로 항 수만 채운 캐스케이드('X = X + 0')는 관계가 아니다 — 자본변동표의
    # '당기초 = 전기말 + 0(조정)'(정당한 이월)만 예외
    strict_nonzero = not C.SOCIE_TITLE_RE.search(C.norm(tv.sec_title))
    for c in range(tv.ncols):
        if c in rate_cols:
            continue
        numeric_rows = [j for j in tv.body if tv.value(j, c) is not None]
        for pos, i in enumerate(numeric_rows):
            if pos < 2 or col.has(tv, i, c):
                continue
            stated = tv.value(i, c)
            if not tv.cell(i, c) or not tv.cell(i, c).text.strip():
                continue
            acc = Decimal(0)
            acc_signed = Decimal(0)
            terms = []
            terms_signed = []
            for j in reversed(numeric_rows[:pos]):
                v = Decimal(str(tv.value(j, c)))
                acc += v
                terms.append((j, c, 1))
                sgn = _pl_sign(tv.row_label(j))
                acc_signed += sgn * v
                terms_signed.append((j, c, sgn))
                # ±1 절사는 금액다운 값(|기재|≥1000)에서만 — 처분 열처럼 0·소액이
                # 몰린 열은 ±1 우연 조합이 흔해 노이즈가 된다 (AI검토 발견 실측).
                # exhaustive 프로필이면 게이트 해제 — 필요 시 AI가 후속 판정
                dec_stated = Decimal(str(stated))
                tol_ok = RELAXED or abs(dec_stated) >= 1000
                # 실질(0 아닌) 항 2개 이상만 ('주당손익 205 = 0 + 205' — 삼보 연결 PL 실측 오탐)
                if strict_nonzero and sum(1 for j2, _, _ in terms if tv.value(j2, c)) < 2:
                    continue
                if (len(terms) >= 2 and acc != 0
                        and (acc == dec_stated
                             or (tol_ok and abs(acc - dec_stated) <= C.FOOTING_TOL))):
                    col.add(tv, "cascade", "col", i, c,
                            [(acc, list(reversed(terms)))], stated)
                    break
                # 부호 변형: 차감성 행(법인세·비용 등)은 빼는 캐스케이드 —
                # 부문정보 '당기순손익 = 영업이익 + 금융손익 − 법인세' 실측
                if (len(terms_signed) >= 2 and acc_signed != 0
                        and (acc_signed == dec_stated
                             or (tol_ok and abs(acc_signed - dec_stated) <= C.FOOTING_TOL))
                        and terms_signed != terms):
                    col.add(tv, "cascade", "col", i, c,
                            [(acc_signed, list(reversed(terms_signed)))], stated)
                    break
                if len(terms) >= 10:
                    break


def _check_running_totals(tv: TableView, col: Collector):
    """데이터 주도 소계 발견(세로, 정확 일치만) — 라벨 표식 없는 '…계' 행 대응.

    위에서부터 열을 훑으며 '아직 소계에 접히지 않은 최근 단위들의 합 = 현재 행 값'이면
    소계로 확정하고 그 구간을 접는다(스택). 대형 이연법인세 명세의
    '계 = 항목들 / 계상액 = 계 + 조정 + 미인식' 체인 재현 (연결 실측 — 기준 양식 동일).
    cascade가 소형 표 전용(주석 ≤20행·항 ≤6)이라 못 덮는 대형 표를 담당한다.
    같은 값 합이 여러 구간 길이로 성립하면 최장 구간 채택 (기준 양식 구간 전체 합 실측)."""
    if tv.note_no is None or len(tv.body) > 80:
        return
    rate_cols = _rate_cols(tv)
    for c in range(tv.ncols):
        if c in rate_cols:
            continue
        units: list[tuple[Decimal, int]] = []   # (단위 합, 대표 행) — 접힌 소계 포함
        for j in tv.body:
            v = tv.value(j, c)
            if v is None:
                continue
            dv = Decimal(str(v))
            matched = 0
            matched_acc = None
            acc = Decimal(0)
            for k in range(1, len(units) + 1):   # 뒤에서 k개 단위
                acc += units[-k][0]
                if k >= 2 and acc != 0 and (
                        acc == dv or ((RELAXED or abs(dv) >= 1000)
                                      and abs(acc - dv) <= C.FOOTING_TOL)):
                    matched = k          # 최장 k 선호 (±1 절사는 |기재|≥1000만)
                    matched_acc = acc
            if matched:
                terms = [(units[-i][1], c, 1) for i in range(matched, 0, -1)]
                if not col.has(tv, j, c):
                    col.add(tv, "cascade", "col", j, c, [(matched_acc, terms)], v)
                del units[-matched:]
            units.append((dv, j))


_RATIO_COL_RE = re.compile(r"구성비|비중|지분율")


def _check_ratio_columns(tv: TableView, col: Collector):
    """구성비율 열: 비율 = 행 금액 ÷ 합계행 금액 (실측 관계 — 사용자 지시로 채택).

    빨간 차이도 방출한다(force_certain) — 표시 자릿수 ROUND로 흡수되지 않는 비율
    불일치는 작더라도 사람이 확인할 가치가 있다. 비율 기재 스케일 이원 대응:
    '83.6%'(분수 0.836 저장)는 factor 1, 헤더만 (%)이고 셀은 83.6이면 factor 100."""
    headers = _headers_of(tv)
    for rc_ in range(tv.ncols):
        if not _RATIO_COL_RE.search(headers[rc_]):
            continue
        # 지분율은 주식수 비율이다 — 분자·분모는 '주식' 머리글 열만. 종속기업 투자 표
        # (지분율 · 장부금액)처럼 금액 열뿐이면 지분율 식을 세우지 않는다(분기 DSD 실측).
        share_only = "지분율" in headers[rc_]
        amt = next((cc for cc in range(rc_ - 1, -1, -1)
                    if not _RATIO_COL_RE.search(headers[cc])
                    and (not share_only or "주식" in headers[cc])
                    and any(tv.value(i, cc) is not None for i in tv.body)), None)
        if amt is None:
            continue
        total_row = next((j for j in tv.body
                          if tv.row_is_total(j) and tv.value(j, amt) is not None), None)
        if total_row is None:
            continue
        tot = Decimal(str(tv.value(total_row, amt)))
        if not tot:
            continue
        total_ratio = tv.value(total_row, rc_)
        column_factor = (100 if total_ratio is not None
                         and abs(Decimal(str(total_ratio))) > Decimal("1.5") else 1)
        for i in tv.body:
            stated = tv.value(i, rc_)
            av = tv.value(i, amt)
            if stated is None or av is None:
                continue
            sd = Decimal(str(stated))
            # 헤더만 (%)이고 셀 값은 0.39, 39.65, ..., 합계 100인 표에서는
            # 개별 값 크기가 아니라 합계행 스케일로 열 전체의 factor를 정한다.
            factor = column_factor
            if "지분율" in headers[rc_]:
                # 소유주식수에서 계산한 92.8105%를 92.82%로 기재한 실문서처럼
                # 주주 지분율은 원천 집계의 미세 절사/반올림 차이를 허용한다.
                digits = 1
            else:
                # 구성비·비중은 기존 표시 정밀도를 유지해야 실제 불일치가 숨지 않는다.
                exp = sd.normalize().as_tuple().exponent
                digits = max(1, min(6, -exp if isinstance(exp, int) else 1))
            col.add(tv, "ratio", "row", i, rc_,
                    [(Decimal(str(av)) / tot * factor,
                      [(i, amt, 1), (total_row, amt, 1)])],
                    stated, op="/", force_certain=True, digits=digits, factor=factor)


# 라벨 기반 나눗셈 패턴: (대상 라벨, 분자 라벨, 분자 제외, 분모 라벨, 배율 후보)
# 미러 TRUE만 방출(추측) — 분모·분자 오인 시 일치가 안 나와 자연 침묵.
# 배율: 주당이익은 순이익이 천원 단위 표에서 원 단위로 기재되는 관행 대응 (×1000)
_LABEL_RATIO_PATTERNS = [
    (re.compile(r"유효(법인)?세율"), re.compile(r"법인세비용"), re.compile(r"차감전"),
     re.compile(r"차감전"), (1,)),
    (re.compile(r"설정(비)?율"), re.compile(r"충당금"), None,
     re.compile(r"(합계|계$|채권)"), (1,)),
    (re.compile(r"주당(순)?(이익|손익|손실)"),
     re.compile(r"(순(이익|손익|손실)|(계속|중단)영업(손익|이익|손실))"), None,
     re.compile(r"주식수"), (1, 1000)),
]
# 주당손익의 분자 한정어: '계속영업주당순손익'은 계속영업손익, '중단영업주당순손익'은
# 중단영업손익, 무한정 주당손익은 한정어 없는 순손익 (연결 주당손익 주석 실측 —
# 한정어를 무시하면 계속영업 주당손익이 총순손익÷주식수와 우연 일치해 오짝이 된다)
_EPS_QUALIFIERS = ("계속영업", "중단영업")


def _eps_qualifier_ok(target_label: str, num_label: str) -> bool:
    q = next((x for x in _EPS_QUALIFIERS if x in target_label), None)
    if q is not None:
        return q in num_label
    return not any(x in num_label for x in _EPS_QUALIFIERS)


def _check_label_ratios(tv: TableView, col: Collector):
    """라벨 기반 비율: 유효세율=법인세비용÷차감전순이익, 설정비율=충당금÷채권,
    주당이익=순이익÷주식수 (나눗셈, 미러 TRUE만 — 기준 양식 동일 관계 실측)."""
    if tv.note_no is None or len(tv.body) > 80:
        return
    for i in tv.body:
        label = C.norm(tv.row_label(i))
        pat = next((p for p in _LABEL_RATIO_PATTERNS if p[0].search(label)), None)
        if pat is None:
            continue
        _, num_re, num_excl, den_re, factors = pat
        above = [j for j in tv.body if j < i]
        is_eps = "주당" in label
        num = next((j for j in reversed(above)
                    if num_re.search(C.norm(tv.row_label(j)))
                    and not (num_excl and num_excl.search(C.norm(tv.row_label(j))))
                    and not pat[0].search(C.norm(tv.row_label(j)))
                    and (not is_eps
                         or _eps_qualifier_ok(label, C.norm(tv.row_label(j))))), None)
        den = next((j for j in reversed(above)
                    if den_re.search(C.norm(tv.row_label(j))) and j != num), None)
        if num is None or den is None:
            continue
        for c in range(tv.ncols):
            stated = tv.value(i, c)
            nv, dv = tv.value(num, c), tv.value(den, c)
            if stated is None or nv is None or not dv:
                continue
            sd = Decimal(str(stated))
            exp = sd.normalize().as_tuple().exponent
            digits = max(0, min(6, -exp if isinstance(exp, int) else 1))
            terms = [(num, c, 1), (den, c, 1)]
            for factor in factors:
                col.add(tv, "combo", "col", i, c,
                        [(Decimal(str(nv)) / Decimal(str(dv)) * factor, terms)],
                        stated, op="/", digits=digits, factor=factor)


def _find_header(headers: list[str], pattern: str, exclude: str = "") -> int | None:
    rx = re.compile(pattern)
    ex = re.compile(exclude) if exclude else None
    return next((c for c, h in enumerate(headers)
                 if rx.search(h) and not (ex and ex.search(h))), None)


def _check_variance_columns(tv: TableView, col: Collector):
    """당기/전기/증감액 표: 증감액 = 당기 − 전기 (기준 양식 별첨 분석표 실측)."""
    headers = _headers_of(tv)
    cur = _find_header(headers, r"당기(말)?")
    prev = _find_header(headers, r"전기(말)?")
    change = _find_header(headers, r"증\(?감\)?액|증감액")
    rate = _find_header(headers, r"증\(?감\)?률|증감률")
    if None in (cur, prev, change) or len({cur, prev, change}) < 3:
        return
    for i in tv.body:
        cv, pv, stated = tv.value(i, cur), tv.value(i, prev), tv.value(i, change)
        if cv is None or pv is None or stated is None:
            continue
        col.add(tv, "flow_h", "row", i, change,
                [(Decimal(str(cv)) - Decimal(str(pv)),
                  [(i, cur, 1), (i, prev, 1)])],
                stated, op="-")
        if rate is None:
            continue
        rv = tv.value(i, rate)
        if rv is None or not pv:
            continue
        rd = Decimal(str(rv))
        exp = rd.normalize().as_tuple().exponent
        digits = max(1, min(6, -exp if isinstance(exp, int) else 1))
        col.add(tv, "ratio", "row", i, rate,
                [(Decimal(str(stated)) / Decimal(str(pv)),
                  [(i, change, 1), (i, prev, 1)])],
                rv, op="/", force_certain=True, digits=digits)


def _check_contract_rollforward(tv: TableView, col: Collector):
    """분양계약 표의 명시적 가로 관계 두 가지.

    총분양금액 = 기초분양계약 + 신규계약액
    기말예약매출잔액 = 총분양금액 − 누적수익인식액
    """
    headers = _headers_of(tv)
    opening = _find_header(headers, r"기초.*분양계약")
    new = _find_header(headers, r"신규계약액?")
    total = _find_header(headers, r"총분양금액")
    recognized = _find_header(headers, r"누적.*(수익|매출).*인식")
    ending = _find_header(headers, r"기말.*(예약)?매출잔액")
    for i in tv.body:
        if None not in (opening, new, total):
            a, b, stated = tv.value(i, opening), tv.value(i, new), tv.value(i, total)
            if a is not None and b is not None and stated is not None:
                computed = Decimal(str(a)) + Decimal(str(b))
                if abs(computed - Decimal(str(stated))) <= C.FOOTING_TOL:
                    col.add(tv, "flow_h", "row", i, total,
                            [(computed, [(i, opening, 1), (i, new, 1)])], stated)
        if None not in (total, recognized, ending):
            a, b, stated = (tv.value(i, total), tv.value(i, recognized),
                            tv.value(i, ending))
            if a is not None and b is not None and stated is not None:
                computed = Decimal(str(a)) - Decimal(str(b))
                if abs(computed - Decimal(str(stated))) <= C.FOOTING_TOL:
                    col.add(tv, "flow_h", "row", i, ending,
                            [(computed, [(i, total, 1), (i, recognized, 1)])],
                            stated, op="-")


def _check_weighted_shares(tv: TableView, col: Collector):
    """가중평균주식수 명세: 누적주식수=일수×주식수, 평균=누적합÷일수."""
    if tv.note_no is None:
        return
    headers = _headers_of(tv)
    days = _find_header(headers, r"(해당)?일수")
    shares = _find_header(headers, r"유통.*주식수", r"누적|가중평균")
    cumulative = _find_header(headers, r"누적.*유통.*주식수")
    if None in (days, shares, cumulative):
        return
    for i in tv.body:
        d, q, stated = tv.value(i, days), tv.value(i, shares), tv.value(i, cumulative)
        if d is None or q is None or stated is None or col.has(tv, i, cumulative):
            continue
        col.add(tv, "combo", "row", i, cumulative,
                [(Decimal(str(d)) * Decimal(str(q)),
                  [(i, days, 1), (i, shares, 1)])], stated, op="*")

    for i in tv.body:
        if "가중평균" not in C.norm(tv.row_label(i)):
            continue
        stated = tv.value(i, cumulative)
        if stated is None or col.has(tv, i, cumulative):
            continue
        above = [j for j in tv.body if j < i]
        total_row = next((j for j in reversed(above)
                          if tv.row_is_total(j) and tv.value(j, cumulative) is not None), None)
        # 라벨에 '=합계÷365'처럼 일수가 명시되면 그 값과 같은 일수 셀을 우선한다.
        m = re.search(r"(?:÷|/)(\d{2,4})", C.norm(tv.row_label(i)))
        wanted = Decimal(m.group(1)) if m else None
        day_row = next((j for j in above if tv.value(j, days) is not None
                        and (wanted is None or Decimal(str(tv.value(j, days))) == wanted)), None)
        if total_row is None or day_row is None or not tv.value(day_row, days):
            continue
        total = Decimal(str(tv.value(total_row, cumulative)))
        den = Decimal(str(tv.value(day_row, days)))
        col.add(tv, "combo", "row", i, cumulative,
                [(total / den, [(total_row, cumulative, 1), (day_row, days, 1)])],
                stated, op="/", digits=0)


def _check_combo(tv: TableView, col: Collector):
    """산술조합(주석 소형 표): 대상 행 = 위 두 행의 차·곱·몫 (같은 열).

    법인세 유효세율(② ÷ ①)·자본금(주식수 × 주당금액)·납부세액(산출 − 공제) 실측 대응.
    방출 조건: 수식 판정 미러가 TRUE인 조합만 (12자리 정확 일치 — 우연 일치 무시 가능).
    이미 다른 검사가 잡은 대상 셀은 건너뛴다."""
    if tv.note_no is None or len(tv.body) > 20:
        return
    ratio_hint = re.compile(r"\[?\((\d)\)\s*/\s*\((\d)\)\]?")   # '유효세율[(2)/(1)]'
    # 행 태그: '(A)'·'(C=A-B)' — 등호 좌변이 그 행의 태그 (재무비율 표 실측)
    tag_re = re.compile(r"\(([A-Z\d])(?:=[^)]*)?\)$")
    # 태그 수식: '(C=A-B)'·'(F=C/E)'·'영업권(C+B-A)' — 태그들의 사칙 체인
    eq_re = re.compile(r"\((?:[A-Z\d]=)?([A-Z\d](?:\s*[+\-*/]\s*[A-Z\d])+)\)$")
    tok_re = re.compile(r"([+\-*/])\s*([A-Z\d])")
    row_tag = {}
    for j in tv.body:
        m = tag_re.search(_label_core(tv.row_label(j)) or C.norm(tv.row_label(j)))
        m = m or tag_re.search(C.norm(tv.row_label(j)))
        if m:
            row_tag.setdefault(m.group(1), j)
    for c in range(tv.ncols):
        numeric_rows = [j for j in tv.body if tv.value(j, c) is not None]
        for pos, i in enumerate(numeric_rows):
            if pos < 2 or col.has(tv, i, c):
                continue
            stated = tv.value(i, c)
            if not tv.cell(i, c) or not tv.cell(i, c).text.strip():
                continue
            # 라벨 힌트 나눗셈: '…[(2)/(1)]' → (2)행 ÷ (1)행 (ROUND(d,1) 흡수 인정)
            hm = ratio_hint.search(C.norm(tv.row_label(i)))
            if hm and hm.group(1) in row_tag and hm.group(2) in row_tag:
                ja, jb = row_tag[hm.group(1)], row_tag[hm.group(2)]
                if tv.value(ja, c) is not None and tv.value(jb, c) not in (None, 0):
                    val = Decimal(str(tv.value(ja, c))) / Decimal(str(tv.value(jb, c)))
                    col.add(tv, "combo", "col", i, c,
                            [(val, [(ja, c, 1), (jb, c, 1)])], stated, op="/")
                    continue
            # 태그 수식 힌트: '순부채(C=A-B)'·'자본조달비율(F=C/E)'·'영업권(C+B-A)' —
            # 라벨이 관계식을 자기서술하는 재무비율 표 (공단·연결 실측)
            em = eq_re.search(C.norm(tv.row_label(i)))
            if em:
                expr = em.group(1).replace(" ", "")
                toks = [("+", expr[0])] + tok_re.findall(expr[1:])
                tags = [t for _, t in toks]
                if (all(t in row_tag and row_tag[t] != i for t in tags)
                        and all(tv.value(row_tag[t], c) is not None for t in tags)):
                    ops = {o for o, _ in toks[1:]}
                    if ops <= {"+", "-"}:
                        terms = [(row_tag[t], c, 1 if o == "+" else -1)
                                 for o, t in toks]
                        total = sum(s * tv.value(j, cc) for j, cc, s in terms)
                        col.add(tv, "combo", "col", i, c, [(total, terms)], stated)
                        continue
                    if len(toks) == 2 and toks[1][0] in ("/", "*"):
                        ja, jb = row_tag[toks[0][1]], row_tag[toks[1][1]]
                        op2 = toks[1][0]
                        if op2 == "/" and not tv.value(jb, c):
                            pass
                        else:
                            va = Decimal(str(tv.value(ja, c)))
                            vb = Decimal(str(tv.value(jb, c)))
                            val = va / vb if op2 == "/" else va * vb
                            sd = Decimal(str(stated))
                            exp = sd.normalize().as_tuple().exponent
                            digits = max(1, min(6, -exp if isinstance(exp, int) else 1))
                            col.add(tv, "combo", "col", i, c,
                                    [(val, [(ja, c, 1), (jb, c, 1)])], stated,
                                    op=op2, digits=digits)
                            continue
            prev = numeric_rows[:pos]
            found = False
            # 블라인드 2항 조합(합·차·곱)은 정확 일치만 — 가까운 행 우선, 첫 일치 방출
            for a_idx in range(len(prev) - 1, -1, -1):
                for b_idx in range(len(prev) - 1, -1, -1):
                    if a_idx == b_idx or found:
                        continue
                    ja, jb = prev[a_idx], prev[b_idx]
                    va = Decimal(str(tv.value(ja, c)))
                    vb = Decimal(str(tv.value(jb, c)))
                    for op, val in (("+", va + vb if a_idx > b_idx else None),
                                    ("-", va - vb),
                                    ("*", va * vb if a_idx > b_idx else None)):
                        if val is None:
                            continue
                        if op == "*" and (abs(va) < 2 or abs(vb) < 2):
                            continue  # ×0/×1 류 퇴화 방지
                        if op == "+" and (va == 0 or vb == 0):
                            continue  # +0 퇴화(단순 복제) 방지
                        d = Decimal(str(stated)) - val
                        if d == 0 and not (stated == 0 and va == 0 and vb == 0):
                            col.add(tv, "combo", "col", i, c,
                                    [(val, [(ja, c, 1), (jb, c, 1)])], stated, op=op)
                            found = True
                            break
                    if found:
                        break
                if found:
                    break


def _check_nested_partition(tv: TableView, col: Collector):
    """같은 들여쓰기 깊이에 소계 행과 세부 행이 섞인 구간의 값 기반 계층 발견.

    연결 포괄손익계산서 실측(삼보 반기연결): 'XI.기타포괄손익' 아래 '후속적으로 …
    포괄손익'(소계)과 '1.확정급여…'(세부)가 같은 깊이, 'XIII.반기순이익의 귀속' 아래
    '1.지배기업 소유주지분'(소계)·'계속영업 소유주지분'(세부)이 같은 깊이라 들여쓰기
    그룹 합(rollup)이 이중 합산으로 빗나간다. 기준 양식은 소계=세부합, 상위=소계합을 모두
    검증한다.

    방법: 구간(같은 최소 깊이 행들)을 앞에서부터 '머리 행 = 뒤따르는 행들의 누적합'이
    성립하는 순간 닫히는 연속 세그먼트로 탐욕 분할한다. 값 열 전체를 벡터로 비교하므로
    우연 일치는 사실상 없다. 전 세그먼트가 닫힐 때만 방출(추측형 — 정확 일치 게이트).
    상위 행에 값이 있으면 상위 = 세그먼트 머리 합도 방출한다. 값이 없는 상위(주당손익
    섹션 머리)는 세그먼트만 방출한다."""
    body = tv.body
    ncols = tv.ncols
    labeled = [i for i in body if tv.row_label(i)]

    def has_val(j):
        return any(tv.value(j, c) is not None for c in range(ncols))

    for gi, i in enumerate(labeled):
        d = tv.label_depth(i)
        span = []
        for j in labeled[gi + 1:]:
            if tv.label_depth(j) <= d:
                break
            span.append(j)
        data = [j for j in span if has_val(j)]
        if len(data) < 3:
            continue
        min_d = min(tv.label_depth(j) for j in data)
        rows = [j for j in data if tv.label_depth(j) == min_d]
        if len(rows) < 3:
            continue
        cols = [c for c in range(ncols) if any(tv.value(j, c) is not None for j in rows)]
        if not cols:
            continue

        def vec(j):
            return tuple(Decimal(str(tv.value(j, c) or 0)) for c in cols)

        zero = tuple(Decimal(0) for _ in cols)
        # 이미 평면합이 맞는 정상 그룹은 group/rollup 검사가 담당 — 여기서는 제외
        head_vec = vec(i) if has_val(i) else None
        if head_vec is not None and head_vec == tuple(
                sum(vec(j)[k] for j in rows) for k in range(len(cols))):
            continue
        segments = []
        k = 0
        ok = True
        while k < len(rows):
            h = rows[k]
            hv = vec(h)
            acc = zero
            children = []
            k2 = k + 1
            closed = False
            while k2 < len(rows):
                cv = vec(rows[k2])
                acc = tuple(a + b for a, b in zip(acc, cv))
                children.append(rows[k2])
                k2 += 1
                if acc == hv:
                    closed = True
                    break
            if not closed:
                ok = False
                break
            while k2 < len(rows) and vec(rows[k2]) == zero:  # 후행 0 행 흡수
                children.append(rows[k2])
                k2 += 1
            segments.append((h, children))
            k = k2
        if not ok or not segments:
            continue
        # 세그먼트가 하나뿐이고 상위에 값이 없으면 근거가 약하다 (단순 소계 표) — 제외
        if len(segments) < 2 and head_vec is None:
            continue
        for h, children in segments:
            if vec(h) == zero:
                continue
            for c in cols:
                stated = tv.value(h, c)
                if stated is None:
                    continue
                terms = [(j, c, 1) for j in children if tv.value(j, c) is not None]
                if not terms:
                    continue
                col.add(tv, "rollup", "col", h, c,
                        [(sum(tv.value(j, cc) for j, cc, _ in terms), terms)], stated,
                        label=f"{clean_label(tv.row_label(h))} 합계")
        if head_vec is not None:
            for c in cols:
                stated = tv.value(i, c)
                if stated is None:
                    continue
                terms = [(h, c, 1) for h, _ in segments if tv.value(h, c) is not None]
                if len(terms) < 2:
                    continue
                col.add(tv, "rollup", "col", i, c,
                        [(sum(tv.value(j, cc) for j, cc, _ in terms), terms)], stated,
                        label=f"{clean_label(tv.row_label(i))} 합계")


def _looks_like_pl(tv: TableView) -> bool:
    """섹션 제목이 아닌 표 내용만으로 손익계산서형 표를 판별한다.

    지방공기업 별첨에는 재무제표가 주석 TITLE 아래 삽입돼 sec_title이 '주석'인
    사례가 있다. 단계손익 핵심 라벨이 3개 이상일 때만 활성화한다.
    """
    labels = [C.norm(tv.row_label(i)) for i in tv.body]
    pats = (r"매출액", r"매출원가", r"매출총(이익|손실)", r"영업(이익|손실)",
            r"법인세비용차감전", r"(당기|반기|분기)순(이익|손실)")
    return sum(any(re.search(p, label) for label in labels) for p in pats) >= 3


def check_statement_lines(views: list[TableView], col: Collector):
    for tv in views:
        if _skip_section(tv):
            continue
        if "손익계산서" in tv.sec_title or _looks_like_pl(tv):
            _check_pl_anchors(tv, col)
            _check_nested_partition(tv, col)
        _check_cf_anchors(tv, col)
        _check_socie(tv, col)
        _check_section_subtotal(tv, col)
        _check_flow_columns(tv, col)
        _check_asset_cost(tv, col)
        _check_maturity(tv, col)
        _check_fv_equal(tv, col)
        _check_summary_fin(tv, col)
        _check_net_pair(tv, col)
        _check_ratio_columns(tv, col)
        _check_variance_columns(tv, col)
        _check_contract_rollforward(tv, col)
        _check_weighted_shares(tv, col)
        _check_bullet_subtotal(tv, col)
        _check_eq_carry(tv, col)
        _check_hint_lines(tv, col)


# ── 진입점 ───────────────────────────────────────────────────────────

# 단가 × 수량 = 금액 (곱셈형 — AI검토 전제 승격, 정확 일치·±1 절사만)
_PRICE_COL_RE = re.compile(r"(단가|1주당.{0,6}(금액|가액)|액면가액?|주당.{0,4}(금액|가액))")
_QTY_COL_RE = re.compile(r"(수량|주식수|주수|좌수|구좌수)")
_AMT_ONLY_RE = re.compile(r"(금액|자본금|장부금액|취득가액|평가액|가액)$")


def _check_price_qty(tv: TableView, col: Collector):
    """곱셈형: 금액 = 단가 × 수량. 헤더에 세 열이 다 있을 때만, 정확 일치 위주."""
    if tv.note_no is None:
        return
    heads = [C.norm(tv.leaf_header(c)) for c in range(tv.ncols)]
    price_cols = [c for c, h in enumerate(heads) if h and _PRICE_COL_RE.search(h)]
    qty_cols = [c for c, h in enumerate(heads) if h and _QTY_COL_RE.search(h)]
    amt_cols = [c for c, h in enumerate(heads)
                if h and _AMT_ONLY_RE.search(h)
                and not _PRICE_COL_RE.search(h) and not _QTY_COL_RE.search(h)]
    if not (price_cols and qty_cols and amt_cols):
        return
    for i in tv.body:
        for pc in price_cols:
            p = tv.value(i, pc)
            if not p:
                continue
            for qc in qty_cols:
                q = tv.value(i, qc)
                if not q or qc == pc:
                    continue
                prod = Decimal(str(p)) * Decimal(str(q))
                for ac in amt_cols:
                    a = tv.value(i, ac)
                    if a is None or ac in (pc, qc) or col.has(tv, i, ac):
                        continue
                    if Decimal(str(a)) == prod:   # 정확 일치만 방출 (추측형)
                        col.add(tv, "combo", "row", i, ac,
                                [(prod, [(i, pc, 1), (i, qc, 1)])], a, op="*")


def analyze(views: list[TableView], relaxed: bool = False) -> list[FootingSpec]:
    """relaxed: exhaustive 검출 프로필 — 소액 ±1 절사 게이트를 풀어
    검출 재현율을 높인다. AI 후속 판정 여부와는 독립적이다."""
    global RELAXED
    RELAXED = relaxed
    try:
        col = Collector()
        check_autosum(views, col)
        check_group_footing(views, col)
        check_total_rows(views, col)
        check_total_columns(views, col)
        check_statement_lines(views, col)
        # 캐스케이드·산술조합·소계 발견은 다른 검사가 놓친 셀만 — 마지막에 수행
        for tv in views:
            if not _skip_section(tv):
                _check_cascade(tv, col)
                _check_combo(tv, col)
                _check_label_ratios(tv, col)
                _check_running_totals(tv, col)
                _check_price_qty(tv, col)
        return col.finish()
    finally:
        RELAXED = False
