"""대사(크로스) 검증 엔진 — 재무제표↔주석·재무제표↔재무제표·자본변동표·주석↔주석.

차이 노출 방식(사용자 확정): 근거 있는 후보 쌍은 금액이 정확히 맞지 않아도 수식화해
엑셀에 빨간 차이를 노출한다. 근거의 종류:
- 참조 스코프: 재무제표 라벨 "(주석4,13)" / IFRS 주석열이 가리키는 주석 ↔ 그 주석의
  모든 표 셀 (참조 자체가 근거 — 금액 정확 일치 시 쌍 확정, 실측: 세부 행 포함)
- 전역 검색: 참조가 없어도 재무제표에서 같은 금액이 유일하게 발견되면 쌍 확정
- 재무제표 간: 현금흐름표 당기순이익 ↔ 손익계산서, 기초·기말 현금 ↔ 재무상태표 현금
- 자본변동표: 잔액행 ↔ BS 자본 항목, 거래행(당기순이익·배당·유상증자 등) ↔ BS/PL/CF
- 주석 ↔ 주석: 라벨 대응(동일·포함) + 동일 기간 + 금액 일치 (주석쌍·라벨당 1건)

'기초' 행·열의 값은 전기말 잔액이므로 재무제표 전기 열과 대조한다 (버킷 이동).
배치는 dst(주석/CE/CF) 행 오른쪽. 하이퍼링크 라벨은 "시트:상위경로 -> 라벨".
"""
from __future__ import annotations

import re
from decimal import ROUND_DOWN, ROUND_HALF_UP, Decimal

from . import constants as C
from .checks import CrossSpec
from .tableview import TableView, is_total_label

MAX_PER_CELL = 4          # 대상 셀 하나에 붙는 대사 슬롯 상한 (실측 2~3)
MAX_PER_CELL_SHEET = 2    # 같은 재무제표로의 슬롯 상한 (소계열+세부열)
_TOTAL_SUFFIX_RE = re.compile(r"(합계|총계)$")
_OPENING_RE = re.compile(r"기초|전기이월")
# 반기·분기 검토보고서는 '반기순이익'·'분기순이익' (연결은 '연결반기순이익')
_NI_RE = re.compile(r"(당기|반기|분기)순(이익|손실|손익)")
_CASH_RE = re.compile(r"현금및(현금성자산|예치금)")
_CF_CASH_RE = re.compile(r"기[초말]의?현금")


# ── 재무제표 행 인덱스 ───────────────────────────────────────────────

class FsRow:
    __slots__ = ("tv", "i", "refs", "path", "label")

    def __init__(self, tv: TableView, i: int, refs: set[int], path: str, label: str):
        self.tv = tv
        self.i = i
        self.refs = refs
        self.path = path
        self.label = label


def _row_refs(tv: TableView, i: int) -> set[int]:
    """행의 주석 참조: K-GAAP 라벨 '(주석4,13)' + IFRS 주석 열 텍스트."""
    refs: set[int] = set()
    m = C.NOTE_REF_RE.search(tv.row_label(i))
    if m:
        refs |= {int(x) for x in re.findall(r"\d+", m.group(1))}
    for c in tv.note_ref_cols():
        for x in re.findall(r"\d+", tv.text(i, c)):
            refs.add(int(x))
    return refs


def _ancestor_path(tv: TableView, i: int) -> str:
    """들여쓰기 계층 상위 경로 'Ⅱ. 비유동자산 -> (2) 유형자산(주석4,13) -> 선박 '."""
    chain = [tv.row_label(i)]
    d = tv.label_depth(i)
    j = i - 1
    while j >= 0 and d > 0:
        if j in tv.body and tv.row_label(j):
            dj = tv.label_depth(j)
            if dj < d:
                chain.append(tv.row_label(j))
                d = dj
        j -= 1
    return " -> ".join(reversed(chain))


def build_fs_index(stmts: list[TableView]) -> list[FsRow]:
    rows: list[FsRow] = []
    for tv in stmts:
        if tv.sheet == "CE":
            continue  # 자본변동표는 대사의 dst — src로 쓰지 않는다
        # 참조의 들여쓰기 자식 상속: 그룹 행의 (주석n)은 하위 세부 행에도 적용
        stack: list[tuple[int, set[int]]] = []  # (depth, refs)
        for i in tv.body:
            if not tv.row_label(i):
                continue
            d = tv.label_depth(i)
            while stack and stack[-1][0] >= d:
                stack.pop()
            own = _row_refs(tv, i)
            inherited = set().union(*(r for _, r in stack)) if stack else set()
            rows.append(FsRow(tv, i, own | inherited, _ancestor_path(tv, i),
                              C.norm(tv.row_label(i))))
            stack.append((d, own))
    return rows


# ── 값·단위 비교 ─────────────────────────────────────────────────────

def _quant(d: Decimal, rounder: str) -> Decimal:
    return d.quantize(Decimal(1), ROUND_HALF_UP if rounder == "ROUND" else ROUND_DOWN)


def _match(src_sum, src_unit: int, dst_val, dst_unit: int):
    """단위 환산 비교. 반환: (일치 여부, scale, dst_scale, rounder)."""
    a, b = Decimal(str(abs(src_sum))), Decimal(str(abs(dst_val)))
    if src_unit == dst_unit:
        return a == b, 1, 1, "ROUND"
    if src_unit < dst_unit:            # src가 미세 단위 (원 vs 천원) → src를 축소
        k = dst_unit // src_unit
        for rounder in ("ROUND", "TRUNC"):
            if _quant(a / k, rounder) == b:
                return True, k, 1, rounder
        return False, k, 1, "ROUND"
    k = src_unit // dst_unit           # dst가 미세 단위 → dst를 축소
    for rounder in ("ROUND", "TRUNC"):
        if _quant(b / k, rounder) == a:
            return True, 1, k, rounder
    return False, 1, k, "ROUND"


_COMBO_CONTRA_RE = re.compile(r"(누계액|충당금|현재가치할인차금|보조금)$")


def _combo_rows(tv: TableView, i: int) -> list[int] | None:
    """인접 (원가, 차감행들) 조합 — 취득원가+감가상각누계액(+국고보조금)=장부가액.

    차감행이 여러 개(누계액·보조금 3단 표시 — AI검토 캘리브레이션 실측)면
    연속 차감행 전부를 묶는다."""
    rows = [i]
    j = i + 1
    while j in tv.body and _COMBO_CONTRA_RE.search(C.norm(tv.row_label(j))):
        rows.append(j)
        j += 1
    return rows if len(rows) > 1 else None


def _fs_bucket_for(label: str, header: str, bucket: str) -> str | None:
    """주석 셀의 재무제표 대조 기수. '기초' 행/열은 전기말 잔액이므로 한 기수 이동:
    당기→전기, 전기→대조 불가(전전기)."""
    if (_OPENING_RE.search(label) or _OPENING_RE.search(header)) \
            and not C.INVENTORY_RE.match(label):
        return "전기" if bucket == "당기" else None
    return bucket


def _spec(fr: FsRow, cc: int, rows: list[int], tv: TableView, i: int, c: int,
          matched=None) -> CrossSpec:
    vals = [Decimal(str(fr.tv.value(j, cc))) for j in rows]
    total = sum(vals)
    ok, k, dk, rounder = matched if matched is not None else _match(
        total, fr.tv.unit, tv.value(i, c), tv.unit)
    path = fr.path if len(rows) == 1 else ", ".join(
        _ancestor_path(fr.tv, j) for j in rows)
    return CrossSpec(
        src_table=fr.tv.table.id, src_cells=[(j, cc) for j in rows],
        dst_table=tv.table.id, dst_cell=(i, c),
        label_path=f"{fr.tv.sheet}:{path}",
        scale=k, dst_scale=dk, rounder=rounder)


# ── FS ↔ 주석 ────────────────────────────────────────────────────────

def _note_cell_candidates(tv: TableView, include_eps: bool = False):
    """대사 대상 주석 셀: 본문·숫자·0 아님·주석열 아님·비율/주당 아님.

    주당이익 행은 기본 제외 — 표 단위(천원)가 적용되지 않는 원단위 값이라
    단위 환산 대사가 성립하지 않는다 (실샘플 실측 노이즈). include_eps=True면
    포함하되 호출 측(FS↔주석)이 '주당' 재무제표 행과 동일 단위로만 대조한다."""
    for i in tv.body:
        if not include_eps and "주당" in C.norm(tv.row_label(i)):
            continue
        for c in range(tv.ncols):
            v = tv.value(i, c)
            if v is None or v == 0:
                continue
            if C.NONSUM_COL_RE.search(C.norm(tv.leaf_header(c))):
                continue
            yield i, c, v


def _label_hit(scope: list[FsRow], label: str) -> list[FsRow]:
    """라벨 매칭되는 스코프 행 (정확 일치 또는 4자 이상 포함 관계)."""
    core = re.sub(r"\([^)]*\)", "", label)  # (주석n)·괄호 접미 제거
    out = []
    for fr in scope:
        fcore = re.sub(r"\([^)]*\)", "", fr.label)
        fcore = re.sub(r"^[0-9]{1,2}\.", "", fcore)  # '1.예적금담보대출' 번호 접두
        if not core or not fcore:
            continue
        if core == fcore or (len(core) >= 4 and core in fcore) \
                or (len(fcore) >= 4 and fcore in core):
            out.append(fr)
    return out


def _best_label_match(scope: list[FsRow], label: str) -> FsRow | None:
    """dsdxlsx ties._best_note_match 스코어 포팅: 합계류 +4, 정확 일치 +3, 포함 +2."""
    best, best_score = None, 0
    for fr in scope:
        score = 0
        if is_total_label(fr.label) or C.DEDUCT_TOTAL_RE.search(fr.label):
            score += 4
        if fr.label == label:
            score += 3
        elif label and (label in fr.label or fr.label in label):
            score += 2
        if score > best_score:
            best, best_score = fr, score
    return best if best_score >= 2 else None


# 담보·지급보증 내역 주석 — 표의 금액(채권최고액 등)이 BS 장부금액과 다른 게
# 정상이라 금액 대사가 무의미하다 (AI검토 캘리브레이션 실측: 오짝 노이즈만 생산)
_COLLATERAL_NOTE_RE = re.compile(r"담보|지급보증")


def _adjacent_contra_sum(tv: TableView, i: int, c: int):
    """dst 행 바로 아래 연속 차감행(충당금 등) 합 — 총액 셀 오짝 설명용."""
    total = Decimal(0)
    j = i + 1
    while j in tv.body and C.CONTRA_RE.search(C.norm(tv.row_label(j))):
        av = tv.value(j, c)
        if av is not None:
            total += Decimal(str(av))
        j += 1
    return total


def _implied_scale(sv: Decimal, dv: Decimal):
    """단위 미선언 오짝 감지 — 값 비율이 10^3·10^6 근방이면 함의 배율 반환.

    (원 단위 BS ↔ 천원 주석인데 한쪽 단위 선언이 빠져 원값끼리 비교된 경우:
    차이가 금액 전체로 보이는 노이즈 대신 환산 후의 의미 있는 차이를 노출)"""
    if not sv or not dv:
        return None
    r = abs(sv) / abs(dv)
    for k in (Decimal(1000), Decimal(1000000)):
        if k / Decimal("1.1") <= r <= k * Decimal("1.1"):
            return int(k), 1
        if 1 / (k * Decimal("1.1")) <= r <= 1 / (k / Decimal("1.1")):
            return 1, int(k)
    return None


def _fs_note_pairs(fs_rows: list[FsRow], notes: dict[int, list[TableView]]) -> list[CrossSpec]:
    specs: list[CrossSpec] = []
    by_ref: dict[int, list[FsRow]] = {}
    for r in fs_rows:
        for n in r.refs:
            by_ref.setdefault(n, []).append(r)

    for note_no, tvs in notes.items():
        note_title = next((t.note_title for t in tvs if t.note_title), "")
        if _COLLATERAL_NOTE_RE.search(C.norm(note_title)):
            continue
        scope = by_ref.get(note_no, [])
        main_table = next((tv for tv in tvs
                           if any(True for _ in _note_cell_candidates(tv))), None)
        for tv in tvs:
            cells = list(_note_cell_candidates(tv, include_eps=True))
            single_row = len(cells) <= 2
            last_col_of_row: dict[int, int] = {}
            label_count: dict[str, int] = {}
            for i, c, v in cells:
                last_col_of_row[i] = c
            for i in {i for i, _, _ in cells}:
                lb = C.norm(tv.row_label(i))
                label_count[lb] = label_count.get(lb, 0) + 1
            for i, c, v in cells:
                label = C.norm(tv.row_label(i))
                header = C.norm(tv.leaf_header(c))
                total_ish = (tv.row_is_total(i) or is_total_label(label)
                             or bool(C.FLOW_END_RE.match(label))
                             or bool(C.DEDUCT_TOTAL_RE.search(label))
                             or bool(_TOTAL_SUFFIX_RE.search(label)))
                labeled = _label_hit(scope, label)
                bucket = tv.bucket_of(c) or "당기"
                fs_bucket = _fs_bucket_for(label, header, bucket)
                if fs_bucket is None:
                    continue  # 전기의 '기초'(=전전기)는 대조 대상 없음
                hits: list[CrossSpec] = []
                per_sheet: dict[str, int] = {}
                if "주당" in label:
                    # 주당손익 행: 손익계산서 '주당' 행과만, 단위 환산 없이(둘 다 원) 대조.
                    # 소액(205원 등)이라 전역 금액 검색은 하지 않는다 (연결 주당손익 주석 실측)
                    core_l = re.sub(r"\([^)]*\)", "", label)
                    for fr in fs_rows:
                        if fr.tv.sheet not in ("PL", C.SECOND_PL_CODE) or "주당" not in fr.label:
                            continue
                        if fr.tv.unit != tv.unit:
                            continue
                        core_f = re.sub(r"\([^)]*\)", "", fr.label)
                        # 한정어(계속영업/중단영업/희석/기본)가 어긋나는 짝은 제외
                        if any((q in core_l) != (q in core_f)
                               for q in ("계속영업", "중단영업", "희석")):
                            continue
                        for cc in range(fr.tv.ncols):
                            if fr.tv.bucket_of(cc) != fs_bucket:
                                continue
                            fv = fr.tv.value(fr.i, cc)
                            if fv in (None, 0) or Decimal(str(fv)) != Decimal(str(v)):
                                continue
                            if len(hits) >= MAX_PER_CELL_SHEET:
                                break
                            specs.append(_spec(fr, cc, [fr.i], tv, i, c,
                                               matched=(True, 1, 1, "ROUND")))
                            hits.append(specs[-1])
                    continue
                won = abs(Decimal(str(v))) * tv.unit
                if won < C.GLOBAL_SEARCH_MIN:
                    continue  # 소액은 우연 일치 위험 — 대사 비대상

                def consider(fr: FsRow, cc: int, rows: list[int]):
                    if len(hits) >= MAX_PER_CELL:
                        return
                    if per_sheet.get(fr.tv.sheet, 0) >= MAX_PER_CELL_SHEET:
                        return
                    vals = [fr.tv.value(j, cc) for j in rows]
                    if any(x is None for x in vals):
                        return
                    total = sum(Decimal(str(x)) for x in vals)
                    if total == 0:
                        return
                    ok, k, dk, rounder = _match(total, fr.tv.unit, v, tv.unit)
                    if not ok:
                        return
                    specs.append(_spec(fr, cc, rows, tv, i, c,
                                       matched=(ok, k, dk, rounder)))
                    hits.append(specs[-1])
                    per_sheet[fr.tv.sheet] = per_sheet.get(fr.tv.sheet, 0) + 1

                # 1) 참조 스코프 — 모든 행이 후보 (참조 자체가 근거, 실측).
                #    라벨 매칭 행 우선, 조합(원가+누계) 포함
                for fr in (labeled + [x for x in scope if x not in labeled]):
                    for cc in range(fr.tv.ncols):
                        if (fr.tv.bucket_of(cc) != fs_bucket
                                or fr.tv.value(fr.i, cc) is None):
                            continue
                        consider(fr, cc, [fr.i])
                        combo = _combo_rows(fr.tv, fr.i)
                        if combo:
                            consider(fr, cc, combo)
                # 2) 전역 검색 (스코프에서 못 찾았을 때) — 유일 매칭만
                if not hits:
                    globals_hit: list[tuple[FsRow, int]] = []
                    for fr in fs_rows:
                        for cc in range(fr.tv.ncols):
                            if fr.tv.bucket_of(cc) != fs_bucket:
                                continue
                            fv = fr.tv.value(fr.i, cc)
                            if fv in (None, 0):
                                continue
                            ok, *_ = _match(Decimal(str(fv)), fr.tv.unit, v, tv.unit)
                            if ok:
                                globals_hit.append((fr, cc))
                    # 소계열+세부열 이중 히트(같은 라벨·같은 금액)는 유일로 간주
                    if globals_hit and len({fr.label for fr, _ in globals_hit}) <= 2:
                        for fr, cc in globals_hit[:MAX_PER_CELL_SHEET]:
                            consider(fr, cc, [fr.i])
                # 3) 강근거 불일치(빨강): 스코프가 있는데 금액이 안 맞는 행 —
                #    합계류이거나, 주석의 본 표에서 라벨이 유일하게 정확 대응할 때.
                #    행당 1건(마지막 열 = 잔액 성격). 차이 노출 방식: 차이를 노출.
                core = re.sub(r"\([^)]*\)", "", label)
                exact_labeled = [fr for fr in labeled
                                 if re.sub(r"\([^)]*\)", "", fr.label) == core]
                strong_label = (exact_labeled and tv is main_table
                                and label_count.get(label, 0) <= 1)
                # 기간 열(당기/전기 …)의 합계류 행은 열마다 별개 대조이므로 마지막 열
                # 제한을 두지 않는다 — 당기 열의 불일치가 전기 열에 가려지던 문제
                # (반기 CF 조정 합계 ↔ 주석 합계 ±1 실측)
                period_total = total_ish and tv.bucket_of(c) is not None
                if (not hits and scope
                        and (c == last_col_of_row.get(i) or period_total)
                        and (total_ish or strong_label)
                        and not C.CONTRA_RE.search(header)):    # 누계액·충당금 열 셀 제외
                    best = (exact_labeled[0] if exact_labeled
                            else _best_label_match(scope, label))
                    if best is not None:
                        fr = best
                        cc = next((x for x in range(fr.tv.ncols)
                                   if fr.tv.bucket_of(x) == fs_bucket
                                   and fr.tv.value(fr.i, x) is not None), None)
                        # 같은 라벨이 유동·비유동으로 나뉜 재무제표 행(매출채권및기타채권 등)
                        # 의 합이 주석 총액과 맞으면 조합 일치로 방출 — 유동 행 하나와의
                        # 빨간 차이(=비유동 금액) 오짝 방지 (연결 금융상품 주석 실측)
                        same_tbl = [x for x in exact_labeled if x.tv is fr.tv and x.i != fr.i]
                        if cc is not None and same_tbl:
                            rows = [fr.i] + [x.i for x in same_tbl]
                            vals = [fr.tv.value(j, cc) for j in rows]
                            if all(x is not None for x in vals):
                                ok_c, *_ = _match(sum(Decimal(str(x)) for x in vals),
                                                  fr.tv.unit, v, tv.unit)
                                if ok_c:
                                    specs.append(_spec(fr, cc, rows, tv, i, c))
                                    continue
                        if cc is not None:
                            sv = Decimal(str(fr.tv.value(fr.i, cc)))
                            dv = Decimal(str(v))
                            # 부호 불일치 짝(총액↔누계액 셀 등)은 무의미 — 억제
                            if (sv > 0) != (dv > 0):
                                continue
                            # 인접 차감행(충당금 등)으로 설명되는 총액 셀 오짝 — 억제
                            adj = _adjacent_contra_sum(tv, i, c)
                            if adj:
                                ok2, *_ = _match(sv, fr.tv.unit, dv + adj, tv.unit)
                                if ok2:
                                    continue
                            # 단위 미선언 오짝: 함의 배율로 보정해 의미 있는 차이 노출
                            matched = None
                            ok3, k, dk, rounder = _match(sv, fr.tv.unit, dv, tv.unit)
                            if not ok3 and k == 1 and dk == 1:
                                imp = _implied_scale(sv, dv)
                                if imp:
                                    matched = (False, imp[0], imp[1], "ROUND")
                            specs.append(_spec(fr, cc, [fr.i], tv, i, c,
                                               matched=matched))
    return specs


# ── 재무제표 간 대사 (현금흐름표 ↔ PL·BS) ────────────────────────────

def _cf_statement_pairs(stmts: list[TableView], fs_rows: list[FsRow]) -> list[CrossSpec]:
    """CF 당기순이익 ↔ PL, CF 기초·기말 현금 ↔ BS 현금및현금성자산.

    기준 양식은 IFRS 샘플에만 일부 방출하지만 관계 자체는 모든 문서에서 성립하므로
    대칭·전 문서 방출한다 (기초 현금은 BS 전기 열과 대조)."""
    specs: list[CrossSpec] = []
    ni_rows = [fr for fr in fs_rows
               if fr.tv.sheet in ("PL", C.SECOND_PL_CODE)
               and _NI_RE.search(re.sub(r"\([^)]*\)", "", fr.label))
               and "주당" not in fr.label]
    cash_rows = [fr for fr in fs_rows
                 if fr.tv.sheet == "BS" and _CASH_RE.search(fr.label)]
    for tv in stmts:
        if tv.sheet != "CF":
            continue
        for i in tv.body:
            label = C.norm(tv.row_label(i))
            if "주당" in label:
                continue
            is_ni = bool(_NI_RE.search(re.sub(r"\([^)]*\)", "", label)))
            is_cash = bool(_CF_CASH_RE.search(label))
            if not (is_ni or is_cash):
                continue
            for c in range(tv.ncols):
                v = tv.value(i, c)
                if v in (None, 0):
                    continue
                bucket = tv.bucket_of(c)
                if bucket is None:
                    continue
                if is_cash and _OPENING_RE.search(label):
                    fs_bucket = "전기" if bucket == "당기" else None
                else:
                    fs_bucket = bucket
                if fs_bucket is None:
                    continue
                n = 0
                for fr in (ni_rows if is_ni else cash_rows):
                    if n >= MAX_PER_CELL_SHEET:
                        break
                    for cc in range(fr.tv.ncols):
                        if fr.tv.bucket_of(cc) != fs_bucket:
                            continue
                        fv = fr.tv.value(fr.i, cc)
                        if fv in (None, 0):
                            continue
                        ok, *_ = _match(Decimal(str(fv)), fr.tv.unit, v, tv.unit)
                        if ok:
                            specs.append(_spec(fr, cc, [fr.i], tv, i, c))
                            n += 1
                            break
    return specs


# ── 자본변동표 매트릭스 ─────────────────────────────────────────────

_CE_BALANCE_RE = re.compile(r"(전기초|전기말|당기초|당기말|기초|기말)")


def _ce_matrix(ce_views: list[TableView], fs_rows: list[FsRow]) -> list[CrossSpec]:
    specs: list[CrossSpec] = []
    bs_rows = [fr for fr in fs_rows if fr.tv.sheet == "BS"]
    for tv in ce_views:
        # 행 블록 버킷: 전기초~전기말 = 전기, 당기초(또는 전기말 이후) = 당기
        current_block = "전기"
        for i in tv.body:
            label = C.norm(tv.row_label(i))
            date_like = re.match(r"20\d\d[.\s]", label)
            bal = _CE_BALANCE_RE.search(label)
            if bal and re.search(r"당\d?(반|분)?기", label):   # 당기초·당반기말·당3분기말
                current_block = "당기"

            if bal or date_like:
                # 잔액 행 × 구성요소 열 ↔ BS 자본 항목 (전기초는 제외 — BS에 없음)
                if "전기초" in label or (current_block == "전기" and "기초" in label):
                    continue
                # 반기·분기 보고서의 전반기말(전분기말) 잔액은 전기말 BS(연말)와 다른
                # 시점 — 대조 대상이 없다 (금액이 우연히 같아도 근거 없는 짝)
                if re.search(r"전(반|분)기말", label):
                    continue
                bucket = "당기" if (re.search(r"당\d?(반|분)?기말", label)
                                  or (current_block == "당기" and "기말" in label)) else "전기"
                for c in range(tv.ncols):
                    v = tv.value(i, c)
                    if v in (None, 0):
                        continue
                    n = 0
                    for fr in bs_rows:
                        if n >= 2:
                            break
                        for cc in range(fr.tv.ncols):
                            if fr.tv.bucket_of(cc) != bucket:
                                continue
                            fv = fr.tv.value(fr.i, cc)
                            if fv in (None, 0):
                                continue
                            ok, *_ = _match(Decimal(str(fv)), fr.tv.unit, v, tv.unit)
                            if ok:
                                specs.append(_spec(fr, cc, [fr.i], tv, i, c))
                                n += 1
                                break
            else:
                # 거래 행 (당기순이익·총포괄·배당·유상증자·주식보상 등) ↔ BS/PL/CF —
                # 블록 기수의 금액 일치로 대응 (실측: 유상증자 ↔ BS 주식발행초과금
                # + CF 유상증자)
                if tv.is_section_row(i):
                    continue
                bucket = current_block
                for c in range(tv.ncols):
                    v = tv.value(i, c)
                    if v in (None, 0):
                        continue
                    won = abs(Decimal(str(v))) * tv.unit
                    if won < C.GLOBAL_SEARCH_MIN:
                        continue
                    n = 0
                    per_sheet: dict[str, int] = {}
                    for fr in fs_rows:
                        if n >= 3:
                            break
                        if per_sheet.get(fr.tv.sheet, 0) >= 1:
                            continue
                        for cc in range(fr.tv.ncols):
                            if fr.tv.bucket_of(cc) != bucket:
                                continue
                            fv = fr.tv.value(fr.i, cc)
                            if fv in (None, 0):
                                continue
                            ok, *_ = _match(Decimal(str(fv)), fr.tv.unit, v, tv.unit)
                            if ok:
                                specs.append(_spec(fr, cc, [fr.i], tv, i, c))
                                n += 1
                                per_sheet[fr.tv.sheet] = 1
                                break
    return specs


# ── 주석 ↔ 주석 ─────────────────────────────────────────────────────

def _labels_relate(a: str, b: str) -> bool:
    if a == b:
        return True
    return (len(a) >= 4 and a in b) or (len(b) >= 4 and b in a)


def _emit_pair(specs: list[CrossSpec], src, dst) -> None:
    """주석↔주석 한 방향 방출. src/dst = (note_no, tv, i, c)."""
    _, stv, si, sc = src
    _, dtv, di, dc = dst
    ok, k, dk, rounder = _match(Decimal(str(stv.value(si, sc))), stv.unit,
                                dtv.value(di, dc), dtv.unit)
    if not ok:
        return
    specs.append(CrossSpec(
        src_table=stv.table.id, src_cells=[(si, sc)],
        dst_table=dtv.table.id, dst_cell=(di, dc),
        label_path=f"{stv.sheet}:{stv.row_label(si).strip()}",
        scale=k, dst_scale=dk, rounder=rounder))


def _note_note_pairs(notes: dict[int, list[TableView]]) -> list[CrossSpec]:
    """주석 간 대사 — 라벨 대응(동일·포함) + 동일 기간 + 금액 일치.

    폭주 방지 (실측에서 9배 과다였던 사례): 같은 계정이 여러 주석(금융상품 범주·
    공정가치·개별 주석·위험관리)에 반복되므로 **라벨·기간당 전역 1쌍**만 방출한다.
    소액(GLOBAL_SEARCH_MIN 미만)·일반 라벨(합계·기타 등)은 제외."""
    specs: list[CrossSpec] = []
    entries: list = []  # (note_no, tv, i, c, 라벨, 버킷, 원화금액)
    singles: list = []
    for n, tvs in notes.items():
        for tv in tvs:
            cells = list(_note_cell_candidates(tv))
            for i, c, v in cells:
                label = C.norm(tv.row_label(i))
                bucket = tv.bucket_of(c) or "당기"
                won = abs(Decimal(str(v))) * tv.unit
                if won < C.GLOBAL_SEARCH_MIN:
                    continue
                if len(cells) <= 2:
                    singles.append(((n, tv, i, c), won, bucket))
                if (len(label) < 2 or label in C.GENERIC_LABELS
                        or is_total_label(label)):
                    continue
                entries.append((n, tv, i, c, label, bucket, won))
    done: set = set()
    for ai in range(len(entries)):
        na, tva, ia, ca, la, ba, wa = entries[ai]
        for bi in range(ai + 1, len(entries)):
            nb, tvb, ib, cb, lb, bb, wb = entries[bi]
            if na == nb or ba != bb or wa != wb:
                continue
            if not _labels_relate(la, lb):
                continue
            key = (la if len(la) <= len(lb) else lb, ba)  # 라벨·기간당 전역 1쌍
            if key in done:
                continue
            done.add(key)
            _emit_pair(specs, (na, tva, ia, ca), (nb, tvb, ib, cb))
            _emit_pair(specs, (nb, tvb, ib, cb), (na, tva, ia, ca))
    # 단일 행 표(장기차입금의 유동성대체 등): 다른 주석의 동일 금액과 상호 방출
    amount_index: dict = {}
    for e in entries:
        amount_index.setdefault((e[6], e[5]), []).append(e)
    for entry, won, bucket in singles:
        others = [e for e in amount_index.get((won, bucket), []) if e[0] != entry[0]]
        if not others or len(others) > 2:
            continue
        for e in others:
            key = (min(entry[0], e[0]), max(entry[0], e[0]), "단일행", bucket, won)
            if key in done:
                continue
            done.add(key)
            _emit_pair(specs, (e[0], e[1], e[2], e[3]), entry)
            _emit_pair(specs, entry, (e[0], e[1], e[2], e[3]))
    return specs


# ── 진입점 ───────────────────────────────────────────────────────────

def analyze(views: list[TableView], ctx: dict, note_sheets: dict[int, str]) -> list[CrossSpec]:
    stmts: list[TableView] = ctx["statements"]
    notes: dict[int, list[TableView]] = ctx["notes"]
    fs_rows = build_fs_index(stmts)
    specs: list[CrossSpec] = []
    specs += _fs_note_pairs(fs_rows, notes)
    specs += _cf_statement_pairs(stmts, fs_rows)
    specs += _ce_matrix([tv for tv in stmts if tv.sheet == "CE"], fs_rows)
    specs += _note_note_pairs(notes)
    # 중복 제거: (src표, src셀들, dst표, dst셀) 동일 쌍은 1건
    seen: set = set()
    out: list[CrossSpec] = []
    for s in specs:
        key = (s.src_table, tuple(s.src_cells), s.dst_table, s.dst_cell)
        if key in seen:
            continue
        seen.add(key)
        out.append(s)
    return out
