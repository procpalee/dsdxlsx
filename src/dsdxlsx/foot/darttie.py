"""DART 대사 — 현재 DSD 의 비교 열(전기·전전기)을 DART 에 공시된 직전 보고서 원문과 맞춘다.

원천은 dartsrc.find_prior_reports 가 고른 원문(dart4 XML) 이고, 정렬은 diff.compare_docs 의
구획→표→행 3단 정렬을 그대로 쓴다. 열은 표 헤더의 기간 버킷으로 짝짓는다:
현재 문서 '전기' 열(k번째) ↔ 직전 문서 '당기' 열(k번째). 캡션으로만 기간이 갈리는 표
('(1) 당기말' 표 / '(2) 전기말' 표 쌍)는 v1 대상이 아니다(표 자체가 다른 표라 정렬이
당기↔당기로 붙는다).

반기·분기 검토보고서: 재무상태표는 전기 감사보고서와, 손익·현금흐름·자본변동은 전년
동기 검토보고서와 대사한다. 주석은 연차 보고서끼리만.

결과 배치(apply): 만든 워크북에 덧붙이는 후처리 — 각 시트 오른쪽 슬롯에 'DART: …' 링크(연녹 채움)
+ 살아있는 비교 수식(대상은 'DART원천' 시트의 값 셀) + 근거 메모, 총괄표 크로스 카운터에
합산. 'DART원천' 시트에는 원천 보고서·접수번호·조회시각·값 전량이 남는다(감사추적).
"""
from __future__ import annotations

import warnings
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from . import constants as C
from . import docdiff as diff
from . import sections
from .checks import CellRef, CrossRef
from .crossref import _match
from .dartsrc import DartClient, DartError, DartUnavailable, PriorDoc, find_prior_reports, identify
from .tableview import TableView, build_views

DART_SHEET = "DART원천"
_STMT_INTERIM_ONLY = ("PL", "CI", "CF", "CE", "RE")   # 전년 동기 검토보고서와 대사


@dataclass
class DartTie:
    sheet: str
    row: int
    col: int
    cur_value: object
    prior_value: object
    label: str            # 현재 문서 행 라벨
    header: str           # 현재 문서 열 헤더
    prior_label: str      # 직전 문서 행 라벨
    prior_header: str
    source: str           # 원천 보고서 표시
    scale: int
    dst_scale: int
    rounder: str
    ok: bool
    diff: Decimal
    dart_row: int = 0     # DART원천 시트 행 (apply 가 채움)


@dataclass
class DartTieResult:
    status: str                    # complete | partial | skipped
    reason: str = ""
    ties: list[DartTie] = field(default_factory=list)
    sources: list[str] = field(default_factory=list)
    unmatched: list[str] = field(default_factory=list)   # 표는 짝지었지만 열/값이 없어 건너뛴 사유

    @property
    def n_ok(self) -> int:
        return sum(1 for t in self.ties if t.ok)

    @property
    def n_bad(self) -> int:
        return sum(1 for t in self.ties if not t.ok)


# ── 분석 ─────────────────────────────────────────────────────────────

def _period_cols(tv: TableView, bucket: str) -> list[int]:
    """헤더 기반 기간 버킷 열 (캡션 폴백 제외). 주석 참조 열 제외."""
    out = []
    note_cols = tv.note_ref_cols()
    for c in range(tv.ncols):
        if c in note_cols:
            continue
        if tv.period_bucket(c) == bucket:
            out.append(c)
    return out


def _sheet_allowed(sheet: str, kind_cur: str, prior: PriorDoc) -> bool:
    if kind_cur == "annual":
        return True
    # 분·반기: 전기 감사보고서 → BS(재무상태표)만, 전년 동기 검토 → 손익·현금흐름·자본변동만
    if prior.kind == "annual":
        return sheet == "BS"
    return sheet in _STMT_INTERIM_ONLY


def analyze(doc, plans, ctx: dict, priors: list[PriorDoc], placed_rows: dict,
            kind_cur: str) -> tuple[list[DartTie], list[str]]:
    """placed_rows: table_id → (sheet, start_row) — 현재 문서 표의 시트 배치."""
    ties: list[DartTie] = []
    unmatched: list[str] = []
    for prior in priors:
        p_plans = sections.build(prior.doc)
        _, p_ctx = build_views(p_plans)
        rep = diff.compare_docs(prior.doc, doc)
        for sd in rep.sections:
            for bd in sd.blocks:
                if bd.kind != "table" or bd.old is None or bd.new is None:
                    continue
                ta, tb = bd.old.item, bd.new.item
                tv_a, tv_b = p_ctx["by_table"].get(ta.id), ctx["by_table"].get(tb.id)
                if tv_a is None or tv_b is None or tb.id not in placed_rows:
                    continue
                sheet, start_row = placed_rows[tb.id]
                if not _sheet_allowed(sheet, kind_cur, prior):
                    continue
                cols_b = _period_cols(tv_b, "전기")
                cols_a = _period_cols(tv_a, "당기")
                if not cols_b or not cols_a:
                    continue
                if len(cols_a) != len(cols_b):
                    if len(cols_a) == 1 or len(cols_b) == 1:
                        cols_a, cols_b = cols_a[:1], cols_b[:1]
                    else:
                        unmatched.append(f"{sheet} 표 {tb.id}: 기간 열 수 불일치 (직전 {len(cols_a)} vs 현재 {len(cols_b)})")
                        continue
                pairs = [(rd.old_idx, rd.new_idx) for rd in bd.rows
                         if rd.old_idx is not None and rd.new_idx is not None and not rd.in_thead]
                for ca, cb in zip(cols_a, cols_b):
                    for ia, ib in _repair_pairs(tv_a, tv_b, ca, cb, pairs):
                        va, vb = tv_a.value(ia, ca), tv_b.value(ib, cb)
                        ok, k, dk, rounder = _match(vb, tv_b.unit, va, tv_a.unit)
                        a = Decimal(str(abs(vb))) / k if k > 1 else Decimal(str(abs(vb)))
                        b = Decimal(str(abs(va))) / dk if dk > 1 else Decimal(str(abs(va)))
                        d = (a - b).copy_abs()
                        ties.append(DartTie(
                            sheet=sheet, row=start_row + ib, col=cb + C.CONTENT_START_COL,
                            cur_value=vb, prior_value=va,
                            label=tv_b.row_label(ib), header=tv_b.col_header(cb),
                            prior_label=tv_a.row_label(ia), prior_header=tv_a.col_header(ca),
                            source=prior.label, scale=k, dst_scale=dk, rounder=rounder,
                            ok=ok, diff=d))
    return ties, unmatched


def _repair_pairs(tv_a: TableView, tv_b: TableView, ca: int, cb: int,
                  pairs: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """행 정렬 짝(직전 ia ↔ 현재 ib) 중 값이 안 맞는 짝을 값으로 다시 맞춘다.

    DART 쪽 행 이름이 달라 정렬이 한 줄씩 밀리는 경우(현재 '단기금융자산의 증가' ↔ DART
    '장기금융상품의 취득', 분기 DSD 실측): DART 값이 같은 열의 다른 현재 행 값과 정확히
    같고 그 행이 유일하며 그 행의 기존 짝도 틀렸을 때만 그 행으로 옮긴다. 옮겨진 원래 행은
    짝 없음(대사 생략)이 된다. 0 은 우연 일치가 흔해 옮기지 않는다."""
    def val(tv, i, c):
        return tv.value(i, c)

    def ok_of(ia: int, ib: int) -> bool:
        return _match(val(tv_b, ib, cb), tv_b.unit, val(tv_a, ia, ca), tv_a.unit)[0]

    live = [(ia, ib) for ia, ib in pairs
            if val(tv_a, ia, ca) is not None and val(tv_b, ib, cb) is not None]
    by_b = {ib: ia for ia, ib in live}
    bad = {ib for ia, ib in live if not ok_of(ia, ib)}
    if not bad:
        return live
    moves: dict[int, int] = {}       # 현재 행 ib2 → 직전 행 ia (새 짝)
    taken: dict[int, int] = {}       # ib2 → 몇 번 지목됐나
    for ia, ib in live:
        if ib not in bad or not val(tv_a, ia, ca):
            continue
        hits = [ib2 for ib2 in tv_b.body
                if ib2 != ib and val(tv_b, ib2, cb) is not None
                and (ib2 not in by_b or ib2 in bad) and ok_of(ia, ib2)]
        if len(hits) == 1:
            moves[hits[0]] = ia
            taken[hits[0]] = taken.get(hits[0], 0) + 1
    moves = {ib2: ia for ib2, ia in moves.items() if taken[ib2] == 1}
    if not moves:
        return live
    moved_from = {ib for ia, ib in live if ia in moves.values() and ib not in moves}
    out = [(ia, ib) for ia, ib in live if ib not in moves and ib not in moved_from]
    out += [(ia, ib2) for ib2, ia in moves.items()]
    return sorted(out, key=lambda p: p[1])


def run(package, doc, plans, ctx: dict, placed_rows: dict, client: DartClient | None = None,
        fs_div: str | None = None) -> DartTieResult:
    """DSD 식별 → 직전 보고서 원문 → 대사. 키 없음·비공시 등은 skipped."""
    ident = identify(package.contents, package.meta)
    if fs_div:
        ident.consolidated = fs_div.upper() == "CFS"
    client = client or DartClient()
    try:
        priors = find_prior_reports(client, ident)
    except DartUnavailable as e:
        return DartTieResult("skipped", str(e))
    except DartError as e:
        return DartTieResult("skipped", f"OpenDART 오류: {e}")
    ties, unmatched = analyze(doc, plans, ctx, priors, placed_rows, ident.kind)
    res = DartTieResult("complete" if ties else "partial", "" if ties else "정렬된 표에서 대사 가능한 기간 열을 찾지 못함",
                        ties, [p.label for p in priors], unmatched)
    return res


# ── 워크북 반영 ───────────────────────────────────────────────────────

def _fmt(v) -> str:
    try:
        d = Decimal(str(v))
        return f"{d:,.0f}" if d == d.to_integral_value() else f"{d:,}"
    except Exception:  # noqa: BLE001
        return str(v)


def apply(out_path: Path, res: DartTieResult, right_starts: dict) -> None:
    """워크북에 DART원천 시트 + 시트별 오른쪽 슬롯 대사 수식 + 총괄표 카운터."""
    from openpyxl import load_workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    from . import styles as S
    from .slots import slot_for as _slot_for
    from .slots import summary_add as _summary_add
    from .emit import attach_comment

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(str(out_path))
    if DART_SHEET in wb.sheetnames:
        del wb[DART_SHEET]
    ws = wb.create_sheet(DART_SHEET)
    ws.cell(1, 1, "DART 대사 원천").font = S.FONT_BOLD
    ws.cell(2, 1, f"조회 {datetime.now():%Y-%m-%d %H:%M} · OpenDART 공시서류원본파일(document.xml)")
    ws.cell(3, 1, "상태 " + res.status + (f" — {res.reason}" if res.reason else ""))
    r = 4
    for src in res.sources:
        ws.cell(r, 1, "원천: " + src)
        r += 1
    for u in res.unmatched[:50]:
        ws.cell(r, 1, "건너뜀: " + u)
        r += 1
    r += 1
    heads = ("시트", "셀", "행 라벨", "열 헤더", "현재 값", "DART 값", "판정", "차이", "직전 문서 행", "직전 문서 열", "원천")
    for c, h in enumerate(heads, start=1):
        cell = ws.cell(r, c, h)
        cell.fill, cell.font = S.FILL_HEAD, S.FONT_HEAD
    r += 1
    first_data = r
    for t in res.ties:
        t.dart_row = r
        ws.cell(r, 1, t.sheet)
        ws.cell(r, 2, f"{get_column_letter(t.col)}{t.row}")
        ws.cell(r, 3, t.label)
        ws.cell(r, 4, t.header)
        ws.cell(r, 5, float(t.cur_value) if isinstance(t.cur_value, Decimal) else t.cur_value).number_format = S.NUMFMT_INT
        ws.cell(r, 6, float(t.prior_value) if isinstance(t.prior_value, Decimal) else t.prior_value).number_format = S.NUMFMT_INT
        ws.cell(r, 7, "일치" if t.ok else "불일치")
        ws.cell(r, 8, float(t.diff)).number_format = S.NUMFMT_INT
        ws.cell(r, 9, t.prior_label)
        ws.cell(r, 10, t.prior_header)
        ws.cell(r, 11, t.source)
        if not t.ok:
            for c in range(1, 12):
                ws.cell(r, c).fill = S.FILL_BAD
        r += 1
    for col, w in (("A", 8), ("B", 8), ("C", 30), ("D", 22), ("E", 16), ("F", 16), ("G", 8), ("H", 12),
                   ("I", 30), ("J", 22), ("K", 60)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = ws.cell(first_data, 1)

    added: dict[str, list[str]] = {}
    for t in res.ties:
        if t.sheet not in wb.sheetnames:
            continue
        wsd = wb[t.sheet]
        rs = right_starts.get(t.sheet) or (wsd.max_column + C.RIGHT_START_GAP)
        col = _slot_for(wsd, t.row, rs)
        link = wsd.cell(t.row, col)
        disp = f"DART: {t.prior_label or t.label}".replace('"', "'")[:60]
        link.value = f'=HYPERLINK("#\'{DART_SHEET}\'!F{t.dart_row}","{disp}")'
        link.font = S.FONT_BOLD
        link.fill = S.FILL_DART
        link.alignment = Alignment(horizontal="left")
        fcell = wsd.cell(t.row, col + 1)
        xr = CrossRef(src=[CellRef(t.sheet, t.row, t.col, t.cur_value)],
                      dst=CellRef(DART_SHEET, t.dart_row, 6, t.prior_value),
                      label_path=f"DART:{t.prior_label}", scale=t.scale, dst_scale=t.dst_scale,
                      rounder=t.rounder)
        fcell.value = xr.compare_formula()
        fcell.number_format = S.NUMFMT_CHECK
        fcell.border = S.BORDER_CHECK
        conv = (f" (÷{t.scale:,} {t.rounder})" if t.scale > 1
                else (f" (DART ÷{t.dst_scale:,} {t.rounder})" if t.dst_scale > 1 else ""))
        attach_comment(fcell, (f"[DART 대사] {t.source}\n"
                               f"직전 문서 {t.prior_label} / {t.prior_header} = {_fmt(t.prior_value)}\n"
                               f"현재 {t.label} / {t.header} = {_fmt(t.cur_value)}{conv}\n"
                               f"판정: {'일치' if t.ok else '불일치 (차이 ' + _fmt(t.diff) + ')'}"), height=120)
        wsd.conditional_formatting.add(
            fcell.coordinate, FormulaRule(formula=[f"ISNUMBER({fcell.coordinate})"], fill=S.FILL_BAD))
        for cc, w in ((col, 22.0), (col + 1, S.CHECK_NUM_WIDTH)):
            dim = wsd.column_dimensions[get_column_letter(cc)]
            dim.width = max(dim.width or 0, w)
        added.setdefault(t.sheet, []).append(fcell.coordinate)
    if "총괄표" in wb.sheetnames:
        for sheet, addrs in added.items():
            for a in addrs:
                _summary_add(wb["총괄표"], sheet, a)
        ws_sum = wb["총괄표"]
        note_row = ws_sum.max_row + 2
        ws_sum.cell(note_row, 2, "DART 대사").font = S.FONT_BOLD
        ws_sum.cell(note_row, 3, f"상태 {res.status} · 대사 {len(res.ties)}건 · 일치 {res.n_ok} · 불일치 {res.n_bad}"
                    + (f" — {res.reason}" if res.reason else "")
                    + (f" · 원천: {'; '.join(res.sources)}" if res.sources else ""))
    wb.save(str(out_path))


def print_summary(res: DartTieResult) -> None:
    if res.status == "skipped":
        print(f"  [DART대사] skipped — {res.reason}")
        return
    print(f"  DART대사: 일치 {res.n_ok} · 불일치 {res.n_bad}"
          + (f" ({res.reason})" if res.reason else ""))
    for s in res.sources:
        print(f"    원천 {s}")
    for t in res.ties:
        if not t.ok:
            print(f"    ✗ {t.sheet}!{t.label} / {t.header}: 현재 {_fmt(t.cur_value)} vs DART {_fmt(t.prior_value)}")
