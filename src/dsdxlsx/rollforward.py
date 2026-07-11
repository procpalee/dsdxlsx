"""기간 롤포워드 — 전기(직전 보고서) DSD에서 당기 표준 정산표를 만든다.

'결산 시즌 시작 버튼': 전기 DSD 하나로 ① 기수 라벨 +1 ② TU 자동필드 연도 +1
(스파이크 V4 검증 방식) ③ 각 표의 당기 열 값 → 전기 열 이동 + 당기 열 클리어
④ THEAD의 기간 날짜 +1년 을 수행하고, 결과를 3등급 리포트로 표면화한다.

등급 원칙 (플랜 확정):
- 자동확정: 결정론 치환 — 기수 라벨, TU, 버킷 페어가 명확한 표의 값 이동.
- 확인필요: 건드리지 않았지만 사람이 봐야 하는 것 — 문단·본문 셀의 날짜,
  meta 문서정보(편집기 소관), 클리어된 당기 열(결산 수치 입력 대상).
- 미처리: 구조를 확신할 수 없어 **절대 무변경**한 표 — 자본변동표(행 블록 구조),
  당/전 열 수가 안 맞는 표, 당기 열만 있는 표.

보고서 종류 전환(1분기→반기 등)은 범위 밖 — 날짜는 일괄 +1년(연차·동일분기 롤).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from . import splice
from .check import TableView, build_views, norm, use_profile
from .model import Document

_GISU_RE = re.compile(r"(제\s*)(\d{1,3})(\s*\(?[당전]?\)?\s*기)")
_DATE_RE = re.compile(r"(19|20)\d{2}(?=\s*[년.\-/])")
_YEAR_RE = re.compile(rb"(19|20)\d{2}")
# 세로형 구분 문단의 마지막 줄: "1) 당기말" / "2) 전기말" 류
_CUR_MARK_RE = re.compile(r"당\s*(분|반)?\s*기\s*말?\s*$")
_PREV_MARK_RE = re.compile(r"전\s*(분|반)?\s*기\s*말?\s*$")
# 자본변동표 블록 라벨 플립: 구 당기 블록이 신 전기 블록이 되므로
# (긴 패턴 먼저 — '당분기말'이 '당기말'보다 앞서 치환되도록)
_SOCIE_FLIP = [("당분기말", "전분기말"), ("당반기말", "전반기말"),
               ("당기초", "전기초"), ("당기말", "전기말")]


@dataclass
class RollItem:
    grade: str      # 자동확정 | 확인필요 | 미처리
    where: str
    item: str
    before: str = ""
    after: str = ""
    note: str = ""


@dataclass
class RollReport:
    items: list[RollItem] = field(default_factory=list)
    # 자동 이동 표시(색칠) 대상: ("cells",[(row_id,gc)..]) | ("table",tid) | ("socie",tid)
    painted: list[tuple] = field(default_factory=list)

    def add(self, grade, where, item, before="", after="", note=""):
        self.items.append(RollItem(grade, where, item, str(before)[:40],
                                   str(after)[:40], note))

    def count(self, grade: str) -> int:
        return sum(1 for i in self.items if i.grade == grade)


def _gisu_mapping(doc: Document) -> dict[int, int]:
    """문서에 등장하는 기수 중 상위 3개만 +1 — '제 1 기'류 서술(설립 연혁 등)
    오치환 방지."""
    nums = set()
    for node in doc.nodes.values():
        for m in _GISU_RE.finditer(node.text):
            nums.add(int(m.group(2)))
    top = sorted(nums, reverse=True)[:3]
    return {n: n + 1 for n in top}


def _period_years(doc: Document) -> set[int]:
    """보고기간 연도 집합 — TU 자동필드(기계값 YYYYMMDD)에서 추출한 최신 2개
    연도만. 이 연도의 날짜만 +1년 대상 — 설립·상장 등 역사적 연도 오염 방지."""
    years: set[int] = set()
    for node in doc.nodes.values():
        if node.tag == "TU":
            for m in _YEAR_RE.finditer(doc.data[node.start:node.end]):
                years.add(int(m.group()))
    if not years:  # TU 없는 문서: 본문 날짜 중 최신 연도로 추정
        for node in doc.nodes.values():
            for m in _DATE_RE.finditer(node.text):
                years.add(int(m.group()))
    if not years:
        return set()
    top = max(years)
    return {top, top - 1}  # 당기말·전기말 연도만 — 그 이전(역사적) 연도는 제외


def _bump_year(text: str, allowed: set[int]) -> str:
    return _DATE_RE.sub(
        lambda m: str(int(m.group()) + 1) if int(m.group()) in allowed else m.group(),
        text)


def _numeric_cells(tv: TableView, i: int):
    """행의 값 성격 셀(숫자·'-' 플레이스홀더)만 — 라벨·주석 텍스트는 제외."""
    from .check import cell_value
    for c in range(tv.ncols):
        cell = tv.cell(i, c)
        if cell is None or cell.tag == "TU":
            continue
        t = cell.text.strip()
        if t and (t == "-" or cell_value(cell.text) is not None):
            yield c, cell


def _roll_paired_tables(doc: Document, report: RollReport, ops: list,
                        touched: set, painted: list) -> set[str]:
    """세로형 표 쌍: 'N) 당기말' 문단 아래 표 ↔ 'N) 전기말' 문단 아래 표.

    기수 간 행 구성이 다를 수 있으므로(파생상품 신규 등) 셀 이동이 아니라
    **전기 표를 구 당기 표의 원시 바이트 클론으로 통째 교체**하고, 당기 표의
    값 셀만 비운다 — clone-only 원칙(편집기 수용 검증 방식)."""
    from .export_xlsx import refined_sections

    handled: set[str] = set()
    for sec in refined_sections(doc):
        pending = None  # (table, 당기마커 문단 요약)
        last_mark = ""
        for block in sec.blocks:
            if block.kind == "p":
                lines = [l.strip() for l in block.item.text.split("\n") if l.strip()]
                if lines:
                    last_mark = lines[-1]
                continue
            if block.kind != "table":
                continue
            t = block.item
            tags = {c.tag for r in t.rows for c in r.cells}
            eligible = ("TU" not in tags and not t.fixtable
                        and t.group_aclass is None and not t.in_insertion)
            if _CUR_MARK_RE.search(last_mark) and eligible:
                pending = (t, last_mark)
            elif (_PREV_MARK_RE.search(last_mark) and eligible
                  and pending is not None and pending[0].ncols == t.ncols):
                cur_t, cur_mark = pending
                clone = doc.data[cur_t.start:cur_t.end]
                ops.append(splice.Op(pos=t.start, start=t.start, end=t.end,
                                     payload=clone))
                for row in t.rows:
                    for cell in row.cells:
                        touched.add(cell.id)
                cur_tv = TableView(cur_t, sec.title)
                n_clear = 0
                for i in cur_tv.body:
                    for _, cell in _numeric_cells(cur_tv, i):
                        ops.append(splice.replace_inner(cell, ""))
                        touched.add(cell.id)
                        n_clear += 1
                handled.update({cur_t.id, t.id})
                painted.append(("table", t.id))
                report.add("자동확정", sec.title if not sec.note_no
                           else f"주석{sec.note_no}",
                           f"{t.id}(전기) ← {cur_t.id}(구 당기) 표 교체, "
                           f"당기 표 값 {n_clear}칸 비움",
                           note=f"'{cur_mark}'/'{last_mark}' 쌍 — 각주(*) 문단은 확인")
                pending = None
            else:
                pending = None
    return handled


def _roll_socie(doc: Document, views, report: RollReport, ops: list,
                touched: set, painted: list, allowed_years: set) -> set[str]:
    """자본변동표: [전기초..전기말] 블록 ← [당기초..당기말] 블록 통째 클론
    (라벨의 당기↔전기 플립), 신 당기초 값 ← 구 당기말 값(이월), 당기 변동·기말
    클리어. 블록 경계가 불명확하면 절대 무변경(미처리)."""
    from .check import norm as _norm

    handled: set[str] = set()
    for tv in views:
        if "자본변동표" not in _norm(tv.sec_title):
            continue
        t = tv.table
        body = tv.body
        opens = [i for i in body if "기초" in _norm(tv.row_label(i))
                 and any(tv.value(i, c) is not None for c in range(tv.ncols))]
        closes = [i for i in body if "기말" in _norm(tv.row_label(i))
                  and any(tv.value(i, c) is not None for c in range(tv.ncols))]
        if len(opens) != 2 or len(closes) != 2:
            continue
        o1, o2 = opens
        c1, c2 = closes
        if not (o1 < c1 < o2 < c2):
            continue
        if any(cell.tag == "TU" for i in body for cell in t.rows[i].cells):
            continue
        rows1 = t.rows[o1:c1 + 1]   # 전기 블록
        rows2 = t.rows[o2:c2 + 1]   # 당기 블록
        clone = doc.data[rows2[0].start:rows2[-1].end]
        for a, b in _SOCIE_FLIP:
            clone = clone.replace(a.encode(), b.encode())
        ops.append(splice.Op(pos=rows1[0].start, start=rows1[0].start,
                             end=rows1[-1].end, payload=clone))
        for row in rows1:
            for cell in row.cells:
                touched.add(cell.id)
        # 신 당기초 ← 구 당기말 값 (이월), 나머지 당기 행 값 클리어
        close_vals = {c: tv.cell(c2, c).text for c, _ in _numeric_cells(tv, c2)}
        for c, cell in _numeric_cells(tv, o2):
            new = close_vals.get(c, "")
            if cell.text != new:
                ops.append(splice.replace_inner(cell, new))
            touched.add(cell.id)
        n_clear = 0
        for i in body:
            if o2 < i <= c2:
                for c, cell in _numeric_cells(tv, i):
                    ops.append(splice.replace_inner(cell, ""))
                    touched.add(cell.id)
                    n_clear += 1
        # 당기 블록 라벨의 보고기간 연도 +1 ('2025.1.1(당기초)' 스타일 대응)
        for i in body:
            if o2 <= i <= c2:
                label_cell = t.rows[i].cells[0] if t.rows[i].cells else None
                if label_cell is not None and label_cell.id not in touched:
                    new_label = _bump_year(label_cell.text, allowed_years)
                    if new_label != label_cell.text:
                        ops.append(splice.replace_inner(label_cell, new_label))
                        touched.add(label_cell.id)
        handled.add(t.id)
        painted.append(("socie", t.id))
        report.add("자동확정", tv.context,
                   f"{t.id} 자본변동표 블록 교체 (전기 ← 구 당기, "
                   f"당기초 ← 구 당기말, 변동 {n_clear}칸 비움)",
                   note="당기 실적 행 구성은 결산 내용에 맞게 확인·수정")
    return handled


def roll_ops(doc: Document, profile=None) -> tuple[list[splice.Op], RollReport]:
    ops: list[splice.Op] = []
    report = RollReport()
    mapping = _gisu_mapping(doc)
    allowed_years = _period_years(doc)
    touched: set[str] = set()  # op가 걸린 노드 — 중복 치환·중복 보고 방지
    painted: list[tuple] = []  # 자동 이동 표시(색칠) 대상 — report.painted로 전달

    with use_profile(profile):
        views, _ = build_views(doc)

        # ── ⓪ 구조 단위 이동: 세로형 표 쌍 + 자본변동표 블록 ─────────
        handled = _roll_paired_tables(doc, report, ops, touched, painted)
        handled |= _roll_socie(doc, views, report, ops, touched, painted,
                               allowed_years)

        # ── ① 표별 당기 → 전기 값 이동 (열 페어) ────────────────────
        for tv in views:
            if _skip_table(tv) or tv.table.id in handled:
                continue
            if "자본변동표" in norm(tv.sec_title):
                # ⓪에서 블록 교체에 실패한 자본변동표만 — 항상 표면화
                if any(tv.value(i, c) is not None for i in tv.body
                       for c in range(tv.ncols)):
                    report.add("미처리", tv.context, f"{tv.table.id} 자본변동표",
                               note="블록 경계 인식 실패 — 당기 실적을 수기 작성")
                continue
            cur = [c for c in range(tv.ncols) if tv.period_bucket(c) == "당기"]
            prev = [c for c in range(tv.ncols) if tv.period_bucket(c) == "전기"]
            if not cur:
                continue  # 기간 열 없는 표 — 롤 무관
            if len(cur) != len(prev):
                report.add("미처리", tv.context,
                           f"{tv.table.id} 당기 {len(cur)}열/전기 {len(prev)}열",
                           note="열 구성이 대칭이 아니어서 자동 이동 안 함")
                continue
            if any(tv.cell(i, c) is not None and tv.cell(i, c).tag == "TU"
                   for i in tv.body for c in cur + prev):
                report.add("미처리", tv.context, f"{tv.table.id} TU 열 포함",
                           note="자동 필드 열은 이동 대상 아님")
                continue
            n_moved = 0
            moved_cells: list[tuple[str, int]] = []
            for i in tv.body:
                for cc, pc in zip(cur, prev):
                    c_cell, p_cell = tv.cell(i, cc), tv.cell(i, pc)
                    if c_cell is None or p_cell is None:
                        continue
                    if c_cell.text != p_cell.text:
                        ops.append(splice.replace_inner(p_cell, c_cell.text))
                        touched.add(p_cell.id)
                        if c_cell.text.strip():
                            moved_cells.append((tv.table.rows[i].id, pc))
                    if c_cell.text:
                        ops.append(splice.replace_inner(c_cell, ""))
                        touched.add(c_cell.id)
                    n_moved += 1
            if n_moved:
                if moved_cells:
                    painted.append(("cells", moved_cells))
                report.add("자동확정", tv.context,
                           f"{tv.table.id} 당기값 {n_moved}칸 → 전기 열",
                           note="당기 열은 빈 값 — 결산 수치 입력 대상")

        # ── ② 기수 라벨 +1 (TU·값이동 셀 제외) ──────────────────────
        n_label = 0
        for node in doc.nodes.values():
            if node.tag == "TU" or node.id in touched:
                continue
            new_text = _GISU_RE.sub(
                lambda m: (m.group(1)
                           + str(mapping.get(int(m.group(2)), int(m.group(2))))
                           + m.group(3)),
                node.text)
            if new_text == node.text:
                continue
            # 같은 노드의 보고기간 날짜도 함께 +1년 (역사적 연도는 제외)
            new_text = _bump_year(new_text, allowed_years)
            ops.append(splice.replace_inner(node, new_text))
            touched.add(node.id)
            n_label += 1
        if mapping:
            desc = ", ".join(f"제{a}기→제{b}기" for a, b in sorted(mapping.items()))
            report.add("자동확정", "문서 전체", f"기수 라벨 {n_label}곳", note=desc)

        # ── ③ THEAD 날짜 +1년 (기수 없는 헤더 — '2026.3.31 현재'류) ──
        n_head_date = 0
        for tv in views:
            if _skip_table(tv):
                continue
            for i in [i for i in range(tv.n) if tv.table.rows[i].in_thead]:
                for cell in tv.table.rows[i].cells:
                    if cell.tag == "TU" or cell.id in touched:
                        continue  # 기수 셀은 ②에서 이미 처리
                    new_text = _bump_year(cell.text, allowed_years)
                    if new_text != cell.text:
                        ops.append(splice.replace_inner(cell, new_text))
                        touched.add(cell.id)
                        n_head_date += 1
        if n_head_date:
            report.add("자동확정", "표 헤더", f"기간 날짜 +1년 {n_head_date}곳")

        # ── ④ TU 자동필드: AUNITVALUE·표시 연도 +1 (스파이크 V4 방식) ─
        n_tu = 0
        for node in doc.nodes.values():
            if node.tag != "TU":
                continue
            frag = doc.data[node.start:node.end]
            rolled = _YEAR_RE.sub(lambda m: str(int(m.group()) + 1).encode(), frag)
            if rolled != frag:
                ops.append(splice.Op(pos=node.start, start=node.start,
                                     end=node.end, payload=rolled))
                n_tu += 1
        if n_tu:
            report.add("자동확정", "자동 필드", f"TU 기간 +1년 {n_tu}곳")

        # ── ⑤ 손대지 않은 날짜 — 확인필요 목록 (역사적 날짜 오염 방지) ─
        for node in doc.nodes.values():
            if node.tag == "TU" or node.id in touched:
                continue
            for m in _DATE_RE.finditer(node.text):
                ctx = node.text[max(0, m.start() - 12):m.end() + 10]
                report.add("확인필요", node.id, "날짜 발견 (자동 치환 안 함)",
                           before=ctx.replace("\n", " ").strip(),
                           note="보고기간 표기면 +1년으로 직접 수정")
                break  # 노드당 1건만

    report.add("확인필요", "meta/문서정보", "보고서 기간·문서명",
               note="DART 편집기 [문서정보]에서 기수·기간 확인 (도구는 미변경)")
    report.painted = painted
    return ops, report


def _skip_table(tv: TableView) -> bool:
    from .check import _P
    return bool(_P.skip_section_re.search(norm(tv.sec_title)))


# ── 자동 이동 셀 색칠 ────────────────────────────────────────────────

MOVED_FILL = "DDEBF7"  # 연한 파랑 — 검증 색(빨강/노랑)과 구분


def paint_moved(xlsx_path: str, rolled_doc: Document, painted: list,
                row_positions: dict) -> int:
    """롤포워드가 자동으로 채운 값 셀에 배경색 — 이동분을 한눈에 식별."""
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill

    from .annotate_xlsx import _cell_positions
    from .check import use_profile as _up

    pos_of = _cell_positions(rolled_doc, row_positions)
    targets: set[tuple[str, int, int]] = set()

    def add(row_id: str, gc: int):
        loc = pos_of.get(row_id)
        if loc is not None:
            targets.add((loc[0], loc[1], 2 + gc))

    with _up(None):
        for spec in painted:
            kind = spec[0]
            if kind == "cells":
                for row_id, gc in spec[1]:
                    add(row_id, gc)
                continue
            t = rolled_doc.tables.get(spec[1])
            if t is None:
                continue
            tv = TableView(t, "")
            if kind == "table":
                rows = tv.body
            else:  # socie: 전기 블록(첫 기초~첫 기말) + 당기초 행
                opens = [i for i in tv.body if "기초" in norm(tv.row_label(i))
                         and any(tv.value(i, c) is not None
                                 for c in range(tv.ncols))]
                closes = [i for i in tv.body if "기말" in norm(tv.row_label(i))
                          and any(tv.value(i, c) is not None
                                  for c in range(tv.ncols))]
                if len(opens) < 2 or not closes:
                    continue
                rows = [i for i in tv.body if opens[0] <= i <= closes[0]]
                rows.append(opens[1])
            for i in rows:
                for gc, _cell in _numeric_cells(tv, i):
                    add(t.rows[i].id, gc)

    if not targets:
        return 0
    wb = load_workbook(xlsx_path)
    fill = PatternFill("solid", fgColor=MOVED_FILL)
    n = 0
    for sheet, r, c in targets:
        if sheet in wb.sheetnames:
            wb[sheet].cell(row=r, column=c).fill = fill
            n += 1
    wb.save(xlsx_path)
    return n


# ── 리포트 시트 ──────────────────────────────────────────────────────

ROLL_SHEET = "롤포워드"
_GRADE_ORDER = {"미처리": 0, "확인필요": 1, "자동확정": 2}


def write_roll_sheet(xlsx_path: str, report: RollReport):
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, PatternFill

    from .export_xlsx import FILL_HEAD, FONT, FONT_HEAD, FONT_TITLE

    wb = load_workbook(xlsx_path)
    if ROLL_SHEET in wb.sheetnames:
        del wb[ROLL_SHEET]
    ws = wb.create_sheet(ROLL_SHEET, 0)
    ws.column_dimensions["A"].hidden = True
    for col, w in (("B", 9), ("C", 26), ("D", 34), ("E", 32), ("F", 32), ("G", 44)):
        ws.column_dimensions[col].width = w
    rid = iter(range(1, 100000))

    def put(r, c, v, font=FONT, fill=None):
        ws.cell(row=r, column=1, value=f"!r{next(rid):05d}").font = FONT
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = font
        if fill:
            cell.fill = fill
        cell.alignment = Alignment(vertical="center", wrap_text=(c in (6, 7)))
        return cell

    put(2, 2, "롤포워드 결과 — 미처리·확인필요부터 위에서 아래로 처리하세요", FONT_TITLE)
    put(3, 2, "미처리 = 도구가 손대지 않은 표(수기 작성) · 확인필요 = 사람 확인 대상 · "
              "자동확정 = 결정론 치환 완료. 파란 배경 셀 = 자동으로 이동·채워진 값. "
              "당기 칸은 비어 있음 — 결산 수치를 입력한 뒤 check로 검산하세요.")
    r = 5
    for c, h in enumerate(["등급", "위치", "항목", "변경 전", "변경 후", "비고"], start=2):
        cell = put(r, c, h, FONT_HEAD)
        cell.fill = FILL_HEAD
    r += 1
    fill_of = {"미처리": PatternFill("solid", fgColor="FDE9E9"),
               "확인필요": PatternFill("solid", fgColor="FFF4DE")}
    for item in sorted(report.items, key=lambda x: _GRADE_ORDER.get(x.grade, 9)):
        put(r, 2, item.grade, fill=fill_of.get(item.grade))
        put(r, 3, item.where)
        put(r, 4, item.item)
        put(r, 5, item.before)
        put(r, 6, item.after)
        put(r, 7, item.note)
        r += 1
    wb.save(xlsx_path)
