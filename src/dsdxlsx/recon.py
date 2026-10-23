"""DSD ↔ 조서(최종분석적검토 등) 대사 — recon.

DSD 재무제표·주석의 숫자 칸마다 조서 엑셀에서 같은 값을 찾는다. 조서 양식은 회사마다
달라(공시형 BS·IS 시트 / 계정형 BS + 공시형 FS·FN 시트 …) 시트·행 위치에 기대지 않는다.

찾는 순서 (칸 하나):
  1) 같은 이름 행(정규화 라벨 키 겹침) × 같은 기간 열(머리글 날짜·3개월) → 일치
  2) 같은 이름 행 여러 개의 합(하위 계정: '지분법적용투자주식(갑)' + '(을)')  → 일치(하위 합)
  3) 같은 이름 행은 있는데 값이 다름 (가장 가까운 조서 값을 보여 줌)          → 불일치
  4) 같은 이름 행이 없을 때만: 같은 범위의 다른 행(재무제표는 같은 기간 열,
     주석은 같은 주석 구획)                                                → 일치(다른 행)
  5) 아무것도 없음 / 같은 이름 행의 그 칸이 비어 있음                       → 상대 없음
비교는 단위(원/천원 반올림 ±1)·비율(27.27 ↔ 0.2727)·손실 표기 부호('분기순손실'을
양수로 적는 관례)를 맞춘다. 조서에 그 기간 열이 없으면(주석 시트에는 보통 당기만 있음)
전기 칸은 '대사 대상 아님'으로 따로 센다 (--prior 로 값 찾기 강제).
"""
from __future__ import annotations

import datetime as _dt
import difflib
import re
from dataclasses import dataclass, field
from decimal import Decimal

from openpyxl import load_workbook

from . import dart_client as D
from . import period as P
from .check import norm
from .model import Document
from .rollinterim import _CUR_MARK_RE, _PREV_MARK_RE, TableRec, _leaf, scan

# 조서 시트 자동 분류 (시트 이름 정규화 후 비교)
FS_SHEETS = {"bs", "is", "pl", "cf", "sce", "fs", "재무상태표", "손익계산서",
             "포괄손익계산서", "현금흐름표", "자본변동표", "재무제표"}
# DSD 재무제표 종류 → 조서 시트 이름 (같은 종류 시트가 없으면 재무제표 시트 전체)
KIND_SHEET_RE = {"BS": re.compile(r"^bs$|재무상태표", re.I),
                 "PL": re.compile(r"^(is|pl)$|손익계산서", re.I),
                 "CF": re.compile(r"^cf$|현금흐름표", re.I),
                 "SCE": re.compile(r"^sce$|자본변동표", re.I)}
GENERIC_FS_RE = re.compile(r"^fs$|^재무제표$", re.I)
NOTE_SHEET_RE = re.compile(r"주석|^fn$|^notes?$", re.I)

_DATE_RE = re.compile(r"(20\d{2})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})")
_3M_RE = re.compile(r"3\s*개\s*월")
_NOTE_TITLE_RE = re.compile(r"^\s*(\d{1,2})\s*\.\s*(\S.*)$")
_PAREN_RE = re.compile(r"\(([^()]*)\)")
_LOSS_RE = re.compile(r"손실|결손|차감|손$|감소|누계액|\(-\)")
_RESULT_RE = re.compile(r"(영업|순|법인세비용차감전순|법인세차감전순)(이익|손실)$")
# 기간이 아니라 계정인 머리글 ('당기순손익' 열이 당기 열로 잡히지 않게)
_ACCT_HDR_RE = re.compile(r"당기순|당기손익|전기오류")
# 금액이 아닌 열 (공정가치 수준 등) — 대사하지 않음
_SKIP_HDR_RE = re.compile(r"수준|레벨|level|주석", re.I)
_CF_DIR_RE = re.compile(r"취득|처분|지급|상환|증가|감소|배당|차입|납부|수취|유입|유출|회수|예치|인출")
# 행 라벨의 기간 표시 — '당기순이익' 같은 계정 이름은 빼고 '당기초·전분기말'만
_ROW_CUR_RE = re.compile(r"당\s*\d?\s*(분|반)?\s*기\s*(초|말)|\(당\)")
_ROW_PREV_RE = re.compile(r"전\s*\d?\s*(분|반)?\s*기\s*(초|말)|\(전\)")


# ── 라벨 키 ─────────────────────────────────────────────────────────

def _canon(text: str) -> str:
    t = re.sub(r"\(\s*주석[^)]*\)", "", text)          # '(주석3)', '(주석12,13)'
    t = _DATE_RE.sub("", t)
    t = re.sub(r"당\s*(분|반)\s*기", "당기", t)
    t = re.sub(r"전\s*(분|반)\s*기", "전기", t)
    k = D.canon(t)
    k = re.sub(r"의?(증가|감소)\((증가|감소)\)$", "", k)      # 현금흐름표 '미수금의 감소(증가)'
    k = k.replace("으로인한", "").replace("에의한", "")
    m = _RESULT_RE.search(k)
    if m:                                               # 영업손실 = 영업이익(손실) 행
        k = k[: m.start(2)] + "이익"
    return k


def label_keys(texts: list[str]) -> set[str]:
    """행의 글자 칸들 → 비교 키 집합: 전체 · 괄호 밖 · 괄호 안 각각의 정규화."""
    keys: set[str] = set()
    for raw in texts:
        s = raw.replace("\xa0", " ").replace("　", " ").strip()
        if not s:
            continue
        for p in [s, _PAREN_RE.sub("", s)]:
            k = _canon(p)
            if k and not k.isdigit():
                keys.add(k)
        for p in _PAREN_RE.findall(s):          # 괄호 안: '(주)'·'(유)' 회사 형태 표기는 버림
            k = _canon(p)
            if k and k not in ("주", "유", "재", "사", "합") and not k.isdigit():
                keys.add(k)
    return keys


# ── 조서 읽기 ───────────────────────────────────────────────────────

@dataclass
class WpRow:
    sheet: str
    row: int
    label: str
    keys: set[str]
    nums: dict[int, float]          # 열 번호 → 값


@dataclass
class WpSheet:
    name: str
    kind: str                        # fs | note
    rows: list[WpRow] = field(default_factory=list)
    col_date: dict[int, tuple[tuple[int, int, int], bool]] = field(default_factory=dict)
    blocks: list[tuple[int, str, int, int]] = field(default_factory=list)  # (번호, 제목키, 시작, 끝)


def _as_date(v) -> tuple[int, int, int] | None:
    if isinstance(v, (_dt.datetime, _dt.date)):
        return (v.year, v.month, v.day)
    if isinstance(v, str):
        m = _DATE_RE.search(v)
        if m:
            return (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def _is_num(v) -> bool:
    return isinstance(v, (int, float, Decimal)) and not isinstance(v, bool)


def _header_dates(grid: list[list]) -> dict[int, tuple[tuple[int, int, int], bool]]:
    """머리글(위 6행)의 날짜 → {열: (날짜, 3개월 여부)}. 날짜 없는 '3개월' 열은
    오른쪽(없으면 왼쪽) 가장 가까운 날짜 열의 날짜를 쓴다 ('6월 누적 | 최종3개월 | 9월 누적')."""
    dates: dict[int, tuple[int, int, int]] = {}
    three: set[int] = set()
    for r in grid[:6]:
        for c, v in enumerate(r):
            if c not in dates and (d := _as_date(v)):
                dates[c] = d
            if isinstance(v, str) and _3M_RE.search(v):
                three.add(c)
    out = {c: (d, c in three) for c, d in dates.items()}
    for c in three:
        if c in out:
            continue
        right = [x for x in sorted(dates) if x > c and x not in three]
        left = [x for x in sorted(dates, reverse=True) if x < c and x not in three]
        src = right[0] if right else (left[0] if left else None)
        if src is not None:
            out[c] = (dates[src], True)
    return out


def _sheet_kind(name: str, fs: set[str] | None, notes: set[str] | None) -> str | None:
    if fs is not None or notes is not None:
        if fs and name in fs:
            return "fs"
        if notes and name in notes:
            return "note"
        return None
    if norm(name).lower() in FS_SHEETS:
        return "fs"
    if NOTE_SHEET_RE.search(norm(name)):
        return "note"
    return None


def load_workpaper(path: str, fs_sheets: list[str] | None = None,
                   note_sheets: list[str] | None = None) -> list[WpSheet]:
    """조서 엑셀(계산값)을 읽어 재무제표·주석 시트만 행 단위로 색인."""
    import warnings
    with warnings.catch_warnings():          # 인쇄 영역·데이터 유효성 확장 경고는 대사와 무관
        warnings.simplefilter("ignore")
        wb = load_workbook(path, data_only=True, read_only=True)
    fs = set(fs_sheets) if fs_sheets else None
    notes = set(note_sheets) if note_sheets else None
    out: list[WpSheet] = []
    for ws in wb.worksheets:
        kind = _sheet_kind(ws.title, fs, notes)
        if kind is None:
            continue
        grid = [list(r) for r in ws.iter_rows(values_only=True)]
        sh = WpSheet(ws.title, kind, col_date=_header_dates(grid))
        ref_cols = {c for r in grid[:6] for c, v in enumerate(r)
                    if isinstance(v, str) and re.fullmatch(r"주석|note|notes?", norm(v), re.I)}
        for ri, r in enumerate(grid, start=1):
            texts = [str(v) for v in r if isinstance(v, str) and v.strip()]
            nums = {c: float(v) for c, v in enumerate(r) if _is_num(v) and c not in ref_cols}
            if not nums:
                for t in texts:
                    m = _NOTE_TITLE_RE.match(t.replace("\xa0", " "))
                    if m and not _DATE_RE.search(t):
                        sh.blocks.append((int(m.group(1)), _canon(m.group(2)), ri, 0))
                        break
                continue
            first = min(nums)
            label_texts = [str(v) for c, v in enumerate(r) if c < first and isinstance(v, str)]
            sh.rows.append(WpRow(ws.title, ri, " ".join(t.strip() for t in label_texts if t.strip()),
                                 label_keys(label_texts), nums))
        sh.blocks = [(n, k, s, (sh.blocks[i + 1][2] - 1 if i + 1 < len(sh.blocks) else len(grid)))
                     for i, (n, k, s, _) in enumerate(sh.blocks)]
        out.append(sh)
    wb.close()
    return out


# ── 값 비교 ─────────────────────────────────────────────────────────

def compare(v, w: float | None, unit: int, label: str) -> str | None:
    """DSD 값 v 와 조서 값 w → '같음' | '반올림' | '부호' | None."""
    if w is None:
        return None
    v = float(v)
    cands = [w]
    if unit > 1:
        cands.append(w / unit)                    # 조서 원 단위, DSD 천원
    for x in cands:
        if x == v:
            return "같음"
        if v and (abs(x - v) <= (1 if unit > 1 else 0) or (v == int(v) and round(x) == v)):
            return "반올림"
    if v != int(v) or abs(v) <= 100:              # 비율: 27.27 ↔ 0.2727, 4.6 ↔ 0.046
        if w and abs(w * 100 - v) <= 0.006:
            return "같음"
    for x in cands:
        if abs(x) == abs(v) or (unit > 1 and abs(abs(x) - abs(v)) <= 1):
            return "부호"
    return None


# ── DSD 칸 순회 ─────────────────────────────────────────────────────

@dataclass
class Item:
    sheet: str           # DSD 구획 (재무상태표 / 주석4 …)
    table: str
    row_id: str
    grid_col: int
    label: str
    col: str
    period: str          # 당기 | 전기
    dsd: float
    wp: float | None
    verdict: str         # 일치 | 불일치 | 부호 다름 | 상대 없음 | 대사 대상 아님
    how: str = ""        # 같은 행 · 하위 합 · 다른 행 · 반올림 · 부호 표기
    where: str = ""      # 조서 위치 '시트!B12'
    wp_label: str = ""
    unit: int = 1        # DSD 표 단위 (원=1, 천원=1000)

    @property
    def wp_in_unit(self) -> float | None:
        """조서 값을 DSD 단위로 (조서가 원 단위면 나눔)."""
        if self.wp is None:
            return None
        if self.unit > 1 and abs(self.wp / self.unit - self.dsd) < abs(self.wp - self.dsd):
            return self.wp / self.unit
        return self.wp


@dataclass
class ReconReport:
    period: str = ""
    items: list[Item] = field(default_factory=list)
    zero_skipped: int = 0
    warnings: list[str] = field(default_factory=list)

    def count(self, verdict: str) -> int:
        return sum(1 for x in self.items if x.verdict == verdict)


def _table_marker(rec: TableRec) -> str | None:
    tv = rec.tv
    rows = (tv.head or []) + tv.body[:1]
    for i in rows:
        t = norm(tv.text(i, 0))
        if not t or t.startswith("(단위"):
            continue
        if _CUR_MARK_RE.search(t) or t.startswith("당"):
            return "당기"
        if _PREV_MARK_RE.search(t) or t.startswith("전"):
            return "전기"
    return None


def _cell_periods(rec: TableRec, year: int | None = None) -> dict[tuple[int, int], str]:
    """(행, 열) → 당기/전기. 열 머리글 > 표 위 캡션 > 표 왼쪽 위 칸 > 행 라벨 순.
    행 라벨은 '당기초'·'전분기말' 같은 말 또는 '2025.01.01(기초자본)'의 연도로 본다."""
    tv = rec.tv
    out: dict[tuple[int, int], str] = {}
    numeric = {c for i in tv.body for c in range(tv.ncols) if tv.value(i, c) is not None}
    acct = {c for c in range(tv.ncols) if _ACCT_HDR_RE.search(norm(tv.col_header(c)))}
    cur = [c for c in rec.cur if c in numeric and c not in acct]
    prev = [c for c in rec.prev if c in numeric and c not in acct]
    table_mark = None
    if cur and not prev:                # '전기초 | 지분법손익 | 전기말' — 머리글 없는 열도 같은 기간
        table_mark = "당기"
    elif prev and not cur:
        table_mark = "전기"
    elif not cur and not prev:
        if rec.caption:
            table_mark = "당기" if rec.caption_is_cur else "전기"
        elif (lc := [c for c in rec.cur if c not in acct]) or [c for c in rec.prev if c not in acct]:
            table_mark = "당기" if lc else "전기"   # 왼쪽 위 칸이 '전기말'인 세로형 표
        else:
            table_mark = _table_marker(rec)
    row_mark = None
    for i in tv.body:
        lab = norm(tv.row_label(i))
        dm = _DATE_RE.search(lab)
        if dm and year:
            row_mark = "당기" if int(dm.group(1)) >= year else "전기"
        elif _ROW_PREV_RE.search(lab):
            row_mark = "전기"
        elif _ROW_CUR_RE.search(lab):
            row_mark = "당기"
        for c in range(tv.ncols):
            if c in cur:
                out[(i, c)] = "당기"
            elif c in prev:
                out[(i, c)] = "전기"
            else:
                out[(i, c)] = table_mark or row_mark or "당기"
    return out


def _target_date(per: P.Period | None, period: str, rec: TableRec):
    if per is None:
        return None
    y, m, d = per.end
    if period == "당기":
        return (y, m, d)
    if per.is_annual or rec.is_balance:
        return (y - 1, 12, 31)
    return (y - 1, m, d)


def _match_block(sheets: list[WpSheet], note_no: int | None, title: str) -> list[tuple[WpSheet, int, int]]:
    """DSD 주석 → 조서 주석 시트 구획(들). 제목 유사도 우선, 없으면 번호."""
    key = _canon(title)
    best, best_r = [], 0.0
    by_no = []
    for sh in sheets:
        if sh.kind != "note":
            continue
        for n, k, s, e in sh.blocks:
            r = difflib.SequenceMatcher(None, key, k).ratio() if key and k else 0.0
            if r > best_r + 1e-9:
                best, best_r = [(sh, s, e)], r
            elif abs(r - best_r) < 1e-9 and r > 0:
                best.append((sh, s, e))
            if note_no is not None and n == note_no:
                by_no.append((sh, s, e))
    if best_r >= 0.6:
        return best
    return by_no


def reconcile(doc: Document, sheets: list[WpSheet], include_prior: bool = False) -> ReconReport:
    per = P.detect_period(doc)
    rep = ReconReport(period=per.label if per else "(기간 미상)")
    if per is None:
        rep.warnings.append("DSD 보고기간을 못 읽어 조서 머리글 날짜를 쓰지 않고 행 안에서만 찾음")
    fs_rows = [r for sh in sheets if sh.kind == "fs" for r in sh.rows]
    if sheets and not any(r.nums for sh in sheets for r in sh.rows):
        rep.warnings.append("조서에서 숫자를 못 읽음 — 계산값이 없는 파일이면 엑셀에서 열어 저장한 뒤 다시 실행")
    sheet_of = {sh.name: sh for sh in sheets}

    recs, _ = scan(doc)
    for rec in recs:
        if rec.kind in ("OTHER", "RE"):
            continue
        tv = rec.tv
        if rec.kind == "NOTE":
            scope = []
            for sh, s, e in _match_block(sheets, rec.sec.note_no, rec.sec.note_title or ""):
                scope += [r for r in sh.rows if s <= r.row <= e]
            ctx = f"주석{rec.sec.note_no}" if rec.sec.note_no else "주석"
        else:
            kre = KIND_SHEET_RE.get(rec.kind)
            mine = [sh for sh in sheets if sh.kind == "fs" and kre and kre.search(norm(sh.name))]
            mine += [sh for sh in sheets if sh.kind == "fs" and GENERIC_FS_RE.search(norm(sh.name))]
            scope = [r for sh in mine for r in sh.rows] if mine else fs_rows
            ctx = rec.sec.title
        periods = _cell_periods(rec, per.year if per else None)
        unit = tv.unit
        for i in tv.body:
            texts = [tv.text(i, c) for c in range(tv.ncols)
                     if tv.text(i, c).strip() and tv.value(i, c) is None]
            label = " ".join(t.strip() for t in texts)
            keys = label_keys(texts)
            own = label_keys([tv.cell(i, c).text for c in range(tv.ncols)    # 병합 칸 제외
                              if tv.cell(i, c) is not None and tv.value(i, c) is None])
            for c in range(tv.ncols):
                v = tv.value(i, c)
                if v is None or _SKIP_HDR_RE.search(norm(tv.col_header(c))):
                    continue
                period = periods[(i, c)]
                target = _target_date(per, period, rec)
                is3m = bool(_3M_RE.search(_leaf(tv, c)))
                item = Item(ctx, tv.table.id, tv.table.rows[i].id, c, label[:40],
                            tv.col_header(c)[:30], period, float(v), None, "", unit=unit)
                if not scope:
                    if v != 0:
                        item.verdict = "조서에 없음"
                        rep.items.append(item)
                    continue
                _judge(item, v, keys, scope, sheet_of, target, is3m, unit,
                       include_prior or period == "당기", rec.kind, own)
                if item.verdict == "0 생략":
                    rep.zero_skipped += 1
                    continue
                rep.items.append(item)
    return rep


def _cols_for(sh: WpSheet, target, is3m: bool) -> list[int]:
    if target is None:
        return []
    cols = [c for c, (d, t3) in sh.col_date.items() if d == target and t3 == is3m]
    return sorted(cols)


def _ref(row: WpRow, col: int) -> str:
    from openpyxl.utils import get_column_letter
    return f"{row.sheet}!{get_column_letter(col + 1)}{row.row}"


def _judge(item: Item, v, keys: set[str], scope: list[WpRow], sheet_of: dict[str, WpSheet],
           target, is3m: bool, unit: int, search_undated: bool, kind: str = "",
           own: set[str] | None = None) -> None:
    def val_cols(row: WpRow) -> tuple[list[int], bool]:
        cols = _cols_for(sheet_of[row.sheet], target, is3m)
        if cols:
            return cols, True
        return (sorted(row.nums), False) if search_undated else ([], False)

    def finish(verdict, how, row=None, col=None, w=None):
        item.verdict, item.how = verdict, how
        if row is not None:
            item.where = _ref(row, col) if col is not None else f"{row.sheet}!{row.row}행"
            item.wp_label = row.label[:40]
        item.wp = w
        if verdict == "일치" and how.endswith("부호"):
            lab = norm(item.label)
            ok = _LOSS_RE.search(lab) or (kind in ("CF", "SCE") and _CF_DIR_RE.search(lab))
            item.verdict = "일치" if ok else "부호 다름"
            item.how = how.replace("부호", "부호 표기")

    # 1) 같은 이름 행 — 시트마다 키 겹침 최대인 행들. 병합 칸(세로 병합된 회사명 등)을 뺀
    #    그 행 자신의 이름으로도 한 번 더 고른다 ('(주)갑 | 기초' 아래 '증감' 행)
    def best_rows(ks: set[str]) -> list[WpRow]:
        best: dict[str, int] = {}
        by_sheet: dict[str, list[WpRow]] = {}
        for row in scope:
            n = len(ks & row.keys)
            if n and n >= best.get(row.sheet, 0):
                if n > best.get(row.sheet, 0):
                    best[row.sheet], by_sheet[row.sheet] = n, []
                by_sheet[row.sheet].append(row)
        return [r for rows in by_sheet.values() for r in rows]

    cands = best_rows(keys)
    if own and own != keys:
        cands += [r for r in best_rows(own) if r not in cands]
    any_dated = any(_cols_for(sheet_of[sh], target, is3m) for sh in {r.sheet for r in scope})
    for row in cands:
        cols, _ = val_cols(row)
        for col in cols:
            k = compare(v, row.nums.get(col), unit, item.label)
            if k:
                return finish("일치", "같은 행" + ("" if k == "같음" else f" · {k}"), row, col,
                              row.nums.get(col))

    if not any_dated and not search_undated:
        item.verdict = "대사 대상 아님"
        return

    if v == 0:
        item.verdict = "0 생략"
        for row in cands:
            cols, dated = val_cols(row)
            w = next((row.nums[c] for c in cols if c in row.nums), None) if dated else None
            if w:
                return finish("불일치", "같은 행", row, cols[0], w)
        return

    # 2) 하위 합 — 같은 이름 행이 여럿이면 기간 열 합
    if len(cands) > 1:
        tot, used = 0.0, []
        for row in cands:
            cols, dated = val_cols(row)
            if dated and cols[0] in row.nums:
                tot += row.nums[cols[0]]
                used.append(row)
        if len(used) > 1 and (k := compare(v, tot, unit, item.label)):
            return finish("일치", f"하위 합 {len(used)}행" + ("" if k == "같음" else f" · {k}"),
                          used[0], None, tot)

    # 3) 같은 이름 행은 있는데 값이 다름 — 가장 가까운 값을 보여 줌
    near = []
    for row in cands:
        cols, _ = val_cols(row)
        for c in cols:
            if c in row.nums:
                w = row.nums[c]
                gap = min(abs(w - float(v)), abs(w / unit - float(v)) if unit > 1 else abs(w - float(v)))
                near.append((gap, row, c, w))
    if near:
        _, row, c, w = min(near, key=lambda x: x[0])
        return finish("불일치", "같은 행", row, c, w)
    # 4) 같은 이름 행이 없거나 비어 있을 때만 — 같은 범위 다른 행
    for row in scope:
        if row in cands:
            continue
        cols, _ = val_cols(row)
        for col in cols:
            k = compare(v, row.nums.get(col), unit, item.label)
            if k in ("같음", "반올림"):
                return finish("일치", "다른 행" + ("" if k == "같음" else f" · {k}"), row, col,
                              row.nums.get(col))

    if cands:
        return finish("상대 없음", "조서 칸 빈칸", cands[0])
    finish("상대 없음", "")


# ── 결과 엑셀 ───────────────────────────────────────────────────────

VERDICTS = ["일치", "불일치", "부호 다름", "상대 없음", "조서에 없음", "대사 대상 아님"]
ISSUE_VERDICTS = ("불일치", "부호 다름", "상대 없음")


def summary_rows(rep: ReconReport) -> list[tuple[str, dict[str, int]]]:
    order: list[str] = []
    counts: dict[str, dict[str, int]] = {}
    for x in rep.items:
        if x.sheet not in counts:
            order.append(x.sheet)
            counts[x.sheet] = {v: 0 for v in VERDICTS}
        counts[x.sheet][x.verdict] += 1
    return [(k, counts[k]) for k in order]


def write_recon_report(rep: ReconReport, out_path: str, dsd_path: str, wp_path: str) -> dict:
    import os

    from openpyxl import Workbook
    from openpyxl.styles import Alignment

    from .report_xlsx import (
        _BORDER, _FILL_BAD, _FONT_BAD, _NUMFMT, FONT, FONT_LINK, FONT_TITLE, _write_header,
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "요약"
    ws.cell(row=2, column=2, value="DSD ↔ 조서 대사 결과").font = FONT_TITLE
    info = [("DSD", os.path.basename(dsd_path)), ("조서", os.path.basename(wp_path)),
            ("보고기간", rep.period), ("0 칸 생략", rep.zero_skipped)]
    r = 4
    for k, v in info:
        ws.cell(row=r, column=2, value=k).font = FONT
        ws.cell(row=r, column=3, value=v).font = FONT
        r += 1
    r += 1
    _write_header(ws, r, ["", "구획"] + VERDICTS)
    ws.cell(row=r, column=1).value = None
    r += 1
    for sheet, cnt in summary_rows(rep):
        vals = [sheet] + [cnt[v] for v in VERDICTS]
        for c, v in enumerate(vals, start=2):
            cell = ws.cell(row=r, column=c, value=v)
            bad = c - 2 in (2, 3) and isinstance(v, int) and v > 0
            cell.font = _FONT_BAD if bad else FONT
            cell.border = _BORDER
            if c > 2:
                cell.alignment = Alignment(horizontal="center")
        r += 1
    r += 1
    for line in ["일치: 조서에서 같은 값을 찾음(찾은 방법 열 참고) / 불일치: 같은 이름 행은 있는데 값이 다름",
                 "부호 다름: 크기만 같음 / 상대 없음: 조서에 같은 이름 행이 없거나 칸이 비어 있음",
                 "조서에 없음: 조서에 그 주석 구획이 없음 / 대사 대상 아님: 조서에 그 기간 열이 없음(주로 전기)",
                 *rep.warnings]:
        ws.cell(row=r, column=2, value=line).font = FONT
        r += 1
    ws.column_dimensions["B"].width = 18
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 12

    heads = ["구획", "표", "행", "열", "기간", "DSD 값", "조서 값", "차이", "결과", "찾은 방법",
             "조서 위치", "조서 행"]
    widths = [12, 7, 30, 22, 6, 18, 18, 14, 10, 18, 16, 30]
    wp_abs = os.path.abspath(wp_path)
    n_issue = 0
    for title, items in (("확인 필요", [x for x in rep.items if x.verdict in ISSUE_VERDICTS]),
                         ("대사결과", rep.items)):
        ws = wb.create_sheet(title)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + i)].width = w
        _write_header(ws, 1, heads)
        ws.freeze_panes = "A2"
        rr = 2
        for x in items:
            wpv = x.wp_in_unit
            diff = (x.dsd - wpv) if wpv is not None and x.verdict == "불일치" else None
            vals = [x.sheet, x.table, x.label, x.col.replace("\n", " "), x.period, x.dsd, wpv,
                    diff, x.verdict, x.how, x.where, x.wp_label]
            issue = x.verdict in ISSUE_VERDICTS
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=rr, column=c, value=v)
                cell.font = _FONT_BAD if issue and c == 9 else FONT
                cell.border = _BORDER
                if c in (6, 7, 8) and isinstance(v, (int, float)):
                    cell.number_format = _NUMFMT if float(v) == int(v) else "#,##0.00"
                if issue:
                    cell.fill = _FILL_BAD
            if "!" in x.where and not x.where.endswith("행"):
                sh, addr = x.where.split("!", 1)
                link = ws.cell(row=rr, column=11)
                link.hyperlink = f"{wp_abs}#'{sh}'!{addr}"
                link.font = FONT_LINK
            rr += 1
        ws.auto_filter.ref = f"A1:L{max(rr - 1, 1)}"
        if title == "확인 필요":
            n_issue = len(items)
    from .export_xlsx import _T
    if _T.get("modernize"):
        _T["modernize"](wb)
    wb.properties.creator = "이재현"
    wb.save(out_path)
    return {"items": len(rep.items), "issues": n_issue}
