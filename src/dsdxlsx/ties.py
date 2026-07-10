"""대사맵(Tie-Map) — 명시적 대사쌍 선언·제안·검증 (C8).

'이 두 위치의 금액은 같아야 한다'를 셀 주소가 아니라 **(구획, 행 라벨 n번째)**
selector로 선언한다 — 행 추가·삭제·기수 이월에도 살아남는 좌표계. 확정(confirmed)
쌍만 엄격 검증하고(불일치=warn), 발견(discovery)은 제안(suggested)만 만든다 —
기존 C4(금액 탐색)·C5(라벨 인덱스) 로직을 후보 생성기로 강등, 확정은 사람이
'대사맵' 시트에서 상태를 바꿔 `dsdxlsx ties save`로 저장한다.

- 저장소: profiles/<회사명>/ties.json (도구 관리, 고객사명 포함 → 커밋 금지)
- 행 단위 대사: selector에 열을 고정하지 않으면 양쪽 행이 공유하는 기간
  버킷(당기/전기)별로 각각 비교한다 — "FS 매출채권 = 주석5 차감후 장부금액"
  한 줄이 당기·전기 두 검증을 의미.
- 해석 불가(라벨 소멸·중복)는 침묵하지 않고 warn으로 표면화한다.
"""
from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from .check import (
    CheckReport, TableView, build_views, fmt, is_total_label, norm, use_profile,
)
from .model import Document
from .profile import Profile, company_dir

TIES_SCHEMA = 1
CHECK_KEY = "C8"

# 시트 표시 상태 ↔ 내부 상태
STATUS_KO = {"confirmed": "확정", "suggested": "제안", "suppressed": "제외"}
STATUS_OF_KO = {v: k for k, v in STATUS_KO.items()}


@dataclass(frozen=True)
class Selector:
    """대사 지점: 구획(재무제표명 또는 주석 번호 문자열) 안의 라벨 n번째 행."""

    section: str        # "재무상태표" | "4" (주석 번호)
    row: str            # 정규화(공백 제거) 행 라벨
    row_nth: int = 1    # 구획 안 같은 라벨의 n번째 (1-base)
    col: str = ""       # ""=행 단위(공유 버킷별 비교) | "당기"/"전기" | 정규화 열 헤더

    def key(self) -> tuple:
        return (self.section, self.row, self.row_nth, self.col)

    def human(self) -> str:
        sec = f"주석{self.section}" if self.section.isdigit() else self.section
        nth = f" #{self.row_nth}" if self.row_nth > 1 else ""
        col = f" × {self.col}" if self.col else ""
        return f"{sec} · {self.row}{nth}{col}"

    def to_dict(self) -> dict:
        d = {"section": self.section, "row": self.row}
        if self.row_nth != 1:
            d["row_nth"] = self.row_nth
        if self.col:
            d["col"] = self.col
        return d

    @staticmethod
    def from_dict(d: dict) -> "Selector":
        return Selector(section=str(d["section"]), row=d["row"],
                        row_nth=int(d.get("row_nth", 1)), col=d.get("col", ""))


@dataclass
class Tie:
    a: Selector
    b: Selector
    status: str = "suggested"      # confirmed | suggested | suppressed
    relation: str = "equal"        # equal | negate | abs
    note: str = ""
    origin: str = "manual"         # FS↔주석 | 주석↔주석 | manual
    confidence: str = ""           # high | low (제안용)

    @property
    def id(self) -> str:
        """selector 쌍에서 유도한 안정 ID — 시트 병합·중복 제거의 키."""
        raw = json.dumps(sorted([self.a.key(), self.b.key()]), ensure_ascii=False)
        return "T" + hashlib.sha1(raw.encode("utf-8")).hexdigest()[:8].upper()

    def pair_key(self) -> frozenset:
        return frozenset({self.a.key(), self.b.key()})

    def to_dict(self) -> dict:
        d = {"id": self.id, "status": self.status, "a": self.a.to_dict(),
             "b": self.b.to_dict(), "relation": self.relation}
        for k in ("note", "origin", "confidence"):
            if getattr(self, k):
                d[k] = getattr(self, k)
        return d

    @staticmethod
    def from_dict(d: dict) -> "Tie":
        return Tie(a=Selector.from_dict(d["a"]), b=Selector.from_dict(d["b"]),
                   status=d.get("status", "suggested"),
                   relation=d.get("relation", "equal"), note=d.get("note", ""),
                   origin=d.get("origin", "manual"),
                   confidence=d.get("confidence", ""))


# ── 저장소 ───────────────────────────────────────────────────────────

def ties_path(company: str, base_dir: Path | None = None) -> Path:
    return company_dir(company, base_dir) / "ties.json"


def load_ties(path: Path) -> list[Tie]:
    if not path.is_file():
        return []
    data = json.loads(path.read_text(encoding="utf-8"))
    return [Tie.from_dict(d) for d in data.get("ties", [])]


def save_ties(path: Path, ties: list[Tie], company: str = ""):
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"schema": TIES_SCHEMA, "company": company,
               "ties": [t.to_dict() for t in ties]}
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=1),
                    encoding="utf-8")


# ── 구획 맵·해석 ─────────────────────────────────────────────────────

def section_views(doc: Document) -> dict[str, list[TableView]]:
    """selector 구획명 → TableView 목록. 재무제표는 표준 명칭, 주석은 번호 문자열."""
    views, ctx = build_views(doc)
    out: dict[str, list[TableView]] = {}
    for tv in ctx["statements"]:
        out.setdefault(tv.sec_title, []).append(tv)
    for n, tvs in ctx["notes"].items():
        out[str(n)] = list(tvs)
    return out


@dataclass
class Hit:
    tv: TableView
    i: int

    @property
    def row_id(self) -> str:
        return self.tv.table.rows[self.i].id

    def value(self, c: int):
        return self.tv.value(self.i, c)

    def buckets(self) -> dict[str, int]:
        """기간 버킷 → 값이 있는 grid 열 (버킷당 첫 값 열)."""
        out: dict[str, int] = {}
        for c in range(self.tv.ncols):
            b = self.tv.period_bucket(c)
            if b and b not in out and self.value(c) is not None:
                out[b] = c
        return out

    def col_by_header(self, header_norm: str) -> int | None:
        for c in range(self.tv.ncols):
            if norm(self.tv.col_header(c)) == header_norm and self.value(c) is not None:
                return c
        return None


def resolve(sections: dict[str, list[TableView]], sel: Selector
            ) -> tuple[Hit | None, str]:
    """selector → 행 하나. (hit, 실패 사유) — 0건/중복은 해석 불가로 표면화."""
    tvs = sections.get(sel.section)
    if not tvs:
        return None, f"구획 '{sel.section}' 없음"
    matches: list[Hit] = []
    for tv in tvs:
        for i in tv.body:
            if norm(tv.row_label(i)) == sel.row:
                matches.append(Hit(tv, i))
    if not matches:
        return None, f"행 '{sel.row}' 없음"
    if sel.row_nth > len(matches):
        return None, f"행 '{sel.row}' {sel.row_nth}번째 없음 (총 {len(matches)}개)"
    return matches[sel.row_nth - 1], ""


def _nth_of(sections: dict, section: str, row_norm: str, hit: Hit) -> int:
    n = 0
    for tv in sections.get(section, []):
        for i in tv.body:
            if norm(tv.row_label(i)) == row_norm:
                n += 1
                if tv is hit.tv and i == hit.i:
                    return n
    return max(n, 1)


# ── 검증 (C8) ────────────────────────────────────────────────────────

@dataclass
class TiePoint:
    """검증된 비교 지점 하나 — 대사맵 시트의 한 행."""

    tie: Tie
    bucket: str                 # "당기"/"전기" 또는 열 헤더 (표시용)
    a: Hit | None = None
    b: Hit | None = None
    a_col: int | None = None
    b_col: int | None = None
    ok: bool | None = None      # None = 해석 불가
    diff: int | Decimal | None = None
    reason: str = ""


def _scaled(hit: Hit, c: int):
    v = hit.value(c)
    return None if v is None else v * hit.tv.unit


def _compare(tie: Tie, va, vb, unit_a: int, unit_b: int, tol: int):
    """단위 환산 후 관계 비교. 반환 (ok, diff)."""
    if tie.relation == "negate":
        vb = -vb
    if tie.relation == "abs":
        va, vb = abs(va), abs(vb)
    diff = va - vb
    allow = max(unit_a, unit_b) if unit_a != unit_b else tol
    return abs(diff) <= allow, diff


def verify_ties(doc: Document, ties: list[Tie], profile: Profile | None = None,
                report: CheckReport | None = None) -> list[TiePoint]:
    """confirmed 대사쌍 검증. report가 오면 C8 findings/counts에 반영."""
    points: list[TiePoint] = []
    with use_profile(profile) as prof:
        sections = section_views(doc)
        for tie in ties:
            if tie.status != "confirmed":
                continue
            hit_a, why_a = resolve(sections, tie.a)
            hit_b, why_b = resolve(sections, tie.b)
            if hit_a is None or hit_b is None:
                why = " / ".join(w for w in (why_a and f"A: {why_a}",
                                             why_b and f"B: {why_b}") if w)
                points.append(TiePoint(tie, "", hit_a, hit_b, reason=why))
                if report is not None:
                    report.tick(CHECK_KEY)
                    report.add(CHECK_KEY, "warn",
                               f"{tie.a.human()} ↔ {tie.b.human()}",
                               f"대사쌍 해석 불가 — {why} (라벨 변경·행 삭제 여부 확인)")
                continue
            pairs = _paired_cols(tie, hit_a, hit_b)
            if not pairs:
                points.append(TiePoint(tie, "", hit_a, hit_b,
                                       reason="비교할 공통 기간/열 없음"))
                if report is not None:
                    report.tick(CHECK_KEY)
                    report.add(CHECK_KEY, "warn",
                               f"{tie.a.human()} ↔ {tie.b.human()}",
                               "대사쌍에 비교할 공통 기간 열이 없음")
                continue
            for bucket, ca, cb in pairs:
                pt = TiePoint(tie, bucket, hit_a, hit_b, ca, cb)
                va, vb = _scaled(hit_a, ca), _scaled(hit_b, cb)
                pt.ok, pt.diff = _compare(tie, va, vb, hit_a.tv.unit,
                                          hit_b.tv.unit, prof.footing_tol)
                points.append(pt)
                if report is not None:
                    report.tick(CHECK_KEY)
                    if not pt.ok:
                        report.add(
                            CHECK_KEY, "warn",
                            f"{tie.a.human()} ↔ {tie.b.human()} ({bucket})",
                            f"확정 대사쌍 불일치: {fmt(hit_a.value(ca))} ≠ "
                            f"{fmt(hit_b.value(cb))} (차이 {fmt(pt.diff)})",
                        )
    return points


def resolve_points(doc: Document, ties_list: list[Tie],
                   profile: Profile | None = None) -> list[TiePoint]:
    """검증 판정 없이 표시용 지점만 해석 (대사맵 시트의 제안 행)."""
    points: list[TiePoint] = []
    with use_profile(profile):
        sections = section_views(doc)
        for tie in ties_list:
            hit_a, why_a = resolve(sections, tie.a)
            hit_b, why_b = resolve(sections, tie.b)
            if hit_a is None or hit_b is None:
                why = " / ".join(w for w in (why_a, why_b) if w)
                points.append(TiePoint(tie, "", hit_a, hit_b, reason=why))
                continue
            pairs = _paired_cols(tie, hit_a, hit_b)
            if not pairs:
                points.append(TiePoint(tie, "", hit_a, hit_b,
                                       reason="공통 기간 열 없음"))
                continue
            for bucket, ca, cb in pairs:
                points.append(TiePoint(tie, bucket, hit_a, hit_b, ca, cb))
    return points


def _paired_cols(tie: Tie, hit_a: Hit, hit_b: Hit) -> list[tuple[str, int, int]]:
    """비교 열 쌍 [(표시명, a열, b열)]. col 고정 selector면 그 열, 아니면 공유 버킷별."""
    if tie.a.col or tie.b.col:
        ca = _fixed_col(hit_a, tie.a.col)
        cb = _fixed_col(hit_b, tie.b.col)
        if ca is None or cb is None:
            return []
        return [(tie.a.col or tie.b.col, ca, cb)]
    ba, bb = hit_a.buckets(), hit_b.buckets()
    return [(k, ba[k], bb[k]) for k in ("당기", "전기") if k in ba and k in bb]


def _fixed_col(hit: Hit, col: str) -> int | None:
    if not col:
        buckets = hit.buckets()
        return next(iter(buckets.values()), None)
    if col in ("당기", "전기"):
        return hit.buckets().get(col)
    return hit.col_by_header(col)


# ── 제안 생성 (C4/C5 로직의 후보 생성기 강등) ─────────────────────────

def suggest_ties(doc: Document, existing: list[Tie],
                 profile: Profile | None = None) -> list[Tie]:
    """FS↔참조주석·주석↔주석의 금액 일치 쌍을 suggested tie로 제안.

    정밀 우선: FS는 '참조 표기된' 주석 안에서만, 주석 간은 (동일 정규화 라벨,
    동일 버킷, 동일 금액)일 때만. 전역 금액 탐색(우연 일치 위험)은 제안하지
    않는다 — 그건 C4 경고의 몫."""
    seen = {t.pair_key() for t in existing}
    out: list[Tie] = []
    with use_profile(profile) as prof:
        sections = section_views(doc)
        _, ctx = build_views(doc)
        _suggest_fs_note(sections, ctx, seen, out, prof)
        _suggest_note_note(sections, ctx, seen, out, prof)
    return out


def _add(out: list[Tie], seen: set, tie: Tie):
    if tie.pair_key() in seen or tie.a.key() == tie.b.key():
        return
    seen.add(tie.pair_key())
    out.append(tie)


def _sel_of(sections: dict, section: str, hit: Hit) -> Selector:
    row = norm(hit.tv.row_label(hit.i))
    return Selector(section=section, row=row,
                    row_nth=_nth_of(sections, section, row, hit))


def _suggest_fs_note(sections, ctx, seen: set, out: list[Tie], prof: Profile):
    note_nos = sorted(int(k) for k in sections if k.isdigit())
    for tv in ctx["statements"]:
        note_col = next((c for c in range(tv.ncols)
                         if "주석" in norm(tv.col_header(c))), None)
        for i in tv.body:
            refs = ([int(x) for x in re.findall(r"\d+", tv.text(i, note_col))
                     if str(int(x)) in sections] if note_col is not None else [])
            fs_hit = Hit(tv, i)
            fs_buckets = fs_hit.buckets()
            if not fs_buckets:
                continue
            bucket, c = next(iter(fs_buckets.items()))
            target = fs_hit.value(c) * tv.unit
            if target == 0:
                continue
            if refs:
                # 참조 열이 있는 회사: 참조된 주석의 최적 일치만 (정밀)
                for n in refs:
                    best = _best_note_match(sections[str(n)], bucket, target,
                                            norm(tv.row_label(i)), prof, tv.unit)
                    if best is None:
                        continue
                    score, note_hit, relation = best
                    _add(out, seen, Tie(
                        a=_sel_of(sections, tv.sec_title, fs_hit),
                        b=_sel_of(sections, str(n), note_hit),
                        relation=relation, origin="FS↔주석",
                        confidence="high" if score >= 2 else "low",
                        note=f"주석 {n} 참조 금액 일치",
                    ))
                    break  # 첫 참조 주석의 최적 일치만 — 다중 참조는 사람이 추가 확정
            else:
                # 참조 열이 없는 회사(K-GAAP 표준 재무제표 등): 전 주석 탐색 폴백.
                # 우연 일치 방지 — 소액 제외 + 라벨·합계행 신호 필수 + 유일 최고점만
                if abs(target) < prof.global_search_min:
                    continue
                ranked = []
                for n in note_nos:
                    best = _best_note_match(sections[str(n)], bucket, target,
                                            norm(tv.row_label(i)), prof, tv.unit)
                    if best is not None:
                        ranked.append((best[0], n, best[1], best[2]))
                if not ranked:
                    continue
                ranked.sort(key=lambda x: x[0], reverse=True)
                score, n, note_hit, relation = ranked[0]
                ambiguous = len(ranked) > 1 and ranked[1][0] == score
                if score < 2 or ambiguous:
                    continue  # 신호 부족·동점 다중 주석 — 사람/AI 판단 몫
                _add(out, seen, Tie(
                    a=_sel_of(sections, tv.sec_title, fs_hit),
                    b=_sel_of(sections, str(n), note_hit),
                    relation=relation, origin="FS↔주석",
                    confidence="high" if score >= 4 else "low",
                    note=f"주석 {n} 금액 일치 (참조 열 없는 양식 — 전역 탐색)",
                ))


def _best_note_match(tvs: list[TableView], bucket: str, target,
                     fs_label: str, prof: Profile,
                     target_unit: int = 1) -> tuple[int, Hit, str] | None:
    """주석 안에서 target과 일치하는 행 중 최선. (점수, hit, 관계) —
    점수: 합계·차감후류 +4, 라벨 정확 일치 +3, 포함 관계 +2.
    교차 단위(원 FS ↔ 천원 주석)는 굵은 단위의 절사 한도까지 일치로 인정."""
    ranked: list[tuple[int, Hit, str]] = []
    for tv in tvs:
        allow = max(tv.unit, target_unit) if tv.unit != target_unit else 0
        for i in tv.body:
            hit = Hit(tv, i)
            c = hit.buckets().get(bucket)
            if c is None:
                continue
            v = hit.value(c) * tv.unit
            if v == 0:
                continue
            if abs(v - target) <= allow:
                relation = "equal"
            elif abs(-v - target) <= allow:
                relation = "negate"
            else:
                continue
            label = norm(tv.row_label(i))
            score = 0
            if (is_total_label(label) or tv.row_is_total(i)
                    or prof.deduct_total_re.search(label)):
                score += 4
            if label == fs_label:
                score += 3
            elif fs_label and (fs_label in label or label in fs_label):
                score += 2
            ranked.append((score, hit, relation))
    if not ranked:
        return None
    return max(ranked, key=lambda x: x[0])


def _suggest_note_note(sections, ctx, seen: set, out: list[Tie], prof: Profile):
    """같은 (정규화 라벨, 버킷)이 서로 다른 주석에서 같은 금액 → 대사쌍 제안."""
    index: dict[str, list[tuple[int, Hit, dict]]] = {}
    for n, tvs in ctx["notes"].items():
        for tv in tvs:
            for i in tv.body:
                label = norm(tv.row_label(i))
                if (len(label) < 2 or label in prof.generic_labels
                        or is_total_label(label)):
                    continue
                hit = Hit(tv, i)
                buckets = hit.buckets()
                if buckets:
                    index.setdefault(label, []).append((n, hit, buckets))
    for label, entries in index.items():
        by_note: dict[int, tuple[Hit, dict]] = {}
        for n, hit, buckets in entries:
            by_note.setdefault(n, (hit, buckets))  # 주석당 첫 행만
        notes = sorted(by_note)
        for ai in range(len(notes)):
            for bi in range(ai + 1, len(notes)):
                na, nb = notes[ai], notes[bi]
                hit_a, ba = by_note[na]
                hit_b, bb = by_note[nb]
                shared = [k for k in ("당기", "전기") if k in ba and k in bb]
                if not shared:
                    continue
                ua, ub = hit_a.tv.unit, hit_b.tv.unit
                allow = max(ua, ub) if ua != ub else 0

                def _match(k):
                    va = hit_a.value(ba[k]) * ua
                    vb = hit_b.value(bb[k]) * ub
                    return va != 0 and abs(va - vb) <= allow

                if all(_match(k) for k in shared):
                    _add(out, seen, Tie(
                        a=_sel_of(sections, str(na), hit_a),
                        b=_sel_of(sections, str(nb), hit_b),
                        origin="주석↔주석", confidence="high",
                        note=f"'{label}' {'/'.join(shared)} 금액 일치",
                    ))


# ── 대사맵 시트 읽기·상태 병합 (쓰기는 annotate_xlsx.write_tie_sheet) ──
#
# 시트 열 배치 (annotate와 계약): A=!t행ID, B=상태, C=ID, D=대사A, E=A금액,
# F=관계, G=대사B, H=B금액, I=기간, J=판정, K=출처, L=비고, M=_data(JSON, 숨김)

SHEET_NAME = "대사맵"
_COL_STATUS, _COL_ID, _COL_NOTE, _COL_DATA = 2, 3, 12, 13


def read_tie_sheet(xlsx_path: str) -> list[Tie]:
    """대사맵 시트에서 사용자가 고친 상태·비고를 읽어 Tie 목록으로 복원.

    같은 tie의 여러 기간 행에 상태가 갈리면: 확정 > 제외 > 제안 우선."""
    from openpyxl import load_workbook

    wb = load_workbook(xlsx_path, read_only=True)
    try:
        if SHEET_NAME not in wb.sheetnames:
            return []
        ws = wb[SHEET_NAME]
        rank = {"confirmed": 2, "suppressed": 1, "suggested": 0}
        by_id: dict[str, Tie] = {}
        for row in ws.iter_rows(min_col=1, max_col=_COL_DATA):
            rid = row[0].value
            if not (isinstance(rid, str) and rid.startswith("!t")):
                continue
            raw = row[_COL_DATA - 1].value
            if not raw:
                continue
            try:
                tie = Tie.from_dict(json.loads(raw))
            except (json.JSONDecodeError, KeyError):
                continue
            status_ko = str(row[_COL_STATUS - 1].value or "").strip()
            tie.status = STATUS_OF_KO.get(status_ko, tie.status)
            tie.note = str(row[_COL_NOTE - 1].value or "").strip() or tie.note
            prev = by_id.get(tie.id)
            if prev is None or rank[tie.status] > rank[prev.status]:
                by_id[tie.id] = tie
        return list(by_id.values())
    finally:
        wb.close()


def merge_sheet_state(saved: list[Tie], sheet: list[Tie]) -> list[Tie]:
    """시트의 상태 편집을 반영한 유효 목록 — 시트가 이기고, 시트에 없는 저장분은 유지."""
    by_id = {t.id: t for t in saved}
    for t in sheet:
        by_id[t.id] = t
    return list(by_id.values())


def persistable(ties: list[Tie]) -> list[Tie]:
    """ties.json에 남길 것: 확정·제외만. 제안은 매 check마다 재생성되므로 저장 안 함."""
    return [t for t in ties if t.status in ("confirmed", "suppressed")]
