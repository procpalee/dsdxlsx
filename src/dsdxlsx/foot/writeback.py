"""writeback — 풋팅 워크북에서 고친 셀 값을 원본 DSD 에 반영한다.

전제: 워크북에 mapsheet(_map/_cellmap)가 심겨 있다(dsdxlsx foot 으로 만든 워크북). 원본 DSD 를 그
안에서 꺼내 같은 파서로 다시 읽으므로(결정론 ID) 노드가 그대로 해석된다.
바뀐 셀만 splice.replace_inner 로 바이트 치환 — 무수정 구간은 1바이트도
건드리지 않는다(dsdxlsx import 와 같은 보존 계약).

범위: **셀 값 편집만**. 엑셀에서 행·표를 추가/삭제/재배열한 것은 반영하지 않고
(행 밀림 가드) 오류로 막는다 — 구조 편집은 dsdxlsx `export → 마커 → import`.

종료코드: 0 = 반영 완료(0건 포함), 1 = 반영했으나 일부 셀 거부·평문 강등 경고,
2 = 오류(셀맵 없음·행 밀림·노드 충돌·파싱 실패) → 파일 미생성.
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from . import mapsheet
from .emit import cell_content
from .tableview import classify_ext
from .. import numbers, splice
from ..dsd_io import read_dsd, write_dsd
from ..parse import LABEL_TAGS, parse_contents


@dataclass
class Edit:
    addr: str        # 'BS!D12'
    node_id: str
    old: str
    new: str


@dataclass
class WritebackResult:
    rc: int
    out: Path | None
    source: str
    edits: list[Edit] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


class WritebackError(Exception):
    pass


# ── 값 비교·변환 ─────────────────────────────────────────────────────

def _addr(sheet: str, row: int, col: int) -> str:
    return f"{sheet}!{get_column_letter(col)}{row}"


def _as_decimal(v) -> Decimal | None:
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float, Decimal)):
        try:
            return Decimal(str(v))
        except InvalidOperation:
            return None
    return None


def unchanged(cur, orig_text: str, note_col: bool) -> bool:
    """셀 현재값이 '기록 당시 값'과 같은가 — emit.cell_content 와 같은 규칙으로 기대값을
    만들어 비교한다('-'↔0, %(분수 저장), '원' 접미가 정확히 왕복)."""
    expected, _ = cell_content(orig_text, note_col)
    ce, cc = _as_decimal(expected), _as_decimal(cur)
    if ce is not None or cc is not None:
        if ce is None or cc is None:
            # 한쪽만 숫자: 텍스트 표현으로 재비교 (예: 라벨 셀에 숫자 입력)
            return (str(cur) if cur is not None else "") == (str(expected) if expected is not None else "")
        return ce == cc
    # 공백뿐인 문단/셀은 None 으로 기록되므로(emit) 공백 ↔ None 은 같은 것으로 본다
    a = "" if cur is None else str(cur)
    b = "" if expected is None else str(expected)
    if not a.strip() and not b.strip():
        return True
    return a == b


def _fmt_number(v: Decimal, places: int, like: str = "") -> str:
    """숫자 → 문서 관례 표기. places>0 이면 소수 자릿수를 고정(끝 0 유지 — '5.00%').
    음수는 원문 관례(△/▲ 또는 괄호) — numbers.to_dsd_text 와 같은 규칙."""
    if places <= 0 and v != v.to_integral_value():
        places = _places_of(v)
    negative = v < 0
    mag = -v if negative else v
    text = f"{mag:,.{places}f}" if places > 0 else f"{int(mag):,}"
    if not negative:
        return text
    sym = like.strip()[:1]
    return f"{sym}{text}" if sym in ("△", "▲") else f"({text})"


def to_node_text(value, orig_text: str, note_col: bool) -> str:
    """엑셀 셀 값 → DSD 노드 텍스트 (원문 표기 관례 유지)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    orig = orig_text.strip()
    d = _as_decimal(value)
    if d is None or note_col:
        return str(value)
    # 원문이 숫자 표기였는가 (텍스트 라벨에 숫자를 넣은 경우는 그대로 콤마 표기)
    orig_num = classify_ext(orig) if orig else None
    suffix = numbers.suffix_of(orig)
    if suffix == "%":
        pct = d * 100
        places = max(numbers.decimal_places(orig), _places_of(pct))
        return _fmt_number(pct.quantize(Decimal(1).scaleb(-places)) if places else pct,
                           places, orig) + "%"
    if orig_num is not None and numbers.classify(orig) is None and orig.endswith("원"):
        # '10,000원' 형 확장 표기
        return _fmt_number(d, numbers.decimal_places(orig[:-1]), orig) + "원"
    if d == 0 and (orig == "-" or orig_num is not None):
        return "-"
    places = numbers.decimal_places(orig) if orig_num is not None else 0
    return _fmt_number(d, max(places, _places_of(d)) if places else _places_of(d), orig) + suffix


def _places_of(d: Decimal) -> int:
    t = format(d.normalize(), "f")
    return len(t.split(".")[1]) if "." in t else 0


# ── 본체 ─────────────────────────────────────────────────────────────

def _read_current(xlsx: Path, entries: list[mapsheet.MapEntry]) -> dict[tuple[str, int, int], object]:
    """매핑된 (sheet,row,col) 의 현재 값만 읽는다 (read_only 순회)."""
    wanted: dict[str, dict[int, set[int]]] = {}
    for e in entries:
        wanted.setdefault(e.sheet, {}).setdefault(e.row, set()).add(e.col)
    out: dict[tuple[str, int, int], object] = {}
    wb = load_workbook(str(xlsx), read_only=True, data_only=True)
    try:
        for sheet, rows in wanted.items():
            if sheet not in wb.sheetnames:
                raise WritebackError(f"시트 '{sheet}' 가 워크북에 없습니다 (시트 삭제/이름 변경은 반영 미지원)")
            ws = wb[sheet]
            max_row, max_col = max(rows) + 40, max(max(cs) for cs in rows.values())
            # 매핑 범위 전체(+여유 40행)를 읽어 둔다 — 행 밀림 가드가 오프셋 위치를 살핀다
            for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_row,
                                                     max_col=max_col, values_only=True), start=1):
                for c in range(1, max_col + 1):
                    v = row[c - 1] if c - 1 < len(row) else None
                    if v is not None or c in rows.get(r_idx, ()):
                        out[(sheet, r_idx, c)] = v
    finally:
        wb.close()
    return out


def _shift_guard(entries: list[mapsheet.MapEntry], current: dict, doc) -> list[str]:
    """행 밀림 감지 — 표별로 라벨(첫 열 텍스트) 시퀀스가 통째로 어긋나면 삽입/삭제로 본다.
    같은 시트 안에서 원래 라벨 시퀀스가 ±k 행 오프셋에 나타나면 그 k 를 알려준다."""
    by_table: dict[tuple[str, str], list[mapsheet.MapEntry]] = {}
    for e in entries:
        if e.kind != "cell":
            continue
        tid = e.node_id.split(".")[0]
        by_table.setdefault((e.sheet, tid), []).append(e)
    errors: list[str] = []
    for (sheet, tid), es in by_table.items():
        first_col = min(e.col for e in es)
        labels = [e for e in es if e.col == first_col and e.orig_text.strip()
                  and e.orig_text.strip() != "-"
                  and classify_ext(e.orig_text.strip()) is None]
        if len(labels) < 3:
            continue
        mism = sum(1 for e in labels
                   if not unchanged(current.get((sheet, e.row, e.col)), e.orig_text, e.note_col))
        if mism < 3 or mism < len(labels) * 0.5:
            continue
        # 오프셋 탐색: 어긋난 라벨들이 어디로 갔나 (삽입 지점 아래 행만 밀린다)
        moved = [e for e in labels
                 if not unchanged(current.get((sheet, e.row, e.col)), e.orig_text, e.note_col)][:3]
        head = moved[0]
        found = None
        for k in range(-30, 31):
            if k == 0:
                continue
            ok = all(unchanged(current.get((sheet, e.row + k, e.col)), e.orig_text, e.note_col)
                     for e in moved)
            if ok:
                found = k
                break
        where = f"{sheet} 표 {tid} (라벨 '{head.orig_text.strip()[:20]}' 부근)"
        if found is not None:
            errors.append(f"{where}: 행이 {found:+d}행 밀렸습니다 — 행 삽입/삭제는 반영 미지원 "
                          f"(구조 편집은 dsdxlsx export → 마커 → import)")
        else:
            errors.append(f"{where}: 라벨 열이 원본과 대부분 다릅니다({mism}/{len(labels)}) — "
                          f"행 이동·삭제로 보여 반영을 중단합니다")
    return errors


def default_output(xlsx: str | Path) -> Path:
    """출력 DSD 기본 경로 - 워크북 폴더의 <원본 DSD 이름>_수정.dsd."""
    meta, _, _ = mapsheet.read(str(xlsx))
    return Path(xlsx).with_name(f"{Path(str(meta.get('source', '원본.dsd'))).stem}_수정.dsd")


def writeback(xlsx: str | Path, out: str | Path | None = None) -> WritebackResult:
    xlsx = Path(xlsx)
    meta, dsd_bytes, entries = mapsheet.read(str(xlsx))
    source = str(meta.get("source", "원본.dsd"))
    result = WritebackResult(rc=0, out=None, source=source)

    package = read_dsd(io.BytesIO(dsd_bytes))
    contents = package.contents
    doc = parse_contents(contents)
    current = _read_current(xlsx, entries)

    shift_errors = _shift_guard(entries, current, doc)
    if shift_errors:
        result.errors.extend(shift_errors)
        result.rc = 2
        return result

    # 노드별로 묶어 충돌 검사 (원문 시트 + 개별 시트 이중 매핑)
    pending: dict[str, tuple[str, str, mapsheet.MapEntry]] = {}  # node_id → (addr, new_text, entry)
    for e in entries:
        cur = current.get((e.sheet, e.row, e.col))
        if unchanged(cur, e.orig_text, e.note_col):
            continue
        addr = _addr(e.sheet, e.row, e.col)
        node = doc.nodes.get(e.node_id)
        if node is None:
            result.errors.append(f"{addr}: 노드 {e.node_id} 를 원본에서 찾지 못함 (셀맵 손상)")
            continue
        if isinstance(cur, str) and cur.startswith("="):
            result.warnings.append(f"{addr}: 수식 셀 — 값 캐시가 없어 반영 불가 (엑셀에서 저장 후 재시도)")
            continue
        new_text = to_node_text(cur, e.orig_text, e.note_col)
        if node.readonly or node.tag in LABEL_TAGS:
            result.warnings.append(f"{addr}: 자동 필드/제목({node.tag})은 수정 불가 — 무시 ('{e.orig_text[:20]}')")
            continue
        if e.partial:
            # 문단을 잘라 썼던 경우: 원 문단 안의 그 조각만 치환
            if e.orig_text and node.text.count(e.orig_text) == 1:
                new_text = node.text.replace(e.orig_text, new_text, 1)
            else:
                result.warnings.append(f"{addr}: 분할 문단({e.node_id})의 원 조각을 유일하게 찾지 못해 무시")
                continue
        if e.node_id in pending:
            prev_addr, prev_text, _ = pending[e.node_id]
            if prev_text != new_text:
                result.errors.append(
                    f"{prev_addr} 와 {addr} 가 같은 칸({e.node_id})을 서로 다르게 고쳤습니다: "
                    f"'{prev_text[:20]}' vs '{new_text[:20]}' — 한 곳만 고치세요")
            continue
        pending[e.node_id] = (addr, new_text, e)

    if result.errors:
        result.rc = 2
        return result

    ops = []
    for node_id, (addr, new_text, e) in pending.items():
        node = doc.nodes[node_id]
        if node.complex:
            result.warnings.append(f"{addr}: 인라인 서식(SPAN 등)이 있던 칸 — 평문으로 강등해 반영")
        ops.append(splice.replace_inner(node, new_text))
        result.edits.append(Edit(addr, node_id, node.text, new_text))

    new_contents = splice.apply(contents, ops) if ops else contents
    if ops:
        try:
            parse_contents(new_contents)  # 정형성 재검증
        except Exception as e:  # noqa: BLE001
            result.errors.append(f"반영 결과 XML 파싱 실패: {e}")
            result.rc = 2
            return result

    out_path = Path(out) if out else xlsx.with_name(f"{Path(source).stem}_수정.dsd")
    write_dsd(str(out_path), package, new_contents)
    result.out = out_path
    result.rc = 1 if result.warnings else 0
    return result


def print_report(xlsx: Path, r: WritebackResult) -> None:
    print(f"\n[{xlsx.name}] 원본 {r.source}")
    for e in r.edits:
        print(f"  {e.addr}  {e.node_id}  '{e.old[:30]}' → '{e.new[:30]}'")
    for w in r.warnings:
        print(f"  ⚠ {w}")
    for er in r.errors:
        print(f"  ✖ {er}")
    if r.rc == 2:
        print("  반영 중단 — 파일을 만들지 않았습니다")
    else:
        print(f"  수정 {len(r.edits)}건 · 경고 {len(r.warnings)}건")
        print(f"  → {r.out.name}")
