"""조각 합성기 — 신규 표/문단의 바이트를 '같은 문서 조각의 복제'로만 만든다.

Phase 0 스파이크 확정 규칙(docs/DART호환성체크리스트.md):
- 같은 문서 표의 원시 바이트 복제 + inner 텍스트 치환은 편집기가 그대로 수용.
- WIDTH/HEIGHT 재계산 없음 → 템플릿 값 복제로 충분. 리터럴 태그 문자열 생성 금지.
- 동일 열수 한정 (열 수 변경 합성은 2차 스파이크 후 개방).

합성 구조: 템플릿 표를 경계 슬라이스로 재조립 —
    [<TABLE…><COLGROUP…><THEAD> 여는 부분] + 헤더 행들 + [</THEAD><TBODY> 구간]
    + 본문 행들 + [</TBODY></TABLE> 닫는 부분]
행 바이트는 splice.clone_row(원시 TR 복제 + 셀 텍스트 치환)로 만든다.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from . import splice
from .model import Document, Row, Table


class ComposeError(Exception):
    pass


@dataclass
class TableTemplate:
    table: Table
    unit_row: Row | None   # 전폭 헤더 행(단위 표기 등) — 셀 1개/전폭 COLSPAN
    head_row: Row          # ncols 셀 · 병합 없는 헤더 행
    body_row: Row          # TBODY · ACOPY=Y · 병합 없는 본문 행


@dataclass
class NewTableSpec:
    header_rows: list[dict[int, str]]  # grid 열 → 텍스트
    body_rows: list[dict[int, str]]
    unit: str | None = None
    where: str = ""                     # 에러 메시지용 위치
    # 합성되는 행과 1:1 시트 위치 [(시트, 행)] — row_positions 재대응용
    sheet_rows: list[tuple[str, int]] = field(default_factory=list)


def _full_width(row: Row, ncols: int) -> bool:
    return len(row.cells) == 1 or (row.cells and row.cells[0].colspan >= ncols)


def harvest(doc: Document, table: Table, where: str = "") -> TableTemplate:
    """템플릿 표에서 합성 원자를 수확. 부적격이면 사유와 함께 ComposeError."""
    t = table

    def bad(reason: str):
        raise ComposeError(f"{where}템플릿 표 {t.id} 부적격 — {reason}")

    if t.group_aclass is not None:
        bad(f"TABLE-GROUP({t.group_aclass}) 래핑 표 (표준 재무제표·표지류)")
    if t.in_insertion:
        bad("INSERTION/LIBRARY 삽입물 내부 표")
    if t.fixtable:
        bad("고정 표(AFIXTABLE=Y)")
    if t.irregular:
        bad("그리드 불규칙")
    tags = {c.tag for r in t.rows for c in r.cells}
    if "TU" in tags:
        bad("자동 필드(TU) 포함")
    if "TE" in tags:
        bad("K-GAAP 표준 셀(TE) 포함 — ACODE 오염 위험")
    head_rows = [r for r in t.rows if r.in_thead]
    body_rows = [r for r in t.rows if not r.in_thead]
    if not head_rows or not body_rows:
        bad("THEAD 또는 본문 행 없음")
    unit_row = next((r for r in head_rows if _full_width(r, t.ncols)), None)
    head_row = next(
        (r for r in reversed(head_rows)
         if len(r.cells) == t.ncols and all(c.colspan == 1 for c in r.cells)),
        None)
    if head_row is None:
        bad("병합 없는 전열 헤더 행이 없음 (원본=t### 로 다른 표 지정)")
    body_row = next(
        (r for r in body_rows
         if r.acopy and len(r.cells) == t.ncols
         and all(c.colspan == 1 and c.rowspan == 1 for c in r.cells)),
        None)
    if body_row is None:
        bad("복제 가능한 본문 행(ACOPY=Y·병합 없음)이 없음")
    return TableTemplate(table=t, unit_row=unit_row, head_row=head_row,
                         body_row=body_row)


def compose_table(data: bytes, tpl: TableTemplate, spec: NewTableSpec) -> bytes:
    """새 표 바이트 합성 — 모든 바이트는 템플릿 복제 출신."""
    t = tpl.table
    used = max((max(r) + 1 for r in spec.header_rows + spec.body_rows if r),
               default=0)
    if used > t.ncols:
        raise ComposeError(
            f"{spec.where}새 표가 {used}열을 쓰지만 템플릿 {t.id}은 {t.ncols}열 — "
            f"열 수가 다른 표는 아직 미지원 (동일 열수 템플릿을 원본=으로 지정)")
    if not spec.body_rows:
        raise ComposeError(f"{spec.where}새 표에 본문 행이 없음")
    if spec.unit and tpl.unit_row is None:
        raise ComposeError(
            f"{spec.where}단위 행을 만들 수 없음 — 템플릿 {t.id}에 전폭 헤더 행이 없음")

    head_rows = [r for r in t.rows if r.in_thead]
    body_all = [r for r in t.rows if not r.in_thead]
    prefix = data[t.start:head_rows[0].start]           # <TABLE><COLGROUP><THEAD>…
    mid = data[head_rows[-1].end:body_all[0].start]     # </THEAD><TBODY> 구간
    suffix = data[body_all[-1].end:t.end]               # </TBODY></TABLE>

    def row_bytes(row: Row, texts: dict[int, str]) -> bytes:
        # 템플릿 행이 병합 없는 전열 행이므로 셀 인덱스 == grid 열
        return splice.clone_row(data, row, {i: texts.get(i, "")
                                            for i in range(len(row.cells))})

    parts: list[bytes] = []
    if spec.unit:
        parts.append(row_bytes(tpl.unit_row, {0: f"(단위 : {spec.unit})"}))
    for texts in spec.header_rows:
        parts.append(row_bytes(tpl.head_row, texts))
    head_bytes = splice.ROW_SEP.join(parts) if parts else b""
    body_bytes = splice.ROW_SEP.join(
        row_bytes(tpl.body_row, texts) for texts in spec.body_rows)
    if not head_bytes:  # 헤더 0행: 템플릿 헤더 행을 빈 텍스트로 유지 (THEAD 필수 구조)
        head_bytes = row_bytes(tpl.head_row, {})
    return prefix + head_bytes + mid + body_bytes + suffix


def compose_paragraph(text: str) -> bytes:
    """새 P — 유일한 비복제 합성 (P는 관례 변형이 없음, 전 샘플 실측)."""
    return b"<P>" + splice.escape_text(text) + b"</P>"


def compose_note(no: int, title: str, parts: list[bytes]) -> bytes:
    """새 주석 = 헤딩 P('N. 제목') + 블록들. 블록 사이 구분자는 관례
    (<TABLE 앞만 빈 줄) — export의 _split_notes가 안전하게 인식하는 분리 헤딩."""
    out = compose_paragraph(f"{no}. {title}".strip())
    for block in parts:
        sep = b"\r\n\r\n" if block.lstrip()[:6] == b"<TABLE" else b"\r\n"
        out += sep + block
    return out
