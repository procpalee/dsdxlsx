"""Phase 0 스파이크 — DART 편집기 허용성 검증용 변형 DSD 4종 생성.

Track A(표·주석 편집, 롤포워드)가 쓸 바이트 조작 방식이 DART 편집기에서
수용되는지 확인하는 실험체를 만든다. compose/blockops의 산술 규칙은 이
결과(docs/DART호환성체크리스트.md)로 확정된다.

사용법 (repo 루트에서):
    PYTHONPATH=src uv run --no-editable python spike/make_variants.py <원본.dsd> [출력폴더]

변형:
    V1_표복제삽입  — 주석 표의 원시 바이트 복제 + 셀 텍스트 1개 치환 후 같은 주석 뒤에 삽입
    V2_표삭제      — 표가 2개 이상인 주석의 마지막 표 + 선행 구분자 제거
    V3_주석재번호  — 뒤쪽 주석 3개의 헤딩 번호를 +1 (신규 주석 삽입 시나리오)
    V4_기수롤      — '제 N 기' 라벨 +1 + TU 자동필드(AUNITVALUE·표시 연도) +1년
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from dsdxlsx import splice
from dsdxlsx.dsd_io import read_dsd, write_dsd
from dsdxlsx.export_xlsx import refined_sections
from dsdxlsx.parse import parse_contents


def _note_sections(doc):
    return [s for s in refined_sections(doc) if s.note_no is not None]


def _clean_note_tables(doc):
    """비래핑·TU/TE 없는 주석 표 (스파이크 대상 자격 = Track A 템플릿 자격)."""
    out = []
    for sec in _note_sections(doc):
        for block in sec.blocks:
            if block.kind != "table":
                continue
            t = block.item
            tags = {c.tag for r in t.rows for c in r.cells}
            if t.fixtable or ("TU" in tags) or ("TE" in tags):
                continue
            out.append((sec, t))
    return out


def _prev_sep(data: bytes, start: int) -> int:
    """블록 앞 구분자 측정 — 삭제 시 함께 제거해 관례 유지."""
    for sep in (b"\r\n\r\n", b"\r\n"):
        if data[start - len(sep):start] == sep:
            return start - len(sep)
    return start


def v1_clone_table(doc, data: bytes) -> tuple[bytes, str]:
    sec, t = _clean_note_tables(doc)[-1]
    frag = bytearray(data[t.start:t.end])
    # 본문 첫 라벨 셀에 마커 — Track A의 '복제+텍스트 치환' 합성 방식 그대로
    cell = next(c for r in t.rows if not r.in_thead for c in r.cells if c.text.strip())
    frag[cell.inner_start - t.start:cell.inner_end - t.start] = splice.escape_text(
        "(스파이크 복제) " + cell.text)
    op = splice.Op(pos=t.end, start=t.end, end=t.end,
                   payload=b"\r\n\r\n" + bytes(frag))
    desc = f"주석{sec.note_no} 표 {t.id} 복제(셀 1개 치환) 후 바로 뒤 삽입"
    return splice.apply(data, [op]), desc


def v2_delete_table(doc, data: bytes) -> tuple[bytes, str]:
    by_note: dict[int, list] = {}
    for sec, t in _clean_note_tables(doc):
        by_note.setdefault(sec.note_no, []).append(t)
    note_no, tables = next((n, ts) for n, ts in by_note.items() if len(ts) >= 2)
    t = tables[-1]
    start = _prev_sep(data, t.start)
    op = splice.Op(pos=start, start=start, end=t.end, payload=b"")
    desc = f"주석{note_no}의 마지막 표 {t.id} 삭제 (+선행 구분자)"
    return splice.apply(data, [op]), desc


def v3_renumber_notes(doc, data: bytes) -> tuple[bytes, str]:
    secs = _note_sections(doc)[-3:]  # 뒤쪽 3개 주석: N → N+1
    ops, moved = [], []
    for sec in secs:
        head = next(b.item for b in sec.blocks if b.kind == "p")
        n = sec.note_no
        new_text = re.sub(rf"(^\s*){n}\.", rf"\g<1>{n + 1}.", head.text, count=1)
        assert new_text != head.text, f"주석 {n} 헤딩 치환 실패: {head.text[:30]!r}"
        ops.append(splice.replace_inner(head, new_text))
        moved.append(f"{n}→{n + 1}")
    desc = "주석 헤딩 재번호 " + ", ".join(moved) + " (신규 주석 삽입 시나리오)"
    return splice.apply(data, ops), desc


def v4_roll_period(doc, data: bytes) -> tuple[bytes, str]:
    # ① '제 N 기' 라벨: 문서 최대 기수 N → N+1, N-1 → N (표기 공백 보존)
    gisu = set()
    for node in doc.nodes.values():
        for m in re.finditer(r"제\s*(\d{1,3})\s*기", node.text):
            gisu.add(int(m.group(1)))
    assert gisu, "기수 라벨을 찾지 못함"
    cur = max(gisu)
    mapping = {cur: cur + 1, cur - 1: cur}

    def bump(m: re.Match) -> str:
        n = int(m.group(2))
        return m.group(1) + str(mapping.get(n, n)) + m.group(3)

    ops = []
    n_label = 0
    for node in doc.nodes.values():
        if node.tag == "TU":
            continue  # TU는 아래 원시 바이트로 (readonly + 속성 동기 필요)
        new_text = re.sub(r"(제\s*)(\d{1,3})(\s*기)", bump, node.text)
        if new_text != node.text:
            ops.append(splice.replace_inner(node, new_text))
            n_label += 1

    # ② TU 자동필드: AUNITVALUE(YYYYMMDD)와 표시 텍스트의 연도를 +1
    n_tu = 0
    for node in doc.nodes.values():
        if node.tag != "TU":
            continue
        frag = data[node.start:node.end]
        rolled = re.sub(rb"(19|20)\d{2}",
                        lambda m: str(int(m.group(0)) + 1).encode(), frag)
        if rolled != frag:
            ops.append(splice.Op(pos=node.start, start=node.start,
                                 end=node.end, payload=rolled))
            n_tu += 1
    desc = (f"기수 {cur}→{cur + 1}·{cur - 1}→{cur} (라벨 {n_label}개) "
            f"+ TU 자동필드 연도 +1 ({n_tu}개)")
    return splice.apply(data, ops), desc


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    src = Path(sys.argv[1])
    outdir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).parent
    outdir.mkdir(parents=True, exist_ok=True)

    package = read_dsd(str(src))
    data = package.contents
    doc = parse_contents(data)

    variants = {
        "V1_표복제삽입": v1_clone_table,
        "V2_표삭제": v2_delete_table,
        "V3_주석재번호": v3_renumber_notes,
        "V4_기수롤": v4_roll_period,
    }
    print(f"원본: {src.name} ({len(data):,} bytes)")
    for name, fn in variants.items():
        new_data, desc = fn(doc, data)
        parse_contents(new_data)  # well-formed 재검증 (실패 시 예외)
        out = outdir / f"{name}.dsd"
        write_dsd(str(out), package, new_data)
        print(f"  {out.name}: {desc}")
        print(f"    ({len(new_data):,} bytes, 원본 대비 {len(new_data) - len(data):+,})")
    print("\n각 변형을 DART 편집기로 열어 체크리스트(스파이크 체크리스트.md)를 채워주세요.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
