"""두 DSD 파일 구조 비교 리포트."""
from __future__ import annotations

from .dsd_io import read_dsd
from .parse import parse_contents


def diff_dsd(path_a: str, path_b: str) -> list[str]:
    doc_a = parse_contents(read_dsd(path_a).contents)
    doc_b = parse_contents(read_dsd(path_b).contents)
    lines: list[str] = []

    ids_a, ids_b = set(doc_a.tables), set(doc_b.tables)
    if len(doc_a.tables) != len(doc_b.tables):
        lines.append(f"표 개수: {len(doc_a.tables)} → {len(doc_b.tables)}")

    for tid in sorted(ids_a & ids_b):
        ta, tb = doc_a.tables[tid], doc_b.tables[tid]
        if len(ta.rows) != len(tb.rows):
            lines.append(f"{tid}: 행 수 {len(ta.rows)} → {len(tb.rows)}")
            continue
        for ra, rb in zip(ta.rows, tb.rows):
            if len(ra.cells) != len(rb.cells):
                lines.append(f"{ra.id}: 셀 수 {len(ra.cells)} → {len(rb.cells)}")
                continue
            for ca, cb in zip(ra.cells, rb.cells):
                if ca.text != cb.text:
                    lines.append(f"{ca.id}: '{ca.text[:40]}' → '{cb.text[:40]}'")

    paras_a = {i: n for i, n in doc_a.nodes.items() if not i.startswith("t")}
    paras_b = {i: n for i, n in doc_b.nodes.items() if not i.startswith("t")}
    for pid in sorted(set(paras_a) & set(paras_b)):
        if paras_a[pid].text != paras_b[pid].text:
            lines.append(f"{pid}: 문단 변경 '{paras_a[pid].text[:40]}…' → '{paras_b[pid].text[:40]}…'")

    if not lines:
        lines.append("구조·텍스트 차이 없음")
    return lines
