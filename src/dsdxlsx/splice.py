"""바이트 스플라이스 엔진 — 보존 보장의 핵심.

편집 목록을 오프셋 내림차순으로 원본 bytes에 적용한다.
무수정 구간은 단 1바이트도 건드리지 않는다 (편집 0건 → 원본 그대로).
"""
from __future__ import annotations

from dataclasses import dataclass

from .model import Node, Row

ROW_SEP = b"\r\n\r\n"  # 파일 관례: TR 블록 사이 빈 줄


def escape_text(text: str) -> bytes:
    """수정된 노드의 새 텍스트 → 문서 관례 이스케이프.

    & < > " 이스케이프, '는 원문 그대로, 셀 내 줄바꿈 → 리터럴 &cr; (파일상 &amp;cr;).
    """
    s = text.replace("\r\n", "\n").replace("\r", "\n")
    s = (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
    s = s.replace("\n", "&amp;cr;")
    return s.encode("utf-8")


@dataclass
class Op:
    pos: int  # 적용 위치 (정렬 키)
    start: int
    end: int  # start==end 이면 순수 삽입
    payload: bytes


def replace_inner(node: Node, new_text: str) -> Op:
    payload = escape_text(new_text)
    return Op(pos=node.inner_start, start=node.inner_start, end=node.inner_end, payload=payload)


def delete_row(data: bytes, row: Row) -> Op:
    """TR 블록 제거 + 인접 빈 줄 정리."""
    start, end = row.start, row.end
    if data[end : end + len(ROW_SEP)] == ROW_SEP:
        end += len(ROW_SEP)
    elif data[start - len(ROW_SEP) : start] == ROW_SEP:
        start -= len(ROW_SEP)
    return Op(pos=start, start=start, end=end, payload=b"")


def clone_row(data: bytes, template: Row, new_texts: dict[int, str]) -> bytes:
    """템플릿 TR의 원시 바이트를 복제하고 셀 inner 텍스트만 치환.

    new_texts: 템플릿 셀 인덱스 → 새 텍스트. 속성·공백 관례는 자동 보존된다.
    """
    raw = bytearray(data[template.start : template.end])
    base = template.start
    for i in sorted(new_texts, reverse=True):
        cell = template.cells[i]
        raw[cell.inner_start - base : cell.inner_end - base] = escape_text(new_texts[i])
    return bytes(raw)


def insert_after_row(row: Row, tr_bytes: bytes) -> Op:
    return Op(pos=row.end, start=row.end, end=row.end, payload=ROW_SEP + tr_bytes)


def insert_before_row(row: Row, tr_bytes: bytes) -> Op:
    return Op(pos=row.start, start=row.start, end=row.start, payload=tr_bytes + ROW_SEP)


def apply(data: bytes, ops: list[Op]) -> bytes:
    """오프셋 내림차순 적용. 범위 겹침은 버그이므로 방어적으로 검출."""
    if not ops:
        return data
    ordered = sorted(ops, key=lambda o: (o.pos, o.start), reverse=True)
    for prev, cur in zip(ordered, ordered[1:]):
        if cur.end > prev.start:
            raise ValueError(f"편집 범위 겹침: [{cur.start},{cur.end}) vs [{prev.start},{prev.end})")
    out = bytearray(data)
    for op in ordered:
        out[op.start : op.end] = op.payload
    return bytes(out)
