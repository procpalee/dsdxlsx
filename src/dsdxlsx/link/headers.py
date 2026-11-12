"""정산표 주석 시트의 두 줄 머리글 고치기(`link fix-headers`) - 0.6.1 이하로 만든 정산표용.

주석 표의 머리글이 두 줄이면('당기말' 아래 '소유주식수 · 지분율') 0.6.1 이하의 `link make`(와 그 전 스킬 스크립트)는
둘째 열부터 위 묶음 이름을 빼고 적었다('당기말 소유주식수 | 지분율 (%)'). 그러면 DSD 를 채울 때 그 열이 당기인지
몰라 옆 열 값을 넣는다. 여기서는 뼈대 DSD 의 표마다 정산표 주석 시트에서 그 표의 머리글 줄을 찾아, 묶음 이름이 빠진
칸만 DSD 의 전체 머리글('당기말 지분율 (%)')로 바꾼다. 칸 글에 기간이 이미 있으면('금액' 아래 '당기 · 전기') 그대로 둔다.
값 · 수식 · 서식은 건드리지 않는다.

머리글 줄은 이렇게 알아본다: 칸마다 글이 DSD 의 전체 머리글이나 맨 아래 머리글과 같고(공백 무시), 바로 아래 줄의
첫 칸이 DSD 표의 첫 행 이름과 같다. 고칠 칸이 하나도 없는 줄은 고친 목록에 넣지 않는다.
"""
from __future__ import annotations

from openpyxl import load_workbook

from ..guard import InputError
from .core import needs_group, squeeze, tidy


def _risky_tables(doc: dict) -> list:
    """머리글이 두 줄 이상이고 묶음 아래 열이 둘 이상인 주석 표: (구획, 표, 전체 머리글, 맨 아래 머리글, 첫 행 이름)."""
    out = []
    for sec in doc["sections"]:
        if sec["kind"] != "note":
            continue
        for b in sec["blocks"]:
            if b["type"] != "table" or b["ncols"] < 2 or sum(1 for r in b["rows"] if r.get("head")) < 2:
                continue
            heads = {c["col"]: c.get("header") or "" for c in b["columns"]}
            full = [" ".join(heads.get(i, "").split("/")) for i in range(b["ncols"])]
            leaf = [heads.get(i, "").split("/")[-1] for i in range(b["ncols"])]
            if not any(needs_group(lf, f) for f, lf in zip(full, leaf)):
                continue
            first = next((r for r in b["rows"] if not r.get("head")), None)
            name = ""
            if first:
                c0 = next((c for c in first["cells"] if c["col"] == 0), None)
                name = tidy(c0["text"]) if c0 and c0.get("text") else tidy(first.get("label"))
            out.append((sec["section"], b, full, leaf, name))
    return out


def find(doc: dict, wb) -> list:
    """고칠 머리글 칸. [{'section', 'row', 'cells': [{'cell', 'old', 'new'}]}] - 같은 줄은 한 번만."""
    if "주석" not in wb.sheetnames:
        return []
    ws = wb["주석"]
    rows = list(ws.iter_rows(min_col=2, values_only=True))
    found, done = [], set()
    for section, b, full, leaf, first in _risky_tables(doc):
        n = b["ncols"]
        for i, row in enumerate(rows):
            r = i + 1
            if r in done:
                continue
            vals = [squeeze(v) for v in (list(row) + [None] * n)[:n]]
            if not all(v == squeeze(f) or v == squeeze(lf) for v, f, lf in zip(vals, full, leaf)):
                continue
            nxt = rows[i + 1] if i + 1 < len(rows) else ()
            if first and (not nxt or squeeze(nxt[0]) != squeeze(first)):
                continue
            cells = [{"cell": ws.cell(r, 2 + k).coordinate, "old": ws.cell(r, 2 + k).value, "new": full[k]}
                     for k in range(n) if vals[k] != squeeze(full[k]) and needs_group(ws.cell(r, 2 + k).value, full[k])]
            if cells:
                found.append({"section": section, "row": r, "cells": cells})
                done.add(r)
    return found


def apply(workbook: str, found: list) -> str | None:
    """고칠 칸을 정산표에 쓴다. 쓰기 전에 정산표 옆 _build 폴더에 사본을 둔다. 반환: 사본 경로(고칠 것이 없으면 None)."""
    from .back import backup_copy

    if not found:
        return None
    backup = backup_copy(workbook)
    wb = load_workbook(workbook)
    ws = wb["주석"]
    for x in found:
        for c in x["cells"]:
            ws[c["cell"]].value = c["new"]
    try:
        wb.save(workbook)
    except PermissionError:
        import os

        os.remove(backup)
        raise InputError("정산표가 열려 있어 저장하지 못했습니다 - 엑셀을 닫고 다시 실행하세요: %s" % workbook) from None
    return backup
