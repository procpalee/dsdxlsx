"""칸별 출처 기록 - DSD 칸마다 값을 가져온 정산표 칸.

fill 과 build 는 DSD 당기 칸을 정산표에서 채운다. 그때 어느 시트의 어느 칸에서, 어떤 배율(원 · 천원)과 부호로
왔는지를 DSD 옆 `_build/<DSD 이름>.map.json` 에 남긴다. DSD 에서 숫자를 고쳤을 때 정산표의 어느 칸에
반영해야 하는지를 이 기록으로 찾는다.

  DSD 칸 값 = sign × (sources 칸들의 합) × scale   (정수 칸은 반올림, 율은 소수 둘째 자리)

기록 형식(dsdxlsx.map/1)
  dsd · skeleton · workbook · recipe   파일 이름과 sha256(만든 때의 내용 - 그 뒤 바뀌었는지 볼 때)
  cells[]   cell(DSD 칸 ID) · section · table · row · label · col · value(쓴 값) · sources[{sheet, cell}] · scale · sign · how
            sources 가 비면 정산표 칸 없이 계산한 값(위 칸의 합). row_label 은 최종 DSD 에서 그 칸이 놓인 행의 이름.
            state = "overridden" 이면 채운 뒤 레시피나 글 바꾸기가 그 칸을 다시 고쳤다(final = 최종 값) - 출처가 이 기록과 다르다.
            state = "missing" 이면 최종 DSD 에 그 칸이 없다(레시피가 행 · 표를 바꿈).
  counts    cells · overridden · missing · unfilled
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path

SCHEMA = "dsdxlsx.map/1"


def path_of(dsd: str) -> str:
    p = Path(dsd)
    return str(p.parent / "_build" / (p.stem + ".map.json"))


def _file(path: str | None) -> dict | None:
    if not path or not os.path.isfile(path):
        return None
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return {"file": os.path.basename(path), "sha256": h.hexdigest()}


def _same(a, b) -> bool:
    if a is None or b is None:
        return a is b
    return abs(float(a) - float(b)) < 0.005


def dsd_cells(dsd: str) -> dict:
    """DSD 의 칸 ID → {"value": 숫자 | None, "text": 칸 글, "row_label": 그 행의 이름}."""
    from . import docview
    from .dsd_io import read_dsd
    from .parse import parse_contents

    out = {}
    d = docview.doc_dict(parse_contents(read_dsd(dsd).contents))
    for sec in d["sections"]:
        for b in sec["blocks"]:
            if b["type"] != "table":
                continue
            for r in b["rows"]:
                for c in r["cells"]:
                    out[c["id"]] = {"value": c.get("value"), "text": c.get("text") or "", "row_label": r.get("label") or ""}
    return out


def final_values(dsd: str) -> dict:
    """DSD 의 칸 ID → 숫자 값(숫자가 아닌 칸은 None)."""
    return {k: v["value"] for k, v in dsd_cells(dsd).items()}


def build_map(rep, *, dsd: str, skeleton: str, workbook: str, recipe: str | None = None, side: str = "") -> dict:
    """채움 결과(FillReport)와 최종 DSD 로 출처 기록을 만든다."""
    doc_cells = dsd_cells(dsd)
    final = {k: v["value"] for k, v in doc_cells.items()}
    cells, n_over, n_miss = [], 0, 0
    for x in rep.items:
        if x.verdict != "채움" or not x.cell:
            continue
        e = {"cell": x.cell, "section": x.sheet, "table": x.table, "row": x.row_id, "label": x.label, "col": x.col,
             "value": x.value, "sources": [{"sheet": s, "cell": c} for s, c in x.sources], "scale": x.scale, "sign": x.sign,
             "how": x.how}
        if x.cell in doc_cells:
            e["row_label"] = doc_cells[x.cell]["row_label"]      # 칸 ID 가 그 뒤 밀렸는지 볼 때 맞춰 보는 행 이름
        if x.cell not in final:
            e["state"] = "missing"
            n_miss += 1
        elif not _same(final[x.cell], x.value):
            e["state"], e["final"] = "overridden", final[x.cell]
            n_over += 1
        cells.append(e)
    return {"schema": SCHEMA, "side": side, "dsd": _file(dsd), "skeleton": _file(skeleton), "workbook": _file(workbook),
            "recipe": _file(recipe), "period": rep.period,
            "counts": {"cells": len(cells), "overridden": n_over, "missing": n_miss, "unfilled": rep.count("못 채움")},
            "cells": cells}


def save(rep, *, dsd: str, skeleton: str, workbook: str, recipe: str | None = None, side: str = "") -> tuple[str, dict]:
    """출처 기록을 DSD 옆 _build 폴더에 쓴다. 반환: (경로, counts)."""
    m = build_map(rep, dsd=dsd, skeleton=skeleton, workbook=workbook, recipe=recipe, side=side)
    path = path_of(dsd)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(m, f, ensure_ascii=False, indent=1)
    return path, m["counts"]


def load(dsd_or_map: str) -> dict | None:
    """DSD 경로(그 옆 기록을 찾음)나 map.json 경로 → 기록. 없거나 형식이 다르면 None."""
    path = dsd_or_map if dsd_or_map.lower().endswith(".json") else path_of(dsd_or_map)
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        m = json.load(f)
    return m if isinstance(m, dict) and m.get("schema") == SCHEMA else None
