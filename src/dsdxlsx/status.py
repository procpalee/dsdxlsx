"""정산표와 DSD 의 상태 - build · fill 뒤에 무엇이 바뀌었고 다음에 무엇을 돌릴지.

칸별 출처 기록(srcmap)에는 만든 때의 DSD · 뼈대 · 정산표 · 레시피 파일의 sha256 이 있다. 지금 파일과 견줘
어느 쪽이 바뀌었는지 보고, DSD 가 바뀌었으면 어떤 칸이(link back 의 분류), 정산표가 바뀌었으면 출처 칸 몇 개가
달라졌는지 센다. 아무것도 쓰지 않는다.

  DSD 만 바뀜            → link back 으로 정산표에 반영
  정산표 · 뼈대만 바뀜   → build(또는 fill)를 다시
  양쪽 다 바뀜           → link back 으로 충돌부터 확인(이미 반영된 것이면 다시 build)
"""
from __future__ import annotations

import hashlib
import os

from . import srcmap

ROLES = (("dsd", "DSD"), ("skeleton", "뼈대 DSD"), ("workbook", "정산표"), ("recipe", "레시피"))
STATE = {"same": "그대로", "changed": "바뀜", "missing": "파일을 못 찾음", None: "-"}


def _sha(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _state(rec: dict | None, path: str | None) -> str | None:
    if rec is None:
        return None
    if not path or not os.path.isfile(path):
        return "missing"
    return "same" if _sha(path) == rec.get("sha256") else "changed"


def workbook_changes(m: dict, workbook: str) -> dict:
    """정산표의 출처 칸 가운데 지금 값으로 다시 계산하면 DSD 값이 달라지는 칸 수.
    no_values = 수식인데 저장된 계산값이 없는 칸(엑셀에서 열어 저장하지 않은 정산표)."""
    from openpyxl import load_workbook

    vals, forms = load_workbook(workbook, data_only=True), load_workbook(workbook)
    changed = no_values = 0
    for e in m["cells"]:
        if e.get("state") or len(e["sources"]) != 1:
            continue
        s = e["sources"][0]
        if s["sheet"] not in vals.sheetnames:
            changed += 1
            continue
        v, f = vals[s["sheet"]][s["cell"]].value, forms[s["sheet"]][s["cell"]].value
        if v is None and isinstance(f, str) and f.startswith("="):
            no_values += 1
            continue
        if v not in (None, "") and not (isinstance(v, (int, float)) and not isinstance(v, bool)):
            changed += 1
            continue
        now = (e["sign"] or 1) * float(v or 0) * (e["scale"] or 1.0)
        old = float(e["value"])
        if abs(now - old) > (0.5 if old == int(old) else 0.006):
            changed += 1
    return {"cells": changed, "no_values": no_values}


def check(dsd: str, m: dict, *, skeleton: str | None = None, workbook: str | None = None, recipe: str | None = None) -> dict:
    """지금 파일과 출처 기록을 견준다. 경로를 안 주면 기록에 적힌 경로를 쓴다."""
    from . import build as B
    from .link import back

    paths = {"dsd": dsd, "skeleton": skeleton or (m.get("skeleton") or {}).get("path"),
             "workbook": workbook or (m.get("workbook") or {}).get("path"), "recipe": recipe or (m.get("recipe") or {}).get("path")}
    files = {k: {"file": (m.get(k) or {}).get("file"), "path": paths[k], "state": _state(m.get(k), paths[k])} for k, _n in ROLES}
    st = {k: v["state"] for k, v in files.items()}
    made_by = m.get("made_by") or "build"
    has_wb = st["workbook"] in ("same", "changed")
    out = {"files": files, "made_by": made_by, "dsd_changes": None, "unmapped": 0, "workbook_changes": None, "next": []}

    dsd_ch = st["dsd"] == "changed"
    in_ch = [n for k, n in ROLES[1:] if st[k] == "changed"]
    missing = [n for k, n in ROLES[1:] if st[k] == "missing"]
    counts: dict = {}
    if dsd_ch:
        if has_wb:
            counts = back.plan(dsd, paths["workbook"], m)["counts"]
        else:                                    # 정산표가 없으면 바뀐 칸 수만
            cells = srcmap.dsd_cells(dsd)
            n = sum(1 for e in m["cells"] if e.get("state") != "missing"
                    and (cells.get(e["cell"]) is None or cells[e["cell"]]["value"] is None
                         or abs(float(cells[e["cell"]]["value"]) - float(e.get("final", e["value"]))) >= 0.005))
            counts = {"changed": n} if n else {}
        out["dsd_changes"] = counts
        if not counts.get("moved"):
            out["unmapped"] = len(back.unmapped(dsd, B.last_build_path(dsd), m))
    if st["workbook"] == "changed":
        out["workbook_changes"] = workbook_changes(m, paths["workbook"])

    nxt = out["next"]
    again = "build --force" if made_by == "build" else "fill --force"
    pending = sum(v for k, v in counts.items() if k != "done")
    if missing:
        nxt.append("%s 파일을 못 찾았습니다 - 옮겼으면 경로를 옵션으로 주세요" % ", ".join(missing))
    if not dsd_ch and not in_ch:
        if st["dsd"] == "same" and not missing:
            nxt.append("DSD 와 정산표가 맞습니다. 할 일이 없습니다.")
    elif in_ch and not dsd_ch:
        wc = out["workbook_changes"] or {}
        detail = ""
        if wc.get("cells"):
            detail = " (정산표 출처 칸 %d개의 값이 달라짐)" % wc["cells"]
        elif wc.get("no_values"):
            detail = " (정산표에 계산값이 없음 - 엑셀에서 열어 저장한 뒤)"
        nxt.append("%s 쪽이 바뀌었습니다%s → %s 를 다시 돌려 DSD 에 반영" % (", ".join(in_ch), detail,
                                                                    "build" if made_by == "build" else "fill"))
    elif dsd_ch and not in_ch:
        if counts.get("moved"):
            nxt.append("DSD 에서 행 · 표를 넣고 빼 칸 번호가 밀렸습니다 → 그 수정을 뼈대 DSD 에 옮기고 다시 %s" % again)
        elif pending:
            w = counts.get("value", 0)
            nxt.append("DSD 에서 %d칸이 바뀌었습니다 → link back 으로 정산표에 반영%s" % (pending, " (값 칸 %d개는 --apply)" % w if w else ""))
        else:
            nxt.append("DSD 에서 출처 기록에 없는 곳만 바뀌었습니다(전기 칸 · 글) - 정산표에 반영할 것은 없습니다. "
                       "다시 build 하면 이 수정이 사라지니 뼈대 DSD 에 옮기세요")
    else:
        if counts.get("moved"):
            nxt.append("양쪽이 바뀌었고 DSD 의 칸 번호가 밀렸습니다 → DSD 의 수정을 뼈대 DSD 에 옮기고 다시 %s" % again)
        elif pending:
            c = counts.get("conflict", 0)
            nxt.append("DSD 와 %s 모두 바뀌었습니다 → link back 으로 확인%s" % (", ".join(in_ch), " (서로 다른 값으로 고친 칸 %d개)" % c if c else ""))
        else:
            nxt.append("DSD 에서 고친 값이 정산표에 반영돼 있습니다 → 정산표를 엑셀에서 열어 저장한 뒤 %s 로 DSD 를 다시 만들어 확인" % again)
    out["in_sync"] = st["dsd"] == "same" and not in_ch and not missing
    return out
