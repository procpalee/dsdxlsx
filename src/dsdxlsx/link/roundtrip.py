"""왕복 시험 - 정산표와 뼈대 DSD 의 연결이 끝까지 이어지는지 확인한다.

정산표 사본의 당기 칸에 전기 금액을 넣고 다시 계산한 뒤 build 한다. 당기에 전기와 같은 금액을 넣었으니
결과 DSD 의 당기 칸은 전기 칸과 같아야 한다. 다르거나 비면 그 행의 연결(짝짓기 · 합계 수식)이 끊긴 것이다.
원본 정산표와 뼈대 DSD 는 건드리지 않는다.
"""
from __future__ import annotations

import os
import re
import shutil
import tempfile

from openpyxl import load_workbook

from . import calc
from .core import HDR_ROW, key_of, main_table, read_link, row_vals, squeeze, tidy
from .sheetkit import is_input


def set_prior_as_current(path: str, notes: bool) -> None:
    """사본의 당기 입력에 전기 금액을 넣는다 - TB 당기 = 전기, 입력칸(크림) = 옆 전기 칸, 자본변동표 = 전기 줄."""
    wb = load_workbook(path)
    link = read_link(wb) or {}
    if notes and "주석" in wb.sheetnames:
        ws = wb["주석"]
        for cell, prior in link.get("note_pairs") or []:
            ws[cell].value = ws[prior].value
    ws = wb["TB"]
    for r in range(HDR_ROW + 1, ws.max_row + 1):
        if ws.cell(r, 2).value is None:
            continue
        ws.cell(r, 6).value = ws.cell(r, 9).value
        ws.cell(r, 7).value = 0
    for name in ("BS", "IS", "CF"):
        if name not in wb.sheetnames:
            continue
        s = wb[name]
        for r in range(HDR_ROW + 1, s.max_row + 1):
            c = s.cell(r, 3)                  # 입력칸만 전기 금액으로 - 수식이 빠진 합계 칸은 비워 둬야 시험에 걸린다
            if c.value is None and s.cell(r, 4).value is not None and (is_input(c) or name == "CF"):
                c.value = s.cell(r, 4).value
    if "SCE" in wb.sheetnames:
        s = wb["SCE"]                        # 당기 입력 칸 = 전기 줄의 같은 칸(이름이 같은 앞 줄)
        rows = list(range(HDR_ROW + 1, s.max_row + 1))
        labels = {r: squeeze(s.cell(r, 2).value) for r in rows}
        for r in rows:
            for c in range(3, s.max_column + 1):
                v = s.cell(r, c).value
                if v is None or v == "":
                    twin = next((q for q in rows if q < r and labels[q] == labels[r]), None)
                    if twin:
                        s.cell(r, c).value = s.cell(twin, c).value
    wb.save(path)


def compare(built: dict, skeleton: dict) -> dict:
    """결과 DSD 의 당기 칸 = 뼈대 전기 칸(같은 표 · 행 · 짝 열). 두 인자는 inspect 사전."""
    secs = {s["section"]: s for s in skeleton["sections"]}
    res = {"same": 0, "diff": [], "blank": [], "note_same": 0, "note_diff": [], "note_blank": []}
    for s in built["sections"]:
        base = secs.get(s["section"])
        if not base:
            continue
        if s["kind"] == "note":
            t_a = [x for x in s["blocks"] if x["type"] == "table" and x["ncols"] >= 2]
            t_b = [x for x in base["blocks"] if x["type"] == "table" and x["ncols"] >= 2]
            for j, (ta, tb) in enumerate(zip(t_a, t_b)):
                cc = [c["col"] for c in ta["columns"] if c.get("period") == "당기"]
                pc = [c["col"] for c in ta["columns"] if c.get("period") == "전기"]
                pair = None
                if cc and len(cc) == len(pc):
                    pair = (ta, tb, list(zip(cc, pc)))
                else:
                    h = re.sub(r"\s", "", (ta["rows"][0]["cells"][0].get("text") or "") if ta["rows"] else "")
                    if re.fullmatch(r"당(분|반)?기말?", h) and j + 1 < len(t_b) and t_b[j + 1]["ncols"] == ta["ncols"]:
                        pair = (ta, t_b[j + 1], [(c, c) for c in range(1, ta["ncols"])])
                if not pair:
                    continue
                x_t, y_t, cols = pair
                for r, r0 in zip(x_t["rows"], y_t["rows"]):
                    if r.get("head"):
                        continue
                    v, v0 = row_vals(r), row_vals(r0)
                    for c, p in cols:
                        y = v0.get(p, {}).get("value")
                        if y is None:
                            continue
                        x = v.get(c, {}).get("value")
                        item = {"section": s["section"], "label": tidy(r["label"]), "built": x, "expected": y}
                        if x is None:
                            res["note_blank"].append(item)
                        elif x == y:
                            res["note_same"] += 1
                        else:
                            res["note_diff"].append(item)
            continue
        if s["kind"] != "statement" or s["section"] == "자본변동표":
            continue
        t, t0 = main_table(s), main_table(base)
        if t is None or t0 is None:
            continue
        cc = [c["col"] for c in t["columns"] if c.get("period") == "당기"]
        pc = [c["col"] for c in t["columns"] if c.get("period") == "전기"]
        for r, r0 in zip(t["rows"], t0["rows"]):
            if r.get("head"):
                continue
            if s["section"] in ("손익계산서", "포괄손익계산서") and key_of(r["label"]).startswith(("기초", "당기", "기말")):
                continue                      # 재고 구성 줄(기초 · 매입 · 기말)은 재무상태표에서 오는 수식
            if s["section"] == "현금흐름표" and re.search("기초", squeeze(r["label"])):
                continue
            v, v0 = row_vals(r), row_vals(r0)
            for c, p in zip(cc, pc):
                x, y = v.get(c, {}).get("value"), v0.get(p, {}).get("value")
                tx, ty = tidy(v.get(c, {}).get("text")), tidy(v0.get(p, {}).get("text"))
                if y is None and ty != "-":
                    continue
                y = y or 0
                item = {"section": s["section"], "label": tidy(r["label"]), "built": x, "expected": y}
                if x is None and tx != "-":
                    res["blank"].append(item)
                elif (x or 0) == y:
                    res["same"] += 1
                else:
                    res["diff"].append(item)
    return res


def run(skeleton: str, workbook: str, *, conf: dict | None = None, recipe: dict | None = None, notes: bool = False,
        workdir: str | None = None) -> dict:
    """왕복 시험을 돌리고 결과를 돌려준다. conf = recon 설정의 한쪽('별도'나 '연결')."""
    from .. import build as B
    from .. import docview
    from .. import recon as R
    from ..dsd_io import read_dsd
    from ..parse import parse_contents
    from .core import recon_config

    tmp = None
    if workdir:
        os.makedirs(workdir, exist_ok=True)
        td = workdir
    else:
        tmp = tempfile.TemporaryDirectory()
        td = tmp.name
    try:
        wp = os.path.join(td, "시험_정산표.xlsx")
        shutil.copy(workbook, wp)
        set_prior_as_current(wp, notes)
        calc.freeze(wp)
        conf = conf or recon_config()
        sheets = R.load_workpaper(wp, None, conf.get("주석"), skip=conf.get("빼는 시트"), fs_map=conf.get("재무제표") or None)
        out = os.path.join(td, "시험_결과.dsd")
        if os.path.exists(out):
            os.remove(out)
        rep = B.build(skeleton, wp, out, recipe=recipe or {}, fill_conf=conf, side_sheets=sheets,
                      workdir=os.path.join(td, "wd"))
        dict_of = lambda p: docview.doc_dict(parse_contents(read_dsd(p).contents), source=p)  # noqa: E731
        res = compare(dict_of(out), dict_of(skeleton))
        res["steps"] = list(rep.steps)
        res["warnings"] = list(rep.warnings)
        if not notes:
            res["note_same"], res["note_diff"], res["note_blank"] = 0, [], []
        return res
    finally:
        if tmp is not None:
            tmp.cleanup()
