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


def note_pairs(wb, pairs=None) -> list:
    """주석 시트의 (당기 칸, 전기 칸) 짝 - 따로 준 것이 있으면 그것, 없으면 정산표의 연결 기록."""
    if pairs is not None:
        return [list(x) for x in pairs]
    return (read_link(wb) or {}).get("note_pairs") or []


def set_prior_as_current(path: str, notes: bool, pairs=None) -> int:
    """사본의 당기 입력에 전기 금액을 넣는다 - TB 당기 = 전기, 입력칸(크림) = 옆 전기 칸, 자본변동표 = 전기 줄.
    반환: 주석에서 전기 금액을 넣은 칸 수."""
    wb = load_workbook(path)
    n_notes = 0
    if notes and "주석" in wb.sheetnames:
        ws = wb["주석"]
        for cell, prior in note_pairs(wb, pairs):
            ws[cell].value = ws[prior].value
            n_notes += 1
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
    return n_notes


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
                heads = {c["col"]: c.get("header") for c in x_t["columns"]}
                for r, r0 in zip(x_t["rows"], y_t["rows"]):
                    if r.get("head"):
                        continue
                    v, v0 = row_vals(r), row_vals(r0)
                    for c, p in cols:
                        y = v0.get(p, {}).get("value")
                        if y is None:
                            continue
                        x = v.get(c, {}).get("value")
                        item = {"section": s["section"], "label": tidy(r["label"]), "built": x, "expected": y,
                                "column": tidy(heads.get(c)), "split": x_t is not y_t and cols[0][0] == cols[0][1],
                                "cell": v.get(c, {}).get("id")}
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
        seen: dict[str, int] = {}
        for r, r0 in zip(t["rows"], t0["rows"]):
            if r.get("head"):
                continue
            # 줄 이름은 뼈대 것으로 - 결과 DSD 는 fill 이 '영업이익' → '영업이익(손실)' 처럼 고쳐 쓸 수 있다
            nth = seen[squeeze(r0["label"])] = seen.get(squeeze(r0["label"]), -1) + 1
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
                item = {"section": s["section"], "label": tidy(r0["label"]), "built": x, "expected": y,
                        "cell": v.get(c, {}).get("id"), "nth": nth}
                if x is None and tx != "-":
                    res["blank"].append(item)
                elif (x or 0) == y:
                    res["same"] += 1
                else:
                    res["diff"].append(item)
    return res


FS_SHEET = {"재무상태표": "BS", "손익계산서": "IS", "포괄손익계산서": "IS", "현금흐름표": "CF"}
HEADER_CAUSE = "주석 머리글에 묶음 이름이 빠짐"
CAUSES = ("연결점검 항목", "위 줄을 따라 바뀜", "주석 - 당기 표와 전기 표가 따로", HEADER_CAUSE, "채우지 못함", "확인 필요")


def _checks(wb) -> set:
    """'연결점검' 시트의 (시트, 과목 열쇠) - 정산표를 만들 때 이미 알린 줄."""
    if "연결점검" not in wb.sheetnames:
        return set()
    ws = wb["연결점검"]
    return {(str(ws.cell(r, 2).value), key_of(ws.cell(r, 3).value)) for r in range(HDR_ROW + 1, ws.max_row + 1)
            if ws.cell(r, 2).value not in (None, "-", "TB")}


def _fs_kind(wb, section: str, label: str, nth: int) -> str:
    """재무제표 시트에서 그 줄의 당기 칸 - 'tb'(TB 를 끌어옴) · 'calc'(다른 칸에서 계산) · 'input' · ''(못 찾음)."""
    name = FS_SHEET.get(section)
    if name not in wb.sheetnames:
        return ""
    ws = wb[name]
    hits = [r for r in range(HDR_ROW + 1, ws.max_row + 1) if squeeze(ws.cell(r, 2).value) == squeeze(label)]
    if not hits:                                  # 번호 · 주석 표시가 다르게 적힌 줄('Ⅲ. 영업이익(손실)' ↔ '영업이익(손실)')
        hits = [r for r in range(HDR_ROW + 1, ws.max_row + 1) if key_of(ws.cell(r, 2).value) == key_of(label)]
    if nth >= len(hits):
        return ""
    v = ws.cell(hits[nth], 3).value
    if isinstance(v, str) and v.startswith("="):
        return "tb" if "SUMIF(TB!" in v.replace(" ", "") else "calc"
    return "input"


def explain(res: dict, workbook: str, fill_items=(), skeleton_doc: dict | None = None) -> dict:
    """어긋난 칸마다 원인(cause)과, 채우기가 못 채웠으면 그 이유(why)를 붙이고 원인별 개수를 res["causes"] 에 둔다.

    재무제표: 정산표를 만들 때 '연결점검'에 올린 과목이면 '연결점검 항목', 그 밖의 어긋난 과목 줄은 '확인 필요'
    (못 채운 칸이면 '채우지 못함'). 다른 칸에서 계산하는 줄(합계 · 이어 받는 줄)은 어긋난 과목 줄이 있으면 '위 줄을
    따라 바뀜', 과목 줄이 모두 맞는데 계산 줄만 어긋나면 수식이 끊긴 것이라 '확인 필요'.
    주석: 당기 표와 전기 표가 따로 있는 표(변동표 등)는 '주석 - 당기 표와 전기 표가 따로' 로 따로 세고, 채우기가 못 채운
    칸은 '채우지 못함'(이유 = fill 의 기록), 그 밖은 '확인 필요'. 뼈대 사전을 주면, 0.6.1 이하로 만든 정산표에서 두 줄
    머리글의 묶음 이름이 빠진 표(`link fix-headers` 로 고침)의 칸은 '주석 머리글에 묶음 이름이 빠짐'."""
    wb = load_workbook(workbook, read_only=False)
    try:
        checks = _checks(wb)
        bad_heads = set()
        if skeleton_doc is not None:
            from .headers import find as find_headers
            bad_heads = {x["section"] for x in find_headers(skeleton_doc, wb)}
        why = {it.cell: it.how for it in fill_items if getattr(it, "cell", "") and it.verdict != "채움"}
        fs = res["diff"] + res["blank"]
        for x in fs:
            sheet = FS_SHEET.get(x["section"], "")
            x["cause"] = "연결점검 항목" if (sheet, key_of(x["label"])) in checks else ""
        kinds = {id(x): _fs_kind(wb, x["section"], x["label"], x.get("nth", 0)) for x in fs}
        for x in fs:                              # 과목 줄(TB · 입력)부터
            if x["cause"] or kinds[id(x)] == "calc":
                continue
            if x.get("cell") in why:
                x["cause"], x["why"] = "채우지 못함", why[x["cell"]]
            else:
                x["cause"] = "확인 필요"
        leaf_bad = any(x["cause"] for x in fs)
        for x in fs:                              # 계산 줄: 어긋난 과목 줄이 있으면 그것을 따라 바뀐 것, 없으면 수식이 끊긴 것
            if x["cause"]:
                continue
            x["cause"] = "위 줄을 따라 바뀜" if leaf_bad else "확인 필요"
        for x in res["note_diff"] + res["note_blank"]:
            if x.get("split"):
                x["cause"] = "주석 - 당기 표와 전기 표가 따로"
            elif x["section"] in bad_heads and "/" in (x.get("column") or ""):
                x["cause"] = HEADER_CAUSE
            elif x.get("cell") in why:
                x["cause"], x["why"] = "채우지 못함", why[x["cell"]]
            else:
                x["cause"] = "확인 필요"
        counts = {k: 0 for k in CAUSES}
        for x in fs + res["note_diff"] + res["note_blank"]:
            counts[x["cause"]] += 1
        res["causes"] = {k: v for k, v in counts.items() if v}
        return res
    finally:
        wb.close()


def run(skeleton: str, workbook: str, *, conf: dict | None = None, recipe: dict | None = None, notes: bool = False,
        workdir: str | None = None, pairs=None) -> dict:
    """왕복 시험을 돌리고 결과를 돌려준다. conf = recon 설정의 한쪽('별도'나 '연결').
    pairs = 주석의 (당기 칸, 전기 칸) 짝 - 연결 기록이 없는 정산표(이전 방식으로 만든 것)에 따로 줄 때."""
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
        n_notes = set_prior_as_current(wp, notes, pairs)
        calc.freeze(wp)
        conf = conf or recon_config()
        sheets = R.load_workpaper(wp, None, conf.get("주석"), skip=conf.get("빼는 시트"), fs_map=conf.get("재무제표") or None)
        out = os.path.join(td, "시험_결과.dsd")
        if os.path.exists(out):
            os.remove(out)
        rep = B.build(skeleton, wp, out, recipe=recipe or {}, fill_conf=conf, side_sheets=sheets,
                      workdir=os.path.join(td, "wd"))
        dict_of = lambda p: docview.doc_dict(parse_contents(read_dsd(p).contents), source=p)  # noqa: E731
        sk = dict_of(skeleton)
        res = compare(dict_of(out), sk)
        if not notes:
            res["note_same"], res["note_diff"], res["note_blank"] = 0, [], []
        explain(res, workbook, getattr(rep.fill, "items", ()) or (), sk)
        res["steps"] = list(rep.steps)
        res["warnings"] = list(rep.warnings)
        if notes and not n_notes:
            res["warnings"].append("정산표에 주석 칸 짝 기록이 없어 주석의 당기 칸을 채우지 못했습니다 - 주석 결과는 참고만")
        return res
    finally:
        if tmp is not None:
            tmp.cleanup()
