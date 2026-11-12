"""DSD 에서 고친 숫자를 정산표로 되돌린다(link back).

build · fill 이 남긴 칸별 출처 기록(srcmap)으로, 고친 DSD 의 칸마다 정산표의 어느 칸이 그 값의 출처인지 찾는다.
출처가 값 칸이면 새 값을 제안하고(--apply 로 그 칸에 쓴다), 수식이면 쓰지 않고 무엇을 고쳐야 하는지만 알린다.

  kind       뜻                                                          --apply
  value      값 칸 - 새 값을 그대로 넣으면 된다                          쓴다
  done       정산표에 이미 그 값이 있다                                  -
  scaled     값 칸이지만 단위가 달라(원 · 천원, 율) 정확한 값을 모른다   안 쓴다(제안 값만)
  conflict   정산표의 그 칸도 build 뒤에 바뀌었다                        안 쓴다
  tb         TB 를 끌어오는 수식 - TB 에 얼마를 조정하면 되는지 제안     안 쓴다
  formula    합계나 다른 칸을 참조하는 수식                              안 쓴다
  sum        정산표 여러 칸의 합                                         안 쓴다
  tie        정산표에 값이 같은 칸이 둘 이상이라 그중 하나를 골랐던 칸   안 쓴다
  calc       정산표 칸 없이 계산한 값(위 칸의 합)                        안 쓴다
  recipe     레시피가 넣은 칸                                            안 쓴다
  blank      고친 칸이 비었거나 숫자가 아니다                            안 쓴다
  moved      칸 번호가 밀렸다(DSD 에서 행 · 표를 넣고 뺌)                안 쓴다
"""
from __future__ import annotations

import os
import re
import shutil
import time

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from .. import srcmap
from ..guard import InputError
from .sheetkit import HDR_ROW

WRITE = ("value",)
NOTES = {
    "value": "값 칸 - 새 값을 넣는다",
    "done": "정산표에 이미 반영됨",
    "scaled": "값 칸이지만 단위가 달라 제안 값만 냄",
    "conflict": "정산표의 그 칸도 바뀌어 있음 - 어느 쪽이 맞는지 확인",
    "tb": "TB 를 끌어오는 수식 - TB 에서 조정",
    "formula": "수식 칸 - 구성 칸을 고치면 따라감",
    "sum": "정산표 여러 칸의 합 - 어느 칸인지 정할 수 없음",
    "tie": "정산표에 값이 같은 칸이 둘 이상이라 그중 하나를 골랐음 - 같은 값 칸을 모두 고쳐야 다시 build 할 때 맞음",
    "calc": "정산표 칸 없이 계산한 값 - 구성 칸을 따라감",
    "recipe": "레시피가 넣은 칸 - 레시피의 출처에서 고침",
    "blank": "고친 칸이 비었거나 숫자가 아님",
    "moved": "칸 번호가 밀림 - 뼈대 DSD 를 고치고 다시 build",
}
_SP = re.compile(r"[\s\xa0　]+")
_REF = re.compile(r"^([A-Z]+)(\d+)$")


def _sq(s) -> str:
    return _SP.sub("", str(s or ""))


def _num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _tol(v) -> float:
    return 0.5 if float(v) == int(float(v)) else 0.006


def _clean(v: float):
    return int(round(v)) if abs(v - round(v)) < 1e-9 else round(v, 6)


def _tb_accounts(wb, sheet: str, row: int) -> list:
    """link make 정산표: 그 행의 'TB 연결 과목'(G열)과 'DSD 행'이 같은 TB 계정들 [(코드, 계정과목)]."""
    if "TB" not in wb.sheetnames:
        return []
    key = wb[sheet].cell(row, 7).value
    if key in (None, ""):
        return []
    tb = wb["TB"]
    return [(str(tb.cell(r, 2).value), str(tb.cell(r, 3).value or "")) for r in range(HDR_ROW + 1, tb.max_row + 1)
            if tb.cell(r, 5).value == key and tb.cell(r, 2).value is not None]


def plan(edited_dsd: str, workbook: str, m: dict) -> dict:
    """고친 DSD 와 출처 기록으로 반영 목록을 만든다(아무것도 쓰지 않는다)."""
    cells = srcmap.dsd_cells(edited_dsd)
    wb = load_workbook(workbook)
    changes, same = [], 0
    for e in m["cells"]:
        base = e["final"] if e.get("state") == "overridden" else e["value"]
        item = {"cell": e["cell"], "section": e["section"], "label": e["label"], "col": e["col"], "old": base, "new": None,
                "kind": "", "sheet": None, "ref": None, "proposed": None}
        cur = cells.get(e["cell"])
        if e.get("state") == "missing":
            continue
        if cur is None or (e.get("row_label") is not None and _sq(cur["row_label"]) != _sq(e["row_label"])):
            item["kind"] = "moved"
            changes.append(item)
            continue
        new = _num(cur["value"])
        if new is not None and base is not None and abs(new - float(base)) < 0.005:
            same += 1
            continue
        item["new"] = cur["value"]
        if new is None:
            if base is None:
                same += 1
                continue
            item["kind"] = "blank"
        elif e.get("state") == "overridden":
            item["kind"] = "recipe"
        elif not e["sources"]:
            item["kind"] = "calc"
        elif len(e["sources"]) > 1:
            item["kind"] = "sum"
            item["sheet"], item["ref"] = e["sources"][0]["sheet"], ", ".join(s["cell"] for s in e["sources"])
        else:
            _one(wb, e, item, new)
            if item["kind"] in ("value", "scaled") and "같은 값" in (e.get("how") or ""):
                item["kind"] = "tie"
        changes.append(item)
    _shared_sources(m, changes)
    for x in changes:
        x["note"] = NOTES[x["kind"]] + (" - " + x.pop("extra") if x.get("extra") else "")
    counts: dict = {}
    for x in changes:
        counts[x["kind"]] = counts.get(x["kind"], 0) + 1
    warnings = []
    if counts.get("moved"):
        warnings.append("DSD 의 표 구조가 build 때와 달라 칸 %d개를 짝짓지 못했습니다 - 행 · 표를 넣고 뺐으면 뼈대 DSD 에 옮기고 다시 build 하세요"
                        % counts["moved"])
    wbm = m.get("workbook") or {}
    if wbm.get("file") and wbm["file"] != os.path.basename(workbook):
        warnings.append("출처 기록의 정산표(%s)와 다른 파일입니다 - 시트와 칸 자리가 같은 파일이어야 합니다" % wbm["file"])
    return {"changes": changes, "counts": counts, "unchanged": same, "warnings": warnings}


def _shared_sources(m: dict, changes: list) -> None:
    """같은 정산표 칸에서 오는 DSD 칸이 여럿일 때 - 서로 다른 값으로 고쳤으면 쓰지 않고, 한쪽만 고쳤으면 알린다."""
    users: dict = {}
    for e in m["cells"]:
        if len(e["sources"]) == 1 and not e.get("state"):
            users.setdefault((e["sources"][0]["sheet"], e["sources"][0]["cell"]), []).append(e["cell"])
    by_src: dict = {}
    for x in changes:
        if x["kind"] == "value":
            by_src.setdefault((x["sheet"], x["ref"]), []).append(x)
    for key, items in by_src.items():
        if len({x["proposed"] for x in items}) > 1:
            for x in items:
                x["kind"], x["extra"] = "conflict", "DSD 의 다른 칸도 같은 정산표 칸에서 오는데 고친 값이 서로 다름"
            continue
        rest = len(users.get(key, [])) - len(items)
        for x in items:
            notes = []
            if rest > 0:
                x["also"] = rest
                notes.append("같은 정산표 칸에서 오는 DSD 칸 %d개도 다시 build 하면 같이 바뀜" % rest)
            if not x["old"]:                      # 0 은 채울 때 후보에서 빠지는 값 - 0 이 아니게 되면 같은 행 다른 칸의 짝이 달라질 수 있다
                notes.append("0 이던 칸 - 다시 build 한 뒤 그 표의 값이 맞는지 확인")
            if notes:
                x["extra"] = " / ".join(notes)


def _one(wb, e: dict, item: dict, new: float) -> None:
    """출처가 한 칸인 경우 - 정산표의 그 칸이 값인지 수식인지 보고 종류를 정한다."""
    src = e["sources"][0]
    item["sheet"], item["ref"] = src["sheet"], src["cell"]
    sign, scale, old = e["sign"] or 1, e["scale"] or 1.0, float(e["value"])
    if src["sheet"] not in wb.sheetnames or not _REF.match(src["cell"]):
        item["kind"], item["extra"] = "formula", "정산표에 그 시트나 칸이 없음"
        return
    c = wb[src["sheet"]][src["cell"]]
    v = c.value
    if isinstance(v, str) and v.startswith("="):
        if "SUMIF(TB!" in v.replace(" ", "").upper():
            line = -1 if v.replace(" ", "").startswith("=-") else 1
            delta = (new - old) / (sign * scale) * line          # TB 금액(차변 +)을 이만큼 바꾸면 DSD 가 new 가 된다
            accts = _tb_accounts(wb, src["sheet"], c.row)
            item["kind"], item["tb_delta"] = "tb", _clean(delta)
            item["accounts"] = [{"code": a, "name": n} for a, n in accts]
            side = "차변" if delta > 0 else "대변"
            who = ("계정 %s %s" % accts[0]) if len(accts) == 1 else ("계정 %d개 중에서 정함" % len(accts)) if accts else "TB 계정을 못 찾음"
            item["extra"] = "TB 수정사항에 %s %s (%s)" % (side, format(abs(_clean(delta)), ","), who)
        else:
            item["kind"] = "formula"
        return
    w = _num(v)
    if v not in (None, "") and w is None:
        item["kind"], item["extra"] = "formula", "정산표 칸에 숫자가 아닌 글이 있음"
        return
    w = w or 0.0
    now = sign * w * scale                                        # 지금 정산표 값으로 다시 계산한 DSD 값
    if abs(now - new) <= _tol(new):
        item["kind"] = "done"
        return
    item["proposed"] = _clean(new / (sign * scale))
    if abs(now - old) > _tol(old):
        item["kind"], item["extra"] = "conflict", "정산표 값으로는 %s" % format(_clean(now), ",")
    elif scale != 1.0:
        item["kind"] = "scaled"
    else:
        item["kind"] = "value"


def backup_copy(workbook: str) -> str:
    """정산표에 쓰기 전 사본 - 정산표 옆 _build/<이름>_<시각>.xlsx. 반환: 사본 경로."""
    d = os.path.join(os.path.dirname(os.path.abspath(workbook)), "_build")
    os.makedirs(d, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(workbook))
    backup = os.path.join(d, "%s_%s%s" % (stem, time.strftime("%Y%m%d_%H%M%S"), ext))
    k = 1
    while os.path.exists(backup):
        k += 1
        backup = os.path.join(d, "%s_%s_%d%s" % (stem, time.strftime("%Y%m%d_%H%M%S"), k, ext))
    shutil.copy2(workbook, backup)
    return backup


def apply(workbook: str, changes: list) -> tuple[int, str | None]:
    """value 종류의 제안 값을 정산표에 쓴다. 쓰기 전에 정산표를 _build 폴더에 복사해 둔다. 반환: (쓴 칸 수, 백업 경로)."""
    todo = [x for x in changes if x["kind"] in WRITE]
    if not todo:
        return 0, None
    backup = backup_copy(workbook)
    wb = load_workbook(workbook)
    for x in todo:
        m = _REF.match(x["ref"])
        wb[x["sheet"]].cell(int(m.group(2)), column_index_from_string(m.group(1))).value = x["proposed"]
        x["applied"] = True
    try:
        wb.save(workbook)
    except PermissionError:
        os.remove(backup)
        for x in todo:
            x.pop("applied", None)
        raise InputError("정산표가 열려 있어 저장하지 못했습니다 - 엑셀을 닫고 다시 실행하세요: %s" % workbook)
    return len(todo), backup


def unmapped(edited_dsd: str, last_dsd: str, m: dict) -> list:
    """지난 build 결과와 달라졌는데 출처 기록에 없는 칸(전기 칸 · 못 채운 칸 · 글 칸) - 정산표에 짝이 없다."""
    if not os.path.exists(last_dsd):
        return []
    a, b = srcmap.dsd_cells(last_dsd), srcmap.dsd_cells(edited_dsd)
    known = {e["cell"] for e in m["cells"]}
    out = []
    for cid, x in b.items():
        y = a.get(cid)
        if cid in known or y is None or _sq(y["row_label"]) != _sq(x["row_label"]) or y["text"] == x["text"]:
            continue
        out.append({"cell": cid, "label": x["row_label"][:40], "old": y["text"], "new": x["text"]})
    return out
