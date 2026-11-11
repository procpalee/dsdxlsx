"""TB(시산표) 읽기 - 정산표를 만들 계정별 금액.

세 가지 모양을 받는다.
  1) JSON(dsdxlsx.tb/1): {"period": {"cur", "pri"}, "rows": [{"code", "name", "settle", "cur", "aje", "pri"}]}
  2) 표(엑셀 시트나 CSV): 머리글 '계정과목'이 있는 줄부터. 열은 이름으로 찾는다.
  3) 계정 체계 폴더: levels.csv · coa.csv · tb.csv · aje.csv · tb.json

돌려주는 값: (rows, info)
  rows = [{"code", "name", "settle"(정산 과목), "sub"(공시용 과목), "group", "cur", "aje", "pri", "plan"}] - 금액은 Decimal(차변 +)
  info = {"cur": 당기 기준일 글, "pri": 전기 기준일 글, "plan": 참고 기준일 글, "form": 양식}
"""
from __future__ import annotations

import csv
import json
import os
import re
from decimal import Decimal, InvalidOperation

from ..guard import InputError

SCHEMA = "dsdxlsx.tb/1"
_DATE_RE = re.compile(r"^(\d{4})[.\-/년 ]+(\d{1,2})[.\-/월 ]+(\d{1,2})")


def _dec(v) -> Decimal:
    if v is None or v == "":
        return Decimal(0)
    if isinstance(v, Decimal):
        return v
    if isinstance(v, bool):
        return Decimal(0)
    if isinstance(v, (int, float)):
        return Decimal(str(v))
    s = str(v).replace(",", "").strip()
    if s in ("", "-"):
        return Decimal(0)
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return Decimal(s)
    except InvalidOperation:
        return Decimal(0)


def _row(code, name, settle, *, sub="", group="", cur=0, aje=0, pri=0, plan=None, **extra) -> dict:
    out = {"code": str(code), "name": str(name), "settle": str(settle or name), "sub": str(sub or ""), "group": str(group or ""),
           "cur": _dec(cur), "aje": _dec(aje), "pri": _dec(pri), "plan": None if plan is None else _dec(plan)}
    out.update(extra)
    return out


# ── 1) JSON ──────────────────────────────────────────

def load_json(path: str):
    with open(path, encoding="utf-8-sig") as f:
        d = json.load(f)
    if not isinstance(d, dict) or not isinstance(d.get("rows"), list):
        raise InputError("TB JSON 에 rows 목록이 없습니다: %s" % path)
    rows = []
    for i, r in enumerate(d["rows"], 1):
        if not isinstance(r, dict) or not r.get("name"):
            raise InputError("TB JSON rows %d번째 줄에 name(계정과목)이 없습니다" % i)
        rows.append(_row(r.get("code") or r["name"], r["name"], r.get("settle"), sub=r.get("sub"), group=r.get("group"),
                         cur=r.get("cur"), aje=r.get("aje"), pri=r.get("pri"), plan=r.get("plan")))
    per = d.get("period") or {}
    return rows, {"cur": str(per.get("cur") or "당기"), "pri": str(per.get("pri") or "전기"),
                  "plan": str(per.get("plan") or ""), "form": str(d.get("form") or "")}


# ── 2) 표 ────────────────────────────────────────────

def _table_rows(lines: list[list]):
    """머리글 줄을 찾아 (머리글 → 열 순번, 그 아래 줄들)."""
    for i, r in enumerate(lines):
        if r and any(re.sub(r"\s", "", str(x)) == "계정과목" for x in r if x is not None):
            hdr = {}
            for j, v in enumerate(r):
                if v is not None and str(v).strip():
                    hdr.setdefault(re.sub(r"\s", "", str(v)), j)
            return hdr, lines[i + 1:]
    raise InputError("TB 표에서 머리글 '계정과목'을 찾지 못했습니다")


def _from_table(lines: list[list]):
    hdr, body = _table_rows(lines)

    def col(*names):
        for n in names:
            hit = next((j for h, j in hdr.items() if h.startswith(n)), None)
            if hit is not None:
                return hit
        return None

    dates = sorted((h for h in hdr if _DATE_RE.match(h)), key=lambda h: tuple(int(x) for x in _DATE_RE.match(h).groups()))
    c_name, c_code = hdr["계정과목"], col("코드", "계정코드")
    c_settle, c_sub, c_group = col("정산과목", "정산"), col("공시용계정과목", "공시과목", "공시"), col("구분", "대분류")
    c_cur, c_aje, c_pri = col("회사제시", "당기"), col("수정사항", "수정", "AJE"), col("전기")
    if c_cur is None and dates:
        c_cur = hdr[dates[-1]]
    if c_pri is None and len(dates) >= 2:
        c_pri = hdr[dates[-2]]
    elif c_pri is None and len(dates) == 1 and c_cur != hdr[dates[0]]:
        c_pri = hdr[dates[0]]
    if c_pri is None:
        raise InputError("TB 표에 전기 금액 열이 없습니다(머리글 '전기' 또는 전기 기준일)")
    get = lambda r, j: r[j] if j is not None and j < len(r) else None  # noqa: E731
    rows = []
    for r in body:
        name = get(r, c_name)
        if name is None or str(name).strip() == "":
            continue
        code = get(r, c_code)
        rows.append(_row(code if code not in (None, "") else name, name, get(r, c_settle), sub=get(r, c_sub), group=get(r, c_group),
                         cur=get(r, c_cur), aje=get(r, c_aje), pri=get(r, c_pri)))
    if not rows:
        raise InputError("TB 표에 계정 줄이 없습니다")
    cur_label = dates[-1] if len(dates) >= 2 else "당기"
    pri_label = dates[-2] if len(dates) >= 2 else (dates[0] if len(dates) == 1 else "전기")
    return rows, {"cur": cur_label, "pri": pri_label, "plan": "", "form": ""}


def load_xlsx(path: str, sheet: str | None = None):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        if sheet and sheet not in wb.sheetnames:
            raise InputError("TB 엑셀에 '%s' 시트가 없습니다: %s" % (sheet, path))
        ws = wb[sheet] if sheet else (wb["TB"] if "TB" in wb.sheetnames else wb.worksheets[0])
        lines = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    return _from_table(lines)


def load_csv(path: str):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return _from_table([list(r) for r in csv.reader(f)])


# ── 3) 계정 체계 폴더 ────────────────────────────────

def is_folder(path: str) -> str | None:
    """계정 체계 폴더면 그 경로(건 폴더를 주면 그 안 _engagement/tb)."""
    for d in (path, os.path.join(path, "_engagement", "tb"), os.path.join(path, "tb")):
        if all(os.path.exists(os.path.join(d, n)) for n in ("tb.csv", "coa.csv", "levels.csv", "tb.json")):
            return d
    return None


def load_folder(d: str):
    def rd(name):
        with open(os.path.join(d, name), encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f))

    with open(os.path.join(d, "tb.json"), encoding="utf-8") as f:
        meta = json.load(f)
    roles = {r["역할"]: r["ID"] for r in rd("levels.csv")}
    coa = rd("coa.csv")
    node = {r["코드"]: r for r in coa}

    def up(code, lv):
        c = code
        while c and c in node:
            if node[c]["Level"] == lv:
                return node[c]["이름"]
            c = node[c]["상위코드"]
        return ""

    amt: dict = {}
    for r in rd("tb.csv"):
        key = (r["기준일"], r["코드"])
        amt[key] = amt.get(key, Decimal(0)) + _dec(r["금액"])
    aje: dict = {}
    if os.path.exists(os.path.join(d, "aje.csv")):
        for r in rd("aje.csv"):
            aje[r["코드"]] = aje.get(r["코드"], Decimal(0)) + _dec(r["차변"]) - _dec(r["대변"])
    cur, pri = meta["기간"]["당기"], meta["기간"]["전기"]
    plan = meta.get("계획") or ""
    settle_lv = roles["공시"] if meta.get("양식") == "K-IFRS" and "공시" in roles else roles["표시과목"]
    rows = []
    for r in coa:
        if r["Level"] != roles["계정"]:
            continue
        c = r["코드"]
        rows.append(_row(c, r["이름"], up(c, settle_lv), group=up(c, roles["구분"]), cur=amt.get((cur, c)), aje=aje.get(c),
                         pri=amt.get((pri, c)), plan=amt.get((plan, c), Decimal(0)) if plan else None))
    return rows, {"cur": cur, "pri": pri, "plan": plan, "form": meta.get("양식", "")}


# ── 고르기 ───────────────────────────────────────────

def load(path: str, sheet: str | None = None):
    if not os.path.exists(path):
        raise InputError("TB 를 찾을 수 없습니다: %s" % path)
    if os.path.isdir(path):
        d = is_folder(path)
        if not d:
            raise InputError("TB 폴더에 tb.csv · coa.csv · levels.csv · tb.json 이 없습니다: %s" % path)
        return load_folder(d)
    ext = os.path.splitext(path)[1].lower()
    if ext == ".json":
        return load_json(path)
    if ext == ".csv":
        return load_csv(path)
    if ext in (".xlsx", ".xlsm"):
        return load_xlsx(path, sheet)
    raise InputError("TB 는 JSON · CSV · 엑셀(.xlsx) 파일이나 계정 체계 폴더여야 합니다: %s" % path)


def load_consolidated(parent: str, subs: list[str], adj_json: str, sheet: str | None = None):
    """연결 정산표용 TB: 지배기업(코드 앞 P.) + 종속기업(S1. S2. …) 계정 줄 + 전기 연결조정 줄(ADJ.).

    adj_json = [{"name": 계정과목, "sub": 공시용 과목, "settle": 정산 과목(선택), "pri": 전기 연결조정 금액}]
    또는 {"adj": [...], "map": {"S1.457": "정산 과목"}} - map 은 계정 줄의 정산 과목을 바꾼다.
    """
    a, info = load(parent, sheet)
    rows = [dict(r, code="P." + r["code"]) for r in a]
    for i, s in enumerate(subs, 1):
        b, _ = load(s, sheet)
        rows += [dict(r, code="S%d.%s" % (i, r["code"])) for r in b]
    adj_list, settle_map = [], {}
    if adj_json:
        with open(adj_json, encoding="utf-8-sig") as f:
            cfg = json.load(f)
        adj_list, settle_map = (cfg, {}) if isinstance(cfg, list) else (cfg.get("adj", []), cfg.get("map", {}))
    for r in rows:
        if r["code"] in settle_map:
            r["settle"] = settle_map[r["code"]]
    known = {re.sub(r"\s", "", r["settle"]) for r in rows}
    for i, x in enumerate(adj_list, 1):
        name = str(x["name"])
        settle = x.get("settle") or (name if re.sub(r"\s", "", name) in known else str(x.get("sub") or name))
        rows.append(_row("ADJ.%d" % i, name, settle, pri=x.get("pri"), adj=True))
    return rows, dict(info, cur_label="합산", aje_label="연결조정")
