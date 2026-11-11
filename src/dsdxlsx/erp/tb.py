"""ERP 자료 폴더 → TB(계정별 당기 · 전기 금액과 분류).

필요한 파일은 당기 분개장 · 재무상태표 · 손익계산서, 제조업이면 제조원가명세서다. 전기 금액은 당기 재무제표의 전기 열에서 온다.
파일 이름은 <자료>_<회사>_<꼬리> 모양으로 찾는다(재무상태표_회사명_2025.xlsx, 기준연월 자료는 _202509). 파일을 직접 지정할 수도 있다.

  계정코드   그 회사 분개장의 코드. 분개장에 거래가 없는 계정은 계정과목 목록 · 계정별원장에서 찾고, 못 찾으면 '*정산 과목'을 코드 자리에 둔다.
  분류       재무제표 안 위치 - 재무상태표 자산 · 부채 · 자본 > Ⅰ.소분류 > (1)공시 > 계정, 손익계산서 그룹, 제조원가명세서(제조).
             차감 계정(감가상각누계액 · 대손충당금 …)의 정산 과목은 바로 위 자산을 괄호에 붙인다 - 감가상각누계액(건물).
  이월이익잉여금(마감 전) = 재무상태표 미처분이익잉여금 - 당기순이익 + 전기오류수정손익
  당기 · 전기 금액의 부호합이 0 이 아니면 멈춘다.
"""
from __future__ import annotations

import glob
import json
import os
import re
import warnings

from ..guard import InputError
from . import semusarang
from .coa import build
from .douzone import is_mfg, parse_bs, parse_journal, year_from_filename

NEED = ("분개장", "재무상태표", "손익계산서")
ACCT_CODE = re.compile(r"^[\(\[](\d{3,5})[\)\]]\s*(.+)$")              # 계정별원장 시트 이름 · 머리 칸: (0331)자본금 · [0103]보통예금


def account_book(raw: str, key: str) -> dict:
    """분개장에 거래가 없는 계정(자본금 · 보증금 · 적립금 …)도 회사 코드로 넣도록 {코드: 이름} 을 읽는다.
    ① 계정과목_*  - 계정과목 목록(아무 시트, 줄마다 3~5자리 코드 칸 + 바로 뒤 이름 칸)
    ② 계정별원장_* - 계정마다 시트 하나(시트 이름 (0331)자본금) 또는 머리 칸 [0103]보통예금
    꼬리가 맞는 파일 우선, 계정과목 목록은 꼬리 없는 파일도 읽는다. 못 읽으면 빈 dict."""
    from openpyxl import load_workbook

    def files(kw, any_key=False):
        hits = glob.glob(os.path.join(raw, kw + "_*.xls*")) + glob.glob(os.path.join(raw, "*", kw + "_*.xls*"))
        hits = [h for h in hits if not os.path.basename(h).startswith("~$")]
        own = [h for h in hits if re.search(r"_%s(?!\d)" % key, os.path.basename(h))]
        return own or (hits if any_key else [])
    book: dict = {}

    def add(code, name):
        name, c = re.sub(r"\s+", "", str(name or "")), int(code)
        if c >= 1000:                                                 # 4자리 0331 · 5자리 33100 → 3자리
            c = c // 100 if c >= 10000 and c % 100 == 0 else c
        if name and not re.search(r"\((판|제|도|분)\)$", name):           # 원장 · 목록은 (판)(제) 표시가 없다 - 분개장처럼 붙여
            name += {5: "(제)", 6: "(도)", 7: "(분)", 8: "(판)"}.get(c // 100, "")   # 판관비와 원가 같은 이름(퇴직급여 …)을 가림
        if name and str(c) not in book:
            book[str(c)] = name
    for p in files("계정과목", any_key=True):
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
            for ws in wb.worksheets:
                for r in ws.iter_rows(values_only=True):
                    cells = [c for c in r if c not in (None, "")]
                    for i, c in enumerate(cells[:-1]):
                        if re.fullmatch(r"\d{3,5}", str(c).strip()) and isinstance(cells[i + 1], str) and not cells[i + 1].strip().isdigit():
                            add(str(c).strip(), cells[i + 1])
                            break
        except Exception:                                             # noqa: BLE001 - 목록은 보조 자료(못 읽어도 TB 는 만든다)
            continue
    for p in files("계정별원장"):
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
            for ws in wb.worksheets:
                m = ACCT_CODE.match(ws.title.strip())
                if m:
                    add(m.group(1), m.group(2))
                    continue
                for r in ws.iter_rows(max_row=6, values_only=True):
                    for c in r:
                        m = ACCT_CODE.match(str(c).strip()) if isinstance(c, str) else None
                        if m:
                            add(m.group(1), m.group(2))
        except Exception:                                             # noqa: BLE001
            continue
    return book


def _find_all(raw: str, kw: str, key: str) -> list:
    """<자료>_*_<key> (key 뒤에 숫자 없음) - 폴더와 한 단계 아래."""
    hits = glob.glob(os.path.join(raw, kw + "_*.*")) + glob.glob(os.path.join(raw, "*", kw + "_*.*"))
    return sorted(h for h in hits if re.search(r"_%s(?!\d)" % key, os.path.basename(h)) and not os.path.basename(h).startswith("~$"))


def _year_of(raw: str) -> int:
    hits = glob.glob(os.path.join(raw, "재무상태표_*.*")) + glob.glob(os.path.join(raw, "*", "재무상태표_*.*"))
    ys = []
    for h in hits:
        try:
            ys.append(year_from_filename(h))
        except ValueError:
            continue
    if not ys:
        raise InputError("재무상태표(재무상태표_회사명_연도.xlsx)가 없습니다: %s" % raw)
    return sorted(set(ys))[-1]


def make(raw: str | None = None, *, year: int = 0, key: str = "", date: str = "", book: dict | None = None,
         files: dict | None = None) -> dict:
    """자료 폴더(또는 files 로 지정한 파일)에서 TB 를 만든다.

    key   파일 이름 꼬리 - 2025(결산일 자료) 또는 202509(기준연월 자료)
    date  기준일 2025.09.30 - 재무상태표 날짜가 다르면 멈추고, 분개장은 그날까지 전표만 쓴다
    book  이미 아는 회사 계정 {코드: 이름}(분개장에 안 나온 계정의 코드)
    files {"분개장": 경로, "재무상태표": 경로, "손익계산서": 경로, "제조원가명세서": [경로, …]} - 이름 규칙 없이 직접
    반환  {"year", "key", "cur_date", "pri_date", "erp", "rows": [[코드, 계정과목, 당기, 전기]],
           "classes": {코드: [대분류, 소분류, 공시, 정산]}, "warnings": [...], "files": [...]}
    """
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")                                # 내려받은 엑셀의 서식 경고는 결과와 무관
        return _make(raw, year, key, date, dict(book or {}), dict(files or {}))


def _make(raw, year, key, date, book, files):
    if raw and not os.path.isdir(raw):
        raise InputError("자료 폴더를 찾을 수 없습니다: %s" % raw)
    if not raw and not files.get("재무상태표"):
        raise InputError("자료 폴더나 재무상태표 파일을 지정하세요")
    for k, p in files.items():
        for one in (p if isinstance(p, (list, tuple)) else [p]):
            if one and not os.path.isfile(one):
                raise InputError("%s 파일을 찾을 수 없습니다: %s" % (k, one))
    y = int(key[:4]) if key else year
    if not y and raw and not files.get("재무상태표"):
        y = _year_of(raw)
    key = key or (str(y) if y else "")

    def find_all(kw):
        return _find_all(raw, kw, key) if raw and key else []
    need = {k: files.get(k) or (find_all(k) or [None])[0] for k in NEED}
    semu = bool(need["재무상태표"]) and semusarang.is_file(need["재무상태표"])
    if semu and not need["분개장"]:                                   # 세무사랑: 분개장이 없으면 재무제표 금액 + 코드 자료(없으면 이름 키)
        need.pop("분개장")
    miss = [k for k, p in need.items() if not p]
    if miss:
        tail = key or "연도"
        raise InputError("TB 에 필요한 파일이 없습니다: %s" % ", ".join("%s_회사명_%s" % (k, tail) for k in miss))
    cost_p = files.get("제조원가명세서") or find_all("제조원가명세서")     # 공사원가 · 용역(분양)원가처럼 여럿이면 모두
    cost_p = list(cost_p) if isinstance(cost_p, (list, tuple)) else [cost_p]
    if not cost_p and is_mfg(need["손익계산서"]):
        raise InputError("손익계산서에 당기제품제조원가가 있어 제조원가명세서가 필요합니다")
    closing = parse_bs(need["재무상태표"]).get("closing")
    if not y:
        if not hasattr(closing, "year"):
            raise InputError("재무상태표에서 기준일을 읽지 못했습니다 - 연도를 지정하세요")
        y = closing.year
        key = key or str(y)
    journal = parse_journal(need["분개장"], y) if need.get("분개장") else []
    if date:                                                          # 기준일까지 전표만(기준연월 분개장에 뒤 전표가 섞여 있어도)
        cut = date.replace(".", "-")
        journal = [r for r in journal if str(r.get("전표일자"))[:10].replace(".", "-") <= cut]
    if semu and raw:                                                  # 세무사랑 코드 자료(계정별원장 · 전표 · 결산부속명세서 - 다른 기간 폴더 것도)
        book.update(semusarang.books_near(raw, key))
    if raw:
        book.update(account_book(raw, key))                           # 계정과목 목록 · 계정별원장의 코드(거래 없는 계정도)
    rows, classes, warns, errors = build(need["재무상태표"], need["손익계산서"], journal, cost_p, book)
    if semu and not need.get("분개장"):
        warns.insert(0, "세무사랑 분개장이 없어 계정별원장 · 전표 · 결산부속명세서의 계정 이름으로 코드를 붙임(분개장 금액 대조 없음, 못 찾은 계정은 이름 키)")
    star = sorted({str(r["계정과목"]) for r in journal if str(r["계정코드"]).startswith("*")})
    if star:
        warns.insert(0, "세무사랑 분개장 계정 중 계정별원장에 코드가 없는 것(이름 키로 넣음): " + ", ".join(star[:15]))
    if errors:
        raise InputError("TB 검증 실패: " + " / ".join(errors[:20]))
    if date:                                                          # 그 날짜로 내려받은 재무상태표만(다른 날짜 자료로 대신하지 않음)
        got = closing.strftime("%Y.%m.%d") if hasattr(closing, "strftime") else ""
        if got != date:
            raise InputError("재무상태표(%s) 날짜가 %s 입니다 - 기준일 %s 로 내려받은 파일을 넣으세요"
                             % (os.path.basename(need["재무상태표"]), got or "알 수 없음", date))
    out = {"year": y, "key": key, "rows": rows, "classes": classes, "warnings": warns, "erp": "세무사랑" if semu else "더존"}
    out["cur_date"] = closing.strftime("%Y.%m.%d") if hasattr(closing, "strftime") else "%d.12.31" % y
    out["pri_date"] = "%d.%s" % (int(out["cur_date"][:4]) - 1, out["cur_date"][5:])
    out["files"] = sorted(os.path.basename(p) for p in list(need.values()) + cost_p if p)
    return out


def to_tb(out: dict, rules=None) -> dict:
    """make() 결과 → dsdxlsx.tb/1(link make --tb 가 읽는 모양). rules 를 주면 정산 과목 이름을 묶는다."""
    from . import settle as S

    rows = []
    for code, name, cur, pri in out["rows"]:
        big, sub, disc, raw_settle = (list(out["classes"].get(code) or []) + ["", "", "", ""])[:4]
        name_settle = S.settle(raw_settle or name, big, sub, disc, rules) if rules is not None else (raw_settle or name)
        rows.append({"code": str(code), "name": name, "settle": name_settle, "group": big, "sub_group": sub, "disclosure": disc,
                     "cur": cur, "aje": 0, "pri": pri})
    return {"schema": "dsdxlsx.tb/1", "period": {"cur": out["cur_date"], "pri": out["pri_date"]}, "erp": out["erp"],
            "files": out["files"], "rows": rows}


def load_book(path: str | None) -> dict:
    if not path:
        return {}
    if not os.path.exists(path):
        raise InputError("계정 목록 JSON 을 찾을 수 없습니다: %s" % path)
    with open(path, encoding="utf-8-sig") as f:
        d = json.load(f)
    if not isinstance(d, dict):
        raise InputError("계정 목록 JSON 은 {코드: 이름} 이어야 합니다: %s" % path)
    return d
