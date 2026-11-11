"""세무사랑에서 내려받은 엑셀을 더존(위하고) 자료와 같은 모양으로 읽는다.

세무사랑 출력 양식(더존과 다른 점)
  재무상태표 · 손익계산서 · 제조원가명세서  A 열이 비고 B 열에 번호(Ⅰ. · ① · 1), 과목 이름은 C~F 열 어딘가, 금액은 '제 N(당)기' 머리 칸부터
                                       두 칸(안쪽 · 바깥쪽), '제 N(전)기' 도 두 칸. → 한 줄 [과목(번호 붙임), 당기 안, 당기 밖, 전기 안, 전기 밖] 로 바꾼다.
  분개장   전표 하나 = 여러 줄: 첫 줄 B 월/일 · D 번호, 줄마다 차변(E 계정 · F 금액) 또는 대변(G 계정 · H 금액), 마지막 줄 E 적요 · G 거래처.
           계정코드가 없어 같은 기간 계정별원장(시트 (0103)보통예금)에서 코드를 찾는다 - 이름이 같은 코드가 여럿(대손충당금 · 감가상각누계액)이면
           원장의 같은 날 · 같은 차대 · 같은 금액 줄로 가린다. 못 찾으면 코드 '*계정과목'.
  계정별원장  계정마다 시트(시트 이름 (0103)보통예금) - 코드 목록과 위 코드 찾기에 쓴다.
  전표(사용자 정의양식)  한 줄 = 한 분개(B 전표일자 · C 번호 · D 구분 · F 코드 6자리 · G 계정과목명 · I 차변 · K 대변 · L 거래처코드 · O 거래처명 · R 적요).
           일부 전표만 뽑은 출력이면 금액은 쓰지 않고 코드만. 전 기간 전표(모든 전표가 차대 일치 - 출금 · 입금 줄은 현금 상대 줄을 채워서)면
           분개장으로 읽는다. 시트가 여럿이면 제목 '전표' 시트를 찾는다.
  결산부속명세서  명세서마다 '계정과목코드' 머리 + B 코드 6자리 · E 계정과목명 - 코드만.
  데이터 내보내기 양식  제목 · 회사명 줄 없이 머리 두 줄부터 시작하는 재무제표와 분개장.
"""
import glob
import os
import re
from collections import defaultdict
from datetime import date

KINDS = ("재무상태표", "손익계산서", "제조원가명세서", "분개장", "계정별원장", "전표", "결산부속명세서")
BOOKS = ("계정별원장", "전표", "결산부속명세서")                       # 계정코드만 읽는 자료(금액은 재무제표 · 분개장)
STATEMENTS = ("재무상태표", "손익계산서", "제조원가명세서")
_ROMAN = "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ"
_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
_DATE = re.compile(r"(\d{4})\s*[년.\-/]\s*(\d{1,2})\s*[월.\-/]\s*(\d{1,2})")
ACCT = re.compile(r"^[\(\[](\d{3,5})[\)\]]\s*(.+)$")
TAG = re.compile(r"\((판|제|도|분|공|운|용)\)$")


def _n(s):
    return re.sub(r"\s+", "", str(s or "")).replace("　", "")


def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def title(rows, scan=8):
    """앞 줄의 제목(공백 뺀 글이 KINDS 중 하나) - 없으면 ''."""
    for r in rows[:scan]:
        for v in r:
            s = _n(v)
            if s in KINDS:
                return s
    return ""


def _hdr(rows, scan=20):
    """재무제표 머리 줄(과목 · (당)기 · (전)기) → (행, 과목 열, 당기 열, 전기 열) 또는 None."""
    for i, r in enumerate(rows[:scan]):
        cells = [_n(v) for v in r]
        if "과목" not in cells:
            continue
        cur = next((j for j, s in enumerate(cells) if "(당)기" in s), None)
        pri = next((j for j, s in enumerate(cells) if "(전)기" in s), None)
        if cur is not None:
            return i, cells.index("과목"), cur, pri
    return None


def is_statement(rows):
    """세무사랑 재무제표 출력인가 - 제목이 재무제표이고 '과목' 머리 칸이 A 열이 아님(더존은 A 열)."""
    if title(rows) not in STATEMENTS:
        return False
    h = _hdr(rows)
    return bool(h) and h[1] >= 1 and all(r[0] in (None, "") for r in rows[h[0]:h[0] + 3] if r)


def is_journal(rows):
    """세무사랑 분개장 - 머리 줄 B 열 '월/일' · '번호'(더존은 A 열 '일자')."""
    if title(rows) != "분개장":
        return False
    return any(len(r) > 3 and _n(r[1]) == "월/일" and _n(r[3]) == "번호" for r in rows[:15])


def _marker(v):
    """B 열 번호 → 더존 이름 앞머리: Ⅰ. → 'Ⅰ.', ① → '(1)', 1 → 'Ⅰ.'(원가명세서)."""
    if v in (None, ""):
        return ""
    s = str(v).strip()
    if s and s[0] in _ROMAN:
        return s[0] + "."
    if s and s[0] in _CIRCLED:
        return "(%d)" % (_CIRCLED.index(s[0]) + 1)
    if re.fullmatch(r"\d+(\.0)?\.?", s):
        k = int(float(s.rstrip(".")))
        return (_ROMAN[k - 1] + ".") if 1 <= k <= len(_ROMAN) else "(%d)" % k
    return ""


def statement_rows(rows):
    """세무사랑 재무제표 → 더존 Rawdata 줄. 머리(제목 · 기간 · 회사명) 줄은 글만 B 열부터(이름 칸 A 는 비움 - 계정으로 읽히지 않게)."""
    h, name_col, cur, pri = _hdr(rows)
    out = []
    for i, r in enumerate(rows):
        r = list(r) + [None] * 12
        if i <= h + 1:
            out.append([None] + [v for v in r if v not in (None, "")])
            continue
        texts = [str(v).strip() for v in r[name_col + 1:cur] if isinstance(v, str) and v.strip()]
        mk = _marker(r[name_col])
        name = " ".join(texts)
        if not name and not mk:
            out.append([None, None, None, None, None])
            continue
        amt = [r[cur], r[cur + 1]] + ([r[pri], r[pri + 1]] if pri is not None else [None, None])
        amt = [v if _num(v) else None for v in amt]
        out.append([(mk + " " + name).strip()] + amt)
    if title(rows) == "재무상태표":
        _fix_signs(out)
    return out


def _fix_signs(out):
    """세무사랑은 자본조정(주식할인발행차금 · 자기주식) 같은 차감 항목을 묶음 줄만 음수로, 계정 줄은 양수로 보인다 -
    Ⅰ. 묶음 금액 = -(계정 줄 합)이면 계정 줄 부호를 바꿔 더존처럼 음수로(바깥 칸만 있는 줄, 당기 · 전기 따로)."""
    i = 0
    while i < len(out):
        nm = str(out[i][0] or "")
        if not (nm and nm[0] in _ROMAN):
            i += 1
            continue
        j = i + 1
        while j < len(out) and not (str(out[j][0] or "")[:1] in _ROMAN or "총계" in str(out[j][0] or "")):
            j += 1
        kids = [r for r in out[i + 1:j] if r[0] and not str(r[0]).startswith("(") and not str(r[0]).startswith(("당기", "전기"))]
        for col in (2, 4):
            g = out[i][col] or 0
            vals = [r[col] or 0 for r in kids]
            if kids and all(r[col - 1] in (None, 0) for r in kids) and g and abs(g + sum(vals)) < 1 and abs(g - sum(vals)) >= 1:
                for r in kids:
                    if r[col]:
                        r[col] = -r[col]
        i = j


def period(rows, scan=10):
    """머리의 당기 기간 (시작일, 끝일) - '2026.01.01 ~ 2026.06.30' · '제 27 기 2026년 01월 01일부터 2026년 06월 30일까지' · '… 현재'(시작일 None)."""
    for r in rows[:scan]:
        for v in r:
            s = str(v or "")
            ds = [date(*map(int, m)) for m in _DATE.findall(s)]
            if not ds:
                continue
            if len(ds) >= 2:
                return ds[0], ds[1]
            if "현재" in s:
                return None, ds[0]
    return None, None


def company(rows, scan=10):
    for r in rows[:scan]:
        for v in r:
            m = re.search(r"회\s*사\s*명\s*:\s*(.+)", str(v or ""))
            if m:
                return m.group(1).strip()
    return ""


# ───── 계정별원장 ─────
def _load_wb(path):
    from openpyxl import load_workbook
    return load_workbook(path, read_only=True, data_only=True)


def ledger(path):
    """계정별원장 → ({코드: 이름(원가 구분 표시 붙임)}, {코드: [(월-일, 차변, 대변)]}). 코드는 더존 3자리(0103 → 103)."""
    names, lines = {}, defaultdict(list)
    wb = _load_wb(path)
    for ws in wb.worksheets:
        m = ACCT.match(ws.title.strip())
        rows = ws.iter_rows(values_only=True)
        head = []
        for k, r in enumerate(rows):
            if not m and k < 6:
                for v in r:
                    mm = ACCT.match(str(v).strip()) if isinstance(v, str) else None
                    if mm:
                        m = mm
            head.append(r)
            if k >= 6:
                break
        if not m:
            continue
        code = code3(m.group(1))
        names[code] = tagged(code, m.group(2))
        for r in list(head) + list(rows):
            r = list(r) + [None] * 9
            d = str(r[1] or "").strip()
            if re.fullmatch(r"\d{2}-\d{2}", d) and (_num(r[6]) or _num(r[7])):
                lines[code].append((d.replace("-", "/"), float(r[6] or 0.0), float(r[7] or 0.0)))
    wb.close()
    return names, lines


def code3(c):
    c = int(c)
    return c // 100 if c >= 10000 and c % 100 == 0 else c


def tagged(code, name):
    """원장 · 목록 이름엔 (판)(제) 표시가 없다 - 더존 분개장처럼 코드 대역으로 붙임."""
    name = _n(name)
    if name and not TAG.search(name):
        name += {5: "(제)", 6: "(도)", 7: "(분)", 8: "(판)"}.get(code // 100, "")
    return name


def find_ledger(journal_path, end):
    """분개장 옆(같은 폴더 · 한 단계 위)의 계정별원장 중 기간 끝일이 같은 것."""
    d = os.path.dirname(journal_path)
    hits = []
    for base in (d, os.path.dirname(d)):
        hits += [h for h in glob.glob(os.path.join(base, "*.xls*")) if not os.path.basename(h).startswith("~$") and "원장" in os.path.basename(h)]
    for h in sorted(set(hits), key=lambda p: (not os.path.basename(p).startswith("계정별원장"), p)):
        try:
            wb = _load_wb(h)
            ws = wb.worksheets[0]
            top = [list(r) for r in ws.iter_rows(max_row=8, values_only=True)]
            ok = ACCT.match(ws.title.strip()) and period(top)[1] == end
            wb.close()
        except Exception:                                              # noqa: BLE001 - 원장이 아닌 엑셀
            continue
        if ok:
            return h
    return None


# ───── 분개장 ─────
def parse_journal(path, rows, year=None, ledger_path=None):
    """세무사랑 분개장 → douzone.parse_journal 과 같은 dict 줄. year 는 쓰지 않는다(머리 기간에서 연도 - 결산월이 12월이 아니어도)."""
    start, end = period(rows)
    if not end:
        raise ValueError("세무사랑 분개장 머리에서 기간(2026.01.01 ~ 2026.06.30)을 찾지 못함: %s" % path)
    lp = ledger_path or find_ledger(path, end)
    names, lines = ledger(lp) if lp else ({}, {})
    by_name = defaultdict(list)
    for c in sorted(names):
        by_name[names[c]].append(c)
    pool = {c: list(v) for c, v in lines.items()}

    def code_of(name, md, dr, cr):
        cs = by_name.get(name) or by_name.get(TAG.sub("", name)) or []
        if len(cs) > 1:                                                # 같은 이름 여럿 - 원장의 같은 날 · 같은 금액 줄
            for c in cs:
                for k, (d, a, b) in enumerate(pool.get(c, [])):
                    if d == md and abs(a - dr) < 0.5 and abs(b - cr) < 0.5:
                        pool[c].pop(k)
                        return c
        return cs[0] if cs else "*" + name

    def ymd(md):
        mm, dd = map(int, md.split("/"))
        y = end.year
        if start and start.year != end.year and mm >= start.month:
            y = start.year
        return y, mm, dd
    out, cur = [], None

    def close():
        if cur:
            for x in cur["lines"]:
                x["적요"], x["거래처명"] = cur["memo"], cur["vendor"]
                out.append(x)
    for r in rows:
        r = list(r) + [None] * 9
        b1 = str(r[1] or "").strip()
        if re.fullmatch(r"\d{2}/\d{2}", b1):
            close()
            vno = str(int(r[3])) if _num(r[3]) else str(r[3] or "").strip()
            y, mm, dd = ymd(b1)
            cur = {"md": b1, "key": "%d%02d%02d-%s" % (y, mm, dd, vno.zfill(5)), "date": "%d-%02d-%02d" % (y, mm, dd),
                   "type": "매입매출" if vno.isdigit() and int(vno) >= 50000 else "일반", "lines": [], "memo": "", "vendor": ""}
        elif b1:                                                       # 머리 · 합계 줄
            if _n(b1) == "합계":
                close()
                cur = None
            continue
        if cur is None:
            continue
        posted = False
        for nc, ac, side in ((4, 5, "차변"), (6, 7, "대변")):
            if isinstance(r[nc], str) and r[nc].strip() and _num(r[ac]):
                nm = _n(r[nc])
                dr, cr = (float(r[ac]), 0.0) if side == "차변" else (0.0, float(r[ac]))
                cur["lines"].append({"전표키": cur["key"], "전표일자": cur["date"], "계정코드": code_of(nm, cur["md"], dr, cr), "계정과목": nm,
                                     "거래처코드": "", "거래처명": "", "차변": dr, "대변": cr, "적요": "", "유형": cur["type"], "구분": side})
                posted = True
        if not posted and r[2] in (None, "") and (r[4] not in (None, "") or r[6] not in (None, "")):
            cur["memo"], cur["vendor"] = str(r[4] or "").strip(), str(r[6] or "").strip()     # 전표 끝 줄: E 적요 · G 거래처
    close()
    return out


# ───── 데이터 내보내기 양식(제목 · 회사명 줄 없이 머리 두 줄부터) ─────
#   재무제표   A '과목' · B '제 N(당)기 2026년1월1일 ~ 2026년9월30일' · D '(전)기', 둘째 줄 금액 · 집계금액(원가명세서는 '금액' 뿐),
#             번호는 이름 앞(Ⅰ.매출액 · ① 당좌자산 · 1.원재료비), 빈 금액은 0 → 출력 양식 줄([None, 번호, 이름, 당기 안 · 밖, 전기 안 · 밖])로 바꿔 기존 읽기에 태움.
#   분개장     A 일자 · B 번호 · C 구분(차변 · 대변 · 출금 · 입금 · 결차 · 결대) · D 코드 4자리 · E 계정과목 · F 차변 · G 대변 · H 적요 · I 거래처코드 · J 거래처명
#             출금 · 입금도 현금 상대 줄이 따로 있어 줄 그대로 쓴다(전표 차대 대조 없음 - 회사 장부 그대로).
def _flat_head(rows):
    """앞 두 줄로 데이터 내보내기 양식인가 - 'statement' · 'journal' · ''."""
    if len(rows) < 2:
        return ""
    r0, r1 = [_n(v) for v in rows[0]], [_n(v) for v in rows[1]]
    if r0[:3] == ["일자", "번호", "구분"] and len(r1) > 9 and r1[3] == "코드" and r1[5] == "차변" and r1[6] == "대변":
        return "journal"
    if r0 and r0[0] == "과목" and len(r0) > 1 and "(당)기" in r0[1] and r1 and r1[0] == "과목":
        return "statement"
    return ""


def _flat_kind(rows):
    names = [_n(r[0]) for r in rows[2:] if r]
    if any("부채와자본총계" in s for s in names):
        return "재무상태표"
    if any("당기순이익" in s or "당기순손실" in s or "영업이익" in s for s in names):
        return "손익계산서"
    return "제조원가명세서"


def _split_marker(s, kind):
    s = s.strip()
    if s and (s[0] in _ROMAN or s[0] in _CIRCLED):
        return s[0] + ("." if s[0] in _ROMAN else ""), s[1:].lstrip(". ").strip()
    m = re.match(r"^(\d{1,2})\.\s*(.*)$", s) if kind == "제조원가명세서" else None
    return (int(m.group(1)), m.group(2).strip()) if m else ("", s)        # 원가명세서 번호 = 숫자 칸(출력 양식처럼 - 이름에 붙지 않게)


def from_flat(rows):
    """데이터 내보내기 재무제표 줄 → 출력 양식 줄(제목 · 머리 · 과목 줄). 그 양식이 아니면 그대로."""
    if _flat_head(rows) != "statement":
        return rows
    kind = _flat_kind(rows)
    r0 = list(rows[0]) + [None] * 4
    hd = [str(r0[1] or ""), str(r0[3] or "")]
    if kind == "재무상태표":                                            # 재무상태표 날짜 = 기간 끝일('… 현재' - douzone.parse_bs 는 머리의 첫 날짜를 씀)
        hd = [re.sub(r"\d{4}\s*년.*$", "", h) + ("%d년%d월%d일 현재" % tuple(map(int, _DATE.findall(h)[-1])) if _DATE.findall(h) else "") for h in hd]
    out = [[None, kind], [None, "과목", None, hd[0], None, hd[1] or None, None], [None, "과목", None, "금액", None, "금액", None]]
    for r in rows[2:]:
        r = list(r) + [None] * 5
        mk, name = _split_marker(str(r[0] or ""), kind)
        out.append([None, mk or None, name or None] + [v if _num(v) and v != 0 else None for v in r[1:5]])
    return out


def _flat_rows(path):
    if path.lower().endswith(".xls"):
        from .douzone import need_xlrd
        sh = need_xlrd().open_workbook(path, on_demand=True).sheet_by_index(0)
        return [sh.row_values(i) for i in range(sh.nrows)]
    wb = _load_wb(path)
    try:
        return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()


def _flat_dates(rows):
    ds = sorted({str(r[0])[:10] for r in rows[2:] if r and re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(r[0])[:10])})
    return ds


def _flat_journal_info(path):
    """데이터 내보내기 분개장이면 {kind 분개장, start, end, company, slips} - 기간 = 첫 · 끝 전표일(머리에 기간이 없음)."""
    if not path.lower().endswith((".xls", ".xlsx", ".xlsm")) or os.path.basename(path).startswith("~$"):
        return None
    head, _ = _sheet_rows_raw(path, 3)
    if _flat_head(head) != "journal":
        return None
    ds = _flat_dates(_flat_rows(path))
    if not ds:
        return None
    s, e = (date(*map(int, d.split("-"))) for d in (ds[0], ds[-1]))
    return {"kind": "분개장", "start": s, "end": _sibling_end(path, s, e), "company": "", "slips": True}


def is_flat_journal(path):
    try:
        return _flat_head(_sheet_rows_raw(path, 3)[0]) == "journal"
    except Exception:                                                  # noqa: BLE001
        return False


def parse_flat_journal(path):
    """데이터 내보내기 분개장 → douzone.parse_journal 과 같은 dict 줄(계정코드 4자리 → 더존 3자리 숫자)."""
    out = []
    for r in _flat_rows(path)[2:]:
        r = list(r) + [None] * 10
        d = str(r[0] or "")[:10]
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", d) or not re.fullmatch(r"\d{3,6}", str(r[3] or "").strip()):
            continue
        vno = str(r[1] or "").strip()
        vno = str(int(float(vno))) if re.fullmatch(r"\d+(\.0)?", vno) else vno
        code = code3(str(r[3]).strip())
        dr, cr = (float(r[5]) if _num(r[5]) else 0.0), (float(r[6]) if _num(r[6]) else 0.0)
        vc = str(int(r[8])) if _num(r[8]) else str(r[8] or "").strip()
        out.append({"전표키": "%s-%s" % (d.replace("-", ""), vno.zfill(5)), "전표일자": d, "계정코드": code, "계정과목": tagged(code, r[4]),
                    "거래처코드": vc, "거래처명": str(r[9] or "").strip(), "차변": dr, "대변": cr, "적요": str(r[7] or "").strip(),
                    "유형": "매입매출" if vno.isdigit() and int(vno) >= 50000 else "일반", "구분": _n(r[2])})
    return out


# ───── 파일 알아보기([세무사랑 불러오기]) ─────
def head_rows(path, n=20):
    """첫 시트의 앞 n 줄 · 시트 이름. 첫 시트에 세무사랑 제목이 없으면 제목이 있는 다른 시트(분개장: 앞에 회사 가공 시트 '분개장', 원본은 'Col1')."""
    rows, name = _sheet_rows(path, n)
    return rows, name


def _sheet_rows(path, n=None):
    """(줄, 시트 이름) - 세무사랑 제목(KINDS)이 앞 줄에 있는 첫 시트, 없으면 첫 시트. n = 읽을 줄 수(None = 전부).
    데이터 내보내기 재무제표(제목 줄 없음)는 출력 양식 줄로 바꿔 준다(과목 종류는 끝 줄까지 봐야 알아 전부 읽음)."""
    rows, name = _sheet_rows_raw(path, n)
    if _flat_head(rows) == "statement":
        if n is not None:
            rows, name = _sheet_rows_raw(path)
        rows = from_flat(rows)
        if n is not None:
            rows = rows[:n]
    return rows, name


def _sheet_rows_raw(path, n=None):
    if path.lower().endswith(".xls"):
        from .douzone import need_xlrd
        wb = need_xlrd().open_workbook(path, on_demand=True)
        shs = [wb.sheet_by_index(i) for i in range(wb.nsheets)]
        sh = next((x for x in shs if title([x.row_values(i) for i in range(min(10, x.nrows))])), shs[0])
        return [sh.row_values(i) for i in range(sh.nrows if n is None else min(n, sh.nrows))], sh.name
    wb = _load_wb(path)
    try:
        ws = next((x for x in wb.worksheets if title([list(r) for r in x.iter_rows(max_row=10, values_only=True)])), wb.worksheets[0])
        return [list(r) for r in ws.iter_rows(max_row=n, values_only=True)], ws.title
    finally:
        wb.close()


def identify(path):
    """세무사랑 출력 파일 → {kind, start, end, company} 또는 None(세무사랑 양식이 아님)."""
    try:
        fj = _flat_journal_info(path)                                  # 데이터 내보내기 분개장(코드 · 차대 있음)
        if fj:
            return fj
        rows, sheet = head_rows(path)
    except Exception:                                                  # noqa: BLE001
        return None
    t = title(rows)
    if t in STATEMENTS and not is_statement(rows):
        return None
    if t == "분개장" and not is_journal(rows):
        return None
    if t == "계정별원장" and not (ACCT.match(sheet.strip()) and any(_n(r[1] if len(r) > 1 else "") == "일자" for r in rows)):
        return None
    if t == "전표" and not _is_slip(rows):
        return None
    if t == "전표":                                                    # 전 기간 전표(차대 일치) = 분개장
        full = _sheet_rows(path)[0]
        if is_slip_journal(full):
            t = "분개장"
            s, e = period(full)
            if e:
                e = _sibling_end(path, s, e)
                return {"kind": t, "start": s, "end": e, "company": company(full), "slips": True}
    if not t and _is_schedule(rows):
        t = "결산부속명세서"
    if t not in KINDS:
        return None
    s, e = period(rows)
    return {"kind": t, "start": s, "end": e, "company": company(rows)} if e else None


def is_file(path):
    """세무사랑 출력 파일인가(재무제표 · 분개장 · 계정별원장 · 전표 · 결산부속명세서)."""
    return bool(identify(path))


def _is_slip(rows):
    return title(rows) == "전표" and any(len(r) > 3 and _n(r[1]) == "전표일자" and _n(r[2]) == "번호" for r in rows[:15])


def _is_schedule(rows):
    return any("계정과목코드" in [_n(v) for v in r] for r in rows[:10]) and any(_n(v).endswith("명세서") for r in rows[:4] for v in r)


def slip_lines(rows):
    """전표(사용자 정의양식) 줄 → [(월/일, 번호, 구분, 코드 6자리, 이름, 차변, 대변, 거래처코드, 거래처명, 적요)].
    세무사랑 출금 · 입금 전표는 한쪽 줄만 나온다(출금 = 차변 계정, 입금 = 대변 계정) - 상대 현금(0101) 줄을 같은 금액으로 채운다."""
    out = []
    for r in rows:
        r = list(r) + [None] * 20
        md = r[1].strip() if isinstance(r[1], str) else ""
        if not re.fullmatch(r"\d{2}/\d{2}", md) or not (isinstance(r[5], str) and re.fullmatch(r"\d{5,6}", r[5].strip())):
            continue
        vno = str(int(r[2])) if _num(r[2]) else str(r[2] or "").strip()
        g = _n(r[3])
        dr, cr = (float(r[8]) if _num(r[8]) else 0.0), (float(r[10]) if _num(r[10]) else 0.0)
        vc = r[11]
        vc = (str(int(vc)) if _num(vc) else str(vc or "").strip())
        line = [md, vno.zfill(5), g, r[5].strip(), _n(r[6]), dr, cr, vc, str(r[14] or "").strip(), str(r[17] or "").strip()]
        out.append(tuple(line))
        if g in ("출금", "입금"):                                      # 상대 현금 줄(거래처 · 적요는 같은 줄 것)
            amt = dr if g == "출금" else cr
            out.append((md, line[1], g, "010100", "현금", 0.0 if g == "출금" else amt, amt if g == "출금" else 0.0, "", line[8], line[9]))
    return out


def is_slip_journal(rows):
    """전 기간 전표인가 - 제목 '전표' 양식이고 전표마다 차변 = 대변(출금 · 입금은 현금 줄을 채워서). 일부 계정만 뽑은 출력은 전표가 안 맞는다."""
    if not _is_slip(rows):
        return False
    lines = slip_lines(rows)
    if not lines:
        return False
    bal = defaultdict(float)
    for ln in lines:
        bal[(ln[0], ln[1])] += ln[5] - ln[6]
    return all(abs(v) < 0.5 for v in bal.values())


def _sibling_end(path, start, end):
    """전표 머리 기간은 출력 조건 그대로(2026.01.01 ~ 2026.12.31)라 실제 자료 기준일이 아닐 수 있다 - 같은 폴더 세무사랑 재무상태표의 기준일이
    그 기간 안이면 그 날을 쓴다(분개장: 머리 12.31, 전표는 9월 초까지 · 재무상태표 2026.06.30 - 기준일 뒤 전표는 기준일 옵션이 뺀다)."""
    d = os.path.dirname(path)
    for h in sorted(glob.glob(os.path.join(d, "*.xls*"))):
        if h == path or os.path.basename(h).startswith("~$"):
            continue
        try:
            rows, _ = head_rows(h)
        except Exception:                                              # noqa: BLE001
            continue
        if title(rows) == "재무상태표" and is_statement(rows):
            e = period(rows)[1]
            if e and (start is None or start <= e) and e <= end:
                return e
    return end


def slip_journal_rows(path):
    """전 기간 전표 파일이면 그 시트 줄, 아니면 None(douzone.parse_journal 이 더존 '일자' 머리가 없을 때 부름)."""
    rows = _sheet_rows(path)[0]
    return rows if is_slip_journal(rows) else None


def parse_slips(path, rows=None):
    """전 기간 전표(사용자 정의양식) → douzone.parse_journal 과 같은 dict 줄. 연도는 머리 기간(결산월이 12월이 아니면 시작 월 이후 = 시작 연도)."""
    if rows is None:
        rows = _sheet_rows(path)[0]
    start, end = period(rows)
    if not end:
        raise ValueError("세무사랑 전표 머리에서 기간(2026년 01월 01일 ~ 2026년 12월 31일)을 찾지 못함: %s" % path)
    out = []
    for md, vno, g, code6, name, dr, cr, vc, vn, memo in slip_lines(rows):
        mm, dd = map(int, md.split("/"))
        y = start.year if start and start.year != end.year and mm >= start.month else end.year
        code = code3(code6)
        out.append({"전표키": "%d%02d%02d-%s" % (y, mm, dd, vno), "전표일자": "%d-%02d-%02d" % (y, mm, dd), "계정코드": code,
                    "계정과목": tagged(code, name), "거래처코드": vc, "거래처명": vn, "차변": dr, "대변": cr, "적요": memo,
                    "유형": "매입매출" if vno.isdigit() and int(vno) >= 50000 else "일반", "구분": g})
    return out


def code_book(path):
    """계정코드 자료 → {코드(더존 3자리): 이름}. 계정별원장(시트 이름) · 전표(F 코드 · G 이름) · 결산부속명세서(코드 칸 · 그 뒤 이름 칸)."""
    info = identify(path)
    kind = info["kind"] if info else ""
    if kind == "계정별원장":
        return ledger(path)[0]
    book = {}
    wb = _load_wb(path)
    for ws in wb.worksheets:
        for r in ws.iter_rows(values_only=True):
            r = list(r) + [None] * 8
            if kind == "전표":
                c, nm = r[5], r[6]
            else:
                j = next((k for k, v in enumerate(r[:6]) if isinstance(v, str) and re.fullmatch(r"\d{6}", v.strip())), None)
                if j is None:
                    continue
                c = r[j]
                nm = next((v for v in r[j + 1:j + 4] if isinstance(v, str) and v.strip()), None)
            if isinstance(c, str) and re.fullmatch(r"\d{5,6}", c.strip()) and isinstance(nm, str) and nm.strip():
                code = code3(c.strip())
                book.setdefault(code, tagged(code, nm))
    wb.close()
    return book


def books_near(raw, key):
    """자료 폴더와 그 옆 폴더(다른 기간)의 코드 자료 → {코드: 이름}. 회사 계정코드는 해마다 같으므로 다른 기간 자료도 쓰고, 같은 꼬리 자료가 이긴다."""
    base = os.path.dirname(os.path.abspath(raw))
    own, other = [], []
    for kind in BOOKS:
        for h in glob.glob(os.path.join(base, "*", kind + "_*.xls*")):
            if os.path.basename(h).startswith("~$"):
                continue
            (own if re.search(r"_%s(?!\d)" % key, os.path.basename(h)) else other).append(h)
    book = {}
    for h in sorted(other) + sorted(own):
        try:
            book.update({str(c): n for c, n in code_book(h).items()})
        except Exception:                                              # noqa: BLE001 - 코드 자료는 보조(못 읽어도 TB 는 만든다)
            continue
    return book
