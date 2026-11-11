"""더존(위하고)에서 내려받은 엑셀 읽기 - 분개장, 재무상태표, 손익계산서, 제조원가명세서.

재무제표는 한 줄이 [과목, 당기 안쪽 칸, 당기 바깥 칸, 전기 안쪽 칸, 전기 바깥 칸] 인 모양으로 읽는다.
세무사랑 출력도 같은 모양으로 바꿔 읽는다(semusarang).
- 계정코드: 3자리 정수 기준
- 계정과목 이름: norm_name() 으로 공백을 뺀 것을 열쇠로 쓴다
"""
import os
import re
from datetime import date

from ..guard import InputError
from . import semusarang

FULLWIDTH_SPACE = "\u3000"


def norm_name(s):
    """계정과목명 매칭 키: 모든 공백(전각 포함) 제거."""
    if s is None:
        return ""
    return str(s).replace(FULLWIDTH_SPACE, "").replace(" ", "").strip()


def to_num(v):
    """더존 셀 값 → float. 빈 값/문자는 0. '△1,234'(부호 기호)는 음수."""
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    neg = s.startswith("△") or s.startswith("(") and s.endswith(")")
    s = s.strip("△()")
    if s in ("", "-"):
        return 0.0
    try:
        n = float(s)
    except ValueError:
        return 0.0
    return -n if neg else n


def year_from_filename(path):
    m = re.search(r"_(\d{4})(?!\d)", os.path.basename(path))           # _202509(기준연월 파일)는 연도 파일이 아님
    if not m:
        raise ValueError(f"파일명에서 연도(_YYYY)를 찾을 수 없음: {path}")
    return int(m.group(1))


def need_xlrd():
    """옛 엑셀(.xls)을 읽는 패키지. 없으면 설치 방법을 알린다."""
    try:
        import xlrd
    except ModuleNotFoundError:
        raise InputError("옛 엑셀 형식(.xls)을 읽으려면 xlrd 가 필요합니다 - 엑셀에서 .xlsx 로 저장하거나 xlrd 를 설치하세요")
    return xlrd


def _load_rows(path):
    """xlsx/xls 공통: 시트 전체를 2차원 리스트로."""
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    else:
        xlrd = need_xlrd()
        wb = xlrd.open_workbook(path)
        ws = wb.sheet_by_index(0)
        rows = [[ws.cell_value(r, c) for c in range(ws.ncols)] for r in range(ws.nrows)]
    rows = semusarang.from_flat(rows)                                  # 세무사랑 데이터 내보내기(제목 줄 없음) → 출력 양식 줄
    if semusarang.is_statement(rows):                                  # 세무사랑 재무제표 출력 → 더존 줄 모양
        rows = semusarang.statement_rows(rows)
    return rows



# ──────────────────────────────────────────────────────────────
# 분개장
# ──────────────────────────────────────────────────────────────
def parse_journal(path, year=None):
    """분개장 → [{전표키, 전표일자, 계정코드, 계정과목, 거래처코드, 거래처명,
                 차변, 대변, 적요, 유형, 구분}]"""
    if semusarang.is_flat_journal(path):                                # 세무사랑 데이터 내보내기 분개장(코드 · 차대 있음)
        return semusarang.parse_flat_journal(path)
    if semusarang.title(semusarang.head_rows(path)[0]) == "전표":         # 세무사랑 전 기간 전표(사용자 정의양식, 코드 있음)
        srows = semusarang.slip_journal_rows(path)
        if srows:
            return semusarang.parse_slips(path, srows)
    rows = _load_rows(path)
    if semusarang.is_journal(rows):                                     # 세무사랑 분개장(코드는 같은 기간 계정별원장에서)
        return semusarang.parse_journal(path, rows, year)
    year = year or year_from_filename(path)

    # 헤더 행 탐색: 첫 열이 '일자'
    hdr = next(i for i, r in enumerate(rows) if norm_name(r[0] if r else "") == "일자")
    out = []
    for r in rows[hdr + 1:]:
        date_raw = r[0]
        if not (isinstance(date_raw, str) and re.match(r"^\d{2}/\d{2}$", date_raw)):
            continue  # 합계행·빈 행
        mm, dd = date_raw.split("/")
        vno = str(r[1]).strip().zfill(5)
        code = int(to_num(r[2]))
        out.append({
            "전표키": f"{year}{mm}{dd}-{vno}",
            "전표일자": f"{year}-{mm}-{dd}",
            "계정코드": code,
            "계정과목": str(r[3] or "").strip(),
            "거래처코드": str(r[8] or "").strip(),
            "거래처명": str(r[9] or "").strip(),
            "차변": to_num(r[5]),
            "대변": to_num(r[6]),
            "적요": str(r[4] or "").strip(),
            "유형": "매입매출" if int(vno) >= 50000 else "일반",
            "구분": str(r[7] or "").strip(),
        })
    return out


# ──────────────────────────────────────────────────────────────
# 재무제표 (재무상태표 / 손익계산서 / 제조원가명세서)
# ──────────────────────────────────────────────────────────────
_ROMAN = "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ"


def _is_group_name(raw_name):
    s = str(raw_name).strip()
    return bool(s) and (s[0] in _ROMAN or s.startswith("("))


_PERIOD_RE = re.compile(r"제\s*(\d+)\s*(?:\(\s*당\s*\))?\s*기")


def _period_no(rows, scan=12):
    """재무제표 머리글에서 당기 기수(제 N기)를 읽는다. 없으면 None."""
    for r in rows[:scan]:
        for v in r:
            m = _PERIOD_RE.search(str(v or ""))
            if m:
                return int(m.group(1))
    return None


_DATE_RE = re.compile(r"(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일")
_COMPANY_RE = re.compile(r"회\s*사\s*명\s*:\s*(.+)")


def _head_match(rows, pattern, scan=12):
    """재무제표 머리글(앞 scan 행)에서 처음 맞는 정규식 결과."""
    for r in rows[:scan]:
        for v in r:
            m = pattern.search(str(v or ""))
            if m:
                return m
    return None


def parse_bs(path):
    """재무상태표 → {'leaves': [(계정과목키, 당기, 전기, 차감여부)], 'totals': {...}}
    누계액 행은 gross 금액(당기 col1/전기 col3), 일반 행은 col2/col4.
    부채·자본 leaf 는 양수 그대로 반환(부호 처리는 호출측)."""
    rows = _load_rows(path)
    leaves, totals = [], {}
    section = ""
    for r in rows:
        name = str(r[0] or "").strip()
        if not name:
            continue
        key = norm_name(name)
        c = [to_num(r[i]) if len(r) > i else 0.0 for i in range(5)]
        if key in ("자산총계", "부채총계", "자본총계", "부채및자본총계"):
            totals[key] = (c[2], c[4])
            continue
        if key in ("자산", "부채", "자본") and not any(c[1:]):
            section = key
            continue
        if _is_group_name(name) or key.startswith("당기") or key.startswith("전기") \
                or key.startswith("(당기순이익"):
            continue
        if key == "감가상각누계액":
            leaves.append((key, c[1], c[3], True, section))
        elif c[1] and not c[2]:
            leaves.append((key, c[1], c[3], False, section))
        elif c[2] or c[4]:
            leaves.append((key, c[2], c[4], False, section))
    company = _head_match(rows, _COMPANY_RE)
    closing = _head_match(rows, _DATE_RE)
    return {"leaves": leaves, "totals": totals, "period": _period_no(rows),
            "company": company.group(1).strip() if company else None,
            "closing": date(*map(int, closing.groups())) if closing else None}


NO_COST = {"leaves": [], "summary": {}}      # 제조원가명세서가 없는 회사(제조업 아님)


def is_mfg(is_path):
    """손익계산서에 당기제품제조원가(당기 · 전기 어느 쪽이든 0 아님)가 있으면 제조업 - 제조원가명세서가 있어야 한다."""
    for r in _load_rows(is_path):
        if "제품제조원가" in norm_name(str(r[0] or "")) and any(to_num(v) for v in r[1:5]):
            return True
    return False


def parse_cost(path):
    """제조원가명세서 → {'leaves': [(그룹, 계정과목키, 당기, 전기)], 'summary': {...}}. 파일이 없는(제조업 아닌) 회사는 NO_COST."""
    rows = _load_rows(path)
    leaves, summary = [], {}
    group = ""
    for r in rows:
        name = str(r[0] or "").strip()
        if not name:
            continue
        key = norm_name(name)
        c = [to_num(r[i]) if len(r) > i else 0.0 for i in range(5)]
        if _is_group_name(name):
            g = norm_name(key.lstrip("".join(_ROMAN) + ". ").strip("."))
            summary[g] = (c[2], c[4])
            if g == "원재료비":
                leaves.append(("재료비", "원재료비", c[2], c[4]))
                group = ""
            elif g == "노무비":
                group = "노무비"
            elif g == "경비":
                group = "제조경비"
            else:
                group = ""
            continue
        if key.startswith(("기초", "기말")) or "매입액" in key:
            continue
        if group and (c[1] or c[3]):
            leaves.append((group, key, c[1], c[3]))
    return {"leaves": leaves, "summary": summary}
