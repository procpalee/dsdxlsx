"""자체감사 — 산출 워크북에 기록된 검증 수식을 독립적으로 재평가한다.

단일 진실 보증: CheckSet(Python 판정)과 수식 문자열이 같은 데이터에서 나오지만,
혹시 모를 드리프트(좌표 바인딩 실수·수식 렌더 버그)를 잡기 위해 **기록된 수식
텍스트를 파싱해 워크북의 실제 셀 값으로 재계산**하고 CheckSet 판정과 대조한다.
"""
from __future__ import annotations

import re
from decimal import ROUND_DOWN, ROUND_HALF_UP, Decimal

from openpyxl.utils import column_index_from_string

from .checks import CheckSet

_FOOT_RE = re.compile(r"^=IF\(ABS\(IF\((?P<d>.+?)=TRUNC\(")
_CROSS_RE = re.compile(r"^=IF\(ABS\((?P<src>.+?)\)=ABS\((?P<dst>.+?)\),TRUE")
_REF_RE = re.compile(r"^(?:'(?P<qs>[^']+)'!|(?P<s>[A-Za-z가-힣0-9]+)!)?(?P<col>[A-Z]{1,3})(?P<row>\d+)$")
_ROUNDER_RE = re.compile(r"^(ROUND|TRUNC)\((?P<inner>.+)/(?P<k>\d+),0\)$")


class AuditError(Exception):
    pass


def _cell(wb, sheet: str, ref: str, default_sheet: str):
    m = _REF_RE.match(ref)
    if not m:
        raise AuditError(f"참조 파싱 실패: {ref}")
    sname = m.group("qs") or m.group("s") or default_sheet
    v = wb[sname].cell(int(m.group("row")), column_index_from_string(m.group("col"))).value
    if isinstance(v, str) and v.startswith("="):
        raise AuditError(f"수식 셀을 항으로 참조: {sname}!{ref}")
    if v is None or v == "":
        return Decimal(0)
    if isinstance(v, bool) or isinstance(v, str):
        raise AuditError(f"숫자 아닌 항: {sname}!{ref} = {v!r}")
    return Decimal(str(v))


def _split_terms(expr: str) -> list[tuple[int, str]]:
    """깊이 0의 +,- 로 분리 → [(부호, 토큰)]."""
    out, depth, cur, sign = [], 0, "", 1
    for ch in expr:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if depth == 0 and ch in "+-" and cur:
            out.append((sign, cur))
            sign = 1 if ch == "+" else -1
            cur = ""
            continue
        if depth == 0 and ch in "+-" and not cur:
            sign = sign * (1 if ch == "+" else -1)
            continue
        cur += ch
    if cur:
        out.append((sign, cur))
    return out


def _eval_operand(wb, token: str, sheet: str) -> Decimal:
    token = token.strip()
    if token.startswith("(") and token.endswith(")"):
        return _eval_sum(wb, token[1:-1], sheet)
    for op in ("*", "/"):
        if op in token:
            a, b = token.split(op, 1)
            va, vb = _eval_operand(wb, a, sheet), _eval_operand(wb, b, sheet)
            return va * vb if op == "*" else va / vb
    if re.fullmatch(r"\d+(\.\d+)?", token):
        return Decimal(token)   # 숫자 리터럴 (비율 수식의 배율 *100)
    return _cell(wb, sheet, token, sheet)


def _eval_sum(wb, expr: str, sheet: str) -> Decimal:
    total = Decimal(0)
    for sign, token in _split_terms(expr):
        total += sign * _eval_operand(wb, token, sheet)
    return total


def _eval_footing(wb, formula: str, sheet: str):
    """반환: True(일치) 또는 ABS 차이(Decimal)."""
    m = _FOOT_RE.match(formula)
    if not m:
        raise AuditError(f"풋팅 수식 형태 아님: {formula[:60]}")
    d_expr = m.group("d")
    tm = re.match(r"^(.+?)-\((.+)\)$", d_expr)
    if not tm:
        raise AuditError(f"d 파싱 실패: {d_expr[:60]}")
    target = _cell(wb, sheet, tm.group(1), sheet)
    d = target - _eval_sum(wb, tm.group(2), sheet)
    # ROUND(d, n) 판정 자릿수는 수식 자체에서 읽는다 (비율 검증은 n>1)
    dm = re.search(r"ROUND\(.*,(\d+)\)\)\)<0\.000001", formula)
    digits = int(dm.group(1)) if dm else 1
    q = Decimal(1).scaleb(-digits)
    d2 = d if d == d.to_integral_value() else d.quantize(q, ROUND_HALF_UP)
    return True if abs(d2) < Decimal("0.000001") else abs(d2)


def _eval_side(wb, expr: str, sheet: str) -> Decimal:
    expr = expr.strip()
    rm = _ROUNDER_RE.match(expr)
    if rm:
        inner = _eval_sum(wb, rm.group("inner"), sheet) / Decimal(rm.group("k"))
        rounding = ROUND_HALF_UP if expr.startswith("ROUND") else ROUND_DOWN
        return inner.quantize(Decimal(1), rounding)
    return _eval_sum(wb, expr, sheet)


def _eval_cross(wb, formula: str, sheet: str):
    m = _CROSS_RE.match(formula)
    if not m:
        raise AuditError(f"대사 수식 형태 아님: {formula[:60]}")
    src = abs(_eval_side(wb, m.group("src"), sheet))
    dst = abs(_eval_side(wb, m.group("dst"), sheet))
    return True if src == dst else abs(src - dst)


def audit(wb, checkset: CheckSet) -> list[str]:
    """워크북 수식 재평가 ↔ CheckSet 판정 대조. 반환: 불일치 설명 목록(비면 통과)."""
    problems: list[str] = []
    for fc in checkset.footing:
        cell = wb[fc.formula_cell.sheet].cell(fc.formula_cell.row, fc.formula_cell.col)
        try:
            got = _eval_footing(wb, str(cell.value), fc.formula_cell.sheet)
        except AuditError as e:
            problems.append(f"{fc.formula_cell.sheet}!{fc.formula_cell.addr}: {e}")
            continue
        want = fc.excel_result()
        if got != want:
            problems.append(
                f"{fc.formula_cell.sheet}!{fc.formula_cell.addr}: 수식 재평가 {got} ≠ 판정 {want}")
    for xr in checkset.cross:
        cell = wb[xr.formula_cell.sheet].cell(xr.formula_cell.row, xr.formula_cell.col)
        try:
            got = _eval_cross(wb, str(cell.value), xr.formula_cell.sheet)
        except AuditError as e:
            problems.append(f"{xr.formula_cell.sheet}!{xr.formula_cell.addr}: {e}")
            continue
        want = True if xr.passed() else xr.diff()
        if got != want:
            problems.append(
                f"{xr.formula_cell.sheet}!{xr.formula_cell.addr}: 대사 재평가 {got} ≠ 판정 {want}")
    return problems
