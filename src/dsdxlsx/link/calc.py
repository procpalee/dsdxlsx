"""정산표 수식을 엑셀 없이 계산한다.

정산표가 쓰는 수식만 다룬다: 사칙연산, 괄호, 칸 참조(다른 시트 포함), SUM, SUMIF.
그 밖의 함수가 나오면 CalcError 로 멈춘다(엑셀에서 열어 저장한 파일을 쓰면 된다).
"""
from __future__ import annotations

import re

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, range_boundaries


class CalcError(Exception):
    pass


_TOKEN = re.compile(
    r"\s*(?:"
    r"(?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)"
    r"|(?P<str>\"(?:[^\"]|\"\")*\")"
    r"|(?P<ref>(?:'(?:[^']|'')+'|[^\s'!+\-*/(),=<>&:\"]+)!\$?[A-Za-z]{1,3}\$?\d*(?::\$?[A-Za-z]{1,3}\$?\d*)?"
    r"|\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})"
    r"|(?P<name>[A-Za-z_][A-Za-z0-9_.]*)(?=\s*\()"
    r"|(?P<op>[+\-*/(),])"
    r")")


def _num(v) -> float:
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    if v is None or v == "":
        return 0.0
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return 0.0


def _criteria(crit):
    """SUMIF 조건 → 판정 함수. 글은 대소문자 무시, * ? 는 와일드카드(~ 로 풀기), 앞의 = <> > < >= <= 연산자."""
    if crit is None or crit == "":
        return lambda v: False
    if isinstance(crit, (int, float)) and not isinstance(crit, bool):
        return lambda v: isinstance(v, (int, float)) and not isinstance(v, bool) and float(v) == float(crit)
    s = str(crit)
    m = re.match(r"^(<>|>=|<=|=|>|<)(.*)$", s)
    op, body = (m.group(1), m.group(2)) if m else ("=", s)
    try:
        target = float(body)
        numeric = True
    except ValueError:
        numeric = False
    if op in (">", "<", ">=", "<="):
        if not numeric:
            return lambda v: False
        return {">": lambda v: isinstance(v, (int, float)) and v > target, "<": lambda v: isinstance(v, (int, float)) and v < target,
                ">=": lambda v: isinstance(v, (int, float)) and v >= target, "<=": lambda v: isinstance(v, (int, float)) and v <= target}[op]
    pat, i = "", 0
    while i < len(body):
        ch = body[i]
        if ch == "~" and i + 1 < len(body):
            pat += re.escape(body[i + 1])
            i += 2
            continue
        pat += ".*" if ch == "*" else "." if ch == "?" else re.escape(ch)
        i += 1
    rx = re.compile(pat, re.I | re.S)

    def eq(v):
        if numeric and isinstance(v, (int, float)) and not isinstance(v, bool):
            return float(v) == target
        return v is not None and rx.fullmatch(str(v)) is not None

    return eq if op == "=" else (lambda v: not eq(v))


class Calc:
    """통합문서(수식 그대로 읽은 것)의 칸 값을 계산해 돌려준다."""

    def __init__(self, wb):
        self.wb = wb
        self.cache: dict = {}
        self.busy: set = set()

    # ── 값 ──
    def value(self, sheet: str, row: int, col: int):
        key = (sheet, row, col)
        if key in self.cache:
            return self.cache[key]
        if key in self.busy:
            raise CalcError("순환 참조: %s!%s%d" % (sheet, _letter(col), row))
        ws = self._ws(sheet)
        v = ws.cell(row, col).value
        if isinstance(v, str) and v.startswith("="):
            self.busy.add(key)
            try:
                v = self._eval(v[1:], sheet, "%s!%s%d" % (sheet, _letter(col), row))
            finally:
                self.busy.discard(key)
        self.cache[key] = v
        return v

    def _ws(self, sheet: str):
        if sheet not in self.wb.sheetnames:
            raise CalcError("없는 시트: %s" % sheet)
        return self.wb[sheet]

    def _range(self, ref: str, here: str):
        """참조 글 → (시트, min_col, min_row, max_col, max_row). 열 전체(E:E)는 쓰인 행까지."""
        sheet = here
        if "!" in ref:
            name, ref = ref.rsplit("!", 1)
            sheet = name[1:-1].replace("''", "'") if name.startswith("'") else name
        ref = ref.replace("$", "")
        ws = self._ws(sheet)
        if re.fullmatch(r"[A-Za-z]{1,3}:[A-Za-z]{1,3}", ref):
            a, b = ref.split(":")
            return sheet, column_index_from_string(a), 1, column_index_from_string(b), ws.max_row
        c1, r1, c2, r2 = range_boundaries(ref)
        return sheet, c1, r1, c2 if c2 is not None else c1, r2 if r2 is not None else r1

    def _cells(self, ref: str, here: str):
        sheet, c1, r1, c2, r2 = self._range(ref, here)
        return [self.value(sheet, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]

    def _eval(self, text: str, here: str, where: str):
        toks, pos = [], 0
        while pos < len(text) and text[pos:].strip():
            m = _TOKEN.match(text, pos)
            if not m or m.end() == pos:
                raise CalcError("계산할 수 없는 수식(%s): =%s" % (where, text))
            toks.append((m.lastgroup, m.group(m.lastgroup)))
            pos = m.end()
        return _Parser(self, toks, here, where, text).run()


class _Parser:
    """수식 하나를 읽는다. 칸 참조를 따라가면 그 칸은 새 _Parser 가 읽는다."""

    def __init__(self, calc: Calc, toks, here: str, where: str, text: str):
        self.calc, self.t, self.i, self.here = calc, toks, 0, here
        self.err = CalcError("계산할 수 없는 수식(%s): =%s" % (where, text))
        self.where = where

    def run(self):
        v = self.expr()
        if self.i != len(self.t):
            raise self.err
        return v

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self, kind=None, val=None):
        k, v = self.peek()
        if k is None or (kind and k != kind) or (val is not None and v != val):
            raise self.err
        self.i += 1
        return v

    def cells(self, ref: str):
        return self.calc._cells(ref, self.here)

    def expr(self):
        v = self.term()
        while self.peek() in (("op", "+"), ("op", "-")):
            op = self.take()
            w = self.term()
            v = _num(v) + _num(w) if op == "+" else _num(v) - _num(w)
        return v

    def term(self):
        v = self.factor()
        while self.peek() in (("op", "*"), ("op", "/")):
            op = self.take()
            w = self.factor()
            if op == "/" and _num(w) == 0:
                raise CalcError("0으로 나눔(%s)" % self.where)
            v = _num(v) * _num(w) if op == "*" else _num(v) / _num(w)
        return v

    def factor(self):
        k, v = self.peek()
        if (k, v) == ("op", "-"):
            self.take()
            return -_num(self.factor())
        if (k, v) == ("op", "+"):
            self.take()
            return _num(self.factor())
        if (k, v) == ("op", "("):
            self.take()
            out = self.expr()
            self.take("op", ")")
            return out
        if k == "num":
            self.take()
            return float(v)
        if k == "str":
            self.take()
            return v[1:-1].replace('""', '"')
        if k == "ref":
            self.take()
            vals = self.cells(v)
            return vals[0] if vals else None
        if k == "name":
            return self.func()
        raise self.err

    def args(self):
        """함수 인자 - 참조 하나만 있는 인자는 글 그대로(범위), 그 밖은 값."""
        self.take("op", "(")
        out = []
        if self.peek() != ("op", ")"):
            while True:
                k, v = self.peek()
                nxt = self.t[self.i + 1] if self.i + 1 < len(self.t) else (None, None)
                if k == "ref" and nxt in (("op", ","), ("op", ")")):
                    self.take()
                    out.append(("ref", v))
                else:
                    out.append(("val", self.expr()))
                if self.peek() == ("op", ","):
                    self.take()
                    continue
                break
        self.take("op", ")")
        return out

    def func(self):
        name = self.take("name").upper()
        args = self.args()
        if name == "SUM":
            total = 0.0
            for kind, a in args:
                for v in (self.cells(a) if kind == "ref" else [a]):
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        total += float(v)
            return total
        if name == "SUMIF" and len(args) in (2, 3) and args[0][0] == "ref":
            crit = args[1]
            crit_v = (self.cells(crit[1]) or [None])[0] if crit[0] == "ref" else crit[1]
            test = _criteria(crit_v)
            keys = self.cells(args[0][1])
            vals = self.cells(args[2][1]) if len(args) == 3 and args[2][0] == "ref" else keys
            return sum(float(v) for k, v in zip(keys, vals)
                       if isinstance(v, (int, float)) and not isinstance(v, bool) and test(k))
        raise CalcError("계산할 수 없는 함수 %s(%s)" % (name, self.where))


def _letter(col: int) -> str:
    from openpyxl.utils import get_column_letter
    return get_column_letter(col)


def freeze(path: str, out: str | None = None) -> int:
    """통합문서의 수식 칸을 계산한 값으로 바꿔 저장한다(시험 사본용). 반환: 바꾼 칸 수."""
    wb = load_workbook(path)
    calc = Calc(wb)
    todo = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    todo.append((ws.title, c.row, c.column))
    values = {key: calc.value(*key) for key in todo}
    for (sheet, r, c), v in values.items():
        wb[sheet].cell(r, c).value = v
    wb.save(out or path)
    return len(todo)
