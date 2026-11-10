"""검증 데이터 모델 — 단일 진실 원칙.

FootingSpec/CrossSpec  : 그리드 좌표 공간 (표, 행 i, 열 c) — 분석 엔진 산출물
FootingCheck/CrossRef  : 셀 좌표 공간 (시트, 행, 열, 값) — layout.bind()가 변환

FootingCheck/CrossRef 하나가 Python 평가(Decimal)와 엑셀 수식 문자열을 모두
생성한다 — 수식이 말하는 것과 파이썬이 검산한 것이 어긋날 수 없다.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal

from openpyxl.utils import get_column_letter

from .constants import FOOTING_TOL

_ALPHA_SHEET = re.compile(r"^[A-Za-z][A-Za-z0-9]*$")


def _q(sheet: str) -> str:
    """수식 내 시트 참조 — 영문 이름은 그대로, 숫자·한글 이름은 인용."""
    return sheet if _ALPHA_SHEET.match(sheet) else f"'{sheet}'"


def _dec(v) -> Decimal:
    return v if isinstance(v, Decimal) else Decimal(v)


def _excel_round(d: Decimal, places: int = 0) -> Decimal:
    """Excel ROUND (절반 올림, 0에서 멀어지는 방향)."""
    q = Decimal(1).scaleb(-places)
    return d.quantize(q, rounding=ROUND_HALF_UP)


def _excel_trunc(d: Decimal) -> Decimal:
    return d.to_integral_value(rounding="ROUND_DOWN")


# ── 그리드 공간 Spec ──────────────────────────────────────────────────

@dataclass
class FootingSpec:
    """풋팅 검증 하나 — 대상 셀 = 항들의 부호합 (op='+') 또는 2항 산술식(-,*,/)."""

    table_id: str
    target: tuple[int, int]                  # (행 i, 그리드 열 c)
    terms: list[tuple[int, int, int]]        # [(행 j, 열 cc, 부호 ±1)]
    label: str
    kind: str        # autosum|group|total_row|total_col|pl_anchor|hint|socie|flow_h|…
    axis: str        # 'col' → 표 아래 | 'row' → 오른쪽 (PL 예외는 layout에서)
    ok: bool
    diff: Decimal    # 기재 − 계산 (0이면 일치)
    op: str = "+"    # '+' 부호합 | '-'|'*'|'/' 2항 산술 (법인세 유효세율·자본금 산식 등)
    digits: int = 1  # 수식 ROUND(d, n) 판정 자릿수 — 비율(분수 저장)은 표시 소수 자리
    factor: int = 1  # '/' 결과 배율 — 비율이 퍼센트 수치(83.6)로 기재된 표는 100


@dataclass
class CrossSpec:
    """대사 검증 하나 — src(참조되는 쪽, 보통 재무제표) vs dst(배치되는 쪽 셀)."""

    src_table: str
    src_cells: list[tuple[int, int]]         # 1~2개 (조합매칭)
    dst_table: str
    dst_cell: tuple[int, int]
    label_path: str                          # "BS:Ⅱ. 비유동자산 -> … -> 선박 "
    scale: int = 1                           # src를 나눌 배수 (dst가 굵은 단위)
    dst_scale: int = 1                       # dst를 나눌 배수 (src가 굵은 단위)
    rounder: str = "ROUND"                   # ROUND | TRUNC
    ok: bool = True
    diff: Decimal = Decimal(0)


# ── 셀 좌표 공간 ─────────────────────────────────────────────────────

@dataclass(frozen=True)
class CellRef:
    sheet: str
    row: int          # 1-base
    col: int          # 1-base
    value: object     # 그 셀에 실제 기록되는 값 (int | Decimal | None)

    @property
    def addr(self) -> str:
        return f"{get_column_letter(self.col)}{self.row}"

    def ref(self, from_sheet: str) -> str:
        if self.sheet == from_sheet:
            return self.addr
        return f"{_q(self.sheet)}!{self.addr}"

    def ref_abs(self) -> str:
        """항상 시트명 포함 (대사 수식은 자기 시트도 접두)."""
        return f"{_q(self.sheet)}!{self.addr}"


@dataclass
class FootingCheck:
    target: CellRef
    terms: list[tuple[CellRef, int]]
    label: str
    kind: str
    axis: str
    placement: str                    # below | right
    op: str = "+"                     # '+' 부호합 | '-'|'*'|'/' 2항 산술
    digits: int = 1                   # ROUND(d, n) 판정 자릿수 (비율은 표시 소수 자리)
    factor: int = 1                   # '/' 결과 배율 (퍼센트 수치 기재 표는 100)
    label_cell: CellRef | None = None
    formula_cell: CellRef | None = None

    # ── 평가 (수식과 같은 데이터에서) ──
    def d(self) -> Decimal:
        if self.op == "+":
            total = sum((_dec(cell.value) * sign for cell, sign in self.terms),
                        Decimal(0))
        else:
            a, b = _dec(self.terms[0][0].value), _dec(self.terms[1][0].value)
            if self.op == "-":
                total = a - b
            elif self.op == "*":
                total = a * b
            else:  # '/'
                total = a / b * self.factor
        return _dec(self.target.value) - total

    def excel_result(self):
        """수식 미러: TRUE 또는 ABS(차이)."""
        d = self.d()
        d2 = d if d == _excel_trunc(d) else _excel_round(d, self.digits)
        return True if abs(d2) < Decimal("0.000001") else abs(d2)

    def passed(self) -> bool:
        return self.excel_result() is True

    def rounding_pass(self) -> bool:
        # 원단위 절사 관행은 정수 스케일 검증에만 해당 — 비율(digits>1)은 제외
        return (self.digits == 1 and not self.passed()
                and abs(self.d()) <= FOOTING_TOL)

    # ── 수식 ──
    def d_expr(self) -> str:
        sheet = self.target.sheet
        if self.op != "+":
            a = self.terms[0][0].ref(sheet)
            b = self.terms[1][0].ref(sheet)
            mul = f"*{self.factor}" if (self.op == "/" and self.factor != 1) else ""
            return f"{self.target.ref(sheet)}-({a}{self.op}{b}{mul})"
        parts = []
        for k, (cell, sign) in enumerate(self.terms):
            r = cell.ref(sheet)
            if k == 0:
                parts.append(f"-{r}" if sign < 0 else r)
            else:
                parts.append(f"-{r}" if sign < 0 else f"+{r}")
        return f"{self.target.ref(sheet)}-({''.join(parts)})"

    def formula(self) -> str:
        d = self.d_expr()
        inner = f"IF({d}=TRUNC({d}),{d},ROUND({d},{self.digits}))"
        return f"=IF(ABS({inner})<0.000001,TRUE,ABS({inner}))"


_LABEL_NUM_RE = re.compile(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ0-9]+\s*[.)]\s*")
_LABEL_NOTE_RE = re.compile(r"\(주석[^)]*\)")


def short_label(label_path: str) -> str:
    """대사 링크 표시용 축약 — 'BS:Ⅱ.비유동자산 -> 파생상품자산' → 'BS > 파생상품자산(비유동)'.

    전체 경로(label_path)는 provenance·AI검토용으로 유지하고 표시만 줄인다:
    소스 시트(주석은 '주석N') + 리프 라벨 + 유동/비유동 구분(조상 경로에서).
    조합(원가+누계 등)은 리프들을 '+'로."""
    leaves: list[str] = []
    qual = ""
    src = ""
    for k, seg in enumerate(s.strip() for s in label_path.split(",")):
        sheet, _, rest = seg.partition(":")
        parts = [p.strip() for p in (rest or seg).split("->")]
        leaf = _LABEL_NUM_RE.sub("", parts[-1])
        leaf = _LABEL_NOTE_RE.sub("", leaf).strip()
        if leaf:
            leaves.append(leaf)
        if k == 0:
            if rest:
                src = f"주석{sheet}" if sheet.strip().isdigit() else sheet.strip()
            for p in parts[:-1]:
                if "비유동" in p:
                    qual = "비유동"
                    break
                if "유동" in p:
                    qual = "유동"
                    break
    label = "+".join(leaves) or label_path
    if len(label) > 22 and len(leaves) > 1:
        label = f"{leaves[0]} 외 {len(leaves) - 1}"
    if qual and qual not in label:
        label += f"({qual})"
    return f"{src} > {label}" if src else label


@dataclass
class CrossRef:
    src: list[CellRef]
    dst: CellRef
    label_path: str
    scale: int = 1
    dst_scale: int = 1
    rounder: str = "ROUND"
    link_cell: CellRef | None = None
    formula_cell: CellRef | None = None

    # ── 평가 ──
    def _side(self, cells_sum: Decimal, scale: int) -> Decimal:
        if scale == 1:
            return abs(cells_sum)
        d = cells_sum / Decimal(scale)
        d = _excel_round(d) if self.rounder == "ROUND" else _excel_trunc(d)
        return abs(d)

    def src_value(self) -> Decimal:
        return self._side(sum((_dec(c.value) for c in self.src), Decimal(0)), self.scale)

    def dst_value(self) -> Decimal:
        return self._side(_dec(self.dst.value), self.dst_scale)

    def passed(self) -> bool:
        return self.src_value() == self.dst_value()

    def diff(self) -> Decimal:
        return abs(self.src_value() - self.dst_value())

    # ── 수식 ──
    def _src_expr(self) -> str:
        refs = "+".join(c.ref_abs() for c in self.src)
        if len(self.src) > 1:
            refs = f"({refs})"
        if self.scale > 1:
            return f"{self.rounder}({refs}/{self.scale},0)"
        return refs

    def _dst_expr(self) -> str:
        r = self.dst.ref_abs()
        if self.dst_scale > 1:
            return f"{self.rounder}({r}/{self.dst_scale},0)"
        return r

    def compare_formula(self) -> str:
        s, d = self._src_expr(), self._dst_expr()
        return f"=IF(ABS({s})=ABS({d}),TRUE,ABS(ABS({s})-ABS({d})))"

    def display_label(self) -> str:
        return short_label(self.label_path).replace('"', "'")

    def hyperlink_formula(self) -> str:
        first = self.src[0]
        target = f"#'{first.sheet}'!{first.addr}"
        return f'=HYPERLINK("{target}","{self.display_label()}")'


@dataclass
class CheckSet:
    footing: list[FootingCheck] = field(default_factory=list)
    cross: list[CrossRef] = field(default_factory=list)

    def cells_by_sheet(self) -> dict[str, dict[str, list[str]]]:
        """총괄표 COUNTIF 체인용: 시트 → {footing:[주소…], cross:[주소…]}."""
        out: dict[str, dict[str, list[str]]] = {}
        for fc in self.footing:
            if fc.formula_cell is not None:
                out.setdefault(fc.formula_cell.sheet, {"footing": [], "cross": []})[
                    "footing"].append(fc.formula_cell.addr)
        for xr in self.cross:
            if xr.formula_cell is not None:
                out.setdefault(xr.formula_cell.sheet, {"footing": [], "cross": []})[
                    "cross"].append(xr.formula_cell.addr)
        return out


# 관계 종류(kind) → 도구가 가정한 관계의 뜻. 검증 칸 메모와 결과 요약에 쓴다.
KIND_DESC = {
    "group": "들여쓰기 계층상 소계 = 직계 자식 합",
    "rollup": "손익 계정 롤업(소계 = 세부 합)",
    "total_row": "합계류 라벨 행 = 위쪽 구성 행 합",
    "total_col": "합계 열 = 구성 열 가로합",
    "autosum": "원문 자동합 필드 = 지정 방향 합",
    "pl_anchor": "손익 단계손익(앵커 = 직전 앵커 + 구성)",
    "cf_anchor": "현금흐름 활동합계·순증감 관계",
    "hint": "라벨에 명시된 합산식((Ⅰ+Ⅱ+Ⅲ) 등)",
    "socie": "자본변동 기초+변동=기말",
    "socie_sub": "자본변동 섹션 소계",
    "flow_h": "가로 흐름(기말 = 기초 + 증감 열들)",
    "asset_cost": "장부금액 = 취득원가 + 차감 열(누계액·충당금·보조금)",
    "maturity": "만기분석 가로합(장부금액 = Σ만기구간)",
    "fv_equal": "장부가액 = 공정가치 등가",
    "summary_fin": "요약재무 자본 = 자산 − 부채",
    "net_pair": "이연법인세 순액 = 자산 계 − 부채 계",
    "cascade": "위 연속(또는 발견된 소계) 행들의 합",
    "combo": "두 행의 산술조합(차·곱·몫, 라벨 힌트 포함)",
    "ratio": "구성비율 = 행 금액 ÷ 합계행 금액",
    "eq_carry": "무변동 이월(기초 = 기말)",
}
