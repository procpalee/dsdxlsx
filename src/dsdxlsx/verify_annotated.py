"""annotate 가 심은 살아있는 검산 수식을 Excel 없이 재평가 — 헤드리스 검증.

`check`(→ annotate)는 편집용 xlsx 에 검증을 **살아있는 엑셀 수식**으로 심는다:

    =IF(대상=(구성셀 합), TRUE, 대상-(구성셀 합))          # 합계·footing 검산
    =IF(OR(COUNTIF(주석범위, 값)>0, …), TRUE, "미확인")     # 재무제표↔주석 대사(C4)

이 수식들은 **엑셀로 열어야만** TRUE/차이로 평가된다(openpyxl 은 캐시가 없어 None).
무인 파이프라인·에이전트·CI 에서 "이 조서가 검산을 통과했나?"를 확인하려면 Excel 을
띄우지 않고 평가해야 하는데, OfficeCLI(자체 수식엔진, 헤드리스)가 그 역할을 한다.

식별: 검산 셀은 수식에 마커 ",TRUE," 를 포함한다(위 두 패턴 모두). openpyxl 로
수식을 스캔해 위치·라벨을 모으고, OfficeCLI 로 평가값을 읽어 판정한다.

판정(check.py `_pick` 의 ±1 절사 허용과 일치):
    TRUE(엔진은 1 로 렌더) 또는 |차이| ≤ 1  → 통과
    |차이| > 1                              → 불일치(차이 금액)
    "미확인"                                → 미확인(대사 근거 없음)

의존: officecli 바이너리만(npm i -g @officecli/officecli). 외부 라이브러리 미의존.

주의: C4 대사 수식은 크로스시트 COUNTIF 를 쓴다. OfficeCLI 의 크로스시트 함수 평가가
Excel 과 다를 수 있으므로(미지원 시 0 반환), 대사 '미확인' 은 Excel 재확인을 권한다.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
from dataclasses import dataclass, field

from openpyxl.utils import coordinate_to_tuple, get_column_letter

_VERIFY_MARK = ",TRUE,"   # 검산 수식(=IF(...,TRUE,...)) 식별 마커
_TOL = 1                  # |차이| ≤ 1 은 통과 (원단위 절사 — check.py 와 동일)
_ROW_ID_RE = re.compile(r"^(t\d{3}\.r\d{3}|![sv]\d+)$")  # A열 ID(라벨 아님)


# ── OfficeCLI 호출 ────────────────────────────────────────────────────

def _officecli_bin() -> str | None:
    exe = shutil.which("officecli")
    if exe:
        return exe
    cand = os.path.join(os.environ.get("APPDATA", ""), "npm", "officecli.cmd")
    return cand if os.path.exists(cand) else None


def officecli_available() -> bool:
    return _officecli_bin() is not None


def _ocli(*args) -> str:
    exe = _officecli_bin()
    if not exe:
        raise RuntimeError("officecli 를 찾을 수 없습니다 (npm i -g @officecli/officecli).")
    r = subprocess.run([exe, *args], capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return r.stdout or ""


def _walk_cells(node, acc: dict):
    if node.get("type") == "cell":
        acc[node["path"].split("/")[-1]] = node.get("text")
    for ch in node.get("children", []) or []:
        _walk_cells(ch, acc)


def _eval_range(xlsx_path: str, sheet: str, a1range: str) -> dict[str, str]:
    """OfficeCLI 로 시트 범위를 평가해 {A1: 평가텍스트} 반환 (원본 불변, 읽기만)."""
    out = _ocli("get", os.path.abspath(xlsx_path), f"/{sheet}/{a1range}", "--json")
    try:
        d = json.loads(out)
    except json.JSONDecodeError:
        return {}
    flat: dict[str, str] = {}
    for r in d.get("data", {}).get("results", []):
        _walk_cells(r, flat)
    return flat


# ── 검산 셀 수집 (openpyxl, 평가 안 함) ────────────────────────────────

@dataclass
class VCheck:
    sheet: str
    coord: str
    label: str


def _row_label(ws, row: int) -> str:
    """행의 왼쪽에서 첫 텍스트 라벨(계정명 등). A열 ID·수식·숫자는 건너뜀."""
    for col in range(2, min(ws.max_column, 14) + 1):
        v = ws.cell(row=row, column=col).value
        if isinstance(v, str) and v.strip() and not v.startswith("="):
            if not _ROW_ID_RE.match(v.strip()):
                return v.strip()[:40]
    return ""


def find_checks(xlsx_path: str) -> list[VCheck]:
    """편집용 xlsx 에서 검산 수식 셀(=IF(...,TRUE,...))의 위치·라벨 수집."""
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, data_only=False)
    checks: list[VCheck] = []
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if isinstance(v, str) and v.startswith("=") and _VERIFY_MARK in v:
                        checks.append(VCheck(ws.title, cell.coordinate,
                                             _row_label(ws, cell.row)))
    finally:
        wb.close()
    return checks


# ── 판정 ──────────────────────────────────────────────────────────────

@dataclass
class VResult:
    sheet: str
    coord: str
    label: str
    value: str
    status: str          # PASS | FAIL | UNCONFIRMED | UNKNOWN
    diff: float | None


def _classify(text: str) -> tuple[str, float | None]:
    t = (text or "").strip().replace(",", "")
    if t == "":
        return "UNKNOWN", None
    if t == "미확인":
        return "UNCONFIRMED", None
    if t.upper() == "TRUE":
        return "PASS", 0.0
    try:
        v = float(t)
    except ValueError:
        return "UNKNOWN", None
    return ("PASS", v) if abs(v) <= _TOL else ("FAIL", v)


@dataclass
class AnnotVerifyReport:
    results: list[VResult] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.fails and not self.unconfirmed and not self.unknown

    @property
    def fails(self):
        return [r for r in self.results if r.status == "FAIL"]

    @property
    def unconfirmed(self):
        return [r for r in self.results if r.status == "UNCONFIRMED"]

    @property
    def unknown(self):
        return [r for r in self.results if r.status == "UNKNOWN"]

    @property
    def passes(self):
        return [r for r in self.results if r.status == "PASS"]

    def __str__(self) -> str:
        n = len(self.results)
        head = (f"검산 셀 {n}개 — 통과 {len(self.passes)} / 불일치 {len(self.fails)} / "
                f"미확인 {len(self.unconfirmed)} / 불명 {len(self.unknown)}")
        lines = [head]
        for r in self.fails + self.unconfirmed + self.unknown:
            d = f"차이 {r.diff:,.0f}" if r.diff not in (None, 0.0) else r.value
            lines.append(f"  [{r.status}] {r.sheet}!{r.coord} '{r.label}': {d}")
        lines.append(f"종합: {'PASS' if self.ok else 'FAIL'}")
        return "\n".join(lines)


def verify_annotated(xlsx_path: str) -> AnnotVerifyReport:
    """annotate 된 xlsx 의 검산 수식을 OfficeCLI 로 평가해 PASS/FAIL 판정."""
    checks = find_checks(xlsx_path)
    report = AnnotVerifyReport()
    # 시트별로 묶어 bounding box 를 한 번에 평가 (프로세스 최소화)
    by_sheet: dict[str, list[VCheck]] = {}
    for c in checks:
        by_sheet.setdefault(c.sheet, []).append(c)
    for sheet, items in by_sheet.items():
        rows = [coordinate_to_tuple(c.coord)[0] for c in items]
        cols = [coordinate_to_tuple(c.coord)[1] for c in items]
        rng = (f"{get_column_letter(min(cols))}{min(rows)}:"
               f"{get_column_letter(max(cols))}{max(rows)}")
        vals = _eval_range(xlsx_path, sheet, rng)
        for c in items:
            status, diff = _classify(vals.get(c.coord, ""))
            report.results.append(
                VResult(c.sheet, c.coord, c.label, vals.get(c.coord, ""), status, diff))
    return report


def _cli(argv=None) -> int:
    import sys
    for s in (sys.stdout, sys.stderr):
        if hasattr(s, "reconfigure") and (s.encoding or "").lower() not in ("utf-8", "utf8"):
            s.reconfigure(encoding="utf-8")
    argv = argv if argv is not None else sys.argv[1:]
    if not argv:
        print("사용법: python -m dsdxlsx.verify_annotated <annotate된.xlsx>")
        return 2
    if not officecli_available():
        print("officecli 미설치 — npm i -g @officecli/officecli", file=sys.stderr)
        return 2
    rep = verify_annotated(argv[0])
    print(rep)
    return 0 if rep.ok else 1


if __name__ == "__main__":
    import sys
    sys.exit(_cli())
