"""콘솔 요약 — 방출된 모든 검증을 Python으로 평가해 결과를 출력한다."""
from __future__ import annotations

from .checks import CheckSet


def _num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    return int(f) if f == int(f) else f


def summary_dict(checkset: CheckSet) -> dict:
    """결과 봉투(--json)에 싣는 요약 - 건수와 불일치 위치."""
    foot = {"total": len(checkset.footing), "pass": 0, "rounding": 0, "fail": 0}
    cross = {"total": len(checkset.cross), "pass": 0, "fail": 0}
    mismatches = []
    for fc in checkset.footing:
        if fc.passed():
            foot["pass"] += 1
        elif fc.rounding_pass():
            foot["rounding"] += 1
        else:
            foot["fail"] += 1
            mismatches.append({"kind": "footing", "sheet": fc.formula_cell.sheet,
                               "cell": fc.formula_cell.addr, "label": fc.label, "diff": _num(fc.d())})
    for xr in checkset.cross:
        if xr.passed():
            cross["pass"] += 1
        else:
            cross["fail"] += 1
            mismatches.append({"kind": "cross", "sheet": xr.formula_cell.sheet,
                               "cell": xr.formula_cell.addr, "label": xr.label_path,
                               "diff": _num(xr.diff())})
    return {"footing": foot, "cross": cross, "mismatches": mismatches}


def print_summary(name: str, company: str, n_sheets: int, checkset: CheckSet) -> int:
    """파일 하나의 요약 출력. 반환: 0=전부 통과, 1=FAIL 존재."""
    foot_pass = foot_round = 0
    foot_fail = []
    for fc in checkset.footing:
        if fc.passed():
            foot_pass += 1
        elif fc.rounding_pass():
            foot_round += 1
        else:
            foot_fail.append(fc)
    cross_pass = 0
    cross_fail = []
    for xr in checkset.cross:
        if xr.passed():
            cross_pass += 1
        else:
            cross_fail.append(xr)

    print(f"\n[{name}] {company}")
    print(f"  시트 {n_sheets}개 · 풋팅 {len(checkset.footing)}건 · 대사 {len(checkset.cross)}건")
    line = f"  풋팅: 일치 {foot_pass}"
    if foot_round:
        line += f" · 절사차이(±1) {foot_round}"
    if foot_fail:
        line += f" · 불일치 {len(foot_fail)}"
    print(line)
    line = f"  대사: 일치 {cross_pass}"
    if cross_fail:
        line += f" · 불일치 {len(cross_fail)} (근사 표시 — 엑셀에서 빨간 차이 확인)"
    print(line)
    for fc in foot_fail[:30]:
        print(f"    [풋팅 불일치] {fc.formula_cell.sheet}!{fc.formula_cell.addr} "
              f"'{fc.label}' 차이 {fc.d():,}")
    if len(foot_fail) > 30:
        print(f"    … 외 {len(foot_fail) - 30}건")
    for xr in cross_fail[:20]:
        print(f"    [대사 차이] {xr.formula_cell.sheet}!{xr.formula_cell.addr} "
              f"{xr.label_path[:50]} 차이 {xr.diff():,}")
    if len(cross_fail) > 20:
        print(f"    … 외 {len(cross_fail) - 20}건")
    return 1 if foot_fail else 0
