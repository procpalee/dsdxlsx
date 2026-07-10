"""dsdxlsx CLI: export / import / verify."""
from __future__ import annotations

import argparse
import os
import sys

from .dsd_io import read_dsd, write_dsd
from .export_xlsx import export_xlsx
from .import_xlsx import ImportError_, import_xlsx
from .parse import DsdParseError, parse_contents
from .verify import diff_dsd


def cmd_export(args) -> int:
    package = read_dsd(args.dsd)
    doc = parse_contents(package.contents)
    with open(args.dsd, "rb") as f:
        dsd_bytes = f.read()
    out = args.out or os.path.splitext(args.dsd)[0] + ".xlsx"
    stats = export_xlsx(doc, dsd_bytes, args.dsd, out)
    print(f"내보내기 완료: {out}")
    print(f"  시트 {stats['sheets']}개 / 표 {stats['tables']}개 / 문단 {stats['paras']}개")
    print("  A열(숨김)의 행 ID는 재변환 기준이므로 편집하지 마세요.")
    return 0


def cmd_import(args) -> int:
    package, new_contents, report = import_xlsx(args.xlsx)
    out = args.out or os.path.splitext(args.xlsx)[0] + "_수정.dsd"
    write_dsd(out, package, new_contents)
    print(f"재변환 완료: {out}")
    if not report.changed:
        print("  변경 없음 — contents.xml 원본 바이트 그대로 유지됨")
    else:
        if report.cell_edits:
            print(f"  셀 수정 {len(report.cell_edits)}건: " + ", ".join(report.cell_edits[:10])
                  + (" …" if len(report.cell_edits) > 10 else ""))
        if report.row_inserts:
            print(f"  행 추가 {len(report.row_inserts)}건:")
            for line in report.row_inserts:
                print(f"    - {line}")
            print("  [주의] 재무제표 본문 합계는 DART 편집기가 자동 재계산하지 않음 — "
                  "신설 행 금액을 상위 합계(유동자산 등)에 직접 반영했는지 확인하세요.")
        if report.row_deletes:
            print(f"  행 삭제 {len(report.row_deletes)}건:")
            for line in report.row_deletes:
                print(f"    - {line}")
    for w in report.warnings:
        print(f"  [경고] {w}")
    if report.autosum_tables:
        print(
            f"  [경고] 자동합계(AUTOSUM) 셀이 있는 표를 수정함: "
            f"{', '.join(sorted(report.autosum_tables))} — DART 편집기에서 합계를 재확인하세요."
        )
    print("  최종 확인: 결과 DSD를 DART 편집기로 열어 정상 표시되는지 확인하세요.")
    return 0


def cmd_verify(args) -> int:
    for line in diff_dsd(args.a, args.b):
        print(line)
    return 0


def cmd_recheck(args) -> int:
    """annotate된 xlsx의 살아있는 검산 수식을 Excel 없이(OfficeCLI) 재평가."""
    from .verify_annotated import officecli_available, verify_annotated

    if not officecli_available():
        print("officecli 미설치 — npm i -g @officecli/officecli", file=sys.stderr)
        return 2
    rep = verify_annotated(args.xlsx)
    print(rep)
    return 0 if rep.ok else 1


def cmd_check(args) -> int:
    from .annotate_xlsx import annotate
    from .check import CHECK_NAMES, run_checks
    from .import_xlsx import import_xlsx
    from .report_xlsx import write_report

    path = args.file
    is_xlsx = path.lower().endswith(".xlsx")
    imp = None
    if is_xlsx:
        # 엑셀의 현재 편집 상태를 반영해 검사 (임베드 원본이 아니라)
        _, new_contents, imp = import_xlsx(path)
        contents = new_contents
        if imp.changed:
            print(f"엑셀 편집 반영 상태로 검사: 셀 수정 {len(imp.cell_edits)}건, "
                  f"행 추가 {len(imp.row_inserts)}건, 행 삭제 {len(imp.row_deletes)}건")
    else:
        contents = read_dsd(path).contents
    doc = parse_contents(contents)
    only = set(x.strip().upper() for x in args.only.split(",")) if args.only else None
    report = run_checks(doc, only)

    if is_xlsx:
        # 검증 결과를 편집용 엑셀 안에 기록: 표 오른쪽 검증 열 + '검증요약' 시트
        try:
            stats = annotate(path, report, doc, imp.row_positions)
        except PermissionError:
            print(f"오류: {path} 이(가) 엑셀에서 열려 있어 기록할 수 없습니다. "
                  f"파일을 닫고 다시 실행하세요.", file=sys.stderr)
            return 1
        print(f"검증 결과 기록: {path}")
        print(f"  각 표 오른쪽 '검증' 열 {stats['cells']}개 셀 — 일치 TRUE / 불일치 차이 금액")
        print(f"  '검증요약' 시트: 검사별 집계 + 불일치 {stats['bad']}건 목록(셀 링크)")
    if args.out:
        stats = write_report(report, args.out, target_xlsx=path if is_xlsx else None)
        print(f"별도 검증 리포트 저장: {args.out} ({stats['results']}개 항목)")

    order = {"error": 0, "warn": 1, "info": 2}
    total_issues = 0
    for key, name in CHECK_NAMES.items():
        if only is not None and key not in only:
            continue
        findings = sorted(
            (f for f in report.findings if f.check == key), key=lambda f: order[f.level]
        )
        checked = report.counts.get(key, 0)
        mark = "이상 없음" if not findings else f"{len(findings)}건"
        print(f"\n[{key}] {name} — 검사 {checked}건, {mark}")
        for f in findings:
            print(f"  ({f.level}) {f.where}: {f.message}")
        total_issues += len(findings)
    print(f"\n총 {total_issues}건 검출. 모든 결과는 참고용 — 최종 확인은 DART 편집기에서.")
    return 0


def _force_utf8_io():
    """파이프/리다이렉트 시 cp949 대신 UTF-8 출력 (콘솔 직접 출력에는 영향 없음)."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure") and (stream.encoding or "").lower() not in ("utf-8", "utf8"):
            stream.reconfigure(encoding="utf-8")


def main(argv=None) -> int:
    _force_utf8_io()
    ap = argparse.ArgumentParser(
        prog="dsdxlsx",
        description="DART DSD ↔ XLSX 왕복 변환 (무수정 부분 바이트 보존)",
    )
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("export", help="DSD → XLSX 내보내기")
    p.add_argument("dsd")
    p.add_argument("-o", "--out")
    p.set_defaults(func=cmd_export)

    p = sub.add_parser("import", help="편집한 XLSX → DSD 재변환")
    p.add_argument("xlsx")
    p.add_argument("-o", "--out")
    p.set_defaults(func=cmd_import)

    p = sub.add_parser("verify", help="두 DSD 파일 구조 비교")
    p.add_argument("a")
    p.add_argument("b")
    p.set_defaults(func=cmd_verify)

    p = sub.add_parser("check", help="합계·대사·표기 정합성 검증 (.dsd 또는 export한 .xlsx)")
    p.add_argument("file")
    p.add_argument("--only", help="쉼표 구분 검사 선택 (예: C1,C4)")
    p.add_argument("-o", "--out", help="검증 결과 엑셀 경로 (기본: <대상>_검증.xlsx)")
    p.set_defaults(func=cmd_check)

    p = sub.add_parser(
        "recheck",
        help="annotate된 xlsx의 검산 수식을 Excel 없이 재평가 (헤드리스, OfficeCLI 필요)",
    )
    p.add_argument("xlsx")
    p.set_defaults(func=cmd_recheck)

    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except (ImportError_, DsdParseError, FileNotFoundError, ValueError) as e:
        print(f"오류: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
