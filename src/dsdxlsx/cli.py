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
    from .profile import resolve_profile
    from .report_xlsx import write_report

    path = args.file
    is_xlsx = path.lower().endswith(".xlsx")
    imp = None
    if is_xlsx:
        # 엑셀의 현재 편집 상태를 반영해 검사 (임베드 원본이 아니라)
        package, new_contents, imp = import_xlsx(path)
        contents = new_contents
        if imp.changed:
            print(f"엑셀 편집 반영 상태로 검사: 셀 수정 {len(imp.cell_edits)}건, "
                  f"행 추가 {len(imp.row_inserts)}건, 행 삭제 {len(imp.row_deletes)}건")
    else:
        package = read_dsd(path)
        contents = package.contents
    doc = parse_contents(contents)
    only = set(x.strip().upper() for x in args.only.split(",")) if args.only else None
    try:
        meta = package.meta
    except FileNotFoundError:
        meta = None
    profile = resolve_profile(meta, getattr(args, "profile", None))
    if profile.name != "_default":
        print(f"검증 프로파일: {profile.name}")
    report = run_checks(doc, only, profile)

    # 대사맵(C8): 저장된 대사쌍 + 시트 상태 편집 병합 → 확정 검증 + 제안 생성
    from . import ties as ties_mod
    from .profile import regname_of

    # --only에 C8이 빠지면 None 전달 — annotate가 대사맵 시트를 건드리지 않아
    # 사용자의 미저장 상태 편집이 보존된다
    tie_points = None
    display_points = None
    n_sugg = 0
    if only is None or "C8" in only:
        company = regname_of(meta)
        effective = ties_mod.load_ties(ties_mod.ties_path(company)) if company else []
        if is_xlsx:
            effective = ties_mod.merge_sheet_state(
                effective, ties_mod.read_tie_sheet(path))
        confirmed = [t for t in effective if t.status == "confirmed"]
        tie_points = ties_mod.verify_ties(doc, confirmed, profile, report)
        carried = [t for t in effective if t.status != "confirmed"]
        new_sugg = ties_mod.suggest_ties(doc, effective, profile)
        n_sugg = len([t for t in carried if t.status == "suggested"]) + len(new_sugg)
        display_points = ties_mod.resolve_points(doc, carried + new_sugg, profile)

    if args.json_out:
        import json as _json

        from .jsonout import build_review_json

        payload = build_review_json(doc, report, tie_points or [],
                                    display_points or [], profile, source=path)
        with open(args.json_out, "w", encoding="utf-8") as f:
            _json.dump(payload, f, ensure_ascii=False, indent=1, default=str)
        print(f"리뷰 JSON 저장: {args.json_out} — AI 검토(/dsd-review) 입력용")

    if is_xlsx:
        # 검증 결과를 편집용 엑셀 안에 기록: 표 오른쪽 검증 열 + '검증요약' 시트
        try:
            stats = annotate(path, report, doc, imp.row_positions,
                             tie_points=tie_points, tie_suggestions=display_points,
                             profile=profile)
        except PermissionError:
            print(f"오류: {path} 이(가) 엑셀에서 열려 있어 기록할 수 없습니다. "
                  f"파일을 닫고 다시 실행하세요.", file=sys.stderr)
            return 1
        print(f"검증 결과 기록: {path}")
        print(f"  각 표 오른쪽 '검증' 열 {stats['cells']}개 셀 — 일치 TRUE / 불일치 차이 금액")
        print(f"  '검증요약' 시트: 검사별 집계 + 불일치 {stats['bad']}건 목록(셀 링크)")
        if stats.get("ties"):
            print(f"  '대사맵' 시트: 확정 검증 {len(tie_points or [])}건 + 제안 {n_sugg}건 — "
                  f"상태를 '확정'으로 바꾸고 `dsdxlsx ties save`로 저장하세요")
    elif n_sugg:
        print(f"\n대사 제안 {n_sugg}건 — xlsx로 check하면 '대사맵' 시트에서 확정할 수 있습니다.")
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
    if report.suppressed:
        print(f"\n억제된 발견 {len(report.suppressed)}건 (프로파일 '{report.profile_name}' "
              f"suppress 규칙 — 검토 완료 예외):")
        for f, reason in report.suppressed:
            print(f"  [{f.check}] {f.where}: {f.message} — {reason}")
    if report.coverage:
        n_cov = sum(1 for c in report.coverage if c["checks"])
        n_skip = sum(1 for c in report.coverage if c["skipped"] and not c["checks"])
        uncov = [c for c in report.coverage if not c["skipped"] and not c["checks"]]
        line = (f"\n커버리지: 표 {len(report.coverage)}개 — 검산 적용 {n_cov} / "
                f"제외 구획 {n_skip} / 미적용 {len(uncov)}")
        if uncov:
            line += " (미적용 표는 검증요약 시트 목록 참조)" if is_xlsx else \
                    " — " + ", ".join(f"{c['context']}·{c['table']}" for c in uncov[:5])
        print(line)
    print(f"\n총 {total_issues}건 검출. 모든 결과는 참고용 — 최종 확인은 DART 편집기에서.")
    return 0


def cmd_annotate_ai(args) -> int:
    """AI 검토 소견(findings.json)을 편집용 xlsx의 'AI검토' 시트로 기록."""
    import json as _json

    from .annotate_xlsx import annotate_ai
    from .import_xlsx import import_xlsx

    with open(args.findings, encoding="utf-8-sig") as f:
        data = _json.load(f)
    findings = data.get("findings", [])
    if not isinstance(findings, list):
        print("findings.json 형식 오류 — {'findings': [...]} 필요", file=sys.stderr)
        return 1
    _, contents, imp = import_xlsx(args.xlsx)
    doc = parse_contents(contents)
    try:
        stats = annotate_ai(args.xlsx, findings, doc, imp.row_positions)
    except PermissionError:
        print(f"오류: {args.xlsx} 이(가) 엑셀에서 열려 있어 기록할 수 없습니다.",
              file=sys.stderr)
        return 1
    warns = sum(1 for x in findings if x.get("severity") == "warn")
    print(f"AI 검토 기록: {args.xlsx} — 'AI검토' 시트 {stats['findings']}건 "
          f"(warn {warns} / info {stats['findings'] - warns})")
    print("  전 항목 참고용 — 반영은 본문 셀 수정 후 check 재실행으로 검산하세요.")
    return 0


def cmd_calibrate(args) -> int:
    """새 회사 온보딩: 전기 심리완료 DSD가 전수 통과하는지 확인하고,
    잔여 지적은 suppress 규칙 초안(주석 처리)으로 스캐폴딩한다.

    통과(warn 0) → 이 파일을 회사 골든 샘플로 등록 권장, 종료코드 0.
    잔여 warn → suppress TOML 초안 출력(--write 시 profile.toml에 주석으로 추가,
    사람이 검토 후 주석 해제해야 활성화), 종료코드 1."""
    from datetime import date

    from .check import run_checks
    from .profile import PROFILES_DIR, company_dir, regname_of, resolve_profile

    package = read_dsd(args.dsd)
    doc = parse_contents(package.contents)
    try:
        meta = package.meta
    except FileNotFoundError:
        meta = None
    company = args.profile or regname_of(meta)
    profile = resolve_profile(meta, args.profile)
    report = run_checks(doc, profile=profile)

    warns = [f for f in report.findings if f.level in ("warn", "error")]
    infos = [f for f in report.findings if f.level == "info"]
    print(f"캘리브레이션: {args.dsd}")
    print(f"  프로파일 {report.profile_name} — warn {len(warns)}건 / info {len(infos)}건 / "
          f"억제 {len(report.suppressed)}건")
    if not warns:
        print("  ✅ 전수 통과 — 이 파일을 회사 골든 샘플로 등록하고 매 릴리스 회귀에 쓰세요.")
        return 0

    print(f"\n  ⚠ warn {len(warns)}건 — 실제 문서 이상인지 먼저 확인하세요. 검토 결과")
    print("    '승인된 예외'(전기 심리완료 문서의 알려진 특성)로 판단되면 아래 suppress")
    print("    초안의 주석(#)을 해제하고 사유를 채워 활성화합니다:\n")
    lines = [f"\n# ── calibrate 초안 {date.today()} — 검토 후 주석 해제 ──"]
    for f in warns:
        lines += [
            "# [[suppress]]",
            f'# check = "{f.check}"',
            f'# where_contains = "{f.where[:60]}"',
            '# reason = "TODO: 검토 사유 기입"',
            f"#   ↳ {f.message[:100]}",
        ]
    snippet = "\n".join(lines)
    print(snippet)
    if args.write:
        if not company:
            print("\n오류: 회사명을 알 수 없어 --write 불가 — --profile 회사명 지정",
                  file=sys.stderr)
            return 1
        path = company_dir(company) / "profile.toml"
        path.parent.mkdir(parents=True, exist_ok=True)
        header = "" if path.exists() else f'[meta]\nname = "{company}"\n'
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(header + snippet + "\n")
        print(f"\n  초안 추가됨: {path} (주석 상태 — 검토 후 해제해야 활성화)")
    else:
        print(f"\n  `--write`를 주면 {PROFILES_DIR}/<회사>/profile.toml에 주석으로 추가합니다.")
    return 1


def cmd_ties_save(args) -> int:
    """대사맵 시트의 상태 편집(확정/제외)을 profiles/<회사>/ties.json으로 저장."""
    from . import ties as ties_mod
    from .import_xlsx import import_xlsx
    from .profile import regname_of

    sheet = ties_mod.read_tie_sheet(args.xlsx)
    if not sheet:
        print("대사맵 시트가 없거나 비어 있음 — 먼저 `dsdxlsx check 파일.xlsx`를 실행하세요.",
              file=sys.stderr)
        return 1
    company = args.profile
    if not company:
        package, _, _ = import_xlsx(args.xlsx)
        try:
            company = regname_of(package.meta)
        except FileNotFoundError:
            company = None
    if not company:
        print("회사명을 알 수 없음 — `--profile 회사명`으로 지정하세요.", file=sys.stderr)
        return 1
    path = ties_mod.ties_path(company)
    saved = ties_mod.load_ties(path)
    merged = ties_mod.merge_sheet_state(saved, sheet)
    keep = ties_mod.persistable(merged)
    ties_mod.save_ties(path, keep, company)
    n_conf = sum(1 for t in keep if t.status == "confirmed")
    n_supp = len(keep) - n_conf
    print(f"대사맵 저장: {path}")
    print(f"  확정 {n_conf}건 / 제외 {n_supp}건 (제안은 저장하지 않음 — 매 check마다 재생성)")
    print("  ⚠ profiles/ 는 고객사명 포함 — 커밋 금지(.gitignore 적용됨)")
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
    p.add_argument("--profile",
                   help="검증 프로파일: 회사명 또는 profile.toml 경로 "
                        "(기본: meta.xml 회사명으로 profiles/ 자동 탐색)")
    p.add_argument("--json", dest="json_out",
                   help="리뷰용 구조화 JSON 저장 경로 — AI 검토(/dsd-review) 입력")
    p.add_argument("-o", "--out", help="검증 결과 엑셀 경로 (기본: <대상>_검증.xlsx)")
    p.set_defaults(func=cmd_check)

    p = sub.add_parser(
        "annotate-ai",
        help="AI 검토 소견(findings.json)을 편집용 xlsx 'AI검토' 시트로 기록",
    )
    p.add_argument("xlsx")
    p.add_argument("findings", help="/dsd-review 스킬이 만든 findings.json")
    p.set_defaults(func=cmd_annotate_ai)

    p = sub.add_parser(
        "recheck",
        help="annotate된 xlsx의 검산 수식을 Excel 없이 재평가 (헤드리스, OfficeCLI 필요)",
    )
    p.add_argument("xlsx")
    p.set_defaults(func=cmd_recheck)

    p = sub.add_parser(
        "calibrate",
        help="새 회사 온보딩: 전기 심리완료 DSD 전수 통과 확인 + suppress 초안 스캐폴딩",
    )
    p.add_argument("dsd", help="전기(심리 완료) DSD 파일")
    p.add_argument("--profile", help="회사명 (기본: meta.xml 회사명)")
    p.add_argument("--write", action="store_true",
                   help="suppress 초안을 profiles/<회사>/profile.toml에 주석으로 추가")
    p.set_defaults(func=cmd_calibrate)

    p = sub.add_parser("ties", help="대사맵(대사쌍) 관리")
    tsub = p.add_subparsers(dest="tcmd", required=True)
    ps = tsub.add_parser(
        "save", help="대사맵 시트의 확정/제외 상태를 profiles/<회사>/ties.json으로 저장"
    )
    ps.add_argument("xlsx")
    ps.add_argument("--profile", help="회사명 (기본: 임베드 원본 meta.xml의 회사명)")
    ps.set_defaults(func=cmd_ties_save)

    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except (ImportError_, DsdParseError, FileNotFoundError, ValueError) as e:
        print(f"오류: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
