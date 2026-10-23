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
        for label, items in (("표 신설", report.table_inserts),
                             ("표 삭제", report.table_deletes),
                             ("문단 신설", report.para_inserts),
                             ("문단 삭제", report.para_deletes),
                             ("주석 신설", report.note_inserts),
                             ("주석 삭제", report.note_deletes),
                             ("재번호 자동 치환", report.renumbered)):
            if items:
                print(f"  {label} {len(items)}건:")
                for line in items:
                    print(f"    - {line}")
        if report.structural:
            print("  [주의] 표/문단 구조가 바뀌었습니다 — 이 xlsx로 계속 편집하지 말고, "
                  "생성된 DSD를 다시 export해 새 편집본을 받으세요 (표 ID 재부여).")
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


def cmd_refoot(args) -> int:
    """상위 합계 자동 재계산 — 검증 통과한 전기 열의 구조를 당기 열에 적용.
    기본 미리보기, --apply 시 xlsx 셀에 기입 (DSD는 불변 — import 시 반영)."""
    from . import refoot as rf
    from .check import run_checks
    from .import_xlsx import import_xlsx
    from .profile import resolve_profile

    package, contents, imp = import_xlsx(args.xlsx)
    doc = parse_contents(contents)
    try:
        meta = package.meta
    except FileNotFoundError:
        meta = None
    profile = resolve_profile(meta, getattr(args, "profile", None))
    report = run_checks(doc, profile=profile)
    items, skips = rf.refoot_plan(doc, report, imp.row_positions, profile)

    for s in skips[:12]:
        print(f"[보류] {s.context} '{s.label}': {s.reason}")
    if len(skips) > 12:
        print(f"[보류] … 외 {len(skips) - 12}건")
    if not items:
        print("재계산 제안 없음 — 상위 합계가 전부 일치하거나(정상) 구성 셀이 미입력입니다.")
        return 0
    print(f"재계산 제안 {len(items)}건"
          + ("" if args.apply else " (미리보기 — 실제 기입은 --apply)") + ":")
    for it in items:
        print("  " + it.describe())
    if args.apply:
        try:
            n = rf.apply_refoot(args.xlsx, items)
        except PermissionError:
            print(f"오류: {args.xlsx} 이(가) 엑셀에서 열려 있어 기입할 수 없습니다.",
                  file=sys.stderr)
            return 1
        print(f"\n{n}건 기입 완료 — `dsdxlsx check` 재실행으로 검산하세요.")
    return 0


def cmd_rollforward(args) -> int:
    """전기 DSD → 당기 기초 DSD + 표준 정산표(xlsx, '롤포워드' 리포트 시트 포함)."""
    from . import rollforward as rf
    from . import splice
    from .export_xlsx import export_xlsx
    from .profile import resolve_profile

    package = read_dsd(args.dsd)
    doc = parse_contents(package.contents)
    try:
        meta = package.meta
    except FileNotFoundError:
        meta = None
    profile = resolve_profile(meta, getattr(args, "profile", None))

    ops, report = rf.roll_ops(doc, profile)
    rolled = splice.apply(package.contents, ops)
    parse_contents(rolled)  # well-formedness

    base = os.path.splitext(args.out or args.dsd)[0]
    if not args.out:
        base += "_차기기초"
    out_dsd, out_xlsx = base + ".dsd", base + ".xlsx"
    write_dsd(out_dsd, package, rolled)
    doc2 = parse_contents(rolled)
    with open(out_dsd, "rb") as f:
        export_xlsx(doc2, f.read(), out_dsd, out_xlsx)
    n_paint = 0
    if report.painted:
        from .import_xlsx import import_xlsx as _imp
        _, _, imp2 = _imp(out_xlsx)
        n_paint = rf.paint_moved(out_xlsx, doc2, report.painted,
                                 imp2.row_positions)
    rf.write_roll_sheet(out_xlsx, report)

    print(f"롤포워드 완료: {out_xlsx} (당기 기초 DSD: {out_dsd})")
    if n_paint:
        print(f"  자동 이동·채움 값 {n_paint}칸을 파란 배경으로 표시했습니다.")
    print(f"  자동확정 {report.count('자동확정')}건 / 확인필요 {report.count('확인필요')}건 / "
          f"미처리 {report.count('미처리')}건 — '롤포워드' 시트에서 위에서부터 처리")
    for item in report.items:
        if item.grade == "미처리":
            print(f"  [미처리] {item.where}: {item.item} — {item.note}")
    print("  당기 열은 비어 있습니다 — 결산 수치 입력 후 `dsdxlsx check`로 검산하세요.")
    return 0


def _load_doc(path: str):
    package = read_dsd(path)
    return package, parse_contents(package.contents)


def cmd_roll_interim(args) -> int:
    """분·반기 당전기 밀기 — 직전 중간보고서 DSD(골격) → 다음 분기 DSD + 정산표.
    종료코드: 0 완료 / 1 확인필요·미처리 있음 / 2 오류(연차 문서 등)."""
    from . import period as P
    from . import rollforward as rf
    from . import rollinterim as ri
    from . import splice
    from .profile import regname_of, resolve_profile

    package, doc = _load_doc(args.dsd)
    base = P.detect_period(doc)
    if base is None:
        print("오류: 골격 DSD 에서 보고기간(TU PERIODTO)을 읽지 못했습니다", file=sys.stderr)
        return 2
    if base.is_annual:
        print("오류: 감사보고서(연차) DSD 입니다 — 연차 밀기는 `dsdxlsx rollforward` 를 쓰세요",
              file=sys.stderr)
        return 2
    try:
        target = P.parse_target(getattr(args, "target", None), base)
    except ValueError as e:
        print(f"오류: {e}", file=sys.stderr)
        return 2
    try:
        meta = package.meta
    except FileNotFoundError:
        meta = None
    profile = resolve_profile(meta, getattr(args, "profile", None))
    company = regname_of(meta) if meta else ""

    doc_b = doc_c = None
    if getattr(args, "prior_same", None):
        _, doc_b = _load_doc(args.prior_same)
        pb = P.detect_period(doc_b)
        if pb and (pb.quarter != target.quarter or pb.year != target.year - 1):
            print(f"  [주의] 전년 동기 DSD 기간이 {pb.label} — 목표 {target.label}의 전년 동기가 아님",
                  file=sys.stderr)
    if getattr(args, "prior_annual", None):
        _, doc_c = _load_doc(args.prior_annual)
        pc = P.detect_period(doc_c)
        if pc and not (pc.is_annual and pc.year == target.year - 1):
            print(f"  [주의] 전기 감사보고서 DSD 기간이 {pc.label} — {target.year - 1}년 연차가 아님",
                  file=sys.stderr)

    inp = ri.RollInputs(doc=doc, base=base, target=target, doc_b=doc_b, doc_c=doc_c,
                        use_dart=not getattr(args, "no_dart", False), company=company)
    ops, report, dreport = ri.roll_interim_ops(inp, profile)
    rolled = splice.apply(package.contents, ops)
    parse_contents(rolled)  # well-formedness

    out_base = os.path.splitext(args.out or args.dsd)[0]
    if not args.out:
        out_base += "_당전기밀기"
    out_dsd, out_xlsx = out_base + ".dsd", out_base + ".xlsx"
    write_dsd(out_dsd, package, rolled)
    doc2 = parse_contents(rolled)
    with open(out_dsd, "rb") as f:
        export_xlsx(doc2, f.read(), out_dsd, out_xlsx)
    n_paint = 0
    if report.painted:
        from .import_xlsx import import_xlsx as _imp
        _, _, imp2 = _imp(out_xlsx)
        n_paint = rf.paint_moved(out_xlsx, doc2, report.painted, imp2.row_positions)
    rf.write_roll_sheet(out_xlsx, report)
    ri.write_dart_sheet(out_xlsx, dreport)

    print(f"당전기 밀기 완료: {out_xlsx} (다음 분기 DSD: {out_dsd})")
    print(f"  기간: {base.label} → {target.label}"
          + (" / 연결" if base.consolidated else " / 별도"))
    print(f"  원천: 전년 동기 DSD {'있음' if doc_b else '없음'} · 전기 감사보고서 DSD "
          f"{'있음' if doc_c else '없음'} · DART "
          f"{'생략(' + dreport.skipped_reason + ')' if dreport.skipped_reason else '조회'}")
    if n_paint:
        print(f"  자동 채움 값 {n_paint}칸을 파란 배경으로 표시했습니다.")
    print(f"  자동확정 {report.count('자동확정')}건 / 확인필요 {report.count('확인필요')}건 / "
          f"미처리 {report.count('미처리')}건 — '롤포워드' 시트에서 위에서부터 처리")
    if dreport.rows:
        print(f"  DART 대사: 일치 {dreport.count('일치')} · 불일치 {dreport.count('불일치')} · "
              f"미매칭 {dreport.count('미매칭')} · 채움 {dreport.count('채움')} — 'DART대사' 시트")
    for item in report.items:
        if item.grade == "미처리":
            print(f"  [미처리] {item.where}: {item.item} — {item.note}")
    print("  당기 열은 비어 있습니다 — 결산 수치 입력 후 `dsdxlsx check`로 검산하세요.")
    return 1 if (report.count("확인필요") or report.count("미처리")
                 or dreport.count("불일치")) else 0


def cmd_renumber(args) -> int:
    """주석 재번호 — 헤딩·FS 참조 열은 자동, 서술 멘션은 확인필요 목록.
    재번호 전후 C4 대사 결과 동형성 게이트를 통과해야 파일을 만든다."""
    from . import renumber as rn
    from . import splice
    from .profile import resolve_profile

    package = read_dsd(args.dsd)
    doc = parse_contents(package.contents)
    try:
        meta = package.meta
    except FileNotFoundError:
        meta = None
    profile = resolve_profile(meta, args.profile)

    nums = sorted(i.no for i in rn.note_infos(doc))
    if args.map:
        mapping = {}
        for pair in args.map.split(","):
            a, _, b = pair.partition(":")
            mapping[int(a.strip())] = int(b.strip())
    elif args.close_gap:
        mapping = rn.build_close_gap_mapping(nums)
    else:
        print("--map \"15:16,16:17\" 또는 --close-gap 중 하나를 지정하세요.",
              file=sys.stderr)
        return 2
    print(f"주석 현황: {nums}")
    if not mapping:
        print("변경 없음 — 매핑이 비어 있습니다 (번호 공백 없음).")
        return 0
    print(f"재번호 매핑: {dict(sorted(mapping.items()))}")

    sig_before = rn.c4_signature(doc, profile)
    ops, rep = rn.renumber_ops(doc, mapping, profile)  # RenumberError → main에서 출력
    new_contents = splice.apply(package.contents, ops)
    doc2 = parse_contents(new_contents)
    if rn.c4_signature(doc2, profile) != sig_before:
        print("오류: 재번호 후 재무제표↔주석 대사(C4) 결과가 달라짐 — "
              "동형성 게이트 실패, 파일을 만들지 않습니다.", file=sys.stderr)
        return 1
    if rep.mentions:
        print(f"\n확인 필요 — 서술문 멘션 {len(rep.mentions)}건 (자동 치환하지 않음):")
        for m in rep.mentions:
            print(f"  - {m}")
        if args.strict:
            print("--strict: 미확인 멘션이 있어 중단합니다.", file=sys.stderr)
            return 1
    out = args.out or os.path.splitext(args.dsd)[0] + "_재번호.dsd"
    write_dsd(out, package, new_contents)
    print(f"\n재번호 완료: {out}")
    print(f"  헤딩 치환 {len(rep.headings)}건: " + "; ".join(rep.headings[:6]))
    if rep.ref_cells:
        print(f"  재무제표 참조 열 치환 {len(rep.ref_cells)}건:")
        for line in rep.ref_cells[:8]:
            print(f"    - {line}")
    print("  C4 동형성 게이트 통과 ✓ — 최종 확인은 DART 편집기에서.")
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


def cmd_disc(args) -> int:
    """공시사항점검표(한공회 8400) 자동 작성."""
    import json
    from pathlib import Path

    from . import disclosure

    src = args.file
    if src.lower().endswith(".json"):
        review = json.loads(Path(src).read_text(encoding="utf-8"))
    else:
        from .check import run_checks
        from .jsonout import build_review_json
        from .profile import resolve_profile

        package = read_dsd(src)
        doc = parse_contents(package.contents)
        try:
            meta = package.meta
        except FileNotFoundError:
            meta = None
        profile = resolve_profile(meta, None)
        review = build_review_json(doc, run_checks(doc, {"C1"}, profile), profile=profile, source=src)
    out = Path(args.out) if args.out else Path(src).with_name(Path(src).stem + "_8400공시점검표.xlsx")
    listed = {"yes": True, "no": False}.get(args.listed)
    res = disclosure.run(review, out, form=args.form, kicpa_dir=args.kicpa_dir, listed=listed,
                         company=args.company, period_end=args.period_end)
    c = res["counts"]
    print(f"양식: {'K-IFRS' if res['form'] == 'kifrs' else '일반기업회계기준'} · "
          f"상장 {'예' if res['listed'] else '아니오'} · {'중간' if res['interim'] else '연차'}")
    print(f"목차 판정: 해당 {c['Y']} · 해당 없음 {c['NA']} · 확인 필요 {c['확인 필요']}")
    print(f"8400 본문: N/A {res['na_rows']}줄 · 주석번호 {res['y_rows']}줄")
    if res["warns"]:
        print(f"주석 누락 후보: {', '.join(res['warns'])}")
    if res["unmatched_notes"]:
        print("대응 없는 주석: " + ", ".join(f"{n}.{t}" for n, t in res["unmatched_notes"]))
    print(f"저장: {res['out']}")
    return 0


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


def cmd_recon(args) -> int:
    from . import recon as R

    path = args.file
    if path.lower().endswith(".xlsx"):
        from .import_xlsx import import_xlsx
        _, contents, _ = import_xlsx(path)        # 엑셀 편집 상태 그대로
    else:
        contents = read_dsd(path).contents
    doc = parse_contents(contents)
    split = (lambda s: [x.strip() for x in s.split(",") if x.strip()] if s else None)
    sheets = R.load_workpaper(args.workpaper, split(args.fs_sheets), split(args.note_sheets))
    if not sheets:
        print("오류: 조서에서 재무제표·주석 시트를 못 찾음 — --fs-sheets / --note-sheets 로 지정하세요",
              file=sys.stderr)
        return 1
    print("조서 시트: " + ", ".join(f"{s.name}({'재무제표' if s.kind == 'fs' else '주석'})"
                                 for s in sheets))
    rep = R.reconcile(doc, sheets, include_prior=args.prior)
    out = args.out or os.path.splitext(path)[0] + "_조서대사.xlsx"
    try:
        stats = R.write_recon_report(rep, out, path, args.workpaper)
    except PermissionError:
        print(f"오류: {out} 이(가) 엑셀에서 열려 있어 저장할 수 없습니다.", file=sys.stderr)
        return 1
    print(f"\n보고기간 {rep.period} — 대사 {stats['items']}칸 (0 칸 {rep.zero_skipped}개 생략)")
    for w in rep.warnings:
        print(f"  [주의] {w}")
    for sheet, cnt in R.summary_rows(rep):
        parts = [f"{v} {cnt[v]}" for v in R.VERDICTS if cnt[v]]
        print(f"  {sheet:<10} " + " · ".join(parts))
    issues = [x for x in rep.items if x.verdict in R.ISSUE_VERDICTS]
    if issues:
        print(f"\n확인 필요 {len(issues)}칸" + (" (앞 30칸)" if len(issues) > 30 else "") + ":")
        for x in issues[:30]:
            wpv = x.wp_in_unit
            wp_txt = "-" if wpv is None else f"{wpv:,.0f}" if wpv == int(wpv) else f"{wpv:,.2f}"
            dsd_txt = f"{x.dsd:,.0f}" if x.dsd == int(x.dsd) else f"{x.dsd:,.2f}"
            print(f"  ({x.verdict}) [{x.sheet}] {x.label[:24]} × {x.col[:16].replace(chr(10), ' ')}"
                  f" {x.period}: DSD {dsd_txt} / 조서 {wp_txt} {x.where}")
    print(f"\n결과 저장: {out}")
    return 0


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

    p = sub.add_parser("disc", help="공시사항점검표(한공회 8400) 자동 작성 — 해당여부·주석번호")
    p.add_argument("file", help=".dsd 또는 check --json 으로 만든 review.json")
    p.add_argument("-o", "--out", help="저장 경로(기본: <대상>_8400공시점검표.xlsx, 있으면 _2)")
    p.add_argument("--form", choices=["kifrs", "kgaap"], help="양식(기본: 본문에서 자동 판별)")
    p.add_argument("--listed", choices=["yes", "no"], help="상장 여부(기본: 본문에서 자동 판별)")
    p.add_argument("--company", help="회사명(기본: 본문에서 추출)")
    p.add_argument("--period-end", help="결산일 YYYY-MM-DD(기본: 본문에서 추출)")
    p.add_argument("--kicpa-dir", help="한공회 표준조서 폴더(기본: KICPA_FORMS 환경변수)")
    p.set_defaults(func=cmd_disc)

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
        "recon",
        help="DSD ↔ 조서(최종분석적검토 등) 대사 — 재무제표·주석 칸마다 조서에서 같은 값 찾기",
    )
    p.add_argument("file", help=".dsd 또는 export한 .xlsx")
    p.add_argument("workpaper", help="조서 엑셀 (계산값 저장된 파일)")
    p.add_argument("--fs-sheets", help="재무제표 시트 이름(쉼표) — 기본: BS·IS·CF·SCE·FS 등 자동")
    p.add_argument("--note-sheets", help="주석 시트 이름(쉼표) — 기본: '주석'·FN 자동")
    p.add_argument("--prior", action="store_true",
                   help="조서에 기간 열이 없어도 전기 칸까지 값 찾기 (기본: 당기만)")
    p.add_argument("-o", "--out", help="결과 엑셀 경로 (기본: <대상>_조서대사.xlsx)")
    p.set_defaults(func=cmd_recon)

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
        "refoot",
        help="상위 합계 자동 재계산 — 전기 열의 검증된 구조를 당기 열에 적용 (미리보기 기본)",
    )
    p.add_argument("xlsx", help="편집 중인 정산표 (export/rollforward 산출)")
    p.add_argument("--apply", action="store_true", help="제안을 실제 셀에 기입")
    p.add_argument("--profile", help="검증 프로파일 (기본: meta 회사명)")
    p.set_defaults(func=cmd_refoot)

    p = sub.add_parser(
        "rollforward",
        help="전기 DSD → 당기 기초 DSD + 표준 정산표 (기수·TU +1, 당기값→전기 열, 3등급 리포트)",
    )
    p.add_argument("dsd", help="전기(직전 보고서) DSD")
    p.add_argument("--profile", help="검증 프로파일 (기본: meta 회사명)")
    p.add_argument("-o", "--out", help="출력 기본 경로 (기본: <이름>_차기기초.*)")
    p.set_defaults(func=cmd_rollforward)

    p = sub.add_parser(
        "roll-interim",
        help="분·반기 당전기 밀기 — 직전 중간보고서 DSD(골격) → 다음 분기 DSD "
             "(당기 비움, 전기 ← 전년 동기 DSD/DART, 전기말 ← 전기 감사보고서 DSD/DART, 기간 어휘·날짜 치환)",
    )
    p.add_argument("dsd", help="골격: 직전 중간보고서 DSD (2Q 작업이면 1Q, 3Q면 2Q, 1Q면 전년 3Q)")
    p.add_argument("--prior-same", metavar="DSD", help="전년 동기 DSD — 흐름표·흐름형 주석의 전기 값 원천")
    p.add_argument("--prior-annual", metavar="DSD",
                   help="전기 감사보고서 DSD — 3Q→1Q 해 넘김 시 전기말(재무상태) 값 원천")
    p.add_argument("--target", help="목표 분기 1Q|2Q|3Q 또는 2026Q2 (기본: 골격 다음 분기)")
    p.add_argument("--no-dart", action="store_true", help="OpenDART 대사·채움 생략")
    p.add_argument("--profile", help="검증 프로파일 (기본: meta 회사명)")
    p.add_argument("-o", "--out", help="출력 기본 경로 (기본: <이름>_당전기밀기.*)")
    p.set_defaults(func=cmd_roll_interim)

    p = sub.add_parser(
        "renumber",
        help="주석 재번호 (헤딩·FS 참조 열 자동, 서술 멘션은 확인필요 목록, C4 동형성 게이트)",
    )
    p.add_argument("dsd")
    p.add_argument("--map", help='명시 매핑: "15:16,16:17"')
    p.add_argument("--close-gap", action="store_true",
                   help="번호 공백 자동 밀착 (1,2,3,5 → 5를 4로)")
    p.add_argument("--strict", action="store_true",
                   help="서술문 멘션(수동 확인 항목)이 있으면 중단")
    p.add_argument("--profile", help="검증 프로파일 (기본: meta 회사명)")
    p.add_argument("-o", "--out", help="출력 DSD 경로 (기본: <이름>_재번호.dsd)")
    p.set_defaults(func=cmd_renumber)

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
