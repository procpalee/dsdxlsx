"""dsdxlsx CLI: export / import / verify / diff …"""
from __future__ import annotations

import argparse
import contextlib
import io
import os
import sys

from . import guard
from .dsd_io import read_dsd, write_dsd
from .export_xlsx import export_xlsx
from .import_xlsx import ImportError_, import_xlsx
from .parse import DsdParseError, parse_contents
from .verify import diff_dsd

# 명령 → {인자 이름: 와야 하는 파일 형식} - 명령이 돌기 전에 guard.need 로 본다
INPUTS = {
    "export": {"dsd": "dsd"}, "import": {"xlsx": "xlsx"}, "inspect": {"dsd": "dsd"},
    "apply": {"dsd": "dsd"}, "diff": {"a": "dsd", "b": "dsd", "base": "dsd"},
    "verify": {"a": "dsd", "b": "dsd"}, "check": {"file": "either"},
    "recon": {"file": "either", "workpaper": "workbook"},
    "fill": {"file": "dsd", "workpaper": "workbook"},
    "build": {"file": "dsd", "workpaper": "workbook"},
    "annotate-ai": {"xlsx": "xlsx"}, "recheck": {"xlsx": "xlsx"}, "refoot": {"xlsx": "xlsx"},
    "rollforward": {"dsd": "dsd"},
    "roll-interim": {"dsd": "dsd", "prior_same": "dsd", "prior_annual": "dsd"},
    "renumber": {"dsd": "dsd"}, "calibrate": {"dsd": "dsd"}, "ties": {"xlsx": "xlsx"},
}
# 결과 봉투를 스스로 내는 명령(나머지는 main 이 감싼다)
SELF_EMIT = ("apply", "dart")


def _note(args, *, outputs=(), warnings=(), **extra) -> None:
    """봉투(--json)에 실을 내용을 모은다 - 사람용 출력과 따로."""
    res = args.__dict__.setdefault("_res", {"outputs": [], "warnings": [], "extra": {}})
    res["outputs"].extend(o for o in outputs if o)
    res["warnings"].extend(warnings)
    res["extra"].update(extra)


def _envelope(args) -> bool:
    """이 실행이 결과를 JSON 봉투로 내야 하는가."""
    if args.cmd in ("check", "recon"):
        return getattr(args, "json_out", None) == "-"
    if args.cmd == "inspect":
        return False
    return bool(getattr(args, "json", False))


def cmd_export(args) -> int:
    out = args.out or os.path.splitext(args.dsd)[0] + ".xlsx"
    guard.check_out([out], force=getattr(args, "force", False), inputs=[args.dsd])
    package = read_dsd(args.dsd)
    doc = parse_contents(package.contents)
    with open(args.dsd, "rb") as f:
        dsd_bytes = f.read()
    stats = export_xlsx(doc, dsd_bytes, args.dsd, out)
    guard.remember(out)
    _note(args, outputs=[out], sheets=stats["sheets"], tables=stats["tables"], paras=stats["paras"])
    print(f"내보내기 완료: {out}")
    print(f"  시트 {stats['sheets']}개 / 표 {stats['tables']}개 / 문단 {stats['paras']}개")
    print("  A열(숨김)의 행 ID는 재변환 기준이므로 편집하지 마세요.")
    return 0


def _import_changes(report) -> dict:
    out = {}
    for name in ("cell_edits", "row_inserts", "row_deletes", "table_inserts", "table_deletes",
                 "para_inserts", "para_deletes", "note_inserts", "note_deletes", "renumbered"):
        v = getattr(report, name)
        if v:
            out[name] = list(v)
    if report.autosum_tables:
        out["autosum_tables"] = sorted(report.autosum_tables)
    return out


def _import_foot(args) -> int:
    """풋팅 워크북(foot 으로 만든 엑셀)에서 고친 칸 값을 원본 DSD 에 반영 - 칸 값만, 행·표 구조는 안 됨.
    종료코드: 0 반영 / 1 반영했고 무시한 칸이 있음 / 2 오류(파일을 만들지 않음) / 3 멈춤."""
    from pathlib import Path

    from .foot import mapsheet, writeback

    try:
        out = args.out or str(writeback.default_output(args.xlsx))
        guard.check_out([out], force=getattr(args, "force", False), inputs=[args.xlsx])
        r = writeback.writeback(args.xlsx, out)
    except (mapsheet.MapError, writeback.WritebackError) as e:
        raise guard.InputError(str(e)) from None
    writeback.print_report(Path(args.xlsx), r)
    edits = [{"cell": e.addr, "node": e.node_id, "old": e.old, "new": e.new} for e in r.edits]
    if r.rc == 2:
        for msg in r.errors:
            print(f"오류: {msg}", file=sys.stderr)
        _note(args, kind="foot", edits=[])
        return 2
    guard.remember(out)
    _note(args, outputs=[out], warnings=r.warnings, kind="foot", source=r.source,
          changed=bool(edits), structural=False, edits=edits)
    print("  최종 확인: 결과 DSD를 DART 편집기로 열어 정상 표시되는지 확인하세요.")
    return r.rc


def cmd_import(args) -> int:
    from .foot import mapsheet as foot_map

    if foot_map.is_foot_workbook(args.xlsx):
        return _import_foot(args)
    out = args.out or os.path.splitext(args.xlsx)[0] + "_수정.dsd"
    guard.check_out([out], force=getattr(args, "force", False), inputs=[args.xlsx])
    package, new_contents, report = import_xlsx(args.xlsx)
    write_dsd(out, package, new_contents)
    guard.remember(out)
    _note(args, outputs=[out], warnings=report.warnings, changed=report.changed,
          structural=report.structural, changes=_import_changes(report))
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


def cmd_inspect(args) -> int:
    from . import docview
    package = read_dsd(args.dsd)
    doc = parse_contents(package.contents)
    only = [s.strip() for part in (args.section or []) for s in part.split(",") if s.strip()]
    name = os.path.basename(args.dsd)
    if only:
        have = {k for s in docview.sections(doc) for k in (s.key, s.sheet)}
        missing = [s for s in only if s not in have]
        if missing:
            keys = ", ".join(s.key for s in docview.sections(doc))
            print(f"오류: 구획 {', '.join(missing)} 없음 - 있는 구획: {keys}", file=sys.stderr)
            return 2
    if args.find:
        d = docview.find_dict(doc, args.find, only=only or None, source=name, limit=args.limit)
        text = docview.dumps(d, pretty=args.pretty) if args.json else docview.find_markdown(d)
    elif args.outline:
        d = docview.outline_dict(doc, only=only or None, source=name, data=package.contents)
        text = docview.dumps(d, pretty=args.pretty) if args.json else docview.outline_markdown(d)
    else:
        d = docview.doc_dict(doc, only=only or None, source=name, data=package.contents)
        text = docview.dumps(d, pretty=args.pretty) if args.json else docview.to_markdown(d)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"저장: {args.out}")
    else:
        sys.stdout.write(text + "\n")
    return 0


def cmd_apply(args) -> int:
    import json

    from . import result as R
    from .edits import EditError, apply_edits, normalize_edits

    def fail(code: int, msg: str) -> int:
        if args.json:
            return R.emit(R.envelope("apply", code, errors=[msg]))
        print(f"오류: {msg}", file=sys.stderr)
        return code

    try:
        raw = sys.stdin.read() if args.edits == "-" else open(args.edits, encoding="utf-8").read()
        edits = normalize_edits(json.loads(raw))
    except (OSError, json.JSONDecodeError, EditError) as e:
        return fail(R.INPUT_ERROR, f"편집 목록을 읽지 못했다 - {e}")
    out = args.out or os.path.splitext(args.dsd)[0] + "_수정.dsd"
    if not args.dry_run and os.path.exists(out) and not args.force:
        return fail(R.STOPPED, f"출력 파일이 이미 있다: {out} - 다른 -o 를 주거나 --force")
    if not args.dry_run and os.path.abspath(out) == os.path.abspath(args.dsd):
        return fail(R.STOPPED, "입력 DSD 를 그 자리에서 덮어쓰지 않는다 - 다른 -o 를 주라")
    package = read_dsd(args.dsd)
    try:
        res = apply_edits(package.contents, edits)
    except (EditError, DsdParseError) as e:
        return fail(R.INPUT_ERROR, str(e))
    warnings = list(res.report.warnings)
    if res.report.autosum_tables:
        warnings.append("자동합계 칸이 있는 표를 고침: " + ", ".join(sorted(res.report.autosum_tables))
                        + " - DART 편집기에서 합계를 다시 확인")
    outputs = []
    if not args.dry_run:
        write_dsd(out, package, res.contents)
        guard.remember(out)
        outputs.append(out)
    if args.json:
        return R.emit(R.envelope("apply", R.OK, outputs=outputs, warnings=warnings,
                                 dry_run=args.dry_run, changed=res.report.changed,
                                 changes=res.changes(), structural=res.report.structural))
    print(("[미리보기] " if args.dry_run else "") + (f"저장: {out}" if outputs else "쓰지 않음"))
    for k, v in res.changes().items():
        print(f"  {k}: {len(v)}건 - " + ", ".join(map(str, v[:8])) + (" …" if len(v) > 8 else ""))
    if not res.report.changed:
        print("  바뀐 것 없음 - 원본과 같은 바이트")
    for w in warnings:
        print(f"  [경고] {w}")
    if res.report.structural:
        print("  [주의] 표·문단 구조가 바뀌어 ID 가 다시 매겨진다 - 다음 편집 전에 inspect 를 다시 보라")
    return R.OK


def cmd_dart(args) -> int:
    """DART 에 공시된 감사보고서 → DSD (+ 엑셀, 롤포워드)."""
    import re

    from . import dartdoc as D
    from . import result as R

    def fail(code: int, msg: str) -> int:
        if args.json:
            return R.emit(R.envelope("dart", code, errors=[msg]))
        print(f"오류: {msg}", file=sys.stderr)
        return code

    try:
        if args.list:
            corp = D.find_corp(args.query)
            fl = D.filings(corp, args.form, args.year)
            if args.json:
                return R.emit(R.envelope("dart", R.OK, corp={"code": corp.code, "name": corp.name,
                                                            "stock": corp.stock},
                                         filings=[f.to_dict() for f in fl]))
            print(f"{corp.name} ({corp.stock or '비상장'}, 고유번호 {corp.code}) - "
                  f"{D.ACODE_NAMES[D.FORMS[args.form]]}이 들어 있을 공시 {len(fl)}건")
            for f in fl:
                where = "외부감사관련" if f.kind == "F" else "사업보고서 첨부"
                print(f"  {f.rcept_no}  {f.rcept_dt}  {f.report_nm}  [{where}]")
            return R.OK
        got = D.fetch(args.query, args.form, args.year, crlf=not args.keep_lf)
    except D.DartError as e:
        return fail(R.INPUT_ERROR, str(e))

    b = got.built
    try:
        parse_contents(b.contents)
    except DsdParseError as e:
        return fail(R.INPUT_ERROR, f"공시 원문을 DSD 본문으로 읽지 못했다 - {e}")
    m = re.search(rb'AUNIT="PERIODTO"[^>]*AUNITVALUE="(\d{4})', b.contents)
    year = got.filing.year if got.filing and got.filing.year else (int(m.group(1)) if m else None)
    stem = re.sub(r'[\\/:*?"<>|]', "_", f"{b.company or b.corp_code}_{D.ACODE_NAMES[b.acode]}"
                  + (f"_{year}" if year else ""))
    if args.out and os.path.isdir(args.out):
        out = os.path.join(args.out, stem + ".dsd")
    else:
        out = args.out or stem + ".dsd"
    targets = [out] + ([os.path.splitext(out)[0] + ".xlsx"] if args.xlsx else [])
    if args.rollforward:
        rb = os.path.splitext(out)[0] + "_차기기초"
        targets += [rb + ".dsd", rb + ".xlsx"]
    if not args.force:
        exists = [t for t in targets if os.path.exists(t)]
        if exists:
            return fail(R.STOPPED, "이미 있는 파일: " + ", ".join(exists) + " - 다른 -o 를 주거나 --force")

    D.write_dsd(out, b)
    outputs = [out]
    warnings = []
    if b.images:
        warnings.append(f"이미지 {len(b.images)}개는 공시 원문에 파일이 없어 자리만 남았다 - "
                        "DART 편집기에서 다시 넣으라: " + ", ".join(b.images[:5])
                        + (" …" if len(b.images) > 5 else ""))
    if b.fixes.get("escaped"):
        warnings.append(f"원문의 이스케이프 안 된 '&'·'<' {b.fixes['escaped']}곳을 고쳤다")
    if args.xlsx:
        from .export_xlsx import export_xlsx
        xo = os.path.splitext(out)[0] + ".xlsx"
        with open(out, "rb") as f:
            export_xlsx(parse_contents(b.contents), f.read(), out, xo)
        outputs.append(xo)
    roll = None
    if args.rollforward:
        rd, rx, rep, n_paint = run_rollforward(out)
        outputs += [rd, rx]
        roll = {"자동확정": rep.count("자동확정"), "확인필요": rep.count("확인필요"),
                "미처리": rep.count("미처리"), "painted": n_paint}
    guard.remember(*outputs)
    src = got.filing.to_dict() if got.filing else {"rcept_no": None}
    src.update({"file": got.part.filename, "acode": b.acode, "document": got.part.name})
    if args.json:
        return R.emit(R.envelope("dart", R.OK, outputs=outputs, warnings=warnings,
                                 company=b.company, corp_code=b.corp_code, year=year,
                                 source=src, fixes=b.fixes, images=b.images, rollforward=roll))
    print(f"DSD 저장: {out}")
    if got.filing:
        print(f"  원천: {got.filing.report_nm} 접수번호 {got.filing.rcept_no} ({got.filing.rcept_dt}) "
              f"- {got.part.name}")
    for o in outputs[1:]:
        print(f"  저장: {o}")
    if b.fixes.get("crlf"):
        print("  줄끝을 DART 편집기 형식(CRLF)으로 맞췄다(--keep-lf 로 원문 그대로)")
    if roll:
        print(f"  롤포워드: 자동확정 {roll['자동확정']}건 / 확인필요 {roll['확인필요']}건 / "
              f"미처리 {roll['미처리']}건 - '롤포워드' 시트에서 처리")
    for w in warnings:
        print(f"  [경고] {w}")
    print("  최종 확인: 만든 DSD 를 DART 편집기로 열어 보라")
    return R.OK


def cmd_foot(args) -> int:
    """풋팅·대사 검증 워크북 - DSD → 합계·대사 수식이 든 엑셀(총괄표 / 원문 / 재무제표 / 주석).
    종료코드: 0 전부 일치 / 1 풋팅 불일치 있음 / 2 오류 / 3 멈춤(여러 파일이면 가장 나쁜 것)."""
    from pathlib import Path

    from .foot import pipeline, report
    from .foot import sections as fsec

    if args.out and len(args.files) > 1:
        print("오류: -o 는 입력 파일이 하나일 때만 쓸 수 있다", file=sys.stderr)
        return 2
    worst = 0
    results, outputs, warnings = [], [], []
    for name in args.files:
        path = Path(name)
        item = {"file": name}
        try:
            guard.need(name, "dsd")
            if args.sections:
                package = read_dsd(name)
                plans = fsec.build(parse_contents(package.contents))
                print(f"\n[{path.name}] {fsec.company_name(package.meta)}")
                for pl in plans:
                    n_tables = sum(1 for b in pl.sec.blocks if b.kind == "table")
                    print(f"  {pl.name:>4s} | {pl.kind:9s} | 표 {n_tables:2d} | {pl.display_title}")
                item["sheets"] = [{"name": pl.name, "kind": pl.kind, "title": pl.display_title}
                                  for pl in plans]
                rc = 0
            else:
                out = Path(args.out) if args.out else pipeline.default_output(path, args.convert_only)
                guard.check_out([str(out)], force=args.force, inputs=[name])
                res = pipeline.process_file_full(path, out, convert_only=args.convert_only,
                                                 relaxed=args.detection_profile == "exhaustive")
                outputs.append(str(out))
                item.update(output=str(out), company=res.company, sheets=res.n_sheets)
                if not args.convert_only:
                    item.update(report.summary_dict(res.checkset))
                rc = res.rc
                if args.dart_tie and not args.convert_only:
                    rc = max(rc, _foot_dart_tie(args, res, item, warnings))
                guard.remember(str(out))
        except guard.Stopped as e:
            print(f"멈춤: {e}", file=sys.stderr)
            item["error"], rc = str(e), 3
        except Exception as e:  # noqa: BLE001 - 파일 하나의 오류가 나머지 파일을 막지 않게
            msg = _explain(e) if isinstance(e, (guard.InputError, OSError, ValueError)) \
                else f"{type(e).__name__}: {e}"
            print(f"오류: {path.name}: {msg}", file=sys.stderr)
            item["error"], rc = msg, 2
        item["exit_code"] = rc
        results.append(item)
        worst = max(worst, rc)
    _note(args, outputs=outputs, warnings=warnings, files=results)
    return worst


def _foot_dart_tie(args, res, item: dict, warnings: list) -> int:
    """전기 열을 DART 에 공시된 직전 보고서 원문과 대사해 워크북에 덧붙인다. 불일치가 있으면 1.
    키가 없거나 공시가 없으면 건너뛰고 사유만 남긴다(0)."""
    from openpyxl.utils import get_column_letter

    from .foot import darttie

    dres = darttie.run(res.package, res.doc, res.plans, res.ctx, res.placed_rows, fs_div=args.dart_fs)
    darttie.apply(res.out, dres, res.right_starts or {})
    darttie.print_summary(dres)
    item["dart"] = {"status": dres.status, "reason": dres.reason, "total": len(dres.ties),
                    "pass": dres.n_ok, "fail": dres.n_bad, "sources": list(dres.sources),
                    "unmatched": list(dres.unmatched)}
    item.setdefault("mismatches", []).extend(
        {"kind": "dart", "sheet": t.sheet, "cell": f"{get_column_letter(t.col)}{t.row}",
         "label": t.label, "header": t.header, "current": _num(t.cur_value),
         "dart": _num(t.prior_value), "diff": _num(t.diff)} for t in dres.ties if not t.ok)
    if dres.status == "skipped":
        warnings.append(f"{os.path.basename(item['file'])}: DART 대사를 건너뜀 - {dres.reason}")
    return 1 if dres.n_bad else 0


def cmd_verify(args) -> int:
    lines = diff_dsd(args.a, args.b)
    diffs = [ln for ln in lines if ln != "구조·텍스트 차이 없음"]
    _note(args, same=not diffs, differences=diffs)
    for line in lines:
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

    if args.json_out and args.json_out != "-":
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
            return 2
        _note(args, outputs=[path])
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
        _note(args, outputs=[args.out])
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
    shown = [f for f in report.findings if only is None or f.check in only]
    n_bad = sum(1 for f in shown if f.level in ("warn", "error"))
    _note(args, counts={k: report.counts.get(k, 0) for k in CHECK_NAMES if only is None or k in only},
          findings=[{"check": f.check, "level": f.level, "where": f.where, "message": f.message}
                    for f in shown],
          mismatches=[{"check": r.check, "section": r.context, "table": r.table_id, "row": r.row_id,
                       "label": r.row_label, "col": r.grid_col, "header": r.col_header,
                       "stated": _num(r.stated), "computed": _num(r.computed), "note": r.note}
                      for r in report.results if not r.ok and (only is None or r.check in only)],
          suppressed=len(report.suppressed), tie_suggestions=n_sugg, log=False,
          coverage={"tables": len(report.coverage),
                    "checked": sum(1 for c in report.coverage if c["checks"]),
                    "unchecked": [c["table"] for c in report.coverage
                                  if not c["skipped"] and not c["checks"]]})
    return 1 if (_envelope(args) and n_bad) else 0


def _num(v):
    """Decimal 등 → JSON 숫자."""
    if v is None or isinstance(v, (int, float, str, bool)):
        return v
    try:
        f = float(v)
        return int(f) if f == int(f) else f
    except (TypeError, ValueError):
        return str(v)


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
        return 2
    _, contents, imp = import_xlsx(args.xlsx)
    doc = parse_contents(contents)
    try:
        stats = annotate_ai(args.xlsx, findings, doc, imp.row_positions)
    except PermissionError:
        print(f"오류: {args.xlsx} 이(가) 엑셀에서 열려 있어 기록할 수 없습니다.",
              file=sys.stderr)
        return 2
    _note(args, outputs=[args.xlsx], findings=stats["findings"])
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

    _note(args, applied=bool(args.apply and items),
          items=[{"section": it.context, "label": it.label, "sheet": it.sheet, "xlsx_row": it.xlsx_row,
                  "old": _num(it.old), "new": _num(it.new), "check": it.check} for it in items],
          skipped=[{"section": s.context, "label": s.label, "reason": s.reason} for s in skips])
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
            return 2
        _note(args, outputs=[args.xlsx])
        print(f"\n{n}건 기입 완료 — `dsdxlsx check` 재실행으로 검산하세요.")
    return 0


def _roll_paths(dsd_path: str, out: str | None, suffix: str) -> tuple[str, str]:
    base = os.path.splitext(out or dsd_path)[0]
    if not out:
        base += suffix
    return base + ".dsd", base + ".xlsx"


def run_rollforward(dsd_path: str, out: str | None = None, profile_name: str | None = None):
    """전기 DSD → 당기 기초 DSD + 표준 정산표. (out_dsd, out_xlsx, report, n_paint)."""
    from . import rollforward as rf
    from . import splice
    from .export_xlsx import export_xlsx
    from .profile import resolve_profile

    package = read_dsd(dsd_path)
    doc = parse_contents(package.contents)
    try:
        meta = package.meta
    except FileNotFoundError:
        meta = None
    profile = resolve_profile(meta, profile_name)

    ops, report = rf.roll_ops(doc, profile)
    rolled = splice.apply(package.contents, ops)
    parse_contents(rolled)  # well-formedness

    out_dsd, out_xlsx = _roll_paths(dsd_path, out, "_차기기초")
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
    guard.remember(out_dsd, out_xlsx)
    return out_dsd, out_xlsx, report, n_paint


def _roll_counts(report) -> dict:
    return {g: report.count(g) for g in ("자동확정", "확인필요", "미처리")}


def _roll_todo(report) -> list:
    return [{"grade": i.grade, "where": i.where, "item": i.item, "note": i.note}
            for i in report.items if i.grade != "자동확정"]


def cmd_rollforward(args) -> int:
    """전기 DSD → 당기 기초 DSD + 표준 정산표(xlsx, '롤포워드' 리포트 시트 포함)."""
    guard.check_out(_roll_paths(args.dsd, args.out, "_차기기초"), force=getattr(args, "force", False),
                    inputs=[args.dsd])
    out_dsd, out_xlsx, report, n_paint = run_rollforward(
        args.dsd, args.out, getattr(args, "profile", None))
    _note(args, outputs=[out_dsd, out_xlsx], counts=_roll_counts(report), painted=n_paint,
          todo=_roll_todo(report))

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

    out_dsd, out_xlsx = _roll_paths(args.dsd, args.out, "_당전기밀기")
    guard.check_out([out_dsd, out_xlsx], force=getattr(args, "force", False),
                    inputs=[args.dsd, getattr(args, "prior_same", None),
                            getattr(args, "prior_annual", None)])
    inp = ri.RollInputs(doc=doc, base=base, target=target, doc_b=doc_b, doc_c=doc_c,
                        use_dart=not getattr(args, "no_dart", False), company=company)
    ops, report, dreport = ri.roll_interim_ops(inp, profile)
    rolled = splice.apply(package.contents, ops)
    parse_contents(rolled)  # well-formedness

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
    guard.remember(out_dsd, out_xlsx)
    _note(args, outputs=[out_dsd, out_xlsx], counts=_roll_counts(report), painted=n_paint,
          period={"from": base.label, "to": target.label,
                  "side": "연결" if base.consolidated else "별도"},
          dart=({"skipped": dreport.skipped_reason} if dreport.skipped_reason else
                {v: dreport.count(v) for v in ("일치", "불일치", "미매칭", "채움")}),
          todo=_roll_todo(report))

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
    out = args.out or os.path.splitext(args.dsd)[0] + "_재번호.dsd"
    guard.check_out([out], force=getattr(args, "force", False), inputs=[args.dsd])
    _note(args, notes=nums, mapping={str(k): v for k, v in sorted(mapping.items())})
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
        return 2
    if rep.mentions:
        print(f"\n확인 필요 — 서술문 멘션 {len(rep.mentions)}건 (자동 치환하지 않음):")
        for m in rep.mentions:
            print(f"  - {m}")
        if args.strict:
            print("--strict: 미확인 멘션이 있어 중단합니다.", file=sys.stderr)
            _note(args, mentions=list(rep.mentions))
            return 3
    write_dsd(out, package, new_contents)
    guard.remember(out)
    _note(args, outputs=[out], mentions=list(rep.mentions), headings=list(rep.headings),
          ref_cells=list(rep.ref_cells))
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
            return 2
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
    _note(args, outputs=[str(res["out"])], form=res["form"], listed=bool(res["listed"]),
          interim=bool(res["interim"]), counts=dict(c), warnings=list(res["warns"]))
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
        return 2
    company = args.profile
    if not company:
        package, _, _ = import_xlsx(args.xlsx)
        try:
            company = regname_of(package.meta)
        except FileNotFoundError:
            company = None
    if not company:
        print("회사명을 알 수 없음 — `--profile 회사명`으로 지정하세요.", file=sys.stderr)
        return 2
    path = ties_mod.ties_path(company)
    saved = ties_mod.load_ties(path)
    merged = ties_mod.merge_sheet_state(saved, sheet)
    keep = ties_mod.persistable(merged)
    ties_mod.save_ties(path, keep, company)
    n_conf = sum(1 for t in keep if t.status == "confirmed")
    n_supp = len(keep) - n_conf
    _note(args, outputs=[str(path)], confirmed=n_conf, excluded=n_supp)
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
    import json

    from . import period as P
    from . import recon as R

    path = args.file
    if path.lower().endswith(".xlsx"):
        from .import_xlsx import import_xlsx
        _, contents, _ = import_xlsx(path)        # 엑셀 편집 상태 그대로
    else:
        contents = read_dsd(path).contents
    doc = parse_contents(contents)
    per = P.detect_period(doc)
    side = "연결" if per and per.consolidated else "별도"
    # 건별 설정(_engagement/dsd_recon.json) — {'별도': {...}, '연결': {...}}. 없으면 자동 판정 결과를 써 줌
    conf_all: dict = {}
    if args.config and os.path.exists(args.config):
        with open(args.config, encoding="utf-8") as f:
            conf_all = json.load(f)
    conf = conf_all.get(side) or {}
    split = (lambda s: [x.strip() for x in s.split(",") if x.strip()] if s else None)
    fs_map = None if args.fs_sheets or args.note_sheets else (conf.get("재무제표") or None)
    notes = split(args.note_sheets) or (None if args.fs_sheets else conf.get("주석"))
    sheets = R.load_workpaper(args.workpaper, split(args.fs_sheets), notes,
                              skip=conf.get("빼는 시트"), fs_map=fs_map)
    if not sheets:
        print("오류: 조서에서 재무제표·주석 시트를 못 찾음 — --fs-sheets / --note-sheets 로 지정하세요",
              file=sys.stderr)
        return 2
    wrote = False
    if args.config and not conf:
        auto = R.sheet_config(sheets)
        conf_all[side] = auto
        os.makedirs(os.path.dirname(os.path.abspath(args.config)), exist_ok=True)
        with open(args.config, "w", encoding="utf-8") as f:
            json.dump(conf_all, f, ensure_ascii=False, indent=2)
        wrote = True
    print("조서 시트: " + ", ".join(
        f"{s.name}({(s.fs_kind or '재무제표') if s.kind == 'fs' else '주석'})" for s in sheets))
    rep = R.reconcile(doc, sheets, include_prior=args.prior, config=R.ReconConfig.from_dict(conf))
    out = args.out or os.path.splitext(path)[0] + "_조서대사.xlsx"
    try:
        stats = R.write_recon_report(rep, out, path, args.workpaper)
    except PermissionError:
        print(f"오류: {out} 이(가) 엑셀에서 열려 있어 저장할 수 없습니다.", file=sys.stderr)
        return 2
    print(f"\n보고기간 {rep.period}({side}) — 대사 {stats['items']}칸 (0 칸 {rep.zero_skipped}개 생략)")
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
    if wrote:
        print(f"\n건별 설정 저장({side}): {args.config} - 시트 지정 · 빼는 칸 · 이름 다른 행을 고쳐 쓰세요")
    print(f"\n결과 저장: {out}")
    summary = {"period": rep.period, "side": side, "items": stats["items"], "issues": stats["issues"],
               "counts": {v: rep.count(v) for v in R.VERDICTS}, "config_written": wrote,
               "sheets": [s.name for s in sheets]}
    _note(args, outputs=[out] + ([args.config] if wrote else []), warnings=rep.warnings,
          to_check=[{"verdict": x.verdict, "sheet": x.sheet, "label": x.label, "col": x.col,
                     "period": x.period, "dsd": _num(x.dsd), "workpaper": _num(x.wp_in_unit),
                     "where": x.where} for x in issues[:200]], **summary)
    if args.json_out and args.json_out != "-":
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump({"out": os.path.abspath(out), "period": rep.period, "side": side,
                       "items": stats["items"], "issues": stats["issues"],
                       "counts": {v: rep.count(v) for v in R.VERDICTS},
                       "warnings": rep.warnings, "config_written": wrote,
                       "sheets": [s.name for s in sheets]}, f, ensure_ascii=False, indent=2)
    return 1 if (_envelope(args) and issues) else 0


def _recon_inputs(args, doc):
    """recon · fill 공통: 건별 설정(구역 '별도' · '연결')과 조서 시트."""
    import json

    from . import period as P
    from . import recon as R

    per = P.detect_period(doc)
    side = "연결" if per and per.consolidated else "별도"
    conf_all: dict = {}
    if getattr(args, "config", None) and os.path.exists(args.config):
        with open(args.config, encoding="utf-8") as f:
            conf_all = json.load(f)
    conf = conf_all.get(side) or {}
    sheets = R.load_workpaper(args.workpaper, None, conf.get("주석"), skip=conf.get("빼는 시트"),
                              fs_map=conf.get("재무제표") or None)
    return side, conf, sheets


def cmd_fill(args) -> int:
    from . import fill as F
    from . import splice
    from . import recon as R

    package = read_dsd(args.file)
    doc = parse_contents(package.contents)
    side, conf, sheets = _recon_inputs(args, doc)
    if not sheets:
        print("오류: 조서에서 재무제표 · 주석 시트를 못 찾음 - recon --config 설정을 확인하세요", file=sys.stderr)
        return 2
    out = args.out or os.path.splitext(args.file)[0] + "_채움.dsd"
    guard.check_out([out], force=getattr(args, "force", False), inputs=[args.file])
    ops, rep = F.fill_ops(doc, sheets, R.ReconConfig.from_dict(conf), keep_filled=args.keep_filled)
    filled = splice.apply(package.contents, ops)
    parse_contents(filled)                       # well-formedness
    write_dsd(out, package, filled)
    guard.remember(out)
    report = args.report or os.path.splitext(out)[0] + "_채움결과.xlsx"
    try:
        F.write_fill_report(rep, report, args.file, args.workpaper)
    except PermissionError:
        report = "(결과 엑셀이 열려 있어 저장 못 함)"
    unfilled = [x for x in rep.items if x.verdict == "못 채움"]
    _note(args, outputs=[out] + ([report] if os.path.exists(report) else []), warnings=rep.warnings,
          period=rep.period, side=side, counts={"채움": rep.count("채움"), "못 채움": len(unfilled)},
          unfilled=[{"sheet": x.sheet, "label": x.label, "col": x.col, "how": x.how}
                    for x in unfilled[:200]])
    print(f"보고기간 {rep.period}({side}) - 채움 {rep.count('채움')}칸 · 못 채움 {rep.count('못 채움')}칸")
    for w in rep.warnings:
        print(f"  [주의] {w}")
    for x in [x for x in rep.items if x.verdict == "못 채움"][:30]:
        print(f"  (못 채움) [{x.sheet}] {x.label[:24]} × {x.col[:16]}: {x.how}")
    print(f"\nDSD 저장: {out}\n결과: {report}")
    print("다음: dsdxlsx recon 으로 대사 · dsdxlsx check 로 합계 검산")
    return 1 if (_envelope(args) and unfilled) else 0


def cmd_build(args) -> int:
    from . import build as B

    out = args.out or os.path.splitext(args.file)[0] + "_빌드.dsd"
    edits = B.outside_edits(out)
    if edits is None and os.path.exists(out) and not args.force:
        edits = ["(build 로 만든 기록이 없는 파일 - 손으로 만들었거나 이전 방식으로 만든 DSD)"]
    if edits and not args.force:
        print(f"멈춤: {out} 이(가) 지난 build 뒤에 직접 고쳐졌습니다 - 다시 만들면 아래 수정이 사라집니다.",
              file=sys.stderr)
        for ln in edits[:30]:
            print(f"  {ln}", file=sys.stderr)
        print("숫자는 정산표에, 행 · 표 · 문구는 뼈대 DSD 에 옮긴 뒤 다시 build 하세요(그대로 덮으려면 --force).",
              file=sys.stderr)
        return 3                                 # 멈춤 - 오류(1 · 2)와 구분(xlwork 리본이 읽음)
    package = read_dsd(args.file)
    side, conf, sheets = _recon_inputs(args, parse_contents(package.contents))
    if not sheets:
        print("오류: 정산표에서 재무제표 · 주석 시트를 못 찾음 - --config 설정을 확인하세요", file=sys.stderr)
        return 2
    recipe = B.load_recipe(args.recipe)
    rep = B.build(args.file, args.workpaper, out, recipe=recipe, fill_conf=conf, side_sheets=sheets,
                  workdir=args.workdir)
    B.save_last_build(out)
    guard.remember(out)
    _note(args, outputs=[out], warnings=rep.warnings, side=side, steps=list(rep.steps))
    print(f"build({side}) - " + " · ".join(rep.steps))
    for w in rep.warnings[:40]:
        print(f"  [주의] {w}")
    print(f"\nDSD 저장: {out}")
    if args.no_check:
        return 0
    base = os.path.splitext(out)[0]
    print("\n── 대사(recon) ──")
    cmd_recon(argparse.Namespace(cmd="recon", file=out, workpaper=args.workpaper, config=args.config,
                                 fs_sheets=None, note_sheets=None, prior=False,
                                 out=base + "_조서대사.xlsx", json_out=None))
    print("\n── 검산(check) ──")
    return cmd_check(argparse.Namespace(cmd="check", file=out, only=None, profile=None, json_out=None,
                                        out=base + "_검증.xlsx"))


def cmd_diff(args) -> int:
    from . import diffview as D

    docs = [parse_contents(read_dsd(f).contents) for f in (args.a, args.b)]
    base = parse_contents(read_dsd(args.base).contents) if args.base else None
    names = [x.strip() for x in (args.labels or "").split(",") if x.strip()]
    stem = lambda f: os.path.splitext(os.path.basename(f))[0]
    labels = (names[0] if len(names) > 0 else "A", names[1] if len(names) > 1 else "B",
              (names[2] if len(names) > 2 else "기준") if args.base else None)
    d = D.diff(docs[0], docs[1], base)
    changes = d.changes
    out = args.out or os.path.join(os.path.dirname(os.path.abspath(args.b)), f"{stem(args.a)}_vs_{stem(args.b)}_비교.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(D.render_html(d, labels, (os.path.basename(args.a), os.path.basename(args.b),
                                          os.path.basename(args.base) if args.base else None), view=args.view))
    s = D.summary(changes)
    _note(args, outputs=[out], total=s["total"],
          by_kind={D.KINDS[k][0]: n for k, n in s["by_kind"].items()})
    print(f"바뀐 곳 {s['total']}건 - " + " · ".join(f"{D.KINDS[k][0]} {n}" for k, n in s["by_kind"].items()))
    print(f"보고서: {out}")
    return 0


HELP_GROUPS = """\
하는 일로 찾기
  바꾸기        export(DSD → 엑셀)  import(엑셀 → DSD)  dart(DART 공시 → DSD)
  검산          check(합계·대사·표기)  foot(풋팅·대사 검증 워크북)  refoot(상위 합계 다시 계산)
                recheck  calibrate  ties
  다음 기간     rollforward(연차)  roll-interim(분·반기)  renumber(주석 번호)
  조서와 맞추기 recon(대사)  fill(당기 칸 채움)  build(정산표로 다시 만들기)  disc(8400 점검표)
  두 판 비교    diff(HTML 보고서)  verify(구조 비교)
  엑셀 없이     inspect(읽기·목차·찾기)  apply(JSON 편집)  annotate-ai(AI 검토 기록)

명령마다 -h 로 옵션을 본다. --json 을 주면 결과를 JSON 봉투로 낸다(docs/JSON형식.md).
종료코드: 0 정상 / 1 확인할 것 있음 / 2 입력 오류 / 3 안전장치로 멈춤(아무것도 쓰지 않음)
"""


def _fix_json_flag(argv: list[str]) -> list[str]:
    """check·recon 의 --json 은 경로를 받을 수도(옛 방식) 안 받을 수도(봉투) 있다.
    --json 바로 뒤가 .json 경로가 아니면 봉투로 읽는다(뒤따르는 파일 인자를 삼키지 않게)."""
    if not argv or argv[0] not in ("check", "recon"):
        return argv
    out = []
    for i, a in enumerate(argv):
        out.append(a)
        if a == "--json":
            nxt = argv[i + 1] if i + 1 < len(argv) else None
            if nxt is None or not nxt.lower().endswith(".json"):
                out.append("-")
    return out


def _explain(e: Exception) -> str:
    """예외 → 사람이 읽을 한 줄."""
    import json
    import zipfile

    name = getattr(e, "filename", None)
    if isinstance(e, FileNotFoundError):
        return f"파일이 없다: {name or e}"
    if isinstance(e, PermissionError):
        return (f"파일을 쓰거나 읽을 수 없다: {name or e} - 엑셀이나 DART 편집기에서 열려 있으면 "
                "닫고 다시 실행해야 한다")
    if isinstance(e, zipfile.BadZipFile):
        return "파일이 깨졌거나 DSD·엑셀 형식이 아니다"
    if isinstance(e, json.JSONDecodeError):
        return f"JSON 형식 오류: {e}"
    if type(e).__name__ == "InvalidFileException":
        return "엑셀(.xlsx) 파일이 아니다"
    return str(e)


def _check_inputs(args) -> None:
    for name, want in INPUTS.get(args.cmd, {}).items():
        path = getattr(args, name, None)
        if not path:
            continue
        guard.need(path, want)
    if args.cmd == "disc" and not args.file.lower().endswith(".json"):
        guard.need(args.file, "dsd")


def _run(args) -> int:
    """입력 점검 → 명령 실행 → 오류를 한 줄로. --json 이면 결과를 봉투로 감싼다."""
    import zipfile

    from openpyxl.utils.exceptions import InvalidFileException

    from . import result as R

    env = _envelope(args)
    capture = env and args.cmd not in SELF_EMIT
    so, se = io.StringIO(), io.StringIO()
    err = None
    code = R.OK
    known = (guard.InputError, ImportError_, DsdParseError, OSError, ValueError, zipfile.BadZipFile,
             InvalidFileException)
    try:
        _check_inputs(args)
        if capture:
            with contextlib.redirect_stdout(so), contextlib.redirect_stderr(se):
                code = args.func(args)
        else:
            code = args.func(args)
    except guard.Stopped as e:
        code, err = R.STOPPED, f"멈춤: {e}"
    except known as e:
        code, err = R.INPUT_ERROR, f"오류: {_explain(e)}"
    except Exception as e:  # noqa: BLE001 - 봉투를 내는 실행은 표준출력이 JSON 이어야 한다
        if not env:
            raise
        code, err = R.INPUT_ERROR, f"내부 오류: {type(e).__name__}: {e}"
    if not env:
        if err:
            print(err, file=sys.stderr)
        return code
    if args.cmd in SELF_EMIT:
        if err:
            return R.emit(R.envelope(args.cmd, code, errors=[err.split(": ", 1)[-1]]))
        return code
    res = getattr(args, "_res", {"outputs": [], "warnings": [], "extra": {}})
    said = [ln.strip() for ln in se.getvalue().splitlines() if ln.strip()]
    errors = ([err.split(": ", 1)[-1]] if err else []) + (said if code >= R.INPUT_ERROR else [])
    warnings = list(res["warnings"]) + ([] if code >= R.INPUT_ERROR else said)
    extra = dict(res["extra"])
    if extra.pop("log", True):               # 구조화한 결과가 충분한 명령(check)은 사람용 줄을 싣지 않는다
        extra["log"] = [ln.rstrip() for ln in so.getvalue().splitlines() if ln.strip()]
    name = args.cmd + (" " + args.tcmd if args.cmd == "ties" else "")
    return R.emit(R.envelope(name, code, outputs=res["outputs"], warnings=warnings, errors=errors,
                             **extra))


def main(argv=None) -> int:
    _force_utf8_io()
    ap = argparse.ArgumentParser(
        prog="dsdxlsx",
        description="DART DSD ↔ XLSX 왕복 변환 (무수정 부분 바이트 보존)",
        epilog=HELP_GROUPS, formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    from . import __version__
    ap.add_argument("--version", action="version", version=f"dsdxlsx {__version__}")
    sub = ap.add_subparsers(dest="cmd", required=True, metavar="명령")

    p = sub.add_parser("export", help="DSD → XLSX 내보내기")
    p.add_argument("dsd")
    p.add_argument("-o", "--out", help="출력 엑셀(기본: <이름>.xlsx)")
    p.add_argument("--force", action="store_true", help="출력 파일이 고쳐졌거나 기록이 없어도 덮어씀")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_export)

    p = sub.add_parser("import", help="편집한 XLSX → DSD 재변환(export 로 만든 편집용 엑셀, 또는 foot 으로 만든 "
                                      "풋팅 워크북 - 풋팅 워크북은 칸 값 수정만)")
    p.add_argument("xlsx")
    p.add_argument("-o", "--out", help="출력 DSD(기본: <이름>_수정.dsd)")
    p.add_argument("--force", action="store_true", help="출력 파일이 고쳐졌거나 기록이 없어도 덮어씀")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_import)

    p = sub.add_parser("inspect", help="DSD 를 칸 ID·행 이름·열 머리글과 함께 마크다운이나 JSON 으로 펼침(LLM 읽기용)")
    p.add_argument("dsd")
    p.add_argument("--section", action="append",
                   help="이 구획만(재무제표 명칭 또는 주석 번호, 쉼표로 여러 개, 반복 가능)")
    p.add_argument("--outline", action="store_true",
                   help="목차만: 구획마다 표·문단 ID 와 크기(큰 문서에서 먼저 볼 것)")
    p.add_argument("--find", action="append", metavar="글",
                   help="이 글이 든 문단·행·칸의 위치(ID)만(숫자는 값으로도 찾음, 반복 가능)")
    p.add_argument("--limit", type=int, default=200, help="--find 결과 최대 건수(기본 200)")
    p.add_argument("--json", action="store_true", help="JSON 으로(기본: 마크다운)")
    p.add_argument("--pretty", action="store_true", help="JSON 들여쓰기")
    p.add_argument("-o", "--out", help="저장 경로(기본: 표준출력)")
    p.set_defaults(func=cmd_inspect)

    p = sub.add_parser("apply", help="JSON 편집 목록을 DSD 에 반영(엑셀 없이 - import 와 같은 가드)")
    p.add_argument("dsd")
    p.add_argument("edits", help="편집 목록 JSON 파일(- 이면 표준입력)")
    p.add_argument("-o", "--out", help="출력 DSD(기본: <이름>_수정.dsd, 있으면 멈춤)")
    p.add_argument("--dry-run", action="store_true", help="쓰지 않고 바뀔 내용만")
    p.add_argument("--force", action="store_true", help="출력 파일이 있어도 덮어씀")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(종료코드 0/2/3)")
    p.set_defaults(func=cmd_apply)

    p = sub.add_parser("dart", help="DART 에 공시된 감사보고서를 DSD 로(+ 엑셀·롤포워드, OpenDART 키 필요)")
    p.add_argument("query", help="회사명·종목코드(6자리)·고유번호(8자리)·DART 공시 주소·접수번호(14자리)")
    p.add_argument("--form", choices=["감사", "연결감사"], default="감사",
                   help="감사(별도) 또는 연결감사(기본: 감사)")
    p.add_argument("--year", type=int, help="사업연도(기본: 가장 최근 공시)")
    p.add_argument("--list", action="store_true", help="받지 않고 고를 수 있는 공시만 보여 줌")
    p.add_argument("-o", "--out", help="출력 DSD 경로나 폴더(기본: 현재 폴더 <회사>_<서식>_<연도>.dsd)")
    p.add_argument("--xlsx", action="store_true", help="같은 이름의 편집용 엑셀도 만듦(export)")
    p.add_argument("--rollforward", action="store_true",
                   help="당기 작성용 기초 DSD·정산표도 만듦(rollforward, _차기기초)")
    p.add_argument("--keep-lf", action="store_true", help="줄끝을 원문(LF) 그대로 둠(기본: CRLF)")
    p.add_argument("--force", action="store_true", help="출력 파일이 있어도 덮어씀")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(종료코드 0/2/3)")
    p.set_defaults(func=cmd_dart)

    p = sub.add_parser("foot", help="풋팅·대사 검증 워크북 - DSD 를 합계·대사 수식이 든 엑셀로"
                                    "(총괄표 / 원문 / 재무제표 / 주석 시트, 값을 고치면 바로 다시 계산)")
    p.add_argument("files", nargs="+", help=".dsd 파일(여러 개 가능)")
    p.add_argument("-o", "--out", help="출력 엑셀(기본: <이름>_풋팅.xlsx, 입력이 하나일 때만)")
    p.add_argument("--convert-only", action="store_true",
                   help="검증 수식 없이 시트로 옮기기만(출력: <이름>_변환.xlsx)")
    p.add_argument("--sections", action="store_true", help="시트 구성만 보여 주고 끝냄")
    p.add_argument("--detection-profile", choices=("standard", "exhaustive"), default="standard",
                   help="검출 범위: standard(기본) / exhaustive(±1 절사까지 넓게 - 재판정용)")
    p.add_argument("--dart-tie", action="store_true",
                   help="전기 열을 DART 에 공시된 직전 감사·검토보고서 원문과 대사(OpenDART 키 필요, 없으면 "
                        "건너뜀). 각 시트 오른쪽 'DART:' 칸과 'DART원천' 시트에 남긴다")
    p.add_argument("--dart-fs", choices=("OFS", "CFS"),
                   help="DART 대사 대상을 별도(OFS)·연결(CFS)로 강제(기본: 문서에서 판별)")
    p.add_argument("--force", action="store_true", help="출력 파일이 고쳐졌거나 기록이 없어도 덮어씀")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_foot)

    p = sub.add_parser("diff", help="두 DSD 의 보이는 차이를 HTML 보고서로(두 판 나란히 보기 + 바뀐 곳만 카드)")
    p.add_argument("a", help="비교 기준 DSD(예: 우리 판)")
    p.add_argument("b", help="비교 대상 DSD(예: 감사인이 고친 판)")
    p.add_argument("--base", help="기준 DSD(예: 직전 분기 최종) - 바뀐 곳이 이어받은 것인지 표시")
    p.add_argument("--labels", help="이름 'A,B[,기준]' (예: 우리 v1.6,감사인,2Q 최종)")
    p.add_argument("-o", "--out", help="저장 경로(기본: B 폴더에 <A>_vs_<B>_비교.html)")
    p.add_argument("--view", choices=("side", "cards"), default="side",
                   help="처음 여는 화면: side=두 판 나란히(기본) · cards=바뀐 곳만 (보고서 안 탭으로 바꿀 수 있음)")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_diff)

    p = sub.add_parser("verify", help="두 DSD 파일 구조 비교")
    p.add_argument("a")
    p.add_argument("b")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_verify)

    p = sub.add_parser("disc", help="공시사항점검표(한공회 8400) 자동 작성 — 해당여부·주석번호")
    p.add_argument("file", help=".dsd 또는 check --json 으로 만든 review.json")
    p.add_argument("-o", "--out", help="저장 경로(기본: <대상>_8400공시점검표.xlsx, 있으면 _2)")
    p.add_argument("--form", choices=["kifrs", "kgaap"], help="양식(기본: 본문에서 자동 판별)")
    p.add_argument("--listed", choices=["yes", "no"], help="상장 여부(기본: 본문에서 자동 판별)")
    p.add_argument("--company", help="회사명(기본: 본문에서 추출)")
    p.add_argument("--period-end", help="결산일 YYYY-MM-DD(기본: 본문에서 추출)")
    p.add_argument("--kicpa-dir", help="한공회 표준조서 폴더(기본: KICPA_FORMS 환경변수)")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_disc)

    p = sub.add_parser("check", help="합계·대사·표기 정합성 검증 (.dsd 또는 export한 .xlsx)")
    p.add_argument("file")
    p.add_argument("--only", help="쉼표 구분 검사 선택 (예: C1,C4)")
    p.add_argument("--profile",
                   help="검증 프로파일: 회사명 또는 profile.toml 경로 "
                        "(기본: meta.xml 회사명으로 profiles/ 자동 탐색)")
    p.add_argument("--json", dest="json_out", nargs="?", const="-", metavar="경로",
                   help="경로 없이 주면 결과를 JSON 봉투로(표준출력, 불일치가 있으면 종료코드 1). "
                        "경로를 주면 리뷰용 구조화 JSON 을 그 파일로 — AI 검토(/dsd-review) 입력")
    p.add_argument("-o", "--out", help="검증 결과 엑셀 경로 (기본: <대상>_검증.xlsx)")
    p.set_defaults(func=cmd_check)

    p = sub.add_parser(
        "recon",
        help="DSD ↔ 조서(최종분석적검토 등) 대사 — 재무제표·주석 칸마다 조서에서 같은 값 찾기",
    )
    p.add_argument("file", help=".dsd 또는 export한 .xlsx")
    p.add_argument("workpaper", help="조서 엑셀 (계산값 저장된 파일)")
    p.add_argument("--fs-sheets", help="재무제표 시트 이름(쉼표) — 기본: BS · IS · CF · SCE · CE · FS(공시 · 정산 꼬리) 자동")
    p.add_argument("--note-sheets", help="주석 시트 이름(쉼표) — 기본: '주석' · FN · 'N. 제목' 시트 자동")
    p.add_argument("--prior", action="store_true",
                   help="조서에 기간 열이 없어도 전기 칸까지 값 찾기 (기본: 당기만)")
    p.add_argument("--config",
                   help="건별 설정 JSON(시트 지정 · 빼는 칸 · 이름 다른 행, '별도' · '연결' 구역) - "
                        "없으면 자동 판정 결과를 이 파일로 써 줌")
    p.add_argument("--json", dest="json_out", nargs="?", const="-", metavar="경로",
                   help="경로 없이 주면 결과를 JSON 봉투로(표준출력, 확인 필요가 있으면 종료코드 1). "
                        "경로를 주면 요약 JSON 을 그 파일로 (xlaudit 리본이 읽음)")
    p.add_argument("-o", "--out", help="결과 엑셀 경로 (기본: <대상>_조서대사.xlsx)")
    p.set_defaults(func=cmd_recon)

    p = sub.add_parser(
        "fill",
        help="DSD 당기 칸 채우기 - roll-interim 으로 민 DSD 의 빈 당기 칸을 조서(8110 · 정산표)에서 (recon 짝짓기를 거꾸로)",
    )
    p.add_argument("file", help="당전기를 민 .dsd (당기 칸이 빈 골격)")
    p.add_argument("workpaper", help="조서 엑셀 (계산값 저장된 파일)")
    p.add_argument("--config", help="건별 설정 JSON(recon 과 같은 파일 - 시트 지정 · 빼는 칸 · 이름 다른 행)")
    p.add_argument("--keep-filled", action="store_true", help="값이 있는 당기 칸은 그대로 두고 빈 칸만 (기본: 덮어씀)")
    p.add_argument("--report", help="결과 엑셀 경로 (기본: <출력>_채움결과.xlsx)")
    p.add_argument("-o", "--out", help="출력 .dsd (기본: <대상>_채움.dsd)")
    p.add_argument("--force", action="store_true", help="출력 파일이 고쳐졌거나 기록이 없어도 덮어씀")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_fill)

    p = sub.add_parser(
        "build",
        help="정산표 → DSD 다시 만들기 - 뼈대 DSD + 정산표 + 건별 레시피(fill · 표 대응 · 단위 기호 · 문단) → 대사 · 검산",
    )
    p.add_argument("file", help="뼈대 .dsd (당전기를 밀고 행 구조를 맞춘 것)")
    p.add_argument("workpaper", help="정산표 엑셀 (계산값 저장된 파일)")
    p.add_argument("--config", help="건별 설정 JSON(recon · fill 과 같은 파일)")
    p.add_argument("--recipe", help="빌드 레시피 JSON(_engagement/dsd_build.json) - 표 대응 · 문단 · 유지 · 글 바꾸기")
    p.add_argument("--force", action="store_true", help="출력 DSD 를 직접 고친 흔적이 있어도 덮어씀")
    p.add_argument("--no-check", action="store_true", help="빌드 뒤 대사 · 검산 생략")
    p.add_argument("--workdir", help="중간 파일(채움 · 레시피 엑셀)을 남길 폴더 (기본: 임시 폴더)")
    p.add_argument("-o", "--out", help="출력 .dsd (기본: <뼈대>_빌드.dsd)")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_build)

    p = sub.add_parser(
        "annotate-ai",
        help="AI 검토 소견(findings.json)을 편집용 xlsx 'AI검토' 시트로 기록",
    )
    p.add_argument("xlsx")
    p.add_argument("findings", help="/dsd-review 스킬이 만든 findings.json")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_annotate_ai)

    p = sub.add_parser(
        "recheck",
        help="annotate된 xlsx의 검산 수식을 Excel 없이 재평가 (헤드리스, OfficeCLI 필요)",
    )
    p.add_argument("xlsx")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_recheck)

    p = sub.add_parser(
        "refoot",
        help="상위 합계 자동 재계산 — 전기 열의 검증된 구조를 당기 열에 적용 (미리보기 기본)",
    )
    p.add_argument("xlsx", help="편집 중인 정산표 (export/rollforward 산출)")
    p.add_argument("--apply", action="store_true", help="제안을 실제 셀에 기입")
    p.add_argument("--profile", help="검증 프로파일 (기본: meta 회사명)")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_refoot)

    p = sub.add_parser(
        "rollforward",
        help="전기 DSD → 당기 기초 DSD + 표준 정산표 (기수·TU +1, 당기값→전기 열, 3등급 리포트)",
    )
    p.add_argument("dsd", help="전기(직전 보고서) DSD")
    p.add_argument("--profile", help="검증 프로파일 (기본: meta 회사명)")
    p.add_argument("-o", "--out", help="출력 기본 경로 (기본: <이름>_차기기초.*)")
    p.add_argument("--force", action="store_true", help="출력 파일이 고쳐졌거나 기록이 없어도 덮어씀")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
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
    p.add_argument("--force", action="store_true", help="출력 파일이 고쳐졌거나 기록이 없어도 덮어씀")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
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
    p.add_argument("--force", action="store_true", help="출력 파일이 고쳐졌거나 기록이 없어도 덮어씀")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_renumber)

    p = sub.add_parser(
        "calibrate",
        help="새 회사 온보딩: 전기 심리완료 DSD 전수 통과 확인 + suppress 초안 스캐폴딩",
    )
    p.add_argument("dsd", help="전기(심리 완료) DSD 파일")
    p.add_argument("--profile", help="회사명 (기본: meta.xml 회사명)")
    p.add_argument("--write", action="store_true",
                   help="suppress 초안을 profiles/<회사>/profile.toml에 주석으로 추가")
    p.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    p.set_defaults(func=cmd_calibrate)

    p = sub.add_parser("ties", help="대사맵(대사쌍) 관리")
    tsub = p.add_subparsers(dest="tcmd", required=True)
    ps = tsub.add_parser(
        "save", help="대사맵 시트의 확정/제외 상태를 profiles/<회사>/ties.json으로 저장"
    )
    ps.add_argument("xlsx")
    ps.add_argument("--profile", help="회사명 (기본: 임베드 원본 meta.xml의 회사명)")
    ps.add_argument("--json", action="store_true", help="결과를 JSON 봉투로(표준출력, 종료코드 0/1/2/3)")
    ps.set_defaults(func=cmd_ties_save)

    args = ap.parse_args(_fix_json_flag(list(sys.argv[1:] if argv is None else argv)))
    return _run(args)


if __name__ == "__main__":
    sys.exit(main())
