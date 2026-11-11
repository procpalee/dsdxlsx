"""파이프라인 — 파일 하나: 파싱 → 분석 → 레이아웃 → 바인딩 → 방출 → 요약."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from . import constants as C
from . import emit, layout, report, sections, summary
from .tableview import build_views
from ..dsd_io import read_dsd
from ..parse import parse_contents


@dataclass
class ProcessResult:
    """process_file_full 반환 묶음 - 대사 제안·DART 대사 같은 후속 단계가 쓴다."""

    rc: int
    out: Path
    checkset: object       # CheckSet
    company: str
    n_sheets: int
    source: Path
    detection_profile: str = "standard"
    right_starts: dict = None      # 시트 → 오른쪽 검증 슬롯 시작 열 (AI 대사 제안 배치용)
    sheet_units: dict = None       # 시트 → {표 첫 xlsx 행: 단위 배수} (AI 대사 제안 검산용)
    # DART 대사 후처리용 — 파싱 문서·시트 계획·표 뷰·표 배치
    package: object = None
    doc: object = None
    plans: list = None
    ctx: dict = None
    placed_rows: dict = None       # table_id → (sheet, start_row)


def default_output(path: Path, convert_only: bool = False) -> Path:
    suffix = C.CONVERT_SUFFIX if convert_only else C.OUTPUT_SUFFIX
    return Path(path).with_name(f"{Path(path).stem}{suffix}.xlsx")


def analyze(doc, relaxed: bool = False):
    """파싱된 문서 → (plans, views, ctx, footing_specs, cross_specs)."""
    from . import crossref, footing

    plans = sections.build(doc)
    views, ctx = build_views(plans)
    footing_specs = footing.analyze(views, relaxed=relaxed)
    cross_specs = crossref.analyze(views, ctx, sections.note_sheet_map(plans))
    return plans, views, ctx, footing_specs, cross_specs


@dataclass
class Prepared:
    """문서 하나의 분석·배치 결과(워크북을 쓰기 전)."""

    plans: list
    ctx: dict
    layouts: list
    placed_map: dict
    checkset: object
    right_starts: dict
    sheet_units: dict
    placed_rows: dict


def prepare(doc, convert_only: bool = False, relaxed: bool = False) -> Prepared:
    if convert_only:
        plans = sections.build(doc)
        views, ctx = build_views(plans)
        footing_specs, cross_specs = [], []
    else:
        plans, views, ctx, footing_specs, cross_specs = analyze(doc, relaxed=relaxed)
    specs_by_table: dict[str, list] = {}
    for s in footing_specs:
        specs_by_table.setdefault(s.table_id, []).append(s)
    layouts, placed_map = layout.build(plans, specs_by_table)

    def value_of(table_id: str, i: int, c: int):
        return emit.written_value(ctx["by_table"][table_id], i, c)

    checkset = layout.bind(layouts, placed_map, footing_specs, cross_specs, value_of)
    # 대사 칸을 후속으로 심을 때 필요한 배치·단위 정보
    right_starts = {lay.name: lay.right_start for lay in layouts}
    sheet_units: dict[str, dict[int, int]] = {}
    for tid, placed in placed_map.items():
        tv = ctx["by_table"].get(tid)
        if tv is not None:
            sheet_units.setdefault(placed.sheet, {})[placed.start_row] = tv.unit
    placed_rows = {tid: (pl.sheet, pl.start_row) for tid, pl in placed_map.items()}
    return Prepared(plans, ctx, layouts, placed_map, checkset, right_starts, sheet_units, placed_rows)


def process_file(path: Path, output: Path | None = None,
                 convert_only: bool = False) -> int:
    return process_file_full(path, output, convert_only).rc


def process_file_full(path: Path, output: Path | None = None,
                      convert_only: bool = False,
                      relaxed: bool = False) -> ProcessResult:
    package = read_dsd(str(path))
    doc = parse_contents(package.contents)
    company = sections.company_name(package.meta)

    prep = prepare(doc, convert_only=convert_only, relaxed=relaxed)
    plans, ctx, layouts, checkset = prep.plans, prep.ctx, prep.layouts, prep.checkset
    profile = "exhaustive" if relaxed else "standard"

    out = output or default_output(path, convert_only)
    titles = {"원문": plans[0].display_title if plans else ""}
    titles.update({p.name: p.display_title for p in plans})
    emit.write_workbook(
        layouts, checkset, str(out),
        summary_writer=lambda ws, lays, cs: summary.write(ws, lays, cs, titles),
        # 원본 DSD + 셀맵 임베드 → 엑셀에서 고친 값을 DSD 로 되돌릴 때 쓴다
        embed=(Path(path).read_bytes(), Path(path).name, {"detection_profile": profile}),
    )
    if convert_only:
        print(f"\n[{path.name}] {company}")
        print(f"  시트 {len(plans) + 2}개 · 변환 전용(검증 수식 없음)")
        rc = 0
    else:
        rc = report.print_summary(path.name, company, len(plans) + 2, checkset)
    print(f"  → {out.name}")
    return ProcessResult(rc=rc, out=Path(out), checkset=checkset,
                         company=company, n_sheets=len(plans) + 2,
                         source=Path(path),
                         detection_profile=profile,
                         right_starts=prep.right_starts, sheet_units=prep.sheet_units,
                         package=package, doc=doc, plans=plans, ctx=ctx,
                         placed_rows=prep.placed_rows)
