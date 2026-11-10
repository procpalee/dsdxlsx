"""자체감사 — 산출 워크북의 모든 수식을 재평가해 판정과 일치함을 강제.

합성 픽스처는 항상, 실샘플은 존재할 때만 (skip-if-absent).
"""
import warnings

import openpyxl
import pytest

from dsdxlsx.foot import pipeline
from dsdxlsx.foot.selfaudit import audit


def _run(dsd_path, tmp_path):
    out = tmp_path / f"{dsd_path.stem}.xlsx"
    rc = pipeline.process_file(dsd_path, out)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(out))
    return rc, wb, out


def test_synth_selfaudit(tmp_path):
    from fixtures.foot_factory import write_dsd

    p = tmp_path / "synth.dsd"
    write_dsd(p)
    rc, wb, out = _run(p, tmp_path)
    assert rc == 0  # 깨끗한 합성 문서 — FAIL 없음
    checkset = _rebuild_checkset(p)
    assert checkset.footing, "합성 문서에서 풋팅 검증이 생성되어야 한다"
    problems = audit(wb, checkset)
    assert problems == []


def _rebuild_checkset(dsd_path):
    """process_file과 동일 경로로 CheckSet 재구성 (평가 대조용)."""
    from dsdxlsx.foot import emit, layout
    from dsdxlsx.dsd_io import read_dsd
    from dsdxlsx.parse import parse_contents

    package = read_dsd(str(dsd_path))
    doc = parse_contents(package.contents)
    plans, views, ctx, fspecs, xspecs = pipeline.analyze(doc)
    by_table: dict = {}
    for s in fspecs:
        by_table.setdefault(s.table_id, []).append(s)
    layouts, placed = layout.build(plans, by_table)

    def value_of(tid, i, c):
        return emit.written_value(ctx["by_table"][tid], i, c)

    return layout.bind(layouts, placed, fspecs, xspecs, value_of)


def _expected_rc(stem: str) -> int:
    """실샘플의 기대 종료코드 (기본 0 = 전부 통과).

    원본에 실제 합계오류가 있는 샘플은 '검출(rc=1)'이 정답이다. 그 예외 목록은
    파일명(=고객사명)을 키로 하므로 저장소에 두지 않고 tests/sample_exceptions.py
    (git 제외)에 로컬로 둔다 — 없으면 전 샘플이 통과해야 한다고 본다."""
    try:
        from foot_sample_exceptions import EXPECTED_RC  # git 제외 — 없으면 기본값
    except ImportError:
        return 0
    for prefix, rc in EXPECTED_RC.items():
        if stem.startswith(prefix):
            return rc
    return 0


def test_samples_selfaudit(foot_sample_dsd, tmp_path):
    sample_dsd = foot_sample_dsd
    # sample_dsd 파라미터는 conftest.pytest_generate_tests가 실샘플 파일별로 생성
    rc, wb, out = _run(sample_dsd, tmp_path)
    checkset = _rebuild_checkset(sample_dsd)
    problems = audit(wb, checkset)
    assert problems == [], f"{sample_dsd.name}: {problems[:5]}"
    assert rc == _expected_rc(sample_dsd.stem), f"{sample_dsd.name}: 기대 종료코드와 다름"
