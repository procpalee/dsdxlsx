import warnings

import openpyxl

from dsdxlsx.foot import pipeline


def test_convert_only_no_formulas(tmp_path):
    """--convert-only: 시트 구성은 동일하되 검증 수식·검증 라벨이 전혀 없다."""
    from fixtures.foot_factory import write_dsd

    p = tmp_path / "synth.dsd"
    write_dsd(p)
    out = tmp_path / "converted.xlsx"
    rc = pipeline.process_file(p, out, convert_only=True)
    assert rc == 0
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(out))
    visible = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
    assert visible == ["총괄표", "원문", "BS", "PL", "CF", "1", "2", "3"]
    # 셀맵(writeback 전제)은 veryHidden 으로만 존재
    assert wb["_map"].sheet_state == "veryHidden"
    assert wb["_cellmap"].sheet_state == "veryHidden"
    for name in visible:
        if name == "총괄표":
            continue  # 총괄표 하이퍼링크·홈링크는 허용
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("=") and "HYPERLINK" not in v:
                    raise AssertionError(f"{name}!{cell.coordinate} 검증 수식 발견: {v[:60]}")
    # 내용은 그대로 (BS 자산총계 등) — 본문 6행·B열 시작 골격 기준
    assert wb["BS"].cell(16, 4).value == 2000  # 자산총계 당기


def test_convert_only_output_name(tmp_path):
    from fixtures.foot_factory import write_dsd

    p = tmp_path / "synth.dsd"
    write_dsd(p)
    rc = pipeline.process_file(p, None, convert_only=True)
    assert rc == 0
    assert (tmp_path / "synth_변환.xlsx").exists()
