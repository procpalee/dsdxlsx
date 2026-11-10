"""writeback — 풋팅 워크북 셀 편집 → DSD 반영 (보존 계약·변환 규칙·가드)."""
from __future__ import annotations

import hashlib
import warnings

import openpyxl
import pytest

from dsdxlsx.foot import mapsheet, pipeline, writeback
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.parse import parse_contents


def _make(tmp_path, convert_only=True):
    from fixtures.foot_factory import write_dsd

    dsd = tmp_path / "synth.dsd"
    write_dsd(dsd)
    xlsx = tmp_path / "out.xlsx"
    pipeline.process_file(dsd, xlsx, convert_only=convert_only)
    return dsd, xlsx


def _sha(path):
    return hashlib.sha256(read_dsd(str(path)).contents).hexdigest()


def _edit(xlsx, sheet, addr, value):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(xlsx))
    wb[sheet][addr] = value
    wb.save(str(xlsx))


def _find(xlsx, sheet, text):
    """시트에서 text 와 같은 셀 좌표(row, col) 찾기."""
    wb = openpyxl.load_workbook(str(xlsx), read_only=True)
    ws = wb[sheet]
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.strip() == text:
                wb.close()
                return c.row, c.column
    wb.close()
    raise AssertionError(f"{sheet} 에서 '{text}' 없음")


def test_map_embedded_and_readable(tmp_path):
    dsd, xlsx = _make(tmp_path)
    meta, dsd_bytes, entries = mapsheet.read(str(xlsx))
    assert meta["tool"] == "dsdxlsx-foot" and meta["source"] == "synth.dsd"
    assert dsd_bytes == dsd.read_bytes()
    kinds = {e.kind for e in entries}
    assert kinds == {"cell", "para"}
    # 원문 + 개별 시트 이중 매핑 — 같은 노드가 두 시트에
    sheets_of = {}
    for e in entries:
        sheets_of.setdefault(e.node_id, set()).add(e.sheet)
    assert any(len(s) >= 2 and "원문" in s for s in sheets_of.values())


def test_no_edit_roundtrip_identical(tmp_path):
    dsd, xlsx = _make(tmp_path)
    r = writeback.writeback(xlsx)
    assert r.rc == 0 and r.edits == [] and r.out.name == "synth_수정.dsd"
    assert _sha(r.out) == _sha(dsd)


def test_verify_workbook_also_writes_back(tmp_path):
    """검증 수식·오른쪽 슬롯이 있어도 무편집 왕복은 동일."""
    dsd, xlsx = _make(tmp_path, convert_only=False)
    r = writeback.writeback(xlsx)
    assert r.rc == 0 and _sha(r.out) == _sha(dsd)


def test_number_edit_changes_only_that_node(tmp_path):
    dsd, xlsx = _make(tmp_path)
    r, c = _find(xlsx, "BS", "매출채권")
    # 당기 금액 300 → 301 (매출채권 행, 금액 열 = 라벨 오른쪽 첫 열)
    _edit(xlsx, "BS", openpyxl.utils.get_column_letter(c + 1) + str(r), 301)
    res = writeback.writeback(xlsx)
    assert res.rc == 0
    assert len(res.edits) == 1
    e = res.edits[0]
    assert e.old == "300" and e.new == "301"
    doc = parse_contents(read_dsd(str(res.out)).contents)
    assert doc.nodes[e.node_id].text == "301"
    # 그 노드 밖 바이트는 그대로 (앞·뒤 길이 비교)
    before = read_dsd(str(dsd)).contents
    after = read_dsd(str(res.out)).contents
    assert len(after) == len(before)  # '300'→'301' 같은 길이
    assert sum(1 for a, b in zip(before, after) if a != b) == 1


def test_negative_and_zero_conventions(tmp_path):
    dsd, xlsx = _make(tmp_path)
    r, c = _find(xlsx, "CF", "(2) 자산부채의 변동")
    col = openpyxl.utils.get_column_letter(c + 1)
    _edit(xlsx, "CF", f"{col}{r}", -601)      # (600) → (601)
    r2, c2 = _find(xlsx, "BS", "매입채무")
    _edit(xlsx, "BS", f"{openpyxl.utils.get_column_letter(c2 + 1)}{r2}", 0)  # 800 → '-'
    res = writeback.writeback(xlsx)
    assert res.rc == 0
    news = {e.old: e.new for e in res.edits}
    assert news["(600)"] == "(601)"
    assert news["800"] == "-"


def test_dash_left_as_zero_is_unchanged(tmp_path):
    """'-' 플레이스홀더는 0으로 내보내지므로 0 그대로면 무편집."""
    dsd, xlsx = _make(tmp_path)
    # 합성 문서엔 '-' 가 없으므로 to_node_text/unchanged 단위로 검증
    assert writeback.unchanged(0, "-", False)
    assert writeback.to_node_text(0, "-", False) == "-"
    assert writeback.to_node_text(1234, "-", False) == "1,234"


def test_percent_and_suffix_roundtrip():
    assert writeback.unchanged(0.0435, "4.35%", False)
    assert writeback.to_node_text(0.0436, "4.35%", False) == "4.36%"
    assert writeback.to_node_text(0.05, "4.35%", False) == "5.00%"
    assert writeback.unchanged(10000, "10,000원", False)
    assert writeback.to_node_text(12000, "10,000원", False) == "12,000원"
    assert writeback.to_node_text(2978481, "2,978,480주", False) == "2,978,481주"
    assert writeback.to_node_text(-5, "△4", False) == "△5"
    assert writeback.to_node_text("새 라벨", "옛 라벨", False) == "새 라벨"
    # 주석 참조 열은 항상 텍스트
    assert writeback.unchanged("4,5,6", "4,5,6", True)
    assert writeback.to_node_text("4,5,7", "4,5,6", True) == "4,5,7"


def test_paragraph_edit_and_partial(tmp_path):
    dsd, xlsx = _make(tmp_path)
    # 주석 2 시트: 헤딩('2. 현금및현금성자산')이 제거된 partial 문단
    r, c = _find(xlsx, "2", "현금및현금성자산의 내역은 다음과 같습니다.")
    _edit(xlsx, "2", f"{openpyxl.utils.get_column_letter(c)}{r}", "현금및현금성자산의 내역은 아래와 같습니다.")
    res = writeback.writeback(xlsx)
    assert res.rc == 0 and len(res.edits) == 1
    doc = parse_contents(read_dsd(str(res.out)).contents)
    node = doc.nodes[res.edits[0].node_id]
    assert node.text.startswith("2. 현금및현금성자산")  # 헤딩은 보존
    assert node.text.endswith("아래와 같습니다.")


def test_conflicting_edits_error(tmp_path):
    dsd, xlsx = _make(tmp_path)
    r, c = _find(xlsx, "BS", "매출채권")
    col = openpyxl.utils.get_column_letter(c + 1)
    _edit(xlsx, "BS", f"{col}{r}", 301)
    r2, c2 = _find(xlsx, "원문", "매출채권")
    _edit(xlsx, "원문", f"{openpyxl.utils.get_column_letter(c2 + 1)}{r2}", 302)
    res = writeback.writeback(xlsx)
    assert res.rc == 2 and res.out is None
    assert any("서로 다르게" in e for e in res.errors)


def test_row_insert_is_refused(tmp_path):
    dsd, xlsx = _make(tmp_path)
    r, c = _find(xlsx, "BS", "매출채권")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(xlsx))
    wb["BS"].insert_rows(r)
    wb.save(str(xlsx))
    res = writeback.writeback(xlsx)
    assert res.rc == 2 and res.out is None
    assert any("밀렸" in e or "행 이동" in e for e in res.errors)


def test_readonly_tu_and_title_edit_is_warning(tmp_path):
    """TU 자동필드·TITLE 라벨 편집은 반영하지 않고 경고(rc 1), 나머지 편집은 반영."""
    from fixtures.foot_factory import build_dsd, standard_contents

    xml = standard_contents().replace(
        '<TD COLSPAN="2">제 2(당)기말 : 20X2년 12월 31일 현재</TD>'.encode("utf-8"),
        '<TU AUNIT="PERIODTO" AUNITVALUE="20X21231" COLSPAN="2">제 2(당)기말 : 20X2년 12월 31일 현재</TU>'.encode("utf-8"))
    assert b"<TU " in xml
    dsd = tmp_path / "tu.dsd"
    dsd.write_bytes(build_dsd(xml))
    xlsx = tmp_path / "out.xlsx"
    pipeline.process_file(dsd, xlsx, convert_only=True)
    r, c = _find(xlsx, "BS", "제 2(당)기말 : 20X2년 12월 31일 현재")
    _edit(xlsx, "BS", f"{openpyxl.utils.get_column_letter(c)}{r}", "제 3(당)기말")
    r2, c2 = _find(xlsx, "원문", "재 무 상 태 표")
    _edit(xlsx, "원문", f"{openpyxl.utils.get_column_letter(c2)}{r2}", "재무상태표(수정)")
    r3, c3 = _find(xlsx, "BS", "매출채권")
    _edit(xlsx, "BS", f"{openpyxl.utils.get_column_letter(c3 + 1)}{r3}", 301)
    res = writeback.writeback(xlsx)
    assert res.rc == 1 and res.out is not None
    assert len(res.edits) == 1 and res.edits[0].new == "301"
    assert sum("수정 불가" in w for w in res.warnings) == 2


def test_not_a_foot_workbook(tmp_path):
    p = tmp_path / "plain.xlsx"
    openpyxl.Workbook().save(str(p))
    with pytest.raises(mapsheet.MapError):
        writeback.writeback(p)


def test_sample_no_edit_roundtrip(foot_sample_dsd, tmp_path):
    sample_dsd = foot_sample_dsd
    xlsx = tmp_path / "out.xlsx"
    pipeline.process_file(sample_dsd, xlsx, convert_only=True)
    r = writeback.writeback(xlsx, tmp_path / "rt.dsd")
    assert r.rc == 0, r.warnings
    assert _sha(r.out) == _sha(sample_dsd)
