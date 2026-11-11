"""DSD 비교(diff) — 정렬 기반: 표 삽입으로 ID가 밀려도 뒤 표가 변경으로 잡히지 않아야 한다."""
from __future__ import annotations

from pathlib import Path

import pytest

from dsdxlsx import splice
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.foot import docdiff as diff
from dsdxlsx.parse import parse_contents


@pytest.fixture()
def base(tmp_path):
    from fixtures.foot_factory import build_dsd, standard_contents

    xml = standard_contents()
    old = tmp_path / "old.dsd"
    old.write_bytes(build_dsd(xml))
    return old, xml


def _write(tmp_path, name, xml):
    from fixtures.foot_factory import build_dsd
    p = tmp_path / name
    p.write_bytes(build_dsd(xml))
    return p


def _find_cell(doc, text):
    for n in doc.nodes.values():
        if n.tag in ("TD", "TH") and n.text.strip() == text:
            return n
    raise AssertionError(text)


def test_identical_files_no_changes(base, tmp_path):
    old, xml = base
    new = _write(tmp_path, "new.dsd", xml)
    r = diff.compare(old, new)
    assert not r.has_changes and all(s.status == "same" for s in r.sections)


def test_single_value_change_with_delta(base, tmp_path):
    old, xml = base
    doc = parse_contents(xml)
    node = _find_cell(doc, "300")  # BS 매출채권 당기
    new_xml = splice.apply(xml, [splice.replace_inner(node, "350")])
    new = _write(tmp_path, "new.dsd", new_xml)
    r = diff.compare(old, new)
    assert r.counts["값변경"] == 1 and r.counts["텍스트변경"] == 0
    assert r.counts["행추가"] == r.counts["행삭제"] == r.counts["표추가"] == 0
    sec = next(s for s in r.sections if s.status == "changed")
    assert "재무상태표" in sec.title
    cd = [c for bd in sec.blocks for row in bd.rows for c in row.cells][0]
    assert cd.old == "300" and cd.new == "350" and cd.delta == 50


def test_row_delete_and_add(base, tmp_path):
    old, xml = base
    doc = parse_contents(xml)
    row = next(r for r in doc.rows.values() if r.cells and r.cells[0].text.strip() == "매출채권")
    new_xml = splice.apply(xml, [splice.delete_row(xml, row)])
    new = _write(tmp_path, "new.dsd", new_xml)
    r = diff.compare(old, new)
    assert r.counts["행삭제"] == 1 and r.counts["행추가"] == 0
    # 반대 방향은 행추가
    r2 = diff.compare(new, old)
    assert r2.counts["행추가"] == 1 and r2.counts["행삭제"] == 0


def test_table_insert_does_not_cascade(base, tmp_path):
    """주석 2 앞에 표를 하나 끼워도(t### 전부 +1) 뒤 표들은 변경 0."""
    old, xml = base
    from fixtures.foot_factory import _amount_row, _td, table

    extra = table([[_td("구  분"), _td("금액")], _amount_row("신규항목", ["10"])], 2).encode("utf-8")
    marker = "<P>(단위 : 원)</P>".encode("utf-8")
    assert xml.count(marker) == 1
    new_xml = xml.replace(marker, marker + extra)
    new = _write(tmp_path, "new.dsd", new_xml)
    r = diff.compare(old, new)
    assert r.counts["표추가"] == 1
    assert r.counts["값변경"] == 0 and r.counts["텍스트변경"] == 0
    assert r.counts["표삭제"] == 0 and r.counts["행추가"] == 0
    changed = [s for s in r.sections if s.status != "same"]
    assert len(changed) == 1 and "주석 2" in changed[0].title


def test_note_renumber_is_flagged_not_text_change(base, tmp_path):
    old, xml = base
    new_xml = xml.replace("<P>3. 유형자산&amp;cr;".encode("utf-8"), "<P>4. 유형자산&amp;cr;".encode("utf-8"))
    assert new_xml != xml
    new = _write(tmp_path, "new.dsd", new_xml)
    r = diff.compare(old, new)
    assert r.counts["주석재번호"] == 1
    assert r.counts["문단변경"] == 0 and r.counts["값변경"] == 0
    sec = next(s for s in r.sections if s.renumber)
    assert sec.renumber == (3, 4)


def test_paragraph_text_change(base, tmp_path):
    old, xml = base
    new_xml = xml.replace("설립되었습니다.".encode("utf-8"), "설립된 회사입니다.".encode("utf-8"))
    new = _write(tmp_path, "new.dsd", new_xml)
    r = diff.compare(old, new)
    assert r.counts["문단변경"] == 1


def test_sample_self_diff(foot_sample_dsd):
    sample_dsd = foot_sample_dsd
    r = diff.compare(sample_dsd, sample_dsd)
    assert not r.has_changes
