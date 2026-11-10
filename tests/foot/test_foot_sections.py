from dsdxlsx.foot import sections


def test_synth_sheet_plan(synth_doc):
    _, doc = synth_doc
    plans = sections.build(doc)
    names = [p.name for p in plans]
    assert names == ["BS", "PL", "CF", "1", "2", "3"]
    kinds = {p.name: p.kind for p in plans}
    assert kinds["BS"] == "statement" and kinds["2"] == "note"
    titles = {p.name: p.display_title for p in plans}
    assert titles["BS"] == "재 무 상 태 표"
    assert titles["2"] == "현금및현금성자산"


def test_note_sheet_map(synth_doc):
    _, doc = synth_doc
    plans = sections.build(doc)
    assert sections.note_sheet_map(plans) == {1: "1", 2: "2", 3: "3"}


def test_company_name(synth_doc):
    package, _ = synth_doc
    assert sections.company_name(package.meta) == "합성테스트"


def test_positional_note_names_with_gap(tmp_path):
    """원문 주석 번호에 갭(1,3,4)이 있어도 시트명은 순서(1,2,3) — 기준 양식 규칙(실샘플 실측)."""
    from fixtures.foot_factory import build_dsd

    from dsdxlsx.dsd_io import read_dsd
    from dsdxlsx.parse import parse_contents

    xml = (
        '<?xml version="1.0" encoding="UTF-8"?><BODY>'
        '<TITLE ATOC="Y">주석</TITLE>'
        "<P>1. 회사의 개요&amp;cr;내용A</P>"
        "<P>3. 예치금&amp;cr;내용B</P>"
        "<P>4. 유가증권&amp;cr;내용C</P>"
        "</BODY>"
    ).encode("utf-8")
    p = tmp_path / "gap.dsd"
    p.write_bytes(build_dsd(xml))
    doc = parse_contents(read_dsd(str(p)).contents)
    plans = sections.build(doc)
    assert [pl.name for pl in plans] == ["1", "2", "3"]
    assert sections.note_sheet_map(plans) == {1: "1", 3: "2", 4: "3"}


def test_appendix_note_numbers_restart_without_overwriting_refs(tmp_path):
    """별첨 안의 1..N은 후속 시트로 이어지되 정규 '(주석N)' 매핑을 덮지 않는다."""
    from fixtures.foot_factory import build_dsd

    from dsdxlsx.foot import layout
    from dsdxlsx.dsd_io import read_dsd
    from dsdxlsx.parse import parse_contents

    xml = (
        '<?xml version="1.0" encoding="UTF-8"?><BODY>'
        '<TITLE ATOC="Y">주석</TITLE>'
        '<P>1. 회사의 개요&amp;cr;정규 주석 1</P>'
        '<P>2. 현금&amp;cr;정규 주석 2</P>'
        '<P>[별첨1]&amp;cr;감사 기본요구사항&amp;cr;'
        '1. 회계관리에 대한 사항&amp;cr;별첨 내용 1</P>'
        '<P>2. 재무제표 분석&amp;cr;별첨 내용 2</P>'
        '<P>[별첨2]&amp;cr;후속 표 설명</P>'
        '</BODY>'
    ).encode("utf-8")
    p = tmp_path / "appendix.dsd"
    p.write_bytes(build_dsd(xml))
    doc = parse_contents(read_dsd(str(p)).contents)
    plans = sections.build(doc)

    assert [pl.name for pl in plans] == ["1", "2", "3", "4"]
    assert [pl.display_title for pl in plans] == [
        "회사의 개요", "현금", "회계관리에 대한 사항", "재무제표 분석",
    ]
    assert sections.note_sheet_map(plans) == {1: "1", 2: "2"}

    items = layout.compute_render_items(plans)
    texts3 = [x[1] for x in items["3"] if x[0] == "para"]
    texts4 = [x[1] for x in items["4"] if x[0] == "para"]
    assert texts3 == ["별첨 내용 1"]
    assert texts4 == ["별첨 내용 2", "[별첨2]\n후속 표 설명"]
