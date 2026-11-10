from dsdxlsx.foot import crossref, sections
from dsdxlsx.foot.tableview import build_views


def test_fs_note_pair_on_synth(synth_doc):
    """주석2 현금 합계행 ↔ BS 현금및현금성자산(주석2) 대사쌍이 잡힌다."""
    _, doc = synth_doc
    plans = sections.build(doc)
    views, ctx = build_views(plans)
    specs = crossref.analyze(views, ctx, sections.note_sheet_map(plans))
    note2_tables = {tv.table.id for tv in ctx["notes"][2]}
    pairs = [s for s in specs if s.dst_table in note2_tables]
    assert pairs, "주석2 합계행의 재무제표 대사쌍이 있어야 한다"
    assert any("현금및현금성자산" in s.label_path for s in pairs)


def test_cf_statement_pairs(synth_doc):
    """CF 당기순이익 ↔ PL, CF 기초·기말 현금 ↔ BS 현금및현금성자산 (당기/전기 대칭).

    사용자 보고 사례의 회귀 방지 — 재무제표 간 대사가 빠지지 않아야 한다."""
    _, doc = synth_doc
    plans = sections.build(doc)
    views, ctx = build_views(plans)
    specs = crossref.analyze(views, ctx, sections.note_sheet_map(plans))
    cf_tables = {tv.table.id for tv in ctx["statements"] if tv.sheet == "CF"}
    cf_pairs = [s for s in specs if s.dst_table in cf_tables]
    ni = [s for s in cf_pairs if "당기순이익" in s.label_path and s.label_path.startswith("PL:")]
    cash = [s for s in cf_pairs if "현금및현금성자산" in s.label_path
            and s.label_path.startswith("BS:")]
    assert len(ni) >= 2, "CF 당기순이익 ↔ PL 당기/전기 대사가 있어야 한다"
    assert len(cash) >= 3, "CF 기초·기말 현금 ↔ BS 현금 대사(기초당기→전기 포함)가 있어야 한다"


def test_fs_index_ref_inheritance(synth_doc):
    """그룹 행의 (주석n) 참조가 들여쓰기 자식에게 상속된다."""
    _, doc = synth_doc
    plans = sections.build(doc)
    views, ctx = build_views(plans)
    fs_rows = crossref.build_fs_index(ctx["statements"])
    ref3 = [fr.label for fr in fs_rows if 3 in fr.refs]
    assert any("유형자산" in lb for lb in ref3)


def test_cross_formula_shape(synth_doc):
    """대사 수식이 기준 형태(ABS 비교·시트 접두)로 렌더된다."""
    from dsdxlsx.foot import emit, layout, pipeline
    _, doc = synth_doc
    plans, views, ctx, fspecs, xspecs = pipeline.analyze(doc)
    by_table = {}
    for s in fspecs:
        by_table.setdefault(s.table_id, []).append(s)
    layouts, placed = layout.build(plans, by_table)
    cs = layout.bind(layouts, placed, fspecs, xspecs,
                     lambda t, i, c: emit.written_value(ctx["by_table"][t], i, c))
    assert cs.cross
    f = cs.cross[0].compare_formula()
    assert f.startswith("=IF(ABS(") and ",TRUE,ABS(" in f
    link = cs.cross[0].hyperlink_formula()
    assert link.startswith('=HYPERLINK("#')


def test_short_label():
    """대사 링크 표시 축약 — '소스시트 > 리프(유동성)' 형태."""
    from dsdxlsx.foot.checks import short_label

    assert short_label("BS:Ⅱ.비유동자산 -> 파생상품자산") == "BS > 파생상품자산(비유동)"
    assert short_label("BS:유동부채 -> 기타유동금융부채") == "BS > 기타유동금융부채"
    assert short_label("CF:영업활동현금흐름") == "CF > 영업활동현금흐름"
    assert short_label("BS:Ⅰ.유동자산 -> 재고자산(주석5)") == "BS > 재고자산(유동)"
    assert short_label("14:이익잉여금 -> 임의적립금") == "주석14 > 임의적립금"
    combo = short_label("BS:비유동자산 -> 구축물, BS:비유동자산 -> 감가상각누계액")
    assert combo.startswith("BS > 구축물") and "(" in combo
