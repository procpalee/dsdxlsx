from decimal import Decimal

from dsdxlsx.foot import footing, sections
from dsdxlsx.foot.tableview import build_views


def _specs(doc):
    plans = sections.build(doc)
    views, ctx = build_views(plans)
    return footing.analyze(views), ctx


def _by(specs, table_id=None, kind=None):
    return [s for s in specs
            if (table_id is None or s.table_id == table_id)
            and (kind is None or s.kind == kind)]


def test_bs_group_tree(synth_doc):
    """유동자산=현금+매출채권, 자산총계=Ⅰ+Ⅱ, 부채및자본총계 — 전부 정확 일치."""
    _, doc = synth_doc
    specs, ctx = _specs(doc)
    bs_tables = [tv.table.id for tv in ctx["statements"] if tv.sheet == "BS"]
    bs = [s for s in specs if s.table_id in bs_tables]
    assert len(bs) >= 12  # 그룹 5~6종 × 당기/전기
    assert all(s.ok for s in bs)
    assert all(s.diff == 0 for s in bs)


def test_pl_anchor_and_rollup(synth_doc):
    _, doc = synth_doc
    specs, ctx = _specs(doc)
    pl_tables = [tv.table.id for tv in ctx["statements"] if tv.sheet == "PL"]
    anchors = [s for s in specs if s.table_id in pl_tables and s.kind == "pl_anchor"]
    rollups = [s for s in specs if s.table_id in pl_tables and s.kind == "rollup"]
    # 매출총이익·영업이익·당기순이익 × 2기
    assert len(anchors) >= 6
    assert all(s.ok for s in anchors)
    # 매출액=제품매출, 판관비 합계 등
    assert len(rollups) >= 4
    assert all(s.ok for s in rollups)


def test_note_total_row_and_flows(synth_doc):
    _, doc = synth_doc
    specs, ctx = _specs(doc)
    note3 = [tv.table.id for tvs in [ctx["notes"][3]] for tv in tvs]
    flows = [s for s in specs if s.table_id in note3 and s.kind == "flow_h"]
    assets = [s for s in specs if s.table_id in note3 and s.kind == "asset_cost"]
    totals = [s for s in specs if s.table_id in note3 and s.kind == "total_row"]
    assert flows and all(s.ok for s in flows)      # 기말 = 기초+취득+감가상각비
    assert assets and all(s.ok for s in assets)    # 기말 = 취득원가+누계액
    assert totals and all(s.ok for s in totals)    # 합계행 = Σ
    # 같은 대상 셀(기말×합계행)에 세로·가로 검증 공존
    keys = {(s.target, s.axis) for s in specs if s.table_id in note3}
    end_targets = {t for t, a in keys}
    assert any(sum(1 for t2, a in keys if t2 == t) >= 2 for t in end_targets)


def test_error_injection(tmp_path):
    """금액 변조 → 해당 검증이 FAIL로 뒤집히고 차이가 정확히 잡힌다."""
    from fixtures.foot_factory import standard_contents, build_dsd

    from dsdxlsx.dsd_io import read_dsd
    from dsdxlsx.parse import parse_contents

    broken = standard_contents().replace("1,300".encode(), "1,290".encode(), 1)
    p = tmp_path / "broken.dsd"
    p.write_bytes(build_dsd(broken))
    doc = parse_contents(read_dsd(str(p)).contents)
    specs, _ = _specs(doc)
    bad = [s for s in specs if not s.ok]
    assert bad, "변조된 유동자산 합계가 FAIL로 잡혀야 한다"
    # 유동자산 그룹(재무제표 본문 — 확실형)이 -10 차이를 그대로 노출
    assert any(s.diff == Decimal(-10) for s in bad)
