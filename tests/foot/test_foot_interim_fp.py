"""분기 DSD 실측 오탐 회귀 — 문서 종류(기간 월 우선)·자본변동표 소계·지분율 산식·DART 행 짝.

모두 합성 DSD 조각이다(실데이터 없음)."""
from datetime import date

from fixtures.foot_factory import _amount_row, _td, build_dsd, table

from dsdxlsx.foot import footing, sections
from dsdxlsx.foot.tableview import build_views
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.parse import parse_contents


def _views(tmp_path, body_xml: str, name: str = "t.dsd"):
    xml = ('<?xml version="1.0" encoding="UTF-8"?><BODY>' + body_xml + "</BODY>").encode()
    p = tmp_path / name
    p.write_bytes(build_dsd(xml))
    doc = parse_contents(read_dsd(str(p)).contents)
    views, ctx = build_views(sections.build(doc))
    return views, ctx


def test_socie_subtotal_not_double_counted(tmp_path):
    """총포괄손익 아래 같은 깊이의 '소계' 행 — 기말 = 기초 + 소계 + 소유주 거래 (오탐 없음)."""
    ce = table([
        [_td("과목"), _td("자본금"), _td("이익잉여금"), _td("합계")],
        _amount_row("2026년 1월 1일(당기초)", ["500", "100", "600"]),
        _amount_row("총포괄손익 :", ["", "", ""]),
        _amount_row("분기순이익(손실)", ["0", "(70)", "(70)"], 3),
        _amount_row("확정급여제도의 재측정요소", ["0", "(5)", "(5)"], 3),
        _amount_row("소   계", ["0", "(75)", "(75)"], 3),
        _amount_row("자본에 직접 반영된 소유주와의 거래 :", ["", "", ""]),
        _amount_row("연차배당", ["0", "(10)", "(10)"], 3),
        _amount_row("전환권행사", ["20", "0", "20"], 3),
        _amount_row("2026년 9월 30일(당분기말)", ["520", "15", "535"]),
    ], 4)
    views, _ = _views(tmp_path, '<TITLE ATOC="Y">자 본 변 동 표</TITLE>' + ce)
    specs = footing.analyze(views)
    socie = [s for s in specs if s.kind == "socie" and "9월30일" in s.label.replace(" ", "")]
    assert len(socie) >= 3 and all(s.ok for s in socie)
    # 소계와 그 구성 행을 함께 더하지 않는다
    for s in socie:
        rows = {j for j, _, _ in s.terms}
        assert not ({3, 4} & rows and 5 in rows), s.terms


def test_ownership_ratio_needs_share_column(tmp_path):
    """종속기업 투자 표(지분율 · 장부금액)는 지분율 = 장부금액 ÷ 합계 식을 세우지 않는다."""
    note = (
        "<P>5. 종속기업투자&amp;cr;종속기업 투자 내역입니다.</P>"
        + table([
            [_td("회사명"), _td("소재지"), _td("당분기말 지분율"), _td("당분기말 장부금액"),
             _td("전기말 지분율"), _td("전기말 장부금액")],
            _amount_row("가회사", ["한국", "100.00", "3,000", "100.00", "3,000"]),
            _amount_row("나회사", ["미국", "60.00", "1,000", "60.00", "1,000"]),
            _amount_row("합  계", ["", "", "4,000", "", "4,000"]),
        ], 6)
        + "<P>주주 현황입니다.</P>"
        + table([
            [_td("주주"), _td("소유주식수"), _td("장부금액"), _td("지분율")],
            _amount_row("갑", ["600", "9,000", "60.0"]),
            _amount_row("을", ["400", "1,000", "40.0"]),
            _amount_row("합계", ["1,000", "10,000", "100.0"]),
        ], 4)
    )
    views, _ = _views(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    specs = footing.analyze(views)
    ratios = [s for s in specs if s.kind == "ratio"]
    sub_tid = views[0].table.id if "회사명" in views[0].col_header(0) else views[1].table.id
    assert not [s for s in ratios if s.table_id == sub_tid]
    # 주식수 열이 있으면 금액 열을 건너뛰고 주식수로 식을 세운다 (모두 일치)
    assert len(ratios) == 3 and all(s.ok for s in ratios)
