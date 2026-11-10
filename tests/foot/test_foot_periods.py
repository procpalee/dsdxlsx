"""반기·분기 보고서 기간 인식 — '당 반 기/전 반 기'·'당3분기'·'분기순이익' 변형."""
from fixtures.foot_factory import _amount_row, _td, build_dsd, table

from dsdxlsx.foot import crossref, sections
from dsdxlsx.foot.tableview import build_views
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.parse import parse_contents


def _views(xml_body: str, tmp_path):
    xml = ('<?xml version="1.0" encoding="UTF-8"?><BODY>' + xml_body + "</BODY>").encode()
    p = tmp_path / "q.dsd"
    p.write_bytes(build_dsd(xml))
    doc = parse_contents(read_dsd(str(p)).contents)
    plans = sections.build(doc)
    return build_views(plans)


def test_period_bucket_variants(tmp_path):
    body = ('<TITLE ATOC="Y">주석</TITLE>'
            "<P>1. 표</P>"
            + table([[_td("구 분"), _td("당 반 기"), _td("전 반 기"), _td("당3분기"),
                      _td("전 1분기"), _td("제 60(당) 분기"), _td("당분기말")],
                     _amount_row("가", ["1", "2", "3", "4", "5", "6"])], 7))
    views, _ = _views(body, tmp_path)
    tv = views[0]
    assert [tv.period_bucket(c) for c in range(1, 7)] == \
        ["당기", "전기", "당기", "전기", "당기", "당기"]


def test_caption_bucket_variants(tmp_path):
    body = ('<TITLE ATOC="Y">주석</TITLE>'
            "<P>1. 표</P><P>(1) 당반기말</P>"
            + table([[_td("구 분"), _td("금액")], _amount_row("가", ["1"])], 2)
            + "<P>(2) 전3분기말</P>"
            + table([[_td("구 분"), _td("금액")], _amount_row("나", ["1"])], 2))
    views, _ = _views(body, tmp_path)
    assert views[0].caption_bucket == "당기"
    assert views[1].caption_bucket == "전기"


def test_quarter_net_income_regex():
    for lab in ("당기순이익", "반기순이익(손실)", "분기순손익", "연결반기순이익"):
        assert crossref._NI_RE.search(lab), lab
    assert not crossref._NI_RE.search("영업이익")
