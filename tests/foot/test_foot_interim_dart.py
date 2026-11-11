"""분기 DSD 실측 오탐 회귀 — 문서 종류(기간 월 우선)·자본변동표 소계·지분율 산식·DART 행 짝.

모두 합성 DSD 조각이다(실데이터 없음)."""
from datetime import date

from fixtures.foot_factory import _amount_row, _td, build_dsd, table

from dsdxlsx.foot import darttie, dartsrc, sections
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


def _head(doc_name: str, period_to: str) -> bytes:
    return ('<DOCUMENT-NAME ACODE="00760">' + doc_name + '</DOCUMENT-NAME>'
            '<COMPANY-NAME AREGCIK="00123456">합성</COMPANY-NAME>'
            '<TU AUNIT="PERIODTO" AUNITVALUE="' + period_to + '">x</TU>').encode()


def test_doc_kind_follows_period_month():
    """반기 문서를 뼈대로 만든 3분기 DSD: 문서명이 '반기검토보고서'여도 9월 말이면 q3."""
    assert dartsrc.identify(_head("반기검토보고서", "20260930"), None).kind == "q3"
    assert dartsrc.identify(_head("반기검토보고서", "20260331"), None).kind == "q1"
    assert dartsrc.identify(_head("분기검토보고서", "20260630"), None).kind == "half"
    assert dartsrc.identify(_head("반기검토보고서", "20260630"), None).kind == "half"
    ident = dartsrc.identify(_head("감사보고서", "20261231"), None)
    assert ident.kind == "annual" and ident.period_to == date(2026, 12, 31)


def _cf_table(rows: list[tuple[str, str]], head: str) -> str:
    return table([[_td("과목"), _td(head)]] + [_amount_row(lb, [v], 1) for lb, v in rows], 2)


def test_dart_pairs_repaired_by_value(tmp_path):
    """DART 행 이름이 달라 한 줄 밀린 짝: 값이 유일하게 같은 현재 행으로 다시 맞춘다."""
    cur, _ = _views(tmp_path, '<TITLE ATOC="Y">현 금 흐 름 표</TITLE>' + _cf_table([
        ("단기금융자산의 증가", "(300)"),
        ("장기금융상품의 취득", "(450)"),
        ("유형자산의 취득", "(120)"),
    ], "전분기"), "cur.dsd")
    pri, _ = _views(tmp_path, '<TITLE ATOC="Y">현 금 흐 름 표</TITLE>' + _cf_table([
        ("단기금융상품의 증가", "(300)"),
        ("장기금융상품의 취득", "(450)"),
        ("유형자산의 취득", "(120)"),
    ], "당분기"), "pri.dsd")
    tv_b, tv_a = cur[0], pri[0]
    rows_b = [i for i in tv_b.body if tv_b.value(i, 1) is not None]
    rows_a = [i for i in tv_a.body if tv_a.value(i, 1) is not None]
    # 정렬이 한 줄 밀린 상태: DART '장기금융상품의 취득' ↔ 현재 '단기금융자산의 증가' 등
    shifted = [(rows_a[1], rows_b[0]), (rows_a[2], rows_b[1])]
    fixed = darttie._repair_pairs(tv_a, tv_b, 1, 1, shifted)
    assert (rows_a[1], rows_b[1]) in fixed and (rows_a[2], rows_b[2]) in fixed
    assert all(tv_a.value(ia, 1) == tv_b.value(ib, 1) for ia, ib in fixed)
    # 맞는 짝은 그대로
    good = [(rows_a[k], rows_b[k]) for k in range(3)]
    assert darttie._repair_pairs(tv_a, tv_b, 1, 1, good) == good


def test_dart_pairs_ambiguous_value_not_moved(tmp_path):
    """같은 값의 현재 행이 둘 이상이면 옮기지 않는다(불일치 그대로 보고)."""
    cur, _ = _views(tmp_path, '<TITLE ATOC="Y">현 금 흐 름 표</TITLE>' + _cf_table([
        ("가", "(10)"), ("나", "(50)"), ("다", "(50)"),
    ], "전분기"), "cur.dsd")
    pri, _ = _views(tmp_path, '<TITLE ATOC="Y">현 금 흐 름 표</TITLE>' + _cf_table([
        ("라", "(50)"),
    ], "당분기"), "pri.dsd")
    tv_b, tv_a = cur[0], pri[0]
    rb = [i for i in tv_b.body if tv_b.value(i, 1) is not None]
    ra = [i for i in tv_a.body if tv_a.value(i, 1) is not None]
    assert darttie._repair_pairs(tv_a, tv_b, 1, 1, [(ra[0], rb[0])]) == [(ra[0], rb[0])]
