"""DART 대사 — 식별·원문 보정·직전 보고서 선택(모의 서버)·정렬 대사·워크북 반영. 네트워크 불사용."""
from __future__ import annotations

import io
import json
import zipfile
from datetime import date
from pathlib import Path

import openpyxl
import pytest

from dsdxlsx.foot import darttie, dartsrc, pipeline, sections
from dsdxlsx.foot.tableview import build_views
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.parse import parse_contents


def _current_xml() -> bytes:
    """합성 표준 문서를 '제 3(당)기 / 제 2(전)기' 로 한 기수 밀고 회사코드·기간 TU 를 넣는다.
    전기 열 값은 그대로(제 2 기 당기값과 다름) → 대사 불일치가 나야 정상, 한 칸만 맞춰 둔다."""
    from fixtures.foot_factory import standard_contents

    xml = standard_contents()
    xml = xml.replace("제 2(당)기".encode(), "제 3(당)기".encode()).replace("제 1(전)기".encode(), "제 2(전)기".encode())
    # BS 유동자산 전기 1,100 → 1,300 (직전 문서 당기와 일치시켜 '일치' 1건 확보)
    xml = xml.replace(
        '<TD ALIGN="RIGHT"></TD><TD ALIGN="RIGHT">1,300</TD><TD ALIGN="RIGHT"></TD><TD ALIGN="RIGHT">1,100</TD>'.encode(),
        '<TD ALIGN="RIGHT"></TD><TD ALIGN="RIGHT">1,400</TD><TD ALIGN="RIGHT"></TD><TD ALIGN="RIGHT">1,300</TD>'.encode(), 1)
    head = ('<DOCUMENT-NAME ACODE="00760">감사보고서</DOCUMENT-NAME>'
            '<COMPANY-NAME AREGCIK="00123456">합성테스트</COMPANY-NAME>'
            '<TABLE><COLGROUP><COL/></COLGROUP><TBODY><TR>'
            '<TU AUNIT="PERIODFROM" AUNITVALUE="20260101">2026년 01월 01일</TU>'
            '<TU AUNIT="PERIODTO" AUNITVALUE="20261231">2026년 12월 31일</TU></TR></TBODY></TABLE>').encode()
    return xml.replace(b"<BODY>", b"<BODY>" + head, 1)


def _prior_xml() -> bytes:
    from fixtures.foot_factory import standard_contents
    xml = standard_contents()
    head = ('<DOCUMENT-NAME ACODE="00760">감사보고서</DOCUMENT-NAME>'
            '<COMPANY-NAME AREGCIK="00123456">합성테스트</COMPANY-NAME>').encode()
    return xml.replace(b"<BODY>", b"<BODY>" + head, 1)


_META = ('<?xml version="1.0" encoding="UTF-8"?><METAINFO>'
         '<DOCUMENT-HEADER regcik="00123456" regname="합성테스트"/></METAINFO>').encode()


def _write_dsd(path: Path, xml: bytes) -> None:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("contents.xml", xml)
        zf.writestr("meta.xml", _META)
    path.write_bytes(buf.getvalue())


def test_identify():
    ident = dartsrc.identify(_current_xml(), _META)
    assert ident.corp_code == "00123456" and ident.company == "합성테스트"
    assert ident.kind == "annual" and ident.fiscal_year == 2026 and not ident.consolidated
    assert ident.period_to == date(2026, 12, 31)


def test_identify_interim_and_consolidated():
    xml = _current_xml().replace("감사보고서".encode(), "분기검토보고서".encode(), 1)
    xml = xml.replace(b'AUNITVALUE="20261231"', b'AUNITVALUE="20260331"')
    ident = dartsrc.identify(xml, _META)
    assert ident.kind == "q1"
    xml2 = xml.replace('<TITLE ATOC="Y">재 무 상 태 표</TITLE>'.encode(),
                       '<TITLE ATOC="Y">(첨부)연 결 재 무 제 표</TITLE><TITLE ATOC="Y">재 무 상 태 표</TITLE>'.encode())
    assert dartsrc.identify(xml2, _META).consolidated


def test_sanitize_xml_loose_escapes():
    raw = b"<P>IT&\xeb\xaa\xa8 <\xec\x9d\xb4\xec\x82\xac \xeb\xb3\xb4\xec\x88\x98> &amp; &lt; &cr;</P>"
    fixed = dartsrc.sanitize_xml(raw)
    assert b"IT&amp;\xeb\xaa\xa8" in fixed and b"&lt;\xec\x9d\xb4" in fixed
    assert b"&amp;amp;" not in fixed and b"&amp;lt;" not in fixed
    assert b"&amp;cr;" in fixed  # DSD 관례와 같은 형태
    parse_contents(b"<BODY>" + fixed + b"</BODY>")


def test_fs_subdoc_picks_separate_vs_consolidated():
    from fixtures.foot_factory import _amount_row, _td, table
    t = table([[_td("과  목"), _td("제 2(당)기"), _td("제 1(전)기")], _amount_row("자산총계", ["1", "2"])], 3)
    xml = ("<BODY><TITLE ATOC=\"Y\">III. 재무에 관한 사항</TITLE>"
           "<TITLE ATOC=\"Y\">2. 연결재무제표</TITLE><P>연결</P>" + t +
           "<TITLE ATOC=\"Y\">4. 재무제표</TITLE><P>별도</P>" + t +
           "<TITLE ATOC=\"Y\">5. 재무제표 주석</TITLE><P>주석</P></BODY>").encode()
    doc = parse_contents(xml)
    sub = dartsrc.fs_subdoc(doc, consolidated=False)
    assert sub is not None and len(sub.tables) == 1
    assert sub.blocks[0].item.text.strip() == "4. 재무제표"
    sub2 = dartsrc.fs_subdoc(doc, consolidated=True)
    assert sub2.blocks[0].item.text.strip() == "2. 연결재무제표"


def _fake_fetch(prior_xml: bytes, filings: list[dict]):
    """가짜 OpenDART - DartClient 의 fetch(endpoint, params) 자리에 넣는다."""
    def fetch(endpoint: str, params: dict) -> bytes:
        if endpoint == "list.json":
            rows = [f for f in filings if f["_ty"] == params.get("pblntf_ty")]
            if not rows:
                return json.dumps({"status": "013", "message": "no data"}).encode()
            return json.dumps({"status": "000", "total_page": 1,
                               "list": [{k: v for k, v in f.items() if k != "_ty"} for f in rows]}).encode()
        if endpoint == "document.xml":
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w") as zf:
                zf.writestr("20260320000001.xml", '<DOCUMENT><DOCUMENT-NAME ACODE="11011">사업보고서</DOCUMENT-NAME><BODY></BODY></DOCUMENT>'.encode())
                zf.writestr("20260320000001_00760.xml", prior_xml)
            return buf.getvalue()
        raise dartsrc.DartError(f"없는 주소: {endpoint}")
    return fetch


@pytest.fixture()
def client(tmp_path):
    filings = [{"_ty": "A", "rcept_no": "20260320000001", "rcept_dt": "20260320",
                "report_nm": "사업보고서 (2025.12)", "corp_code": "00123456"}]
    return dartsrc.DartClient(key="test", cache_dir=tmp_path / "cache",
                              fetch=_fake_fetch(_prior_xml(), filings))


def test_find_prior_reports_via_attachment(client):
    ident = dartsrc.identify(_current_xml(), _META)
    priors = dartsrc.find_prior_reports(client, ident)
    assert len(priors) == 1 and priors[0].kind == "annual"
    assert "감사보고서" in priors[0].label and "20260320000001" in priors[0].label
    assert len(priors[0].doc.tables) >= 3


def test_run_and_apply_end_to_end(client, tmp_path):
    dsd = tmp_path / "cur.dsd"
    _write_dsd(dsd, _current_xml())
    xlsx = tmp_path / "out.xlsx"
    result = pipeline.process_file_full(dsd, xlsx)
    res = darttie.run(result.package, result.doc, result.plans, result.ctx, result.placed_rows, client=client)
    assert res.status == "complete", res.reason
    assert res.ties, "대사 0건"
    # BS 유동자산: 현재 전기 1,300 == 직전 당기 1,300 → 일치 / 자산총계: 1,600 vs 2,000 → 불일치
    by_label = {(t.sheet, t.label.replace(" ", "")): t for t in res.ties}
    assert by_label[("BS", "Ⅰ.유동자산")].ok
    assert not by_label[("BS", "자산총계")].ok
    assert by_label[("BS", "자산총계")].prior_value == 2000 and by_label[("BS", "자산총계")].cur_value == 1600
    assert res.n_bad >= 1
    darttie.apply(result.out, res, result.right_starts)
    wb = openpyxl.load_workbook(str(xlsx))
    assert darttie.DART_SHEET in wb.sheetnames
    ws = wb[darttie.DART_SHEET]
    vals = [r for r in ws.iter_rows(values_only=True) if r[0] == "BS"]
    assert vals and any(r[6] == "불일치" for r in vals)
    # BS 시트 오른쪽 슬롯에 DART 링크 + 수식
    bs = wb["BS"]
    links = [c for row in bs.iter_rows() for c in row
             if isinstance(c.value, str) and c.value.startswith("=HYPERLINK") and "DART:" in c.value]
    assert links
    f = bs.cell(links[0].row, links[0].column + 1).value
    assert isinstance(f, str) and f.startswith("=IF(") and darttie.DART_SHEET in f
    # 총괄표 크로스 카운터에 합산됨
    top = wb["총괄표"]
    assert any(isinstance(c.value, str) and "DART원천" not in c.value and "COUNTIF" in c.value
               for row in top.iter_rows() for c in row)


def test_run_skipped_without_key(tmp_path):
    dsd = tmp_path / "cur.dsd"
    _write_dsd(dsd, _current_xml())
    result = pipeline.process_file_full(dsd, tmp_path / "o.xlsx")
    res = darttie.run(result.package, result.doc, result.plans, result.ctx, result.placed_rows,
                      client=dartsrc.DartClient(key="", cache_dir=tmp_path / "c"))
    assert res.status == "skipped" and "OPENDART_API_KEY" in res.reason
    darttie.apply(result.out, res, result.right_starts)   # skipped 여도 시트에 상태만 남고 예외 없음
    wb = openpyxl.load_workbook(str(tmp_path / "o.xlsx"))
    assert darttie.DART_SHEET in wb.sheetnames


def test_cli_dart_tie_flag(client, tmp_path, monkeypatch, capsys):
    """foot --dart-tie: 불일치가 있으면 종료코드 1, 워크북에 DART원천 시트, 봉투에 요약."""
    from dsdxlsx.cli import main

    monkeypatch.setattr(darttie, "DartClient", lambda *a, **k: client)
    dsd = tmp_path / "cur.dsd"
    _write_dsd(dsd, _current_xml())
    out = tmp_path / "o.xlsx"
    assert main(["foot", str(dsd), "--dart-tie", "-o", str(out), "--json"]) == 1
    env = json.loads(capsys.readouterr().out)
    tie = env["files"][0]["dart"]
    assert tie["status"] == "complete" and tie["fail"] >= 1 and tie["sources"]
    assert any(m["kind"] == "dart" for m in env["files"][0]["mismatches"])
    assert darttie.DART_SHEET in openpyxl.load_workbook(str(out)).sheetnames


def test_cli_dart_tie_without_key_is_skipped(tmp_path, monkeypatch, capsys):
    from dsdxlsx.cli import main

    monkeypatch.setattr(darttie, "DartClient", lambda *a, **k: dartsrc.DartClient(key="", cache_dir=tmp_path / "c"))
    dsd = tmp_path / "cur.dsd"
    _write_dsd(dsd, _current_xml())
    rc = main(["foot", str(dsd), "--dart-tie", "-o", str(tmp_path / "o.xlsx"), "--json"])
    env = json.loads(capsys.readouterr().out)
    assert rc == env["files"][0]["exit_code"] and env["files"][0]["dart"]["status"] == "skipped"
    assert any("OPENDART_API_KEY" in w for w in env["warnings"])
