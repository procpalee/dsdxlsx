"""dart - DART 공시 감사보고서 → DSD. OpenDART 는 가짜 응답으로 대신한다(키·네트워크 불필요)."""
import io
import json
import zipfile

import pytest

from dsdxlsx import dart_client, dartdoc
from dsdxlsx.cli import main
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.export_xlsx import export_xlsx
from dsdxlsx.import_xlsx import import_xlsx
from dsdxlsx.parse import parse_contents
from fixtures import factory as F

CORP = "00999001"


def audit_xml(acode="00760", name="감사보고서", extra_p="") -> bytes:
    """공시 원문처럼: 문서 머리 + LF 줄끝 + 맨 '&' + 이미지 자리."""
    body = F.standard_contents().decode("utf-8")
    head = (f'<DOCUMENT-HEADER AEXT-CLASS="Y">\n<DOCUMENT-NAME ACODE="{acode}">{name}</DOCUMENT-NAME>\n'
            f'<COMPANY-NAME AREGCIK="{CORP}" AACCOUNTTYPE="A">샘플전자 주식회사</COMPANY-NAME>\n'
            "</DOCUMENT-HEADER>\n"
            '<TABLE><TBODY><TR><TU AUNIT="PERIODTO" AUNITVALUE="20251231">2025년 12월 31일</TU></TR>'
            "</TBODY></TABLE>\n"
            "<P>R&D 비용과 연구개발</P>\n"
            '<IMG WIDTH="600" HEIGHT="800">내부회계관리제도_보고서.jpg</IMG>\n' + extra_p)
    body = body.replace('xsi:noNamespaceSchemaLocation="dart4.xsd">',
                        'xsi:noNamespaceSchemaLocation="dart4.xsd">\r\n' + head, 1)
    return body.replace("\r\n", "\n").encode("utf-8")


def _zip(files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for n, d in files.items():
            zf.writestr(n, d)
    return buf.getvalue()


CORP_XML = _zip({"CORPCODE.xml": (
    "<result>"
    f"<list><corp_code>{CORP}</corp_code><corp_name>샘플전자</corp_name><stock_code> </stock_code></list>"
    "<list><corp_code>00999002</corp_code><corp_name>샘플전자서비스</corp_name><stock_code>123456</stock_code></list>"
    "</result>").encode("utf-8")})

DOCS = {
    # 외부감사관련: 감사보고서 자체
    "20260310000001": _zip({"20260310000001.xml": audit_xml()}),
    # 사업보고서: 본문 + 첨부 감사보고서·연결감사보고서
    "20260320000002": _zip({
        "20260320000002.xml": ('<?xml version="1.0"?><DOCUMENT><DOCUMENT-NAME ACODE="11011">'
                               "사업보고서</DOCUMENT-NAME></DOCUMENT>").encode("utf-8"),
        "20260320000002_00761.xml": audit_xml("00761", "연결감사보고서"),
    }),
}

LISTS = {
    "F": [{"rcept_no": "20260310000001", "rcept_dt": "20260310", "corp_name": "샘플전자",
           "report_nm": "감사보고서 (2025.12)"}],
    "A": [{"rcept_no": "20260320000002", "rcept_dt": "20260320", "corp_name": "샘플전자",
           "report_nm": "사업보고서 (2025.12)"}],
}


@pytest.fixture
def fake_dart(monkeypatch, tmp_path):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "appdata"))
    calls = []

    def fake_get(endpoint, **params):
        calls.append((endpoint, params))
        if endpoint == "corpCode.xml":
            return CORP_XML
        if endpoint == "list.json":
            assert params["corp_code"] == CORP
            items = LISTS.get(params["pblntf_ty"], [])
            return json.dumps({"status": "000" if items else "013", "total_page": 1,
                               "list": items}, ensure_ascii=False).encode("utf-8")
        if endpoint == "document.xml":
            raw = DOCS.get(params["rcept_no"])
            return raw or "<result><status>014</status><message>파일이 존재하지 않습니다.</message></result>".encode()
        raise AssertionError(endpoint)

    monkeypatch.setattr(dartdoc, "_get", fake_get)
    return calls


def test_build_dsd_fixes_and_meta(tmp_path):
    b = dartdoc.build_dsd(audit_xml())
    assert b.acode == "00760" and b.corp_code == CORP and b.company == "샘플전자 주식회사"
    assert b.images == ["내부회계관리제도_보고서.jpg"]
    assert b.fixes["escaped"] == 1 and b.fixes["crlf"] > 0
    assert b"R&amp;D" in b.contents and b"\n" not in b.contents.replace(b"\r\n", b"")
    assert b'doc-id="00760"' in b.meta and f'regcik="{CORP}"'.encode() in b.meta
    doc = parse_contents(b.contents)
    assert doc.tables
    lf = dartdoc.build_dsd(audit_xml(), crlf=False)
    assert "crlf" not in lf.fixes and b"\r\n" not in lf.contents.split(b">", 2)[2][:200]


def test_built_dsd_roundtrips_byte_identical(tmp_path):
    p = tmp_path / "a.dsd"
    dartdoc.write_dsd(str(p), dartdoc.build_dsd(audit_xml()))
    pkg = read_dsd(str(p))
    x = tmp_path / "a.xlsx"
    export_xlsx(parse_contents(pkg.contents), p.read_bytes(), str(p), str(x))
    pkg2, new, rep = import_xlsx(str(x))
    assert not rep.changed and new == pkg.contents


def test_fetch_by_name_prefers_audit_filing(fake_dart):
    got = dartdoc.fetch("샘플전자")
    assert got.filing.rcept_no == "20260310000001" and got.part.acode == "00760"


def test_fetch_consolidated_from_business_report_attachment(fake_dart):
    got = dartdoc.fetch("샘플전자", form="연결감사", year=2025)
    assert got.filing.rcept_no == "20260320000002"
    assert got.part.filename.endswith("_00761.xml") and got.built.acode == "00761"


def test_fetch_by_dart_url(fake_dart):
    got = dartdoc.fetch("https://dart.fss.or.kr/dsaf001/main.do?rcpNo=20260310000001")
    assert got.filing is None and got.part.acode == "00760"
    assert not any(e == "list.json" for e, _ in fake_dart)


def test_fetch_errors_name_the_problem(fake_dart):
    with pytest.raises(dartdoc.DartError, match="비슷한 회사"):
        dartdoc.fetch("샘플")
    with pytest.raises(dartdoc.DartError, match="2019년"):
        dartdoc.fetch("샘플전자", year=2019)
    with pytest.raises(dartdoc.DartError, match="연결감사보고서"):
        dartdoc.fetch("20260310000001", form="연결감사")


def test_document_is_cached(fake_dart):
    dartdoc.document("20260310000001")
    dartdoc.document("20260310000001")
    assert sum(1 for e, _ in fake_dart if e == "document.xml") == 1


def test_cli_dart_xlsx_rollforward_json(fake_dart, tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    assert main(["dart", "샘플전자", "--xlsx", "--rollforward", "--json"]) == 0
    env = json.loads(capsys.readouterr().out)
    assert env["command"] == "dart" and env["ok"] and env["year"] == 2025
    assert [o.rsplit(".", 1)[-1] for o in env["outputs"]] == ["dsd", "xlsx", "dsd", "xlsx"]
    for o in env["outputs"]:
        assert (tmp_path / o).is_file()
    assert env["source"]["rcept_no"] == "20260310000001" and env["images"]
    assert env["rollforward"]["자동확정"] >= 1
    # 다시 돌리면 덮어쓰지 않고 멈춤(3)
    assert main(["dart", "샘플전자", "--json"]) == 3
    assert "이미 있는 파일" in json.loads(capsys.readouterr().out)["errors"][0]


def test_cli_dart_list(fake_dart, capsys):
    assert main(["dart", "샘플전자", "--list", "--json"]) == 0
    env = json.loads(capsys.readouterr().out)
    assert env["corp"]["code"] == CORP
    # 같은 사업연도면 외부감사관련 공시 먼저, 사업보고서(첨부 후보)는 그다음
    assert [(f["rcept_no"], f["kind"]) for f in env["filings"]] == [
        ("20260310000001", "F"), ("20260320000002", "A")]


def test_cli_dart_without_key_is_input_error(monkeypatch, tmp_path, capsys):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "appdata"))
    monkeypatch.setattr(dart_client, "api_key", lambda: None)
    assert main(["dart", "샘플전자", "--json"]) == 2
    assert "OPENDART_API_KEY" in json.loads(capsys.readouterr().out)["errors"][0]
