"""합성 미니 DSD 팩토리 — CI용 (고객 데이터 불사용).

K-GAAP풍 재무상태표 + 손익계산서 + 주석(현금 표·유형자산 흐름표·처분계산서)을
담은 contents.xml/meta.xml을 조립한다. 금액은 전부 가공 값.
"""
from __future__ import annotations

import io
import zipfile


def _td(text, **attrs) -> str:
    a = "".join(f' {k}="{v}"' for k, v in attrs.items())
    return f"<TD{a}>{text}</TD>"


def _tr(cells: list[str]) -> str:
    return "<TR>" + "".join(cells) + "</TR>"


def table(rows: list[list[str]], ncols: int, head_rows: int = 1,
          border: str = "1", aclass: str = "NORMAL") -> str:
    cols = "<COLGROUP>" + "<COL/>" * ncols + "</COLGROUP>"
    trs = [_tr(r) for r in rows]
    head = "<THEAD>" + "".join(trs[:head_rows]) + "</THEAD>" if head_rows else ""
    body = "<TBODY>" + "".join(trs[head_rows:]) + "</TBODY>"
    return f'<TABLE ACLASS="{aclass}" BORDER="{border}" WIDTH="600">{cols}{head}{body}</TABLE>'


def _amount_row(label: str, vals: list[str], indent: int = 0) -> list[str]:
    cells = [_td(" " * indent + label)]
    cells += [_td(v, ALIGN="RIGHT") for v in vals]
    return cells


def standard_contents() -> bytes:
    """재무상태표(2단 열) + 손익계산서 + 주석 1~3."""
    bs_head = [
        [_td("과  목"), _td("제 2(당)기말", COLSPAN=2), _td("제 1(전)기말", COLSPAN=2)],
    ]
    bs_rows = bs_head + [
        _amount_row("자      산", ["", "", "", ""]),
        _amount_row("Ⅰ. 유동자산", ["", "1,300", "", "1,100"]),
        _amount_row("현금및현금성자산(주석2)", ["1,000", "", "900", ""], 1),
        _amount_row("매출채권", ["300", "", "200", ""], 1),
        _amount_row("Ⅱ. 비유동자산", ["", "700", "", "500"]),
        _amount_row("유형자산(주석3)", ["700", "", "500", ""], 1),
        _amount_row("자  산  총  계", ["", "2,000", "", "1,600"]),
        _amount_row("부      채", ["", "", "", ""]),
        _amount_row("Ⅰ. 유동부채", ["", "800", "", "700"]),
        _amount_row("매입채무", ["800", "", "700", ""], 1),
        _amount_row("부  채  총  계", ["", "800", "", "700"]),
        _amount_row("자      본", ["", "", "", ""]),
        _amount_row("Ⅰ. 자본금", ["", "500", "", "500"]),
        _amount_row("보통주자본금", ["500", "", "500", ""], 1),
        _amount_row("Ⅱ. 이익잉여금", ["", "700", "", "400"]),
        _amount_row("미처분이익잉여금", ["700", "", "400", ""], 1),
        _amount_row("자  본  총  계", ["", "1,200", "", "900"]),
        _amount_row("부채및자본총계", ["", "2,000", "", "1,600"]),
    ]
    pl_rows = [
        [_td("과  목"), _td("제 2(당)기", COLSPAN=2), _td("제 1(전)기", COLSPAN=2)],
        _amount_row("Ⅰ. 매출액", ["", "5,000", "", "4,000"]),
        _amount_row("제품매출", ["5,000", "", "4,000", ""], 1),
        _amount_row("Ⅱ. 매출원가", ["", "3,000", "", "2,500"]),
        _amount_row("제품매출원가", ["3,000", "", "2,500", ""], 1),
        _amount_row("Ⅲ. 매출총이익", ["", "2,000", "", "1,500"]),
        _amount_row("Ⅳ. 판매비와관리비", ["", "1,200", "", "1,000"]),
        _amount_row("급여", ["900", "", "800", ""], 1),
        _amount_row("감가상각비", ["300", "", "200", ""], 1),
        _amount_row("Ⅴ. 영업이익", ["", "800", "", "500"]),
        _amount_row("Ⅵ. 법인세등", ["", "100", "", "100"]),
        _amount_row("Ⅶ. 당기순이익", ["", "700", "", "400"]),
    ]
    cf_rows = [
        [_td("과  목"), _td("제 2(당)기", COLSPAN=2), _td("제 1(전)기", COLSPAN=2)],
        _amount_row("Ⅰ. 영업활동으로 인한 현금흐름", ["", "100", "", "80"]),
        _amount_row("(1) 당기순이익", ["700", "", "400", ""], 1),
        _amount_row("(2) 자산부채의 변동", ["(600)", "", "(320)", ""], 1),
        _amount_row("Ⅴ. 기초의 현금", ["", "900", "", "820"]),
        _amount_row("Ⅵ. 기말의 현금(Ⅰ+Ⅴ)", ["", "1,000", "", "900"]),
    ]
    note2 = table([
        [_td("구  분"), _td("당기말"), _td("전기말")],
        _amount_row("보통예금", ["600", "500"]),
        _amount_row("정기예금", ["400", "400"]),
        _amount_row("합  계", ["1,000", "900"]),
    ], 3)
    note3 = table([
        [_td("구  분"), _td("기초"), _td("취득"), _td("감가상각비"), _td("기말"),
         _td("취득원가"), _td("감가상각누계액")],
        _amount_row("기계장치", ["400", "400", "(250)", "550", "950", "(400)"]),
        _amount_row("차량운반구", ["100", "100", "(50)", "150", "250", "(100)"]),
        _amount_row("합  계", ["500", "500", "(300)", "700", "1,200", "(500)"]),
    ], 7)
    body = (
        '<TITLE ATOC="Y">재 무 상 태 표</TITLE>'
        + table([[_td("제 2(당)기말 : 20X2년 12월 31일 현재", COLSPAN=2)],
                 [_td("합성테스트 주식회사"), _td("(단위 : 원)")]],
                2, head_rows=0, border="0", aclass="EXTRACTION")
        + table(bs_rows, 5, head_rows=1, aclass="FINANCE")
        + '<TITLE ATOC="Y">손 익 계 산 서</TITLE>'
        + table(pl_rows, 5, head_rows=1, aclass="FINANCE")
        + '<TITLE ATOC="Y">현 금 흐 름 표</TITLE>'
        + table(cf_rows, 5, head_rows=1, aclass="FINANCE")
        + '<TITLE ATOC="Y">주석</TITLE>'
        + "<P>1. 회사의 개요&amp;cr;&amp;cr;합성테스트 주식회사는 테스트를 위해 설립되었습니다.</P>"
        + "<P>2. 현금및현금성자산&amp;cr;&amp;cr;현금및현금성자산의 내역은 다음과 같습니다.</P>"
        + "<P>(단위 : 원)</P>"
        + note2
        + "<P>3. 유형자산&amp;cr;&amp;cr;유형자산의 변동 내역은 다음과 같습니다(단위: 원).</P>"
        + note3
    )
    xml = f'<?xml version="1.0" encoding="UTF-8"?><BODY>{body}</BODY>'
    return xml.encode("utf-8")


def meta() -> bytes:
    return ('<?xml version="1.0" encoding="UTF-8"?>'
            '<META><DOCUMENT-HEADER regname="합성테스트"/></META>').encode("utf-8")


def build_dsd(contents: bytes | None = None) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("contents.xml", contents or standard_contents())
        zf.writestr("meta.xml", meta())
    return buf.getvalue()


def write_dsd(path, contents: bytes | None = None) -> None:
    with open(path, "wb") as f:
        f.write(build_dsd(contents))
