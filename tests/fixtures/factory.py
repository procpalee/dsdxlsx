"""합성 미니 DSD 팩토리 — 고객 데이터 없이 CI에서 도는 픽스처.

실측 관례(로컬 샘플)의 부분집합을 흉내 낸다: 셀 개행 토큰 &cr;(디스크에는
&amp;cr;), TR 블록 사이 빈 줄(\r\n\r\n), <TABLE 여는 쪽 앞에만 빈 줄.
목적은 check/ties/profile 엔진 테스트 — DART 제출용 정본이 아니다.
"""
from __future__ import annotations

import io
import zipfile

CRLF = "\r\n"


def esc(text: str) -> str:
    out = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return out.replace("\n", "&amp;cr;")


def cell(text: str, tag: str = "TD", **attrs) -> str:
    if "ALIGN" not in attrs:
        stripped = text.strip()
        numeric = stripped == "-" or bool(
            stripped and stripped.strip("()").replace(",", "").replace(".", "").isdigit()
        )
        attrs["ALIGN"] = "RIGHT" if numeric else "LEFT"
    a = "".join(f' {k}="{v}"' for k, v in attrs.items())
    return f"<{tag}{a}>{esc(text)}</{tag}>" if text else f"<{tag}{a}/>"


def tr(cells: list[str], height: int = 30) -> str:
    return (f'<TR ACOPY="Y" ADELETE="Y" HEIGHT="{height}">' + CRLF
            + CRLF.join(cells) + CRLF + "</TR>")


def table(head_rows: list[list[str]], body_rows: list[list[str]],
          ncols: int | None = None, unit: str | None = None, width: int = 600) -> str:
    """head_rows/body_rows = 행별 셀 텍스트. 셀 텍스트에 '||' 접두로 속성 지정 가능
    (예: '||COLSPAN=4||(단위 : 원)'). unit 지정 시 전폭 단위 행을 THEAD 맨 앞에 추가."""
    ncols = ncols or max(len(r) for r in head_rows + body_rows)

    def parse_cell(text: str, tag: str) -> str:
        attrs = {}
        while text.startswith("||"):
            spec, text = text[2:].split("||", 1)
            k, v = spec.split("=", 1)
            attrs[k] = v
        return cell(text, tag, **attrs)

    cols = CRLF.join(f'<COL WIDTH="{width // ncols}"/>' for _ in range(ncols))
    head = list(head_rows)
    rows_h = [tr([parse_cell(t, "TH") for t in r]) for r in head]
    if unit:
        rows_h.insert(0, tr([cell(f"(단위 : {unit})", "TH", COLSPAN=str(ncols),
                                  ALIGN="RIGHT")]))
    rows_b = [tr([parse_cell(t, "TD") for t in r]) for r in body_rows]
    sep = CRLF + CRLF  # TR 블록 사이 빈 줄 (splice.ROW_SEP 관례)
    parts = [f'<TABLE ACLASS="NORMAL" AFIXTABLE="N" BORDER="1" WIDTH="{width}">',
             "<COLGROUP>" + CRLF + cols + CRLF + "</COLGROUP>"]
    if rows_h:
        parts.append("<THEAD>" + CRLF + sep.join(rows_h) + CRLF + "</THEAD>")
    parts.append("<TBODY>" + CRLF + sep.join(rows_b) + CRLF + "</TBODY>")
    parts.append("</TABLE>")
    return CRLF.join(parts)


def title(text: str, atocid: int = 1) -> str:
    return f'<TITLE ATOC="Y" ATOCID="{atocid}">{esc(text)}</TITLE>'


def p(text: str) -> str:
    return f"<P>{esc(text)}</P>"


def pgbrk() -> str:
    return "<PGBRK/>"


def group(inner: str, aclass: str = "COVER") -> str:
    """TABLE-GROUP 래핑 (K-GAAP 표준 재무제표·표지류 — 구조 편집 금지 대상)."""
    return f'<TABLE-GROUP ACLASS="{aclass}">' + CRLF + inner + CRLF + "</TABLE-GROUP>"


def insertion(inner: str) -> str:
    """INSERTION/LIBRARY 래핑 (라이브러리 삽입물 — 구조 편집 금지 대상)."""
    return ("<INSERTION><LIBRARY>" + CRLF + inner + CRLF
            + "</LIBRARY></INSERTION>")


def statement_title_table(name_spaced: str) -> str:
    """재무제표 제목표: 첫 행 첫 셀 '재 무 상 태 표' — _statement_kw 판별용."""
    return table([], [[name_spaced]], ncols=1)


def document(*fragments: str) -> bytes:
    """블록 구분자 관례: <TABLE 여는 쪽 앞에만 빈 줄, 그 외 한 줄."""
    out = ['<?xml version="1.0" encoding="utf-8"?>',
           '<DOCUMENT xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
           'xsi:noNamespaceSchemaLocation="dart4.xsd">']
    body = ""
    for frag in fragments:
        sep = CRLF + CRLF if frag.startswith("<TABLE") else CRLF
        body = frag if not body else body + sep + frag
    out.append(body)
    out.append("</DOCUMENT>")
    return CRLF.join(out).encode("utf-8")


def meta(regname: str = "합성테스트(주)", size: str = "0") -> bytes:
    return (
        '<?xml version="1.0" encoding="utf-8"?>\r\n<METAINFO>\r\n'
        '\t<GENERATOR schema="dart4.xsd" editver="5.110"/>\r\n'
        f'\t<DOCUMENT-HEADER regcik="00000000" regname="{regname}"/>\r\n'
        '\t<DOCUMENT-INFO docver="1.0"/>\r\n'
        f'\t<FILE-INFO date="20261231" size="{size}"/>\r\n'
        "</METAINFO>"
    ).encode("utf-8")


def dsd_bytes(contents: bytes, meta_bytes: bytes | None = None) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("contents.xml", contents)
        zf.writestr("meta.xml", meta_bytes or meta())
    return buf.getvalue()


def write_dsd(path, contents: bytes, meta_bytes: bytes | None = None):
    with open(path, "wb") as f:
        f.write(dsd_bytes(contents, meta_bytes))


# ── 표준 시나리오: 재무상태표 + 주석 3개 (대사·검증 테스트 공용) ──────────
#
# 재무상태표(당기/전기, 주석 참조 열):
#   현금및현금성자산(주4) 1,000/900 · 매출채권(주5) 5,000/4,000 · 자산총계 6,000/4,900
# 주석4 현금: 보유현금 400/300 + 예금 600/600 = 합계 1,000/900
# 주석5 매출채권: 채권 5,500/4,300 − 대손충당금 (500)/(300) = 장부금액(순액) 5,000/4,000
# 주석7 담보제공자산: 매출채권 5,500/4,300 (주석5의 총액과 대사되어야 하는 항목)

def standard_contents(*, cash_note_total: str = "1,000",
                      pledged_ar: str = "5,500") -> bytes:
    bs = table(
        [["과  목", "주  석", "제 2(당) 기말", "제 1(전) 기말"]],
        [
            ["Ⅰ.유동자산", "", "6,000", "4,900"],
            ["   현금및현금성자산", "||ALIGN=CENTER||4", "1,000", "900"],
            ["   매출채권", "||ALIGN=CENTER||5", "5,000", "4,000"],
            ["자산총계", "", "6,000", "4,900"],
        ],
        unit="원",
    )
    note4 = table(
        [["구  분", "당기말", "전기말"]],
        [
            ["보유현금", "400", "300"],
            ["보통예금", "600", "600"],
            ["합  계", cash_note_total, "900"],
        ],
        unit="원",
    )
    note5 = table(
        [["구  분", "당기말", "전기말"]],
        [
            ["매출채권", "5,500", "4,300"],
            ["대손충당금", "(500)", "(300)"],
            ["차감후 장부금액", "5,000", "4,000"],
        ],
        unit="원",
    )
    note7 = table(
        [["구  분", "당기말", "전기말"]],
        [
            ["매출채권", pledged_ar, "4,300"],
        ],
        unit="원",
    )
    return document(
        p("합성 테스트 문서"),
        title("재무제표", 2),
        statement_title_table("재 무 상 태 표"),
        bs,
        title("주석", 3),
        p("제 2(당) 기말 주석"),
        p("1. 일반사항\n합성테스트 주식회사는 테스트를 위해 설립되었습니다."),
        p("4. 현금및현금성자산\n당기말과 전기말 현재 현금및현금성자산의 내역은 다음과 같습니다."),
        note4,
        p("5. 매출채권\n당기말과 전기말 현재 매출채권의 내역은 다음과 같습니다."),
        note5,
        p("7. 담보제공자산\n담보로 제공된 자산의 내역은 다음과 같습니다."),
        note7,
    )
