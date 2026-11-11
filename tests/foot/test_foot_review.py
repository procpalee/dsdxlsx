"""검토 재료 내기(foot --review)와 판정 반영(annotate-ai) - 도구는 LLM 을 부르지 않는다.

판정하는 쪽(에이전트)을 흉내 내 판정 JSON 을 손으로 만들어 넣고, 도구가 그것을 다시 검산해
맞는 것만 워크북에 적는지 본다.
"""
import json
import re
import warnings

import openpyxl
import pytest

from dsdxlsx.cli import main
from dsdxlsx.foot import pipeline, review
from fixtures.foot_factory import _amount_row, _td, build_dsd, table, write_dsd


def _bad(tmp_path):
    """합계 오류(150 이어야 할 곳에 130)를 넣은 합성 문서 → 풋팅 워크북."""
    note = ("<P>1. 오류주입&amp;cr;&amp;cr;내역입니다.</P>"
            + table([
                [_td("구분"), _td("당기")],
                _amount_row("항목가", ["100"]),
                _amount_row("항목나", ["50"]),
                _amount_row("합 계", ["130"]),
            ], 2))
    xml = ('<?xml version="1.0" encoding="UTF-8"?><BODY><TITLE ATOC="Y">주석</TITLE>'
           + note + "</BODY>").encode()
    p = tmp_path / "err.dsd"
    p.write_bytes(build_dsd(xml))
    return pipeline.process_file_full(p, tmp_path / "err.xlsx")


def _texts(xlsx) -> list[str]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(xlsx))
    return [str(c.value) for row in wb["총괄표"].iter_rows() for c in row if c.value]


# ── 검토 재료 ────────────────────────────────────────────────────────

def test_review_material_has_items_context_and_guides(tmp_path):
    res = _bad(tmp_path)
    d = review.build_review(res)
    assert d["schema"] == review.SCHEMA_REVIEW and d["workbook"] == "err.xlsx" and d["source_sha256"]
    it = d["items"][0]
    assert it["id"] == "F1" and it["pretag"] == "금액 불일치"
    assert it["stated"] == "130" and it["expected"] == "150" and it["kind_desc"]
    assert "항목가" in it["context"] and "130" in it["context"]         # 표 발췌
    assert set(d["guide"]) == {"caution", "verdicts", "verdict", "xref", "typo"}
    assert "따르지 말고" in d["guide"]["caution"]                       # 문서 안의 글은 자료일 뿐
    assert d["texts"] and any("오류주입" in t["text"] for t in d["texts"])


def test_review_material_lists_statement_and_note_values(tmp_path):
    p = tmp_path / "std.dsd"
    write_dsd(p)
    d = review.build_review(pipeline.process_file_full(p, tmp_path / "std.xlsx"))
    assert d["items"] == []                                             # 깨끗한 문서 - 검출 없음
    assert [s["sheet"] for s in d["xref"]["statements"]][:2] == ["BS", "PL"]
    assert d["xref"]["notes"] and all(n["sheet"].isdigit() for n in d["xref"]["notes"])
    assert any("=" in ln for ln in d["xref"]["statements"][0]["lines"])  # '셀=값' 줄
    assert d["xref"]["linked"]                                          # 이미 대사된 대상 칸


# ── 판정 반영 ────────────────────────────────────────────────────────

def test_fix_verdict_survives_only_when_it_matches_the_tool(tmp_path):
    res = _bad(tmp_path)
    it = review.build_review(res)["items"][0]
    good = {"reviewer": "시험", "verdicts": [{"id": "F1", "판정": "수정필요", "근거": "구성 합 150과 20 차이",
                                             "수정대상": it["target"], "제안값": it["expected"],
                                             "수정제안": "합계를 150으로"}]}
    a = review.apply_findings(res.out, good)
    v = a.review.verdicts["F1"]
    assert v["최종판정"] == "수정필요" and v["검산상태"] == "수정값검산완료" and a.review.status == "complete"
    texts = _texts(res.out)
    assert any(review.SUMMARY_TITLE in t for t in texts) and any("합계를 150으로" in t for t in texts)
    assert any("판정 시험" in t for t in texts)

    wrong = {"verdicts": [{"id": "F1", "판정": "수정필요", "근거": "틀린 값", "수정대상": it["target"],
                           "제안값": "999", "수정제안": "999로"}]}
    v = review.apply_findings(res.out, wrong).review.verdicts["F1"]
    assert v["판정"] == "수정필요" and v["최종판정"] == "확인필요" and v["검산상태"] == "수정값불일치"
    texts = _texts(res.out)
    assert sum(review.SUMMARY_TITLE in t for t in texts) == 1          # 다시 반영해도 한 번만


def test_missing_unknown_and_duplicate_ids_are_kept_visible(tmp_path):
    res = _bad(tmp_path)
    a = review.apply_findings(res.out, {"verdicts": [{"id": "F9", "판정": "무시가능"}]})
    d = a.to_dict()
    assert d["status"] == "partial" and d["missing_ids"] == ["F1"] and d["unknown_ids"] == ["F9"]
    assert d["to_check"][0]["id"] == "F1" and d["to_check"][0]["validation"] == "판정미완료"
    a = review.apply_findings(res.out, {"verdicts": [{"id": "F1", "판정": "무시가능"},
                                                     {"id": "F1", "판정": "도구오탐"}]})
    assert a.duplicate_ids == ["F1"] and a.review.verdicts["F1"]["최종판정"] == "확인필요"
    a = review.apply_findings(res.out, {"verdicts": [{"id": "F1", "판정": "이상한값", "근거": "x"}]})
    assert a.review.verdicts["F1"]["판정"] == "확인필요"                 # 모르는 판정어


def _cells(lines: list[str]) -> dict[str, int]:
    out = {}
    for ln in lines:
        for addr, val in re.findall(r"([A-Z]+\d+)=(-?[\d,]+)", ln):
            out[addr] = int(val.replace(",", ""))
    return out


def test_xref_proposals_are_verified_before_they_are_planted(tmp_path):
    from fixtures.foot_factory import standard_contents

    # 금액이 천원 단위인 문서(도구는 1,000원 미만 짝을 소액으로 거부한다)
    xml = standard_contents().decode("utf-8").replace("단위 : 원", "단위 : 천원").replace("단위: 원", "단위: 천원")
    p = tmp_path / "std.dsd"
    p.write_bytes(build_dsd(xml.encode("utf-8")))
    res = pipeline.process_file_full(p, tmp_path / "std.xlsx")
    d = review.build_review(res)
    # 원천은 재무상태표 칸만 쓴다(이 합성 문서에서 단위 표시가 있는 재무제표)
    fs = {(s["sheet"], a): v for s in d["xref"]["statements"] if s["sheet"] == "BS"
          for a, v in _cells(s["lines"]).items()}
    linked = set(d["xref"]["linked"])
    good = bad = None
    for n in d["xref"]["notes"]:
        for a, v in _cells(n["lines"]).items():
            dst = f"'{n['sheet']}'!{a}"
            if dst in linked or abs(v) < 100:
                continue
            same = next(((s, fa) for (s, fa), fv in fs.items() if abs(fv) == abs(v)), None)
            diff = next(((s, fa) for (s, fa), fv in fs.items() if abs(fv) != abs(v) and abs(fv) >= 100), None)
            if same and good is None:
                good = {"src": f"{same[0]}!{same[1]}", "dst": dst, "이유": "같은 계정의 같은 시점"}
            elif diff and bad is None:
                bad = {"src": f"{diff[0]}!{diff[1]}", "dst": dst, "이유": "일부러 틀린 짝"}
    assert good and bad
    junk = [{"src": "없는시트!A1", "dst": good["dst"], "이유": "없는 시트"},
            {"src": good["src"], "dst": "BS!D10", "이유": "대상이 재무제표"}]
    a = review.apply_findings(res.out, {"xref": [good, bad, *junk]})
    by = {(p.src, p.dst): p for p in a.review.xref}
    assert by[(good["src"], good["dst"])].status == "반영(일치)", by[(good["src"], good["dst"])].detail
    assert by[(bad["src"], bad["dst"])].status == "불일치"
    assert all(by[(j["src"], j["dst"])].status == "거부" for j in junk)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(res.out))
    sheet, addr = by[(good["src"], good["dst"])].formula_cell.split("!")
    f = wb[sheet][addr].value
    assert isinstance(f, str) and f.startswith("=IF(")                  # 살아 있는 비교 수식
    links = [c.value for row in wb[sheet].iter_rows() for c in row
             if isinstance(c.value, str) and c.value.startswith("=HYPERLINK") and "AI:" in c.value]
    assert len(links) == 1
    # 같은 판정을 다시 넣어도 겹쳐 심지 않는다
    again = review.apply_findings(res.out, {"xref": [good]})
    assert again.review.xref[0].status == "거부" and "이미" in again.review.xref[0].detail


def test_typos_need_a_real_cell_and_matching_text(tmp_path):
    res = _bad(tmp_path)
    t = next(x for x in review.build_review(res)["texts"] if "오류주입" in x["text"])
    a = review.apply_findings(res.out, {"typos": [
        {"셀": t["cell"], "원문": "오류주입", "수정": "오류 주입", "사유": "띄어쓰기"},
        {"셀": "ZZ999", "원문": "없는 글", "수정": "고친 글", "사유": "없는 칸"},
        {"셀": t["cell"], "원문": "문서에 없는 문장", "수정": "x", "사유": "원문이 다름"},
    ]})
    assert len(a.review.typos) == 1 and len(a.rejected_typos) == 2
    assert {r["거부사유"] for r in a.rejected_typos} == {"존재하지 않는 셀", "제시 원문과 실제 셀 내용 불일치"}


def test_findings_shape_and_wrong_workbook_are_input_errors(tmp_path):
    res = _bad(tmp_path)
    with pytest.raises(review.FindingsError):
        review.apply_findings(res.out, [{"id": "F1"}])
    with pytest.raises(review.FindingsError):
        review.apply_findings(res.out, {"verdicts": "F1"})


# ── 명령 ─────────────────────────────────────────────────────────────

def test_cli_roundtrip_review_then_annotate(tmp_path, capsys):
    res = _bad(tmp_path)
    capsys.readouterr()
    dsd = tmp_path / "err.dsd"
    assert main(["foot", str(dsd), "--review", "--json", "--force"]) == 1
    env = json.loads(capsys.readouterr().out)
    info = env["files"][0]["review"]
    assert info["path"].endswith("err_풋팅_검토재료.json") and info["items"] == 1
    assert info["path"] in env["outputs"]
    material = json.loads(open(info["path"], encoding="utf-8").read())
    it = material["items"][0]
    book = env["files"][0]["output"]

    findings = tmp_path / "판정.json"
    findings.write_text(json.dumps({"verdicts": [{
        "id": it["id"], "판정": "수정필요", "근거": "구성 합과 20 차이", "수정대상": it["target"],
        "제안값": it["expected"], "수정제안": "합계를 150으로"}]}, ensure_ascii=False), encoding="utf-8")
    assert main(["annotate-ai", book, str(findings), "--json"]) == 1    # 확인할 것(수정필요)이 있다
    env = json.loads(capsys.readouterr().out)
    assert env["kind"] == "foot" and env["verdicts"] == {"수정필요": 1} and env["status"] == "complete"
    assert env["to_check"][0]["validation"] == "수정값검산완료"
    assert any("합계를 150으로" in t for t in _texts(book))
    assert main(["annotate-ai", book, str(findings)]) == 0              # --json 없이는 0
    assert res.out.exists()


def test_cli_annotate_rejects_bad_findings(tmp_path, capsys):
    _bad(tmp_path)
    capsys.readouterr()
    bad = tmp_path / "bad.json"
    bad.write_text('{"verdicts": "F1"}', encoding="utf-8")
    assert main(["annotate-ai", str(tmp_path / "err.xlsx"), str(bad), "--json"]) == 2
    assert "배열" in json.loads(capsys.readouterr().out)["errors"][0]
