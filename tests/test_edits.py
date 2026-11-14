"""inspect·apply - 엑셀 없이 JSON 으로 DSD 를 읽고 고치는 통로(LLM 용)."""
import json

import pytest

from dsdxlsx import docview, result
from dsdxlsx.check import run_checks
from dsdxlsx.cli import main
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.edits import EditError, apply_edits
from dsdxlsx.parse import parse_contents
from fixtures import factory as F

HEAD = [["과  목", "주  석", "제 2(당) 기말", "제 1(전) 기말"]]
NOTE_HEAD = [["구  분", "당기말", "전기말"]]


def demo_contents(deposit="1,100", fixed=False) -> bytes:
    """샘플전자: 주석 4 보통예금을 1,100 으로 잘못 적어 합계(1,200)와 안 맞는 문서.
    주석 6 에는 '합  계' 행이 있는 표가 둘이다(위치 지정 모호성 시험용)."""
    bs = F.table(HEAD, [
        ["Ⅰ.유동자산", "", "6,200", "4,900"],
        ["   현금및현금성자산", "||ALIGN=CENTER||4", "1,200", "900"],
        ["   매출채권", "||ALIGN=CENTER||5", "5,000", "4,000"],
        ["자산총계", "", "6,200", "4,900"],
    ], unit="백만원")
    note4 = F.table(NOTE_HEAD, [
        ["보유현금", "200", "150"],
        ["보통예금", deposit, "750"],
        ["합  계", "1,200", "900"],
    ], unit="백만원")
    if fixed:
        note4 = note4.replace('AFIXTABLE="N"', 'AFIXTABLE="Y"', 1)
    note5 = F.table(NOTE_HEAD, [
        ["매출채권", "5,500", "4,300"],
        ["대손충당금", "(500)", "(300)"],
        ["차감후 장부금액", "5,000", "4,000"],
    ], unit="백만원")
    note6a = F.table(NOTE_HEAD, [["제품", "100", "90"], ["합  계", "100", "90"]], unit="백만원")
    note6b = F.table(NOTE_HEAD, [["원재료", "30", "20"], ["합  계", "30", "20"]], unit="백만원")
    return F.document(
        F.title("재무제표", 2),
        F.statement_title_table("재 무 상 태 표"),
        bs,
        F.title("주석", 3),
        F.p("1. 일반사항\n샘플전자 주식회사는 시험용 가상 회사입니다."),
        F.p("4. 현금및현금성자산\n현금및현금성자산의 내역은 다음과 같습니다."),
        note4,
        F.p("5. 매출채권\n매출채권의 내역은 다음과 같습니다."),
        note5,
        F.p("6. 기타\n기타 내역은 다음과 같습니다."),
        note6a,
        F.p("기타 둘째 표입니다."),
        note6b,
    )


def c3_warns(contents: bytes) -> list:
    rep = run_checks(parse_contents(contents), only={"C3"})
    return [f for f in rep.findings if f.level == "warn"]


def texts(contents: bytes, tid: str) -> list[list[str]]:
    doc = parse_contents(contents)
    return [[c.text for c in r.cells] for r in doc.tables[tid].rows]


# ── inspect ─────────────────────────────────────────────────────────

def test_inspect_ids_values_and_columns():
    d = docview.doc_dict(parse_contents(demo_contents()))
    keys = [s["section"] for s in d["sections"]]
    assert keys[:2] == ["재무상태표", "주석"] and {"1", "4", "5", "6"} <= set(keys)
    t = next(b for s in d["sections"] if s["section"] == "4"
             for b in s["blocks"] if b["type"] == "table")
    assert [c["period"] for c in t["columns"]] == [None, "당기", "전기"]
    row = next(r for r in t["rows"] if r["label"] == "보통예금")
    assert row["id"] == "t003.r004"
    cell = next(c for c in row["cells"] if c["col"] == 1)
    assert cell["id"] == "t003.r004.c02" and cell["value"] == 1100 and cell["text"] == "1,100"
    assert next(r for r in t["rows"] if r["label"] == "합  계").get("total") is True
    md = docview.to_markdown(d)
    assert "[1] 당기말 (당기)" in md and "| t003.r004 | 보통예금 | 1,100 | 750 |" in md


def test_inspect_cli_section_filter(tmp_path, capsys):
    F.write_dsd(tmp_path / "a.dsd", demo_contents())
    assert main(["inspect", str(tmp_path / "a.dsd"), "--section", "4", "--json"]) == 0
    d = json.loads(capsys.readouterr().out)
    assert [s["section"] for s in d["sections"]] == ["4"]
    assert main(["inspect", str(tmp_path / "a.dsd"), "--section", "99"]) == 2


# ── apply: 결과 ─────────────────────────────────────────────────────

def test_empty_edits_keep_bytes():
    for contents in (demo_contents(), F.standard_contents()):
        assert apply_edits(contents, []).contents == contents


def test_json_only_flow_fix_number_insert_row_then_check_passes():
    """완료 기준: 엑셀 없이 JSON 만으로 숫자 고치기 → 행 넣기 → 검산 통과."""
    contents = demo_contents()
    assert len(c3_warns(contents)) == 1  # 보통예금 1,100 → 합계 안 맞음
    out = apply_edits(contents, [
        {"op": "set", "section": "4", "row": "보통예금", "col": "당기", "value": 950},
        {"op": "insert_row", "after": {"section": "4", "row": "보통예금"},
         "values": {"구분": "외화예금", "당기말": 50, "전기말": 0}},
    ]).contents
    assert texts(out, "t003")[2:] == [["보유현금", "200", "150"], ["보통예금", "950", "750"],
                                       ["외화예금", "50", "-"], ["합  계", "1,200", "900"]]
    assert c3_warns(out) == []


def test_cell_id_and_column_forms_hit_same_cell():
    contents = demo_contents()
    by_id = apply_edits(contents, [{"op": "set", "cell": "t003.r004.c02", "value": 1000}]).contents
    by_num = apply_edits(contents, [{"op": "set", "row": "t003.r004", "col": 1, "value": 1000}]).contents
    by_head = apply_edits(contents, [{"op": "set", "row": "t003.r004", "col": "당기말",
                                      "value": 1000}]).contents
    assert by_id == by_num == by_head
    assert texts(by_id, "t003")[3] == ["보통예금", "1,000", "750"]


def test_negative_number_keeps_parentheses():
    out = apply_edits(demo_contents(), [
        {"op": "set", "section": "5", "row": "대손충당금", "col": "당기", "value": -600}]).contents
    assert texts(out, "t004")[3] == ["대손충당금", "(600)", "(300)"]


def test_two_inserts_after_same_row_keep_order_and_delete_row():
    out = apply_edits(demo_contents(), [
        {"op": "insert_row", "after": "t003.r004", "values": ["외화예금", 50, 0]},
        {"op": "insert_row", "after": "t003.r004", "values": ["정기예금", 7, 0]},
        {"op": "delete_row", "row": "t003.r003"},
    ]).contents
    assert [r[0] for r in texts(out, "t003")[2:]] == ["보통예금", "외화예금", "정기예금", "합  계"]


def test_paragraph_edits():
    res = apply_edits(demo_contents(), [
        {"op": "set_text", "id": "p003", "text": "4. 현금및현금성자산\n바뀐 문장입니다."},
        {"op": "insert_para", "after": "t003", "text": "사용이 제한된 예금은 없습니다."},
        {"op": "delete_para", "id": "p004"},
    ])
    paras = [b.item.text for b in parse_contents(res.contents).blocks if b.kind == "p"]
    assert "4. 현금및현금성자산\n바뀐 문장입니다." in paras
    assert "사용이 제한된 예금은 없습니다." in paras
    assert not any(p.startswith("5. 매출채권") for p in paras)
    assert res.report.structural


def test_text_starting_with_hash_is_not_a_marker():
    out = apply_edits(demo_contents(), [
        {"op": "set_text", "id": "p003", "text": "#1 우선 확인"}]).contents
    assert any(b.item.text == "#1 우선 확인" for b in parse_contents(out).blocks if b.kind == "p")


def test_table_delete_and_insert_in_notes():
    res = apply_edits(demo_contents(), [
        {"op": "delete_table", "id": "t006"},
        {"op": "insert_table", "after": "p004", "template": "t004", "header_rows": 1,
         "rows": [["구  분", "당기말", "전기말"], ["미수금", 10, 20]]},
    ])
    assert res.report.table_deletes and res.report.table_inserts
    doc = parse_contents(res.contents)
    labels = [[c.text for c in r.cells][0] for t in doc.tables.values() for r in t.rows]
    assert "미수금" in labels and "원재료" not in labels and "제품" in labels


# ── apply: 오류·가드 ─────────────────────────────────────────────────

@pytest.mark.parametrize("edit, msg", [
    ({"op": "set", "section": "6", "row": "합계", "col": "당기", "value": 1}, "2개"),
    ({"op": "set", "section": "4", "row": "없는행", "col": "당기", "value": 1}, "없는행"),
    ({"op": "set", "section": "99", "row": "보통예금", "col": "당기", "value": 1}, "99"),
    ({"op": "set", "row": "t003.r004", "col": "c2", "value": 1}, "cell"),
    ({"op": "set", "row": "t003.r004", "col": 7, "value": 1}, "열이 3개"),
    ({"op": "set", "row": "t003.r004", "col": "당기", "value": [1]}, "값"),
    ({"op": "rename", "id": "p001"}, "알 수 없는 편집"),
    ({"op": "delete_table", "id": "t002"}, "재무제표 본문 표"),
    ({"op": "delete_row", "row": "t003.r002"}, "헤더"),
])
def test_errors_name_the_problem(edit, msg):
    with pytest.raises(EditError, match=msg):
        apply_edits(demo_contents(), [edit])


def test_row_nth_picks_one_of_duplicates():
    out = apply_edits(demo_contents(), [
        {"op": "set", "section": "6", "row": "합계", "row_nth": 2, "col": "당기", "value": 31}]).contents
    assert texts(out, "t006")[-1] == ["합  계", "31", "20"]
    assert texts(out, "t005")[-1] == ["합  계", "100", "90"]


def test_fixed_table_refuses_row_insert():
    with pytest.raises(EditError, match="고정 표"):
        apply_edits(demo_contents(fixed=True), [
            {"op": "insert_row", "after": "t003.r004", "values": ["외화예금", 50, 0]}])


# ── CLI: 봉투·덮어쓰기·미리보기 ─────────────────────────────────────

def _cli(tmp_path, edits, *extra):
    src = tmp_path / "a.dsd"
    if not src.exists():
        F.write_dsd(src, demo_contents())
    ef = tmp_path / "e.json"
    ef.write_text(json.dumps(edits, ensure_ascii=False), encoding="utf-8")
    return main(["apply", str(src), str(ef), *extra])


def test_cli_apply_json_envelope_and_output(tmp_path, capsys):
    rc = _cli(tmp_path, {"edits": [{"op": "set", "cell": "t003.r004.c02", "value": 1000}]}, "--json")
    env = json.loads(capsys.readouterr().out)
    assert rc == 0 and env["schema"] == result.SCHEMA and env["ok"] and env["exit_code"] == 0
    assert env["outputs"] == [str(tmp_path / "a_수정.dsd")]
    assert env["changes"]["cell_edits"] == ["t003.r004.c02"]
    assert read_dsd(str(tmp_path / "a_수정.dsd")).contents != read_dsd(str(tmp_path / "a.dsd")).contents


def test_cli_refuses_overwrite_and_same_path(tmp_path, capsys):
    edits = [{"op": "set", "cell": "t003.r004.c02", "value": 1000}]
    assert _cli(tmp_path, edits) == 0
    capsys.readouterr()
    assert _cli(tmp_path, edits, "--json") == result.STOPPED
    env = json.loads(capsys.readouterr().out)
    assert not env["ok"] and env["outputs"] == [] and "이미 있다" in env["errors"][0]
    assert _cli(tmp_path, edits, "--force") == 0
    assert _cli(tmp_path, edits, "-o", str(tmp_path / "a.dsd"), "--force") == result.STOPPED


def test_cli_dry_run_writes_nothing(tmp_path):
    assert _cli(tmp_path, [{"op": "set", "cell": "t003.r004.c02", "value": 1000}], "--dry-run") == 0
    assert not (tmp_path / "a_수정.dsd").exists()


def test_cli_input_errors_are_2(tmp_path, capsys):
    assert _cli(tmp_path, [{"op": "set", "section": "4", "row": "없는행", "col": 1, "value": 1}],
                "--json") == result.INPUT_ERROR
    env = json.loads(capsys.readouterr().out)
    assert not env["ok"] and "없는행" in env["errors"][0]
    (tmp_path / "bad.json").write_text("{not json", encoding="utf-8")
    assert main(["apply", str(tmp_path / "a.dsd"), str(tmp_path / "bad.json")]) == result.INPUT_ERROR
    assert not (tmp_path / "a_수정.dsd").exists()


def test_cli_dry_run_json_shows_before_and_after(tmp_path, capsys):
    from dsdxlsx.parse import parse_contents

    doc = parse_contents(demo_contents())
    before = doc.nodes["t003.r004.c02"].text
    label = next(c.text.strip() for c in doc.rows["t003.r004"].cells if c.text.strip())
    assert _cli(tmp_path, [{"op": "set", "cell": "t003.r004.c02", "value": 1000}], "--dry-run", "--json") == 0
    env = json.loads(capsys.readouterr().out)
    assert env["cells"] == [{"cell": "t003.r004.c02", "label": label, "before": before, "after": "1,000"}]
