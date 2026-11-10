"""입력 오류 안내, 덮어쓰기 보호, 결과 봉투(--json), inspect 목차·찾기, apply 주석 추가·삭제."""
import json

import pytest

from dsdxlsx import docview, guard, result
from dsdxlsx.cli import main
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.edits import EditError, apply_edits
from dsdxlsx.parse import parse_contents
from fixtures import factory as F
from test_edits import demo_contents


@pytest.fixture
def dsd(tmp_path):
    path = tmp_path / "a.dsd"
    F.write_dsd(path, demo_contents())
    return str(path)


def envelope(capsys) -> dict:
    env = json.loads(capsys.readouterr().out)
    assert env["schema"] == result.SCHEMA
    return env


# ── 입력 오류 ────────────────────────────────────────────────────────

def test_missing_file_is_one_line_and_exit_2(tmp_path, capsys):
    assert main(["export", str(tmp_path / "없음.dsd")]) == 2
    err = capsys.readouterr().err
    assert "파일이 없다" in err and "Errno" not in err and "Traceback" not in err


def test_wrong_kind_says_what_belongs_here(dsd, tmp_path, capsys):
    assert main(["import", dsd]) == 2                      # 엑셀 자리에 DSD
    assert "DSD 파일이다" in capsys.readouterr().err
    assert main(["export", dsd]) == 0
    capsys.readouterr()
    assert main(["export", dsd[:-4] + ".xlsx"]) == 2       # DSD 자리에 엑셀
    assert "엑셀 파일이다" in capsys.readouterr().err
    junk = tmp_path / "글.dsd"
    junk.write_text("DSD 가 아님", encoding="utf-8")
    assert main(["check", str(junk)]) == 2
    assert "DSD 도 엑셀도 아니다" in capsys.readouterr().err


def test_old_xls_workpaper_is_explained(dsd, tmp_path, capsys):
    xls = tmp_path / "조서.xls"
    xls.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 64)
    assert main(["recon", dsd, str(xls)]) == 2
    assert ".xlsx 로 다시 저장" in capsys.readouterr().err


def test_input_error_in_json_mode_is_an_envelope(tmp_path, capsys):
    assert main(["export", str(tmp_path / "없음.dsd"), "--json"]) == 2
    env = envelope(capsys)
    assert not env["ok"] and env["exit_code"] == 2 and "파일이 없다" in env["errors"][0]


# ── 덮어쓰기 보호 ────────────────────────────────────────────────────

def test_export_overwrites_own_output_but_not_an_edited_one(dsd, capsys):
    xlsx = dsd[:-4] + ".xlsx"
    assert main(["export", dsd]) == 0
    assert main(["export", dsd]) == 0                      # 만든 그대로면 다시 써도 된다
    with open(xlsx, "ab") as f:                            # 사람이 고쳐 저장한 것으로 친다
        f.write(b"edited")
    before = open(xlsx, "rb").read()
    capsys.readouterr()
    assert main(["export", dsd]) == 3
    assert "고쳐진 파일" in capsys.readouterr().err
    assert open(xlsx, "rb").read() == before               # 손대지 않았다
    assert main(["export", dsd, "--force"]) == 0
    assert open(xlsx, "rb").read() != before


def test_unknown_existing_file_is_not_overwritten(dsd, tmp_path, capsys):
    other = tmp_path / "남의파일.xlsx"
    other.write_bytes(b"someone else's workbook")
    assert main(["export", dsd, "-o", str(other)]) == 3
    assert "만든 기록이 없다" in capsys.readouterr().err
    assert other.read_bytes() == b"someone else's workbook"


def test_never_writes_over_the_input(dsd, capsys):
    assert main(["export", dsd]) == 0
    capsys.readouterr()
    assert main(["import", dsd[:-4] + ".xlsx", "-o", dsd[:-4] + ".xlsx"]) == 3
    assert "입력 파일 자리" in capsys.readouterr().err


def test_rollforward_protects_a_filled_workbook(dsd, capsys):
    assert main(["rollforward", dsd]) == 0
    assert main(["rollforward", dsd]) == 0
    xlsx = dsd[:-4] + "_차기기초.xlsx"
    with open(xlsx, "ab") as f:
        f.write(b"filled")
    capsys.readouterr()
    assert main(["rollforward", dsd, "--json"]) == 3
    env = envelope(capsys)
    assert not env["ok"] and env["outputs"] == [] and "고쳐진 파일" in env["errors"][0]


def test_guard_state_is_per_file(tmp_path):
    a = tmp_path / "a.bin"
    a.write_bytes(b"1")
    assert guard.untouched(str(a)) is None
    guard.remember(str(a))
    assert guard.untouched(str(a)) is True
    a.write_bytes(b"2")
    assert guard.untouched(str(a)) is False


# ── 결과 봉투 ────────────────────────────────────────────────────────

def test_check_json_reports_mismatch_with_ids_and_exit_1(dsd, capsys):
    assert main(["check", dsd, "--json"]) == 1
    env = envelope(capsys)
    assert env["command"] == "check" and env["ok"] and env["exit_code"] == 1
    bad = [m for m in env["mismatches"] if m["check"] == "C3"]
    assert bad and bad[0]["row"] == "t003.r005" and bad[0]["col"] == 1
    assert bad[0]["stated"] == 1200 and bad[0]["computed"] == 1300
    assert any(f["level"] == "warn" for f in env["findings"])
    assert main(["check", dsd]) == 0                       # 옛 방식 종료코드는 그대로
    capsys.readouterr()


def test_check_json_flag_does_not_swallow_the_file(dsd, tmp_path, capsys):
    assert main(["check", "--json", dsd]) == 1             # --json 이 앞에 와도 봉투
    assert envelope(capsys)["command"] == "check"
    review = tmp_path / "review.json"                      # 경로를 주면 옛 방식(리뷰 JSON 파일)
    assert main(["check", dsd, "--json", str(review)]) == 0
    capsys.readouterr()
    assert json.loads(review.read_text(encoding="utf-8"))["schema"] == 1


def test_fix_loop_check_apply_check(dsd, tmp_path, capsys):
    """검산 결과의 위치로 바로 고친다: check --json → apply → check --json."""
    assert main(["check", dsd, "--json"]) == 1
    m = next(x for x in envelope(capsys)["mismatches"] if x["check"] == "C3")
    edits = tmp_path / "e.json"
    edits.write_text(json.dumps([{"op": "set", "row": m["row"], "col": m["col"],
                                  "value": m["computed"]}]), encoding="utf-8")
    assert main(["apply", dsd, str(edits), "--json"]) == 0
    out = envelope(capsys)["outputs"][0]
    assert main(["check", out, "--json", "--only", "C3"]) == 0
    assert envelope(capsys)["mismatches"] == []


def test_export_import_diff_envelopes(dsd, capsys):
    assert main(["export", dsd, "--json"]) == 0
    env = envelope(capsys)
    assert env["outputs"] == [dsd[:-4] + ".xlsx"] and env["tables"] >= 5 and env["log"]
    assert main(["import", env["outputs"][0], "--json"]) == 0
    env = envelope(capsys)
    assert env["changed"] is False and env["changes"] == {}
    assert read_dsd(env["outputs"][0]).contents == read_dsd(dsd).contents
    assert main(["diff", dsd, env["outputs"][0], "--json"]) == 0
    env = envelope(capsys)
    assert env["total"] == 0 and env["outputs"][0].endswith(".html")
    assert main(["verify", dsd, dsd, "--json"]) == 0
    assert envelope(capsys)["same"] is True


def test_rollforward_envelope_counts(dsd, capsys):
    assert main(["rollforward", dsd, "--json"]) == 0
    env = envelope(capsys)
    assert len(env["outputs"]) == 2 and set(env["counts"]) == {"자동확정", "확인필요", "미처리"}
    assert all(t["grade"] != "자동확정" for t in env["todo"])


# ── inspect 목차·찾기 ────────────────────────────────────────────────

def test_outline_lists_blocks_without_cells():
    d = docview.outline_dict(parse_contents(demo_contents()))
    assert d["schema"] == docview.OUTLINE_SCHEMA
    sec = next(s for s in d["sections"] if s["section"] == "4")
    assert sec["tables"] == 1 and sec["paras"] == 1
    t = next(b for b in sec["blocks"] if b["type"] == "table")
    assert t["id"] == "t003" and t["ncols"] == 3 and "rows" not in t
    assert t["labels"][:2] == ["보유현금", "보통예금"] and t["unit"] == "백만원"
    md = docview.outline_markdown(d)
    assert "[section: 4]" in md and "t003 표 3열" in md


def test_find_by_text_and_by_number():
    doc = parse_contents(demo_contents())
    d = docview.find_dict(doc, ["보통 예금"])                # 공백 무시
    assert [(h["type"], h["id"]) for h in d["hits"]] == [("row", "t003.r004")]
    d = docview.find_dict(doc, ["1200"])                   # 숫자는 값으로(쉼표 없이 적어도)
    assert {h["id"] for h in d["hits"]} == {"t002.r004.c03", "t003.r005.c02"}
    assert all(h["type"] == "cell" and h["text"] == "1,200" for h in d["hits"])
    d = docview.find_dict(doc, ["시험용 가상"])
    assert d["hits"][0]["type"] == "para" and d["hits"][0]["section"] == "1"
    d = docview.find_dict(doc, ["합계"], limit=1)
    assert d["total"] > 1 and d["truncated"] and len(d["hits"]) == 1
    assert "건" in docview.find_markdown(d)


def test_inspect_cli_outline_find_and_bad_section(dsd, capsys):
    assert main(["inspect", dsd, "--outline", "--json"]) == 0
    assert json.loads(capsys.readouterr().out)["schema"] == docview.OUTLINE_SCHEMA
    assert main(["inspect", dsd, "--find", "대손충당금", "--json"]) == 0
    d = json.loads(capsys.readouterr().out)
    assert d["schema"] == docview.FIND_SCHEMA and d["hits"][0]["id"] == "t004.r004"
    assert main(["inspect", dsd, "--section", "4,99"]) == 2  # 하나라도 없으면 알린다
    assert "99" in capsys.readouterr().err


# ── apply 주석 추가·삭제 ─────────────────────────────────────────────

NEW_TABLE = {"template": "t004", "header_rows": 2,
             "rows": [["", "", "(단위: 백만원)"], ["구  분", "당기말", "전기말"],
                      ["단기차입금", 10, 20], ["합  계", 10, 20]]}


def test_insert_note_at_end():
    res = apply_edits(demo_contents(), [{"op": "insert_note", "no": 7, "title": "차입금",
                                         "blocks": ["차입금의 내역은 다음과 같습니다.", NEW_TABLE]}])
    assert res.report.structural and res.changes()["note_inserts"]
    d = docview.doc_dict(parse_contents(res.contents), only=["7"])
    blocks = d["sections"][0]["blocks"]
    assert blocks[0]["text"].startswith("7. 차입금")
    table = next(b for b in blocks if b["type"] == "table")
    assert [c["text"] for c in table["rows"][2]["cells"]] == ["단기차입금", "10", "20"]


def test_insert_note_in_the_middle_renumbers_later_notes():
    res = apply_edits(demo_contents(), [{"op": "insert_note", "no": 5, "title": "새 주석",
                                         "blocks": ["새 주석 본문입니다."]}])
    doc = parse_contents(res.contents)
    titles = {s.key: s.title for s in docview.sections(doc) if s.kind == "note"}
    assert titles["5"] == "새 주석" and titles["6"] == "매출채권" and titles["7"] == "기타"
    assert any("매출채권" in line and "6" in line for line in res.changes()["renumbered"])
    # 재무상태표의 주석 참조 칸도 같이 밀린다
    bs = next(b for s in docview.doc_dict(doc, only=["재무상태표"])["sections"]
              for b in s["blocks"] if b["type"] == "table" and b["ncols"] == 4)
    row = next(r for r in bs["rows"] if r["label"].strip() == "매출채권")
    assert [c["text"] for c in row["cells"]][1] == "6"


def test_delete_note_and_guards():
    res = apply_edits(demo_contents(), [{"op": "delete_note", "no": 6}])
    keys = [s.key for s in docview.sections(parse_contents(res.contents))]
    assert "6" not in keys and res.changes()["note_deletes"]
    with pytest.raises(EditError, match="주석 9 없음"):
        apply_edits(demo_contents(), [{"op": "delete_note", "no": 9}])
    with pytest.raises(EditError, match="title"):
        apply_edits(demo_contents(), [{"op": "insert_note", "no": 7}])
    with pytest.raises(EditError, match="template"):
        apply_edits(demo_contents(), [{"op": "insert_note", "no": 7, "title": "차입금",
                                       "blocks": [{"rows": [["구분", "당기말", "전기말"]]}]}])
    with pytest.raises(EditError, match="불연속"):        # 번호를 건너뛰면 막는다
        apply_edits(demo_contents(), [{"op": "insert_note", "no": 12, "title": "먼 번호"}])
