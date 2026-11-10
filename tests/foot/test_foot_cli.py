"""dsdxlsx foot 명령 - 출력 이름, 종료코드, 결과 봉투, 덮어쓰기 보호."""
import json
import warnings

import openpyxl

from dsdxlsx import result
from dsdxlsx.cli import main
from fixtures import factory as F
from fixtures.foot_factory import write_dsd


def _bad_dsd(path):
    """주석 현금 표의 합계(1,000)가 구성(100 + 950)과 안 맞는 문서."""
    F.write_dsd(path, F.standard_contents(cash_note_total="1,050"))


def envelope(capsys) -> dict:
    env = json.loads(capsys.readouterr().out)
    assert env["schema"] == result.SCHEMA and env["command"] == "foot"
    return env


def test_foot_writes_workbook_next_to_input(tmp_path, capsys):
    p = tmp_path / "보고서.dsd"
    write_dsd(p)
    assert main(["foot", str(p)]) == 0
    out = tmp_path / "보고서_풋팅.xlsx"
    assert out.exists() and "풋팅" in capsys.readouterr().out
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(out))
    assert [ws.title for ws in wb.worksheets][:2] == ["총괄표", "원문"]
    for ws in wb.worksheets:                      # 메모에 도구 이름을 남기지 않는다
        for row in ws.iter_rows():
            for c in row:
                if c.comment is not None:
                    assert "dsdfoot" not in c.comment.text and "dsdxlsx" not in c.comment.text


def test_foot_json_and_exit_codes(tmp_path, capsys):
    good, bad = tmp_path / "good.dsd", tmp_path / "bad.dsd"
    write_dsd(good)
    _bad_dsd(bad)
    assert main(["foot", str(good), str(bad), str(tmp_path / "없음.dsd"), "--json"]) == 2
    env = envelope(capsys)
    assert not env["ok"] and len(env["outputs"]) == 2
    by = {f["file"]: f for f in env["files"]}
    assert by[str(good)]["exit_code"] == 0 and by[str(good)]["footing"]["fail"] == 0
    assert by[str(bad)]["exit_code"] == 1 and by[str(bad)]["footing"]["fail"] >= 1
    m = by[str(bad)]["mismatches"][0]
    assert m["kind"] == "footing" and m["sheet"] and m["cell"] and m["diff"] != 0
    assert "파일이 없다" in by[str(tmp_path / "없음.dsd")]["error"]

    assert main(["foot", str(bad), "--json", "--force"]) == 1      # 불일치만 있으면 1
    assert envelope(capsys)["ok"]


def test_foot_convert_only_and_sections(tmp_path, capsys):
    p = tmp_path / "a.dsd"
    write_dsd(p)
    assert main(["foot", str(p), "--convert-only"]) == 0
    assert (tmp_path / "a_변환.xlsx").exists() and not (tmp_path / "a_풋팅.xlsx").exists()
    capsys.readouterr()
    assert main(["foot", str(p), "--sections", "--json"]) == 0
    env = envelope(capsys)
    names = [s["name"] for s in env["files"][0]["sheets"]]
    assert names[:3] == ["BS", "PL", "CF"] and env["outputs"] == []


def test_foot_protects_an_edited_workbook(tmp_path, capsys):
    p = tmp_path / "a.dsd"
    write_dsd(p)
    assert main(["foot", str(p)]) == 0
    assert main(["foot", str(p)]) == 0                    # 만든 그대로면 다시 만든다
    out = tmp_path / "a_풋팅.xlsx"
    with open(out, "ab") as f:
        f.write(b"edited")
    capsys.readouterr()
    assert main(["foot", str(p)]) == 3
    assert "고쳐진 파일" in capsys.readouterr().err
    assert main(["foot", str(p), "--force"]) == 0
    assert main(["foot", str(p), str(p), "-o", str(tmp_path / "x.xlsx")]) == 2   # -o 는 파일 하나일 때만
