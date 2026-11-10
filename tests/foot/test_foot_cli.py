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


# ── 풋팅 워크북에서 고친 값 → DSD (import 가 워크북 종류를 알아서 가른다) ──

def _foot_book(tmp_path, capsys):
    p = tmp_path / "보고서.dsd"
    write_dsd(p)
    assert main(["foot", str(p)]) == 0
    capsys.readouterr()
    return p, tmp_path / "보고서_풋팅.xlsx"


def _set(xlsx, sheet, label, value):
    """sheet 에서 label 행의 바로 오른쪽 칸(당기 금액)을 바꾼다."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(xlsx))
    ws = wb[sheet]
    cell = next(c for row in ws.iter_rows() for c in row
                if isinstance(c.value, str) and c.value.strip() == label)
    ws.cell(cell.row, cell.column + 1).value = value
    wb.save(str(xlsx))


def test_import_takes_a_foot_workbook(tmp_path, capsys):
    from dsdxlsx.dsd_io import read_dsd

    dsd, book = _foot_book(tmp_path, capsys)
    assert main(["import", str(book), "--json", "-o", str(tmp_path / "그대로.dsd")]) == 0
    env = json.loads(capsys.readouterr().out)
    assert env["kind"] == "foot" and env["changed"] is False and env["edits"] == []
    assert read_dsd(env["outputs"][0]).contents == read_dsd(str(dsd)).contents   # 안 고치면 바이트 그대로

    _set(book, "BS", "매출채권", 301)
    assert main(["import", str(book), "--json"]) == 0
    env = json.loads(capsys.readouterr().out)
    assert env["outputs"] == [str(tmp_path / "보고서_수정.dsd")]
    assert [(e["old"], e["new"]) for e in env["edits"]] == [("300", "301")]
    before, after = read_dsd(str(dsd)).contents, read_dsd(env["outputs"][0]).contents
    assert len(before) == len(after) and sum(a != b for a, b in zip(before, after)) == 1


def test_import_refuses_row_insert_in_foot_workbook(tmp_path, capsys):
    _, book = _foot_book(tmp_path, capsys)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(book))
    wb["BS"].insert_rows(9)                           # 표 중간에 행을 끼워 넣음
    wb.save(str(book))
    assert main(["import", str(book), "--json"]) == 2
    env = json.loads(capsys.readouterr().out)
    assert not env["ok"] and env["outputs"] == [] and any("밀렸습니다" in e or "다릅니다" in e for e in env["errors"])
    assert not (tmp_path / "보고서_수정.dsd").exists()


def test_export_workbook_still_imports_the_old_way(tmp_path, capsys):
    p = tmp_path / "a.dsd"
    F.write_dsd(p, F.standard_contents())
    assert main(["export", str(p)]) == 0
    capsys.readouterr()
    assert main(["import", str(tmp_path / "a.xlsx"), "--json"]) == 0
    env = json.loads(capsys.readouterr().out)
    assert "kind" not in env and env["changed"] is False
