"""status - build · fill 뒤에 DSD · 정산표 · 뼈대 가운데 무엇이 바뀌었고 다음에 무엇을 돌릴지(합성 자료)."""
from __future__ import annotations

import json
import os

import openpyxl

from dsdxlsx import srcmap
from dsdxlsx.cli import main
from dsdxlsx.dsd_io import read_dsd

from fixtures.factory import write_dsd
from test_build import _setup
from test_link_back import _cell, _cli, _edit


def _built(tmp_path, capsys, cmd="build"):
    skel, wp = _setup(tmp_path)
    out = str(tmp_path / "보고서.dsd")
    extra = ["--no-check"] if cmd == "build" else []
    main([cmd, skel, wp, "-o", out, *extra])
    capsys.readouterr()
    return skel, wp, out


def _states(env):
    return {k: v["state"] for k, v in env["files"].items()}


def test_in_sync_right_after_build(tmp_path, capsys):
    _skel, wp, out = _built(tmp_path, capsys)
    code, env = _cli(capsys, "status", out, "--json")
    assert code == 0 and env["in_sync"] is True and env["made_by"] == "build"
    assert _states(env) == {"dsd": "same", "skeleton": "same", "workbook": "same", "recipe": None}
    assert env["files"]["workbook"]["file"] == os.path.basename(wp) and "할 일이 없습니다" in env["next"][0]
    assert env["dsd_changes"] is None and env["workbook_changes"] is None


def test_workbook_changed_says_build_again(tmp_path, capsys):
    skel, wp, out = _built(tmp_path, capsys)
    wb = openpyxl.load_workbook(wp)
    wb["BS"]["C2"] = 1700
    wb["BS"]["A9"] = "메모"                                                       # 출처가 아닌 칸은 세지 않는다
    wb.save(wp)
    code, env = _cli(capsys, "status", out, "--json")
    assert code == 1 and _states(env)["workbook"] == "changed" and _states(env)["dsd"] == "same"
    assert env["workbook_changes"] == {"cells": 1, "no_values": 0}
    assert "정산표" in env["next"][0] and "build" in env["next"][0] and "1개" in env["next"][0]
    assert main(["build", skel, wp, "-o", out, "--no-check"]) == 0                 # 안내대로 다시 만들면
    assert _cli(capsys, "status", out, "--json")[1]["in_sync"] is True


def test_dsd_changed_says_link_back_then_rebuild(tmp_path, capsys):
    skel, wp, out = _built(tmp_path, capsys)
    _edit(out, [{"op": "set", "cell": _cell(out, "현금및현금성자산"), "value": 1600},
                {"op": "set", "cell": _cell(out, "합계", "주석1"), "value": 70}])
    code, env = _cli(capsys, "status", out, "--json")
    assert code == 1 and _states(env) == {"dsd": "changed", "skeleton": "same", "workbook": "same", "recipe": None}
    assert env["dsd_changes"] == {"value": 1, "calc": 1}
    assert "link back" in env["next"][0] and "2칸" in env["next"][0] and "값 칸 1개는 --apply" in env["next"][0]

    main(["link", "back", out, wp, "--apply"])                                     # 값 칸을 정산표에 쓰면 양쪽이 바뀐 상태
    capsys.readouterr()
    code, env = _cli(capsys, "status", out, "--json")
    assert _states(env)["workbook"] == "changed" and env["dsd_changes"] == {"done": 1, "calc": 1}
    assert "link back 으로 확인" in env["next"][0]                                  # 합계(계산 값)는 아직 사람이 볼 것

    assert main(["build", skel, wp, "-o", out, "--no-check", "--force"]) == 0
    assert _cli(capsys, "status", out, "--json")[1]["in_sync"] is True


def test_both_changed_reports_conflict_and_reflected_case(tmp_path, capsys):
    _skel, wp, out = _built(tmp_path, capsys)
    cash = _cell(out, "현금및현금성자산")
    saved = read_dsd(out).contents
    _edit(out, [{"op": "set", "cell": cash, "value": 1600}])
    wb = openpyxl.load_workbook(wp)
    wb["BS"]["C2"] = 1700                                                           # 같은 칸을 정산표에서도 다른 값으로
    wb.save(wp)
    code, env = _cli(capsys, "status", out, "--json")
    assert code == 1 and env["dsd_changes"] == {"conflict": 1} and "서로 다른 값으로 고친 칸 1개" in env["next"][0]
    wb["BS"]["C2"] = 1600                                                           # 정산표에 같은 값이 들어가 있으면
    wb.save(wp)
    env = _cli(capsys, "status", out, "--json")[1]
    assert env["dsd_changes"] == {"done": 1} and "반영돼 있습니다" in env["next"][0] and "build --force" in env["next"][0]
    write_dsd(out, saved)


def test_shifted_rows_unmapped_only_and_missing_files(tmp_path, capsys):
    skel, wp, out = _built(tmp_path, capsys)
    cash = _cell(out, "현금및현금성자산")
    saved = read_dsd(out).contents
    _edit(out, [{"op": "insert_row", "after": cash.rsplit(".", 1)[0], "values": ["외화예금", 50, 0]}])
    env = _cli(capsys, "status", out, "--json")[1]
    assert env["dsd_changes"]["moved"] == 2 and "뼈대 DSD 에 옮기고" in env["next"][0]

    write_dsd(out, saved)
    prior = cash[:-2] + "%02d" % (int(cash[-2:]) + 1)
    _edit(out, [{"op": "set", "cell": prior, "value": 999}])                        # 전기 칸만 고침 - 정산표에 짝이 없다
    env = _cli(capsys, "status", out, "--json")[1]
    assert env["dsd_changes"] == {} and env["unmapped"] == 1 and "출처 기록에 없는 곳만" in env["next"][0]

    write_dsd(out, saved)
    moved = str(tmp_path / "옮긴_정산표.xlsx")
    os.replace(wp, moved)
    os.remove(skel)
    code, env = _cli(capsys, "status", out, "--json")
    assert code == 1 and _states(env)["workbook"] == "missing" and _states(env)["skeleton"] == "missing"
    assert "못 찾았습니다" in env["next"][0] and env["in_sync"] is False
    env = _cli(capsys, "status", out, "--workbook", moved, "--json")[1]            # 옮긴 경로를 알려 주면 견준다
    assert _states(env)["workbook"] == "same"


def test_fill_made_map_and_no_map(tmp_path, capsys):
    _skel, wp, out = _built(tmp_path, capsys, "fill")
    assert srcmap.load(out)["made_by"] == "fill"
    wb = openpyxl.load_workbook(wp)
    wb["BS"]["C2"] = 1700
    wb.save(wp)
    env = _cli(capsys, "status", out, "--json")[1]
    assert "fill 를 다시" in env["next"][0]
    plain = str(tmp_path / "기록없음.dsd")
    write_dsd(plain, read_dsd(out).contents)
    assert _cli(capsys, "status", plain)[0] == 2
    text = _cli(capsys, "status", out)[1]                                           # 사람용 출력
    assert "정산표" in text and "바뀜" in text and "→ " in text
