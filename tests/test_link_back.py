"""DSD 에서 고친 숫자를 정산표로(link back) - 반영 목록, --apply, 다시 build 하면 같은 DSD(합성 자료)."""
from __future__ import annotations

import json
import os

import openpyxl

from dsdxlsx import srcmap
from dsdxlsx.cli import main
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.edits import apply_edits
from dsdxlsx.link import calc

from fixtures import link_factory as LF
from fixtures.factory import write_dsd
from test_build import _locate, _setup
from test_fill import _skeleton


def _cli(capsys, *argv):
    capsys.readouterr()
    code = main(list(argv))
    out = capsys.readouterr().out
    return code, json.loads(out[out.index("{"):]) if "--json" in argv else out


def _edit(dsd: str, edits: list) -> bytes:
    """DART 편집기에서 고친 것처럼 같은 DSD 파일을 고친다."""
    contents = apply_edits(read_dsd(dsd).contents, edits).contents
    write_dsd(dsd, contents)
    return contents


def _cell(dsd: str, label: str, section: str | None = None, col: str | None = None) -> str:
    return next(c["cell"] for c in srcmap.load(dsd)["cells"] if c["label"].replace(" ", "") == label
                and (section is None or c["section"] == section) and (col is None or col in c["col"]))


def _built(tmp_path, capsys, *extra):
    skel, wp = _setup(tmp_path)
    out = str(tmp_path / "보고서.dsd")
    assert main(["build", skel, wp, "-o", out, "--no-check", *extra]) == 0
    capsys.readouterr()
    return skel, wp, out


def _by(env, cell):
    return next(x for x in env["changes"] if x["cell"] == cell)


# ── 값 칸 ────────────────────────────────────────────

def test_changed_cells_are_listed_with_workbook_cell_and_proposed_value(tmp_path, capsys):
    _skel, wp, out = _built(tmp_path, capsys)
    cash, allow, total = _cell(out, "현금및현금성자산"), _cell(out, "대손충당금"), _cell(out, "합계", "주석1")
    _edit(out, [{"op": "set", "cell": cash, "value": 1600}, {"op": "set", "cell": allow, "value": -150},
                {"op": "set", "cell": total, "value": 70}])
    before = open(wp, "rb").read()
    code, env = _cli(capsys, "link", "back", out, wp, "--json")
    assert code == 1 and env["counts"] == {"value": 2, "calc": 1} and env["unchanged"] == 13 and env["applied"] == 0
    a = _by(env, cash)
    assert (a["old"], a["new"], a["kind"], a["sheet"], a["ref"], a["proposed"]) == (1500.0, 1600, "value", "BS", "C2", 1600)
    assert _by(env, allow)["proposed"] == -150 and _by(env, total)["kind"] == "calc" and _by(env, total)["ref"] is None
    assert open(wp, "rb").read() == before and env["outputs"] == []                # --apply 없이는 아무것도 안 쓴다


def test_apply_writes_value_cells_keeps_backup_and_is_idempotent(tmp_path, capsys):
    _skel, wp, out = _built(tmp_path, capsys)
    cash, total = _cell(out, "현금및현금성자산"), _cell(out, "합계", "주석1")
    _edit(out, [{"op": "set", "cell": cash, "value": 1600}, {"op": "set", "cell": total, "value": 70}])
    code, env = _cli(capsys, "link", "back", out, wp, "--apply", "--json")
    assert code == 1 and env["applied"] == 1 and env["outputs"] == [wp]            # 합계(계산 값)는 사람이 볼 것으로 남는다
    assert openpyxl.load_workbook(wp)["BS"]["C2"].value == 1600
    assert os.path.dirname(env["backup"]) == os.path.join(str(tmp_path), "_build")
    assert openpyxl.load_workbook(env["backup"])["BS"]["C2"].value == 1500         # 쓰기 전 정산표
    assert _by(env, cash)["applied"] is True and "applied" not in _by(env, total)
    code, env = _cli(capsys, "link", "back", out, wp, "--apply", "--json")         # 다시 돌려도 또 쓰지 않는다
    assert env["counts"] == {"done": 1, "calc": 1} and env["applied"] == 0 and env["backup"] is None


def test_dsd_edit_then_back_then_rebuild_gives_the_same_dsd(tmp_path, capsys):
    """끝 기준: DSD 수정 → 정산표 반영 → 다시 build → 고친 DSD 와 같다."""
    skel, wp, out = _built(tmp_path, capsys)
    edited = _edit(out, [{"op": "set", "cell": _cell(out, "현금및현금성자산"), "value": 1600},
                         {"op": "set", "cell": _cell(out, "매출액", col="누적"), "value": 4321}])
    assert main(["build", skel, wp, "-o", out, "--no-check"]) == 3                 # 고친 DSD 는 build 가 덮지 않는다
    code, env = _cli(capsys, "link", "back", out, wp, "--apply", "--json")
    assert (code, env["applied"]) == (0, 2)
    assert main(["build", skel, wp, "-o", out, "--no-check", "--force"]) == 0
    assert read_dsd(out).contents == edited
    assert _cli(capsys, "link", "back", out, wp, "--json")[1]["changes"] == []     # 새 기록 기준으로 바뀐 칸 없음


def test_opposite_sign_workbook_gets_the_sign_back(tmp_path, capsys):
    skel = str(tmp_path / "skel.dsd")
    write_dsd(skel, _skeleton().data)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BS"
    for row in (["계정과목", None, "2026.06.30", "2025.12.31"], ["현금및현금성자산", None, 1500, 1000],
                ["대손충당금", None, 120, 90], ["자산총계", None, 1500, 1000]):
        ws.append(row)
    wp = str(tmp_path / "wp.xlsx")
    wb.save(wp)
    out = str(tmp_path / "out.dsd")
    assert main(["build", skel, wp, "-o", out, "--no-check"]) == 0
    _edit(out, [{"op": "set", "cell": _cell(out, "대손충당금"), "value": -150}])
    code, env = _cli(capsys, "link", "back", out, wp, "--apply", "--json")
    assert (code, env["changes"][0]["proposed"]) == (0, 150) and openpyxl.load_workbook(wp)["BS"]["C3"].value == 150


# ── 쓰지 않는 경우 ───────────────────────────────────

def test_conflict_scaled_recipe_and_blank_are_not_written(tmp_path, capsys):
    skel, _wp = _setup(tmp_path)
    sheet, tid, rid = _locate(skel, tmp_path, "보증금")
    recipe = str(tmp_path / "recipe.json")
    json.dump({"tables": [{"sheet": sheet, "table": tid, "match": "order", "first": rid, "src": "추가", "rows": "3",
                           "cols": {"C": "C"}}]}, open(recipe, "w", encoding="utf-8"), ensure_ascii=False)
    (tmp_path / "b").mkdir()
    _skel, wp, out = _built(tmp_path / "b", capsys, "--recipe", recipe)
    cash, allow, dep, rent = (_cell(out, n) for n in ("현금및현금성자산", "대손충당금", "보증금", "임대수익"))
    _edit(out, [{"op": "set", "cell": cash, "value": 1600}, {"op": "set", "cell": allow, "value": -150},
                {"op": "set", "cell": dep, "value": 80}, {"op": "set", "cell": rent, "value": None}])
    wb = openpyxl.load_workbook(wp)
    wb["BS"]["C2"] = 1700                                                           # 정산표의 그 칸도 사람이 고침
    wb["BS"]["C3"] = -120000                                                        # 원 단위 정산표인 척(아래에서 기록의 배율도 바꿈)
    wb.save(wp)
    mp = srcmap.path_of(out)
    m = json.load(open(mp, encoding="utf-8"))
    next(c for c in m["cells"] if c["cell"] == allow)["scale"] = 0.001
    json.dump(m, open(mp, "w", encoding="utf-8"), ensure_ascii=False)
    before = open(wp, "rb").read()
    code, env = _cli(capsys, "link", "back", out, wp, "--apply", "--json")
    assert code == 1 and env["applied"] == 0 and open(wp, "rb").read() == before
    assert env["counts"] == {"conflict": 1, "scaled": 1, "recipe": 1, "blank": 1}
    assert _by(env, cash)["proposed"] == 1600 and "1,700" in _by(env, cash)["note"]
    assert _by(env, allow)["proposed"] == -150000                                   # 제안 값만
    assert (_by(env, dep)["old"], _by(env, dep)["new"]) == (77, 80)                 # 레시피가 넣은 값 기준


def test_shifted_rows_are_reported_and_nothing_is_applied(tmp_path, capsys):
    _skel, wp, out = _built(tmp_path, capsys)
    cash = _cell(out, "현금및현금성자산")
    row = cash.rsplit(".", 1)[0]
    _edit(out, [{"op": "set", "cell": cash, "value": 1600}, {"op": "insert_row", "after": row, "values": ["외화예금", 50, 0]}])
    before = open(wp, "rb").read()
    code, env = _cli(capsys, "link", "back", out, wp, "--apply", "--json")
    assert code == 1 and env["counts"]["moved"] == 2 and env["counts"]["value"] == 1    # 그 아래 두 행의 번호가 밀림
    assert env["applied"] == 0 and open(wp, "rb").read() == before and any("구조" in w for w in env["warnings"])


def test_unmapped_changes_and_missing_map(tmp_path, capsys):
    _skel, wp, out = _built(tmp_path, capsys)
    cash = _cell(out, "현금및현금성자산")
    prior = cash[:-2] + "%02d" % (int(cash[-2:]) + 1)                               # 같은 행의 전기 칸
    _edit(out, [{"op": "set", "cell": prior, "value": 999}])
    code, env = _cli(capsys, "link", "back", out, wp, "--json")
    assert code == 1 and env["changes"] == [] and [x["cell"] for x in env["unmapped"]] == [prior]
    copy = str(tmp_path / "사본.dsd")
    write_dsd(copy, read_dsd(out).contents)
    assert _cli(capsys, "link", "back", copy, wp)[0] == 2                           # 기록이 옆에 없는 DSD
    code, env = _cli(capsys, "link", "back", copy, wp, "--map", out, "--json")      # 기록을 남긴 DSD 를 알려 주면 된다
    assert code == 1 and len(env["unmapped"]) == 1


# ── link make 정산표: TB 수식 · 합계 수식 · 입력칸 ───

def test_link_workbook_formula_cells_get_suggestions_and_input_cells_get_values(tmp_path, capsys):
    skel = LF.write_skeleton(tmp_path / "skel.dsd")
    conf = str(tmp_path / "conf.json")
    wp = str(tmp_path / "정산표.xlsx")
    assert main(["link", "make", skel, "--tb", LF.write_tb(tmp_path / "tb.json"), "--config", conf, "-o", wp]) == 0
    wb = openpyxl.load_workbook(wp)
    cf = wb["CF"]
    r = next(r for r in range(6, cf.max_row + 1) if cf.cell(r, 2).value == "감가상각비")
    cf.cell(r, 3).value = 100                                                       # 사람이 넣은 현금흐름표 입력칸
    wb.save(wp)
    vals = str(tmp_path / "정산표_값.xlsx")
    calc.freeze(wp, vals)                                                           # 엑셀에서 열어 저장한 것처럼 값으로
    out = str(tmp_path / "보고서.dsd")
    assert main(["build", skel, vals, "--config", conf, "-o", out, "--no-check"]) == 0
    cash, ap, total = (_cell(out, n, "재무상태표") for n in ("현금및현금성자산", "매입채무", "자산총계"))
    dep = _cell(out, "감가상각비", "현금흐름표")
    _edit(out, [{"op": "set", "cell": cash, "value": 600}, {"op": "set", "cell": ap, "value": 700},
                {"op": "set", "cell": total, "value": 2330}, {"op": "set", "cell": dep, "value": 110}])
    code, env = _cli(capsys, "link", "back", out, wp, "--apply", "--json")
    assert code == 1 and env["counts"] == {"tb": 2, "formula": 1, "value": 1} and env["applied"] == 1
    a = _by(env, cash)                                                              # 현금 + 보통예금 → 어느 계정인지는 사람이
    assert a["tb_delta"] == 50 and [x["name"] for x in a["accounts"]] == ["현금", "보통예금"] and "차변 50" in a["note"]
    b = _by(env, ap)                                                                # 부채가 20 늘었다 = 대변 20
    assert b["tb_delta"] == -20 and b["accounts"] == [{"code": "251", "name": "외상매입금"}] and "대변 20" in b["note"]
    assert _by(env, total)["proposed"] is None
    after = openpyxl.load_workbook(wp)
    assert after["CF"].cell(r, 3).value == 110 and str(after["BS"]["C9"].value).startswith("=SUMIF(")   # 수식은 그대로
    assert any("다른 파일" in w for w in env["warnings"])                           # 기록은 값으로 바꾼 사본에서 만든 것


# ── 같은 정산표 칸을 쓰는 DSD 칸, 값이 같은 후보 ─────

def _remap(out, cell, **changes):
    mp = srcmap.path_of(out)
    m = json.load(open(mp, encoding="utf-8"))
    next(c for c in m["cells"] if c["cell"] == cell).update(changes)
    json.dump(m, open(mp, "w", encoding="utf-8"), ensure_ascii=False)


def test_two_dsd_cells_from_one_workbook_cell(tmp_path, capsys):
    _skel, wp, out = _built(tmp_path, capsys)
    cash, total = _cell(out, "현금및현금성자산"), _cell(out, "자산총계")
    _remap(out, total, sources=[{"sheet": "BS", "cell": "C2"}])                    # 두 칸이 같은 정산표 칸에서 온다
    saved = read_dsd(out).contents
    _edit(out, [{"op": "set", "cell": cash, "value": 1600}])
    code, env = _cli(capsys, "link", "back", out, wp, "--json")
    assert _by(env, cash)["kind"] == "value" and _by(env, cash)["also"] == 1 and "DSD 칸 1개도" in _by(env, cash)["note"]
    write_dsd(out, saved)
    _edit(out, [{"op": "set", "cell": cash, "value": 1600}, {"op": "set", "cell": total, "value": 1700}])
    code, env = _cli(capsys, "link", "back", out, wp, "--apply", "--json")         # 서로 다른 값으로 고침 - 어느 쪽도 쓰지 않는다
    assert code == 1 and env["counts"] == {"conflict": 2} and env["applied"] == 0
    assert openpyxl.load_workbook(wp)["BS"]["C2"].value == 1500


def test_tie_pick_and_zero_cell_are_flagged(tmp_path, capsys):
    _skel, wp, out = _built(tmp_path, capsys)
    cash, allow = _cell(out, "현금및현금성자산"), _cell(out, "대손충당금")
    _remap(out, cash, how="같은 값 2개")                                            # 값이 같은 두 칸 중 하나를 골랐던 칸
    _remap(out, allow, value=0)                                                     # 0 으로 채웠던 칸
    wb = openpyxl.load_workbook(wp)
    wb["BS"]["C3"] = 0
    wb.save(wp)
    saved = _edit(out, [{"op": "set", "cell": allow, "value": 0}])
    assert _cli(capsys, "link", "back", out, wp, "--json")[1]["changes"] == []
    write_dsd(out, saved)
    _edit(out, [{"op": "set", "cell": cash, "value": 1600}, {"op": "set", "cell": allow, "value": -30}])
    code, env = _cli(capsys, "link", "back", out, wp, "--apply", "--json")
    assert _by(env, cash)["kind"] == "tie" and "applied" not in _by(env, cash)
    assert _by(env, allow)["applied"] is True and "0 이던 칸" in _by(env, allow)["note"]
    assert openpyxl.load_workbook(wp)["BS"]["C2"].value == 1500 and openpyxl.load_workbook(wp)["BS"]["C3"].value == -30
