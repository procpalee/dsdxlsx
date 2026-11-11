"""칸별 출처 기록(srcmap) - fill · build 가 DSD 칸마다 정산표 칸 · 배율 · 부호를 남긴다(합성 자료)."""
from __future__ import annotations

import hashlib
import json
import os

import openpyxl

from dsdxlsx import build as B
from dsdxlsx import fill as F
from dsdxlsx import recon as R
from dsdxlsx import srcmap
from dsdxlsx.cli import main
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.parse import parse_contents

from fixtures.factory import write_dsd
from test_build import _locate, _setup
from test_fill import _skeleton


def _built(tmp_path, recipe=None):
    skel, wp = _setup(tmp_path)
    out = str(tmp_path / "out.dsd")
    rep = B.build(skel, wp, out, recipe=recipe or {}, fill_conf={}, side_sheets=R.load_workpaper(wp))
    path, counts = srcmap.save(rep.fill, dsd=out, skeleton=skel, workbook=wp, side="별도")
    return skel, wp, out, path, counts


def _by_label(m, label, col=None):
    return next(c for c in m["cells"] if c["label"].replace(" ", "") == label and (col is None or col in c["col"]))


def test_map_records_source_cell_scale_and_sign(tmp_path):
    _skel, wp, out, path, counts = _built(tmp_path)
    m = srcmap.load(out)
    assert path == os.path.join(str(tmp_path), "_build", "out.map.json") and m["schema"] == "dsdxlsx.map/1"
    assert counts == m["counts"] and counts["cells"] == 16 and counts["overridden"] == counts["missing"] == 0
    cash = _by_label(m, "현금및현금성자산")
    assert (cash["section"], cash["value"], cash["sources"], cash["scale"], cash["sign"]) == (
        "재무상태표", 1500.0, [{"sheet": "BS", "cell": "C2"}], 1.0, 1)
    assert _by_label(m, "매출액", "누적")["sources"] == [{"sheet": "IS", "cell": "D3"}]      # 열마다 따로
    total = _by_label(m, "합계", "당반기말")
    assert total["sources"] == [] and total["scale"] is None and "위 칸 합" in total["how"]   # 정산표 칸 없이 계산한 값
    sha = hashlib.sha256(open(wp, "rb").read()).hexdigest()
    assert m["workbook"] == {"file": os.path.basename(wp), "sha256": sha} and m["recipe"] is None
    assert m["dsd"]["sha256"] == hashlib.sha256(open(out, "rb").read()).hexdigest()


def test_value_equals_sign_times_sources_times_scale(tmp_path):
    """기록한 대로 정산표 칸에서 다시 계산하면 DSD 에 쓴 값이 나온다 - 되돌릴 때 쓰는 식."""
    _skel, wp, out, _path, _counts = _built(tmp_path)
    wb = openpyxl.load_workbook(wp, data_only=True)
    final = srcmap.final_values(out)
    n = 0
    for c in srcmap.load(out)["cells"]:
        assert final[c["cell"]] == c["value"]                                    # 칸 ID 가 최종 DSD 의 그 칸
        if c["sources"]:
            total = sum(wb[s["sheet"]][s["cell"]].value for s in c["sources"])
            assert round(c["sign"] * total * c["scale"], 2) == c["value"]
            n += 1
    assert n == 14


def test_opposite_sign_convention_is_recorded(tmp_path):
    """정산표가 차감 계정을 양수로 적어도 DSD 전기 값과 맞춰 부호를 정하고, 그 부호를 남긴다."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BS"
    for row in (["계정과목", None, "2026.06.30", "2025.12.31"], ["현금및현금성자산", None, 1500, 1000],
                ["대손충당금", None, 120, 90], ["자산총계", None, 1500, 1000]):
        ws.append(row)
    wp = str(tmp_path / "wp.xlsx")
    wb.save(wp)
    _ops, rep = F.fill_ops(parse_contents(_skeleton().data), R.load_workpaper(wp))
    item = next(x for x in rep.items if x.label.strip() == "대손충당금" and x.verdict == "채움")
    assert (item.value, item.sign, item.scale, item.sources) == (-120.0, -1, 1.0, [("BS", "C3")])
    assert item.cell.startswith("t") and ".c" in item.cell


def test_recipe_override_is_marked(tmp_path):
    skel, _wp = _setup(tmp_path)
    sheet, tid, rid = _locate(skel, tmp_path, "보증금")
    recipe = {"tables": [{"sheet": sheet, "table": tid, "match": "order", "first": rid, "src": "추가", "rows": "3",
                          "cols": {"C": "C"}}]}
    (tmp_path / "b").mkdir()
    _skel, _wp2, out, _path, counts = _built(tmp_path / "b", recipe)
    c = _by_label(srcmap.load(out), "보증금")
    assert counts["overridden"] == 1 and (c["state"], c["value"], c["final"]) == ("overridden", 12.0, 77)


def test_load_missing_or_foreign_file(tmp_path):
    assert srcmap.load(str(tmp_path / "none.dsd")) is None
    p = tmp_path / "x.json"
    p.write_text('{"schema": "other/1"}', encoding="utf-8")
    assert srcmap.load(str(p)) is None


def _cli(capsys, *argv):
    capsys.readouterr()
    code = main(list(argv))
    return code, json.loads(capsys.readouterr().out)


def test_cli_fill_and_build_write_map(tmp_path, capsys):
    skel, wp = _setup(tmp_path)
    out = str(tmp_path / "f.dsd")
    _code, env = _cli(capsys, "fill", skel, wp, "-o", out, "--json")
    assert env["map"] == srcmap.path_of(out) and env["map"] in env["outputs"]
    assert srcmap.load(out)["counts"]["cells"] == env["counts"]["채움"]

    out2 = str(tmp_path / "b.dsd")
    code, env = _cli(capsys, "build", skel, wp, "-o", out2, "--no-check", "--json")
    assert code == 0 and env["outputs"] == [out2, srcmap.path_of(out2)]
    assert env["map_counts"] == srcmap.load(out2)["counts"] and env["map_counts"]["overridden"] == 0
    assert read_dsd(out2).contents == read_dsd(out).contents                   # 레시피 없는 build = fill
    write_dsd(out2, read_dsd(out2).contents)
    assert srcmap.load(out2)["dsd"]["sha256"] != hashlib.sha256(open(out2, "rb").read()).hexdigest()   # 그 뒤 바뀐 DSD 를 가린다
