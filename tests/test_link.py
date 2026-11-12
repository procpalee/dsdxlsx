"""정산표와 DSD 연결(link) - 뼈대 DSD 행 그대로 정산표 만들기 · TB 금액 다시 넣기 · 왕복 시험(합성 자료)."""
from __future__ import annotations

import csv
import json

import openpyxl
import pytest

from dsdxlsx import docview
from dsdxlsx.cli import main
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.guard import InputError
from dsdxlsx.link import calc, core, roundtrip, tbsrc
from dsdxlsx.parse import parse_contents

from fixtures import link_factory as LF
from fixtures.factory import document, p, statement_title_table, table, title, write_dsd


def _doc(path: str) -> dict:
    return docview.doc_dict(parse_contents(read_dsd(path).contents), source=path)


def _made(tmp_path, rows=LF.TB):
    skel = LF.write_skeleton(tmp_path / "skel.dsd")
    tb, info = tbsrc.load(LF.write_tb(tmp_path / "tb.json", rows))
    wb, summary = core.make(_doc(skel), tb, info, company="합성테스트(주)")
    out = str(tmp_path / "정산표.xlsx")
    wb.save(out)
    return skel, out, summary


def _col(ws, col: int) -> dict:
    """머리글 아래 B열 이름 → 그 행의 col 열 값."""
    return {ws.cell(r, 2).value: ws.cell(r, col).value for r in range(core.HDR_ROW + 1, ws.max_row + 1)}


# ── 만들기 ───────────────────────────────────────────

def test_make_follows_dsd_rows_and_links_tb(tmp_path):
    _skel, out, summary = _made(tmp_path)
    assert summary["sheets"] == ["TB", "BS", "IS", "SCE", "CF", "주석", "연결점검"]
    assert summary["to_check"] == [] and summary["unlinked_tb"] == []
    wb = openpyxl.load_workbook(out)
    bs = wb["BS"]
    assert [bs.cell(r, 2).value for r in range(6, 10)] == ["자산", "Ⅰ.유동자산", "(1) 당좌자산", "현금및현금성자산"]
    assert _col(bs, 3)["현금및현금성자산"] == "=SUMIF(TB!$E:$E,$G9,TB!$H:$H)"
    assert _col(bs, 3)["매입채무"].startswith("=-SUMIF(")              # 부채는 부호를 뒤집어 양수로
    assert _col(bs, 7)["대손충당금"] == "대손충당금(매출채권)"          # 차감 계정은 바로 위 자산과 짝
    assert _col(bs, 3)["자산총계"] == "=C7+C14"
    assert "IS!C24" in _col(bs, 3)["미처분이익잉여금"]                  # 이월 잔액 + 당기순이익
    tb = wb["TB"]
    assert {tb.cell(r, 5).value for r in range(6, 8)} == {"현금및현금성자산"}   # 현금 · 보통예금이 한 DSD 행으로
    assert tb["H6"].value == "=F6+G6"
    assert wb[core.LINK_SHEET].sheet_state == "hidden"
    assert core.read_link(wb)["note_pairs"] == [["C8", "D8"], ["C9", "D9"], ["C10", "D10"]]


def test_income_statement_inventory_lines_come_from_balance_sheet(tmp_path):
    _skel, out, _ = _made(tmp_path)
    is_ = _col(openpyxl.load_workbook(out)["IS"], 3)
    assert is_["기초상품재고액"] == "=BS!D13" and is_["기말상품재고액"] == "=-BS!C13"
    assert is_["당기상품매입액"] == "=C9-C10-C12"
    assert is_["Ⅸ.당기순이익"] == "=C22-C23"


def test_statement_values_calculated_without_excel(tmp_path):
    _skel, out, _ = _made(tmp_path)
    c = calc.Calc(openpyxl.load_workbook(out))
    bs = openpyxl.load_workbook(out)["BS"]
    row = {bs.cell(r, 2).value: r for r in range(6, bs.max_row + 1)}
    assert c.value("BS", row["자산총계"], 3) == 2280                    # 당기 TB
    assert c.value("BS", row["부채및자본총계"], 3) == 2280
    assert c.value("BS", row["자산총계"], 5) == 2000                    # 전기 TB = 뼈대 전기 칸
    assert all(c.value("BS", r, 6) in (0, None) for r in row.values())   # 차이 열 0
    assert c.value("IS", 24, 3) == 300 and c.value("CF", 19, 3) == 550   # 당기순이익, 기말의현금 = BS 현금


def test_account_without_dsd_row_is_reported(tmp_path):
    rows = LF.TB + [("120", "미수금", "미수금", 40, 10), ("255", "미지급금", "미지급금", -40, -10)]
    _skel, out, summary = _made(tmp_path, rows)
    assert sorted(summary["unlinked_tb"]) == ["미수금", "미지급금"]
    assert {(x["sheet"], x["label"]) for x in summary["to_check"]} == {("TB", "미수금"), ("TB", "미지급금")}
    chk = openpyxl.load_workbook(out)["연결점검"]
    assert chk["D6"].value == "DSD 행에 짝 없음(전기 잔액 있음)"
    tb = openpyxl.load_workbook(out)["TB"]
    last = tb.max_row
    assert tb.cell(last, 5).value is None and tb.cell(last, 5).fill.fgColor.rgb.endswith("FFFF00")


def test_make_needs_balance_sheet_and_income_statement(tmp_path):
    path = tmp_path / "bs_only.dsd"
    write_dsd(path, document(p("문서"), title("재무제표", 2), statement_title_table("재 무 상 태 표"),
                             table([LF.HEAD_BS], [["자산총계", "", "10"]])))
    tb, info = tbsrc.load(LF.write_tb(tmp_path / "tb.json"))
    with pytest.raises(InputError, match="손익계산서"):
        core.make(_doc(str(path)), tb, info)


# ── 왕복 시험 ────────────────────────────────────────

def test_roundtrip_every_current_cell_reaches_dsd(tmp_path):
    skel, out, _ = _made(tmp_path)
    res = roundtrip.run(skel, out, notes=True)
    assert (res["diff"], res["blank"], res["note_diff"], res["note_blank"]) == ([], [], [], [])
    assert res["same"] == 49 and res["note_same"] == 3


def test_roundtrip_reports_broken_link(tmp_path):
    skel, out, _ = _made(tmp_path)
    wb = openpyxl.load_workbook(out)
    tb = wb["TB"]
    row = next(r for r in range(6, tb.max_row + 1) if tb.cell(r, 3).value == "외상매출금")
    tb.cell(row, 5).value = None                                          # 사람이 'DSD 행'을 지운 경우
    wb.save(out)
    res = roundtrip.run(skel, out)
    bad = {x["label"] for x in res["diff"] + res["blank"]}
    assert "매출채권" in bad and "자산총계" in bad
    assert res["note_same"] == 0 and res["note_diff"] == []               # --notes 없이는 주석을 보지 않는다


def test_roundtrip_leaves_originals_untouched(tmp_path):
    skel, out, _ = _made(tmp_path)
    before = (open(skel, "rb").read(), open(out, "rb").read())
    roundtrip.run(skel, out, notes=True, workdir=str(tmp_path / "work"))
    assert (open(skel, "rb").read(), open(out, "rb").read()) == before
    assert (tmp_path / "work" / "시험_결과.dsd").exists()


# ── 금액 다시 넣기 ───────────────────────────────────

def test_refresh_updates_amounts_keeps_manual_row_and_appends_new_account(tmp_path):
    _skel, out, _ = _made(tmp_path)
    wb = openpyxl.load_workbook(out)
    tb = wb["TB"]
    row = next(r for r in range(6, tb.max_row + 1) if tb.cell(r, 3).value == "보통예금")
    tb.cell(row, 5).value = "사람이 고친 행"
    n_rows = tb.max_row
    new = [(c, n, s, pri, cur + 7) for c, n, s, pri, cur in LF.TB] + [("131", "선급금", "선급금", 0, 30)]
    tb_rows, _info = tbsrc.load(LF.write_tb(tmp_path / "tb2.json", new))
    res = core.refresh(wb, tb_rows)
    assert res == {"accounts": 20, "added": ["131"]}
    assert tb.cell(row, 5).value == "사람이 고친 행" and tb.cell(row, 6).value == 407
    assert tb.max_row == n_rows + 1
    assert [tb.cell(n_rows + 1, c).value for c in (2, 3, 4, 5, 6, 8)] == ["131", "선급금", "선급금", None, 30, "=F25+G25"]


# ── TB 읽기 ──────────────────────────────────────────

def _plain(rows):
    return [(r["code"], r["name"], r["settle"], r["pri"], r["cur"]) for r in rows]


def test_tb_table_in_xlsx_and_csv(tmp_path):
    want = _plain(tbsrc.load(LF.write_tb(tmp_path / "tb.json"))[0])
    head = ["코드", "계정과목", "정산 과목", "회사제시", "수정사항", "2026.12.31", "2025.12.31"]
    body = [[c, n, s, cur, 0, cur, pri] for c, n, s, pri, cur in LF.TB]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TB"
    ws.append(["합성테스트(주)"])
    for line in [head] + body:
        ws.append(line)
    wb.save(tmp_path / "tb.xlsx")
    rows, info = tbsrc.load(str(tmp_path / "tb.xlsx"))
    assert _plain(rows) == want and (info["cur"], info["pri"]) == ("2026.12.31", "2025.12.31")
    with open(tmp_path / "tb.csv", "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows([["계정과목", "당기", "전기"]] + [[n, cur, pri] for _c, n, _s, pri, cur in LF.TB])
    rows, info = tbsrc.load(str(tmp_path / "tb.csv"))
    assert [(r["code"], r["settle"], r["pri"]) for r in rows[:2]] == [("현금", "현금", 100), ("보통예금", "보통예금", 200)]
    assert (info["cur"], info["pri"]) == ("당기", "전기")


def test_tb_account_folder(tmp_path):
    d = tmp_path / "eng" / "_engagement" / "tb"
    d.mkdir(parents=True)

    def w(name, rows):
        with open(d / name, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows(rows)

    w("levels.csv", [["ID", "이름", "역할"], ["L1", "대분류", "구분"], ["L4", "정산", "표시과목"], ["L5", "입력", "계정"]])
    w("coa.csv", [["코드", "이름", "Level", "상위코드"], ["C1", "자산", "L1", ""], ["F001", "현금및현금성자산", "L4", "C1"],
                  ["101", "현금", "L5", "F001"], ["103", "보통예금", "L5", "F001"]])
    w("tb.csv", [["기준일", "코드", "금액"], ["2026.12.31", "101", "150"], ["2025.12.31", "101", "100"],
                 ["2025.12.31", "103", "200"]])
    w("aje.csv", [["번호", "기준일", "코드", "차변", "대변", "적요", "작성자"], ["1", "2026.12.31", "103", "5", "0", "", ""]])
    (d / "tb.json").write_text(json.dumps({"기간": {"당기": "2026.12.31", "전기": "2025.12.31"}}, ensure_ascii=False),
                               encoding="utf-8")
    rows, info = tbsrc.load(str(tmp_path / "eng"))                        # 건 폴더를 줘도 그 안에서 찾는다
    assert [(r["code"], r["settle"], r["group"], r["cur"], r["aje"], r["pri"]) for r in rows] == [
        ("101", "현금및현금성자산", "자산", 150, 0, 100), ("103", "현금및현금성자산", "자산", 0, 5, 200)]
    assert info["cur"] == "2026.12.31"


def test_tb_consolidated_prefixes_and_adjustments(tmp_path):
    a = LF.write_tb(tmp_path / "p.json", LF.TB[:2])
    b = LF.write_tb(tmp_path / "s.json", LF.TB[2:3])
    (tmp_path / "adj.json").write_text(json.dumps({"adj": [{"name": "영업권", "pri": 30}], "map": {"S1.108": "매출채권및기타채권"}},
                                                  ensure_ascii=False), encoding="utf-8")
    rows, info = tbsrc.load_consolidated(a, [b], str(tmp_path / "adj.json"))
    assert [r["code"] for r in rows] == ["P.101", "P.103", "S1.108", "ADJ.1"]
    assert rows[2]["settle"] == "매출채권및기타채권" and rows[3]["pri"] == 30 and rows[3]["adj"] is True
    assert info["cur_label"] == "합산"


def test_tb_bad_inputs(tmp_path):
    with pytest.raises(InputError, match="찾을 수 없습니다"):
        tbsrc.load(str(tmp_path / "none.json"))
    (tmp_path / "x.csv").write_text("이름,금액\n현금,1\n", encoding="utf-8")
    with pytest.raises(InputError, match="계정과목"):
        tbsrc.load(str(tmp_path / "x.csv"))


# ── 계산기 ───────────────────────────────────────────

def test_calc_formulas():
    wb = openpyxl.Workbook()
    a = wb.active
    a.title = "TB"
    for i, (k, v) in enumerate([("현금", 100), ("매출채권", 250), ("현금", 5), ("매출원가", 40)], 6):
        a.cell(i, 5, k)
        a.cell(i, 8, v)
    b = wb.create_sheet("BS")
    b["G6"] = "현금"
    b["C6"] = "=SUMIF(TB!$E:$E,$G6,TB!$H:$H)"
    b["C7"] = '=-SUMIF(TB!$E:$E,"매출*",TB!$H:$H)+C6'
    b["C8"] = "=-(C6+C7)-SUM(C6:C7)*2/4"
    b["C9"] = "='TB'!H6+BS!C6+C10"                                         # 빈 칸은 0
    c = calc.Calc(wb)
    assert [c.value("BS", r, 3) for r in (6, 7, 8, 9)] == [105, -185, 120, 205]
    b["C11"] = "=VLOOKUP(1,A1:B2,2,0)"
    with pytest.raises(calc.CalcError, match="VLOOKUP"):
        calc.Calc(wb).value("BS", 11, 3)
    b["C12"], b["C13"] = "=C13", "=C12"
    with pytest.raises(calc.CalcError, match="순환"):
        calc.Calc(wb).value("BS", 12, 3)


# ── 명령 ─────────────────────────────────────────────

def _cli(capsys, *argv):
    capsys.readouterr()
    code = main(list(argv))
    return code, capsys.readouterr().out


def test_cli_make_test_refresh(tmp_path, capsys):
    skel = LF.write_skeleton(tmp_path / "skel.dsd")
    tbp = LF.write_tb(tmp_path / "tb.json")
    conf = str(tmp_path / "conf.json")
    code, out = _cli(capsys, "link", "make", skel, "--tb", tbp, "--company", "합성테스트(주)", "--config", conf, "--json")
    env = json.loads(out)
    wp = str(tmp_path / "skel_정산표.xlsx")
    assert (code, env["command"], env["exit_code"], env["outputs"]) == (0, "link make", 0, [wp, conf])
    assert env["side"] == "별도" and env["accounts"] == 19 and env["to_check"] == []
    assert json.load(open(conf, encoding="utf-8"))["별도"]["재무제표"]["PL"] == ["IS"]

    code, out = _cli(capsys, "link", "test", skel, wp, "--config", conf, "--notes", "--json")
    env = json.loads(out)
    assert (code, env["same"], env["diff"], env["blank"], env["note_same"]) == (0, 49, [], [], 3)

    tb2 = LF.write_tb(tmp_path / "tb2.json", LF.TB + [("131", "선급금", "선급금", 0, 30)])
    code, out = _cli(capsys, "link", "refresh", wp, "--tb", tb2, "--json")
    env = json.loads(out)
    assert (code, env["accounts"], env["added"]) == (1, 20, ["131"])       # 새 계정 = 확인할 것

    code, _ = _cli(capsys, "link", "make", skel, "--tb", tbp)              # refresh 가 손댄 파일 = 도구가 만든 그대로
    assert code == 0


def test_cli_make_stops_on_edited_workbook_and_bad_input(tmp_path, capsys):
    skel = LF.write_skeleton(tmp_path / "skel.dsd")
    tbp = LF.write_tb(tmp_path / "tb.json")
    wp = str(tmp_path / "out.xlsx")
    assert _cli(capsys, "link", "make", skel, "--tb", tbp, "-o", wp)[0] == 0
    wb = openpyxl.load_workbook(wp)
    wb["TB"]["E6"] = "사람이 고친 행"
    wb.save(wp)
    assert _cli(capsys, "link", "make", skel, "--tb", tbp, "-o", wp)[0] == 3
    assert openpyxl.load_workbook(wp)["TB"]["E6"].value == "사람이 고친 행"
    assert _cli(capsys, "link", "make", skel, "--tb", tbp, "-o", wp, "--force")[0] == 0
    assert _cli(capsys, "link", "make", skel, "--tb", str(tmp_path / "none.json"), "-o", wp, "--force")[0] == 2
    plain = openpyxl.Workbook()
    plain.active.title = "TB"
    plain.save(tmp_path / "plain.xlsx")
    assert _cli(capsys, "link", "refresh", str(tmp_path / "plain.xlsx"), "--tb", tbp)[0] == 2


def test_cli_make_reports_unlinked_account(tmp_path, capsys):
    skel = LF.write_skeleton(tmp_path / "skel.dsd")
    tbp = LF.write_tb(tmp_path / "tb.json", LF.TB + [("120", "미수금", "미수금", 40, 10), ("255", "미지급금", "미지급금", -40, -10)])
    code, out = _cli(capsys, "link", "make", skel, "--tb", tbp, "--json")
    env = json.loads(out)
    assert code == 1 and env["ok"] is True and sorted(env["unlinked_tb"]) == ["미수금", "미지급금"]


# ── 서식, 머리말 이름표 ──────────────────────────────

def _sides(cell):
    b = cell.border
    return tuple(s.style if s is not None else None for s in (b.left, b.right, b.top, b.bottom))


def test_sheet_look_horizontal_rules_and_number_format(tmp_path):
    _skel, out, _ = _made(tmp_path)
    bs = openpyxl.load_workbook(out)["BS"]
    last = bs.max_row
    assert _sides(bs["B5"]) == (None, None, "medium", "thin") and bs["B5"].font.b          # 머리글: 위 굵게, 세로줄 없음
    assert _sides(bs["C9"]) == (None, None, "thin", "thin")
    assert _sides(bs.cell(last, 3)) == (None, None, "thin", "medium")                      # 마지막 줄 아래 굵게
    assert bs["C9"].number_format == "#,###,##0;[Red](#,###,##0);-" and bs["C9"].alignment.horizontal == "right"
    assert bs["B9"].alignment.horizontal == "left" and bs["G9"].number_format == "General"
    assert [bs.cell(2, c).value for c in range(2, 8)] == ["회사명", "합성테스트(주)", "번호", None, "시트", "재무상태표"]
    assert [bs.cell(3, c).value for c in (2, 3, 4, 5)] == ["기준일", "2026.12.31", "문서", "정산표"]
    assert _sides(bs["B2"]) == (None, "thin", "medium", "thin") and _sides(bs["G3"]) == ("thin", None, "thin", "medium")
    cf = openpyxl.load_workbook(out)["CF"]
    assert cf["C9"].fill.fgColor.rgb.endswith("FFFBEF") and _sides(cf["C9"]) == (None, None, "thin", "thin")   # 입력칸도 줄은 그대로
    note = openpyxl.load_workbook(out)["주석"]
    assert note["B5"].fill.fgColor.rgb.endswith("393939") and note["B5"].font.color.rgb.endswith("FFFFFF")


def test_header_labels_option(tmp_path, capsys):
    skel = LF.write_skeleton(tmp_path / "skel.dsd")
    tbp = LF.write_tb(tmp_path / "tb.json")
    wp = str(tmp_path / "w.xlsx")
    code, _ = _cli(capsys, "link", "make", skel, "--tb", tbp, "-o", wp, "--ref", "8110A", "--title", "정산표",
                   "--labels", "회사명,조서번호,시트명,결산일,조서명,작성자")
    ws = openpyxl.load_workbook(wp)["TB"]
    assert code == 0 and [ws.cell(2, c).value for c in (4, 5, 6)] == ["조서번호", "8110A", "시트명"]
    assert [ws.cell(3, c).value for c in (2, 4, 6)] == ["결산일", "조서명", "작성자"]
    assert _cli(capsys, "link", "make", skel, "--tb", tbp, "-o", wp, "--force", "--labels", "회사명,번호")[0] == 2


# ── 연결 기록이 없는 정산표(이전 방식으로 만든 것) ──

def test_roundtrip_on_workbook_without_link_record(tmp_path, capsys):
    skel, out, _ = _made(tmp_path)
    wb = openpyxl.load_workbook(out)
    pairs = core.read_link(wb)["note_pairs"]
    del wb[core.LINK_SHEET]
    wb.save(out)
    code, text = _cli(capsys, "link", "test", skel, out, "--notes", "--json")
    env = json.loads(text)
    assert env["side"] == "별도" and env["diff"] == [] and env["blank"] == []
    assert code == 1 and len(env["note_blank"]) == 3                                      # 주석 칸 짝을 몰라 비어 있음
    assert any("주석 칸 짝" in w for w in env["warnings"])

    pj = tmp_path / "dsd_link.json"
    pj.write_text(json.dumps({"주석 당기-전기 칸": pairs}, ensure_ascii=False), encoding="utf-8")
    code, text = _cli(capsys, "link", "test", skel, out, "--notes", "--pairs", str(pj), "--json")
    env = json.loads(text)
    assert (code, env["note_same"], env["note_blank"], env["warnings"]) == (0, 3, [], [])
    pj.write_text('{"주석 당기-전기 칸": "C8"}', encoding="utf-8")
    assert _cli(capsys, "link", "test", skel, out, "--pairs", str(pj))[0] == 2
    assert _cli(capsys, "link", "test", skel, out, "--pairs", str(tmp_path / "none.json"))[0] == 2


# ── 기준일 열이 여럿인 TB 표 ─────────────────────────

def _tb_book(path, head, body):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TB"
    for line in [head] + body:
        ws.append(line)
    wb.save(path)
    return str(path)


def test_tb_table_with_many_date_columns(tmp_path):
    head = ["계정과목", "공시용계정과목", "회사제시", "수정사항", "2025.12.31", "2025.06.30", "2024.12.31"]
    p = _tb_book(tmp_path / "pa.xlsx", head, [["현금", "현금및현금성자산", 150, 1, 100, 90, 80], ["보통예금", "현금및현금성자산", 400, 0, 200, 190, 180]])
    rows, info = tbsrc.load(p)                                           # 지정이 없으면 가장 늦은 기준일의 앞 해 말
    assert info["columns"] == {"cur": "회사제시", "pri": "2024.12.31"} and rows[0]["pri"] == 80
    rows, info = tbsrc.load(p, pri_hint="2025.12.31")                    # 뼈대 DSD 의 전기말이 머리글에 있으면 그 열
    assert info["columns"]["pri"] == "2025.12.31" and (info["cur"], info["pri"]) == ("2026.12.31", "2025.12.31")
    assert [(r["code"], r["settle"], r["sub"], r["cur"], r["aje"], r["pri"]) for r in rows] == [
        ("현금", "현금", "현금및현금성자산", 150, 1, 100), ("보통예금", "보통예금", "현금및현금성자산", 400, 0, 200)]
    rows, info = tbsrc.load(p, cur_col="2025.12.31", pri_col="2025.06.30", pri_hint="2024.12.31")   # 직접 지정이 먼저
    assert (rows[0]["cur"], rows[0]["pri"], info["cur"], info["pri"]) == (100, 90, "2025.12.31", "2025.06.30")
    assert tbsrc.load(p, pri_hint="2030.12.31")[1]["columns"]["pri"] == "2024.12.31"   # 없는 머리글이면 원래 규칙
    with pytest.raises(InputError, match="전기 열"):
        tbsrc.load(p, pri_col="2019.12.31")
    q = _tb_book(tmp_path / "dates.xlsx", ["계정과목", "2026.12.31", "2026.09.30", "2025.12.31"], [["현금", 150, 140, 100]])
    assert tbsrc.load(q)[1]["columns"] == {"cur": "2026.12.31", "pri": "2025.12.31"}   # 당기가 기준일 열이면 1년 전 같은 날


def test_cli_make_picks_prior_column_from_skeleton_period(tmp_path, capsys):
    """TB 표에 기준일 열이 여럿이어도, 전기 금액이 뼈대와 맞는 열을 지정하면 짝이 맞는다."""
    skel = LF.write_skeleton(tmp_path / "skel.dsd")
    head = ["코드", "계정과목", "정산 과목", "회사제시", "수정사항", "2025.12.31", "2024.12.31"]
    p = _tb_book(tmp_path / "tb.xlsx", head, [[c, n, s, cur, 0, pri, 1] for c, n, s, pri, cur in LF.TB])
    code, out = _cli(capsys, "link", "make", skel, "--tb", p, "--tb-pri", "2025.12.31", "-o", str(tmp_path / "w.xlsx"), "--json")
    env = json.loads(out)
    assert (code, env["tb_columns"], env["to_check"]) == (0, {"cur": "회사제시", "pri": "2025.12.31"}, [])
    code, out = _cli(capsys, "link", "make", skel, "--tb", p, "--tb-pri", "2024.12.31", "-o", str(tmp_path / "w2.xlsx"), "--json")
    assert code == 1 and json.loads(out)["to_check"]                     # 전기 열을 잘못 고르면 연결점검에 드러난다


# ── 사람용 출력(회사명 · 왕복 시험 결과 엑셀) ───────────────

def test_make_takes_company_from_dsd_when_not_given(tmp_path):
    skel = LF.write_skeleton(tmp_path / "skel.dsd")
    out = tmp_path / "정산표.xlsx"
    assert main(["link", "make", skel, "--tb", LF.write_tb(tmp_path / "tb.json"), "-o", str(out)]) in (0, 1)
    assert openpyxl.load_workbook(out)["BS"]["C2"].value == "합성테스트(주)"
    assert main(["link", "make", skel, "--tb", str(tmp_path / "tb.json"), "-o", str(out), "--company", "다른이름"]) in (0, 1)
    assert openpyxl.load_workbook(out)["BS"]["C2"].value == "다른이름"


def test_test_writes_result_workbook(tmp_path, capsys):
    skel, out, _ = _made(tmp_path)
    rep = tmp_path / "결과" / "왕복시험.xlsx"
    assert main(["link", "test", skel, out, "--notes", "-o", str(rep), "--json"]) == 0
    env = json.loads(capsys.readouterr().out)
    assert env["outputs"] == [str(rep)]
    ws = openpyxl.load_workbook(rep)["왕복시험"]
    assert [ws.cell(6, c).value for c in range(2, 6)] == ["재무제표", 49, 0, 0]
    assert [ws.cell(7, c).value for c in range(2, 6)] == ["주석", 3, 0, 0]
    assert ws.cell(10, 3).value == "해당 없음"

    wb = openpyxl.load_workbook(out)                                       # 연결을 하나 끊으면 어긋난 칸이 목록에 나온다
    tb = wb["TB"]
    row = next(r for r in range(6, tb.max_row + 1) if tb.cell(r, 3).value == "외상매출금")
    tb.cell(row, 5).value = None
    wb.save(out)
    assert main(["link", "test", skel, out, "-o", str(rep), "--json"]) == 1
    capsys.readouterr()
    ws = openpyxl.load_workbook(rep)["왕복시험"]
    listed = {ws.cell(r, 5).value: ws.cell(r, 4).value for r in range(9, ws.max_row + 1) if ws.cell(r, 5).value}
    assert listed.get("매출채권") in ("다름", "빈칸") and "자산총계" in listed
    assert ws.cell(7, 2).value is None                                    # --notes 없이는 주석 줄이 없다
