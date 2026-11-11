"""ERP 자료 → TB(tb) - 더존(위하고) · 세무사랑 양식 읽기, 계정 분류, 정산 과목 규칙(합성 자료)."""
from __future__ import annotations

import json

import pytest

from dsdxlsx.cli import main
from dsdxlsx.erp import settle as S
from dsdxlsx.erp import tb as T
from dsdxlsx.guard import InputError
from dsdxlsx.link import tbsrc

from fixtures import erp_factory as EF
from fixtures import link_factory as LF


def _by_name(out: dict) -> dict:
    return {r[1]: (r[0], r[2], r[3], out["classes"][r[0]]) for r in out["rows"]}


# ── 더존 ─────────────────────────────────────────────

def test_douzone_tb_amounts_codes_and_classes(tmp_path):
    out = T.make(EF.write_douzone(tmp_path))
    assert (out["erp"], out["key"], out["cur_date"], out["pri_date"]) == ("더존", "2026", "2026.12.31", "2025.12.31")
    assert out["warnings"] == []
    rows = _by_name(out)
    assert rows["보통예금"] == ("103", 400.0, 200.0, ["자산", "유동자산", "당좌자산", "보통예금"])
    assert rows["대손충당금"] == ("109", -120.0, -100.0, ["자산", "유동자산", "당좌자산", "대손충당금(외상매출금)"])   # 차감 계정 = 음수, 짝 자산
    assert rows["감가상각누계액"][3][3] == "감가상각누계액(비품)"
    assert rows["외상매입금"] == ("251", -680.0, -700.0, ["부채", "유동부채", "유동부채", "외상매입금"])                # 부채 · 자본 · 수익 = 대변(-)
    assert rows["상품매출"][1:3] == (-6000.0, -5000.0)
    assert rows["법인세등"] == ("998", 250.0, 70.0, ["손익", "법인세등", "법인세등", "법인세등"])                    # 세부 줄 없는 그룹 = 그룹 줄이 계정
    assert "기초상품재고액" not in rows and rows["상품매출원가"][1:3] == (3600.0, 3000.0)                           # 재고 구성 줄은 계정이 아니다
    assert sum(r[2] for r in out["rows"]) == 0 and sum(r[3] for r in out["rows"]) == 0


def test_retained_earnings_before_closing_and_codes_from_account_list(tmp_path):
    rows = _by_name(T.make(EF.write_douzone(tmp_path)))
    # 이월이익잉여금(마감 전) = 재무상태표 미처분이익잉여금 - 당기순이익: 600 - 300 / 300 - 250
    assert rows["이월이익잉여금"] == ("375", -300.0, -50.0, ["자본", "이익잉여금", "이익잉여금", "미처분이익잉여금"])
    assert rows["자본금"][0] == "331"                                   # 거래가 없는 계정의 코드는 계정과목 목록에서
    assert rows["비품"][0] == "*비품"                                   # 어디에도 코드가 없으면 '*정산 과목'
    (tmp_path / "x").mkdir()
    no_book = _by_name(T.make(EF.write_douzone(tmp_path / "x", book=())))
    assert no_book["자본금"][0] == "*자본금" and no_book["이월이익잉여금"][0] == "*미처분이익잉여금"
    with_book = _by_name(T.make(str(tmp_path / "x"), book={"212": "비품", "331": "자본금"}))
    assert (with_book["비품"][0], with_book["자본금"][0]) == ("212", "331")


def test_same_name_accounts_split_by_journal_amounts(tmp_path):
    """판관비와 영업외에 같은 이름 계정이 있으면 분개장 순액이 재무제표 금액과 맞는 코드를 고른다."""
    pl = [list(r) for r in EF.PL]
    i = next(k for k, r in enumerate(pl) if r[0] == "이자수익")
    pl[i - 1][2], pl[i][1] = 40, 30                                    # 영업외수익 40 = 이자수익 30 + 잡이익 10
    pl.insert(i + 1, ["잡이익", 10, None, 0, None])
    for r in pl:
        if r[0] in ("Ⅶ.법인세차감전이익",):
            r[2] = 560
        if r[0] == "Ⅸ.당기순이익":
            r[2] = 310
    bs = [list(r) for r in EF.BS]
    for r in bs:
        if r[0] in ("보통예금",):
            r[2] = 410
        if r[0] in ("(1)당좌자산",):
            r[2] = 1440
        if r[0] in ("Ⅰ.유동자산",):
            r[2] = 1890
        if r[0] in ("자산총계", "부채및자본총계"):
            r[2] = 2290
        if r[0] in ("Ⅱ.이익잉여금", "미처분이익잉여금"):
            r[2] = 610
        if r[0] == "자본총계":
            r[2] = 1610
    slips = EF.SLIPS + [("09/10", 13, [(103, "보통예금", 10, 0), (930, "잡이익", 0, 10)]),
                        ("09/11", 14, [(103, "보통예금", 5, 0), (931, "잡이익", 0, 5)]),
                        ("09/12", 15, [(931, "잡이익", 5, 0), (103, "보통예금", 0, 5)])]
    out = T.make(EF.write_douzone(tmp_path, pl=pl, bs=bs, slips=slips))
    assert _by_name(out)["잡이익"][:2] == ("930", -10.0)                # 931 은 순액 0 - 금액이 맞는 930


# ── 세무사랑 ─────────────────────────────────────────

def test_semusarang_export_gives_same_tb(tmp_path):
    (tmp_path / "dz").mkdir()
    (tmp_path / "sm").mkdir()
    dz = T.make(EF.write_douzone(tmp_path / "dz", book=()))
    sm = T.make(EF.write_semusarang(tmp_path / "sm"))
    assert sm["erp"] == "세무사랑" and sm["cur_date"] == "2026.12.31"
    key = lambda o: sorted((str(r[0]), r[2], r[3], o["classes"][r[0]]) for r in o["rows"])  # noqa: E731
    assert key(sm) == key(dz)
    assert _by_name(sm)["임원급여(판)"][0] == "801"                      # 세무사랑 계정 이름에는 원가 구분 표시를 붙인다


# ── 멈추는 경우 ──────────────────────────────────────

def test_stops_on_missing_file_wrong_date_and_unbalanced(tmp_path):
    raw = EF.write_douzone(tmp_path)
    with pytest.raises(InputError, match="2025"):
        T.make(raw, key="2025")
    with pytest.raises(InputError, match="2026.12.31"):
        T.make(raw, date="2026.09.30")
    with pytest.raises(InputError, match="자료 폴더"):
        T.make(str(tmp_path / "none"))
    (tmp_path / "bad").mkdir()
    bs = [list(r) for r in EF.BS]
    next(r for r in bs if r[0] == "현금")[2] = 190
    with pytest.raises(InputError, match="부호합"):
        T.make(EF.write_douzone(tmp_path / "bad", bs=bs))


def test_cutoff_date_drops_later_slips(tmp_path):
    raw = EF.write_douzone(tmp_path, slips=EF.SLIPS + [("12/31", 99, [(101, "현금", 7, 0), (103, "보통예금", 0, 7)])])
    assert T.make(raw, date="2026.12.31")["warnings"] == []
    files = {"분개장": str(tmp_path / "분개장_합성_2026.xlsx"), "재무상태표": str(tmp_path / "재무상태표_합성_2026.xlsx"),
             "손익계산서": str(tmp_path / "손익계산서_합성_2026.xlsx")}
    assert T.make(None, files=files)["year"] == 2026                    # 이름 규칙 없이 파일을 직접 - 연도는 재무상태표 기준일


# ── 정산 과목 규칙 ───────────────────────────────────

def test_default_rules_group_accounts_into_statement_names(tmp_path):
    doc = T.to_tb(T.make(EF.write_douzone(tmp_path)), S.load())
    settle = {r["name"]: r["settle"] for r in doc["rows"]}
    assert settle["현금"] == settle["보통예금"] == "현금및현금성자산"
    assert (settle["외상매출금"], settle["대손충당금"], settle["외상매입금"]) == ("매출채권", "대손충당금(매출채권)", "매입채무")
    assert settle["감가상각누계액"] == "감가상각누계액(비품)" and settle["상품매출원가"] == "상품매출원가"
    assert doc["schema"] == "dsdxlsx.tb/1" and doc["period"] == {"cur": "2026.12.31", "pri": "2025.12.31"}
    plain = {r["name"]: r["settle"] for r in T.to_tb(T.make(str(tmp_path)), None)["rows"]}
    assert plain["보통예금"] == "보통예금" and plain["대손충당금"] == "대손충당금(외상매출금)"


def test_rules_file_env_and_errors(tmp_path, monkeypatch):
    p = tmp_path / "rules.toml"
    p.write_text('[[rule]]\nmatch = "임원급여|직원급여"\nbig = "손익"\ngroup = "판매비와관리비"\nsettle = "급여"\n'
                 '[[rule]]\nmatch = ".*예금"\nexcept = "정기"\nbig = "자산"\nsettle = "예금"\n', encoding="utf-8")
    r = S.load(str(p))
    assert S.settle("임원급여", "손익", "판매비와관리비", "판매비와관리비", r) == "급여"
    assert S.settle("임원급여", "손익", "매출원가", "매출원가", r) == "임원급여"          # 원가 계정은 묶지 않는다
    assert (S.settle("보통 예금", "자산", "", "", r), S.settle("정기예금", "자산", "", "", r)) == ("예금", "정기예금")
    monkeypatch.setenv(S.ENV, str(p))
    assert S.load().source == str(p)
    monkeypatch.delenv(S.ENV)
    assert S.load().source == "기본 규칙"
    (tmp_path / "bad.toml").write_text('[[rule]]\nmatch = "("\nbig = "자산"\nsettle = "x"\n', encoding="utf-8")
    with pytest.raises(InputError, match="1번째"):
        S.load(str(tmp_path / "bad.toml"))
    with pytest.raises(InputError, match="찾을 수 없습니다"):
        S.load(str(tmp_path / "none.toml"))


# ── 명령, 그리고 정산표까지 ──────────────────────────

def _cli(capsys, *argv):
    capsys.readouterr()
    code = main(list(argv))
    return code, capsys.readouterr().out


def test_cli_tb_then_link_make_and_roundtrip(tmp_path, capsys):
    raw = EF.write_douzone(tmp_path)
    code, out = _cli(capsys, "tb", raw, "--json")
    env = json.loads(out)
    tbp = str(tmp_path / "TB_2026.json")
    assert (code, env["command"], env["outputs"], env["erp"], env["accounts"]) == (0, "tb", [tbp], "더존", 19)
    assert env["no_code"] == ["비품"] and env["rules"] == "기본 규칙" and env["period"]["cur"] == "2026.12.31"
    rows, info = tbsrc.load(tbp)
    assert info["cur"] == "2026.12.31" and len(rows) == 19

    skel = LF.write_skeleton(tmp_path / "skel.dsd")
    code, out = _cli(capsys, "link", "make", skel, "--tb", tbp, "--json")
    env = json.loads(out)
    assert (code, env["to_check"], env["unlinked_tb"]) == (0, [], [])    # 급여 = 임원 + 직원, 보통주자본금 = 자본금(전기 금액으로 짝)
    code, out = _cli(capsys, "link", "test", skel, env["outputs"][0], "--json")
    env = json.loads(out)
    assert (code, env["same"], env["diff"], env["blank"]) == (0, 49, [], [])


def test_cli_tb_options_and_guard(tmp_path, capsys):
    raw = EF.write_douzone(tmp_path)
    out = str(tmp_path / "out" / "tb.json")
    (tmp_path / "out").mkdir()
    assert _cli(capsys, "tb", raw, "--no-rules", "-o", out)[0] == 0
    doc = json.load(open(out, encoding="utf-8"))
    assert {r["name"]: r["settle"] for r in doc["rows"]}["보통예금"] == "보통예금"
    doc["rows"][0]["settle"] = "사람이 고침"
    json.dump(doc, open(out, "w", encoding="utf-8"), ensure_ascii=False)
    assert _cli(capsys, "tb", raw, "-o", out)[0] == 3                    # 고친 TB 는 덮어쓰지 않는다
    assert _cli(capsys, "tb", raw, "-o", out, "--force")[0] == 0
    assert _cli(capsys, "tb", raw, "--key", "2024", "-o", out, "--force")[0] == 2
    assert _cli(capsys, "tb")[0] == 2
    code, text = _cli(capsys, "tb", "--bs", str(tmp_path / "재무상태표_합성_2026.xlsx"), "--pl", str(tmp_path / "손익계산서_합성_2026.xlsx"),
                      "--journal", str(tmp_path / "분개장_합성_2026.xlsx"), "-o", str(tmp_path / "out" / "direct.json"), "--json")
    assert code == 0 and json.loads(text)["no_code"] == ["비품", "자본금", "이월이익잉여금"]
