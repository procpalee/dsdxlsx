"""출력 엑셀 '만든 이' - 환경변수 > excel-theme 스킬 > 지정 안 함."""
from openpyxl import Workbook

from dsdxlsx import theme


def test_author_env_wins(monkeypatch):
    monkeypatch.setenv("DSDXLSX_AUTHOR", "홍길동")
    monkeypatch.setattr(theme, "_skill_author", lambda: "스킬이름")
    assert theme.author() == "홍길동"


def test_author_falls_back_to_skill(monkeypatch):
    monkeypatch.delenv("DSDXLSX_AUTHOR", raising=False)
    monkeypatch.setattr(theme, "_skill_author", lambda: "스킬이름")
    assert theme.author() == "스킬이름"


def test_author_none_keeps_workbook_default(monkeypatch):
    monkeypatch.setenv("DSDXLSX_AUTHOR", "  ")
    monkeypatch.setattr(theme, "_skill_author", lambda: None)
    assert theme.author() is None
    wb = Workbook()
    before = wb.properties.creator
    theme.set_author(wb)
    assert wb.properties.creator == before


def test_set_author(monkeypatch):
    monkeypatch.setenv("DSDXLSX_AUTHOR", "홍길동")
    wb = Workbook()
    theme.set_author(wb)
    assert wb.properties.creator == "홍길동"


def test_cli_outputs_get_author(monkeypatch, tmp_path):
    """새로 만드는 출력 엑셀(export, check 보고서, rollforward)은 만든 이를 정하고,
    사용자가 가진 엑셀에 검증을 적는 check 는 만든 이를 바꾸지 않는다."""
    from openpyxl import load_workbook

    from dsdxlsx.cli import main
    from fixtures import factory

    monkeypatch.chdir(tmp_path)
    factory.write_dsd(tmp_path / "전기.dsd", factory.standard_contents())
    monkeypatch.setenv("DSDXLSX_AUTHOR", "홍길동")
    assert main(["export", "전기.dsd", "-o", "편집.xlsx"]) == 0
    assert main(["check", "전기.dsd", "-o", "보고서.xlsx"]) in (0, 1)
    assert main(["rollforward", "전기.dsd"]) == 0
    for name in ("편집.xlsx", "보고서.xlsx", "전기_차기기초.xlsx"):
        assert load_workbook(tmp_path / name).properties.creator == "홍길동", name

    monkeypatch.setenv("DSDXLSX_AUTHOR", "다른사람")
    assert main(["check", "편집.xlsx"]) in (0, 1)
    assert load_workbook(tmp_path / "편집.xlsx").properties.creator == "홍길동"
