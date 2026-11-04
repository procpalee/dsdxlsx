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
