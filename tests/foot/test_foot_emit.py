from decimal import Decimal

from dsdxlsx.foot.emit import cell_content


def test_note_ref_column_stays_text():
    """주석참조 '4,5,6,32'는 절대 숫자화하지 않는다."""
    value, fmt = cell_content("4,5,6,32", is_note_col=True)
    assert value == "4,5,6,32" and fmt is None


def test_numeric_label_stays_text():
    """'1.예적금담보대출' 같은 라벨은 텍스트 유지 (전체 일치만 숫자)."""
    value, _ = cell_content("1.예적금담보대출", is_note_col=False)
    assert value == "1.예적금담보대출"
    value, _ = cell_content("2021.05.10", is_note_col=False)
    assert value == "2021.05.10"


def test_placeholder_dash_is_zero():
    value, fmt = cell_content("-", is_note_col=False)
    assert value == 0 and fmt is not None


def test_amount_paren_negative():
    value, _ = cell_content("(1,234)", is_note_col=False)
    assert value == -1234


def test_percent_stored_as_fraction():
    value, fmt = cell_content("4.35%", is_note_col=False)
    assert abs(value - 0.0435) < 1e-12
    assert fmt == "0.00%"


def test_won_suffix_numeric():
    value, _ = cell_content("10,000원", is_note_col=False)
    assert value == 10000
    # 단위 표기('천원')는 숫자로 오인하지 않는다
    value, _ = cell_content("14,191천원", is_note_col=False)
    assert value == "14,191천원"


def test_share_suffix():
    value, fmt = cell_content("200,000주", is_note_col=False)
    assert value == 200000 and '"주"' in fmt
