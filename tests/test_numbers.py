from decimal import Decimal

from dsdxlsx import numbers


def test_classify_int():
    assert numbers.classify("12,345,678,900") == 12345678900
    assert numbers.classify("4,440") == 4440
    assert numbers.classify("0") == 0
    assert numbers.classify("28") == 28
    assert numbers.classify(" 123,456,789") == 123456789  # 공백 낀 숫자


def test_classify_paren_negative():
    assert numbers.classify("(232,116,145)") == -232116145


def test_classify_rejects_note_refs_and_text():
    assert numbers.classify("6,28,31") is None  # 주석 참조 — 3자리 그룹 불일치
    assert numbers.classify("-") is None
    assert numbers.classify("") is None
    assert numbers.classify("(053) 582 - 9230") is None
    assert numbers.classify("(단위 : 원)") is None
    assert numbers.classify("1234,567") is None


def test_classify_suffix_units():
    """주석1 주주현황(주식수·지분율) 합계 검증용 접미 단위."""
    assert numbers.classify("2,978,480주") == 2978480
    assert numbers.classify("62.12%") == Decimal("62.12")
    assert numbers.classify("100.00%") == Decimal("100.00")
    assert numbers.suffix_of("2,978,480주") == "주"
    assert numbers.suffix_of("62.12%") == "%"
    assert numbers.suffix_of("1,234") == ""
    assert numbers.values_equal(2978480, "2,978,480주")  # 무수정 왕복 판정


def test_classify_decimal():
    assert numbers.classify("5.49") == Decimal("5.49")
    assert numbers.classify("1,234.50") == Decimal("1234.50")


def test_classify_over_15_digits_is_text():
    assert numbers.classify("1,000,000,000,000,000") is None


def test_to_dsd_text():
    assert numbers.to_dsd_text(12345678900) == "12,345,678,900"
    assert numbers.to_dsd_text(-232116145) == "(232,116,145)"
    assert numbers.to_dsd_text(0) == "0"
    assert numbers.to_dsd_text(5.49) == "5.49"
    assert numbers.to_dsd_text(Decimal("1234.50")) == "1,234.50"


def test_values_equal():
    assert numbers.values_equal(12345678900, "12,345,678,900")
    assert numbers.values_equal(123456789, " 123,456,789")  # 공백 낀 원본 보존 판정
    assert numbers.values_equal(-232116145, "(232,116,145)")
    assert numbers.values_equal(None, "")
    assert numbers.values_equal("-", "-")
    assert numbers.values_equal(0, "-")    # '-'는 0으로 내보냄 (수식 호환) — 무수정 판정
    assert numbers.values_equal(0, "- ")   # 공백 낀 '-'도 동일
    assert not numbers.values_equal(0, "1,000")
    assert not numbers.values_equal(1, "2")
    assert not numbers.values_equal("텍스트", "다른 텍스트")


def test_classify_triangle_negative():
    """일부 회사의 △/▲ 음수 표기 (괄호 관례와 병행 실존)."""
    assert numbers.classify("△1,234") == -1234
    assert numbers.classify("▲ 232,116,145") == -232116145
    assert numbers.classify("△1,234주") == -1234  # 접미 단위와 결합
    assert numbers.classify("△") is None


def test_to_dsd_text_follows_original_notation():
    """수정 셀 재기입 시 원본 셀의 음수 관례를 따른다 — 표기 혼용 방지."""
    assert numbers.to_dsd_text(-1234) == "(1,234)"
    assert numbers.to_dsd_text(-1234, like="△1,000") == "△1,234"
    assert numbers.to_dsd_text(-1234, like="▲1,000") == "▲1,234"
    assert numbers.to_dsd_text(-1234, like="(1,000)") == "(1,234)"
    assert numbers.to_dsd_text(1234, like="△1,000") == "1,234"  # 양수는 기호 없음


def test_values_equal_triangle_roundtrip():
    assert numbers.values_equal(-1234, "△1,234")   # 무수정 판정
    assert not numbers.values_equal(-1235, "△1,234")
