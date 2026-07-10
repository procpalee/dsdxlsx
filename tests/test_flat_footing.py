"""무들여쓰기 총계-먼저(flat totals-first) footing — 흥구형 IFRS 재무상태표.

구조: 들여쓰기·로마자 없이 '유동자산(합계) → 구성항목' 순. 중첩 부모
(매출채권및기타채권 = 매출채권 − 대손충당금 + 미수금)와 단일 자식 부모,
총계 결합(자본과부채총계 = 부채총계 + 자본총계)까지 재현한다.
"""
from dsdxlsx.check import run_checks
from dsdxlsx.parse import parse_contents

from fixtures import factory


def _flat_bs(cash="500", assets_total="10,000"):
    """자산 = 유동 6,000 + 비유동 4,000 = 10,000 / 부채 3,000 / 자본 7,000."""
    bs = factory.table(
        [["과목", "제 26 기 1분기말", "제 25 기말"]],
        [
            ["유동자산", "6,000", "5,800"],
            ["현금및현금성자산", cash, "600"],           # 기본 500
            ["매출채권 및 기타유동채권", "3,500", "3,200"],  # = 3,800-400+100
            ["매출채권", "3,800", "3,500"],
            ["대손충당금", "(400)", "(400)"],
            ["단기미수금", "100", "100"],
            ["재고자산", "2,000", "2,000"],
            ["비유동자산", "4,000", "4,000"],
            ["기타비유동금융자산", "1,000", "1,000"],       # 단일 자식 부모
            ["상각후원가 측정 금융자산", "1,000", "1,000"],
            ["유형자산", "3,000", "3,000"],
            ["자산총계", assets_total, "9,800"],
            ["유동부채", "3,000", "2,900"],
            ["매입채무", "3,000", "2,900"],
            ["부채총계", "3,000", "2,900"],
            ["자본금", "1,500", "1,500"],
            ["이익잉여금", "5,500", "5,400"],
            ["자본총계", "7,000", "6,900"],
            ["자본과부채총계", "10,000", "9,800"],
        ],
        unit="원",
    )
    return factory.document(
        factory.p("표지"),
        factory.title("재무제표", 2),
        factory.statement_title_table("재 무 상 태 표"),
        bs,
    )


def test_flat_pristine_passes_with_full_tree():
    doc = parse_contents(_flat_bs())
    report = run_checks(doc, only={"C2"})
    assert not [f for f in report.findings if f.level == "warn"]
    # 당기·전기 2열 × (자산총계·부채총계·자본총계·총계결합 + 부모 3개) 이상
    assert report.counts.get("C2", 0) >= 14
    kinds = {x.note for x in report.results if x.check == "C2"}
    assert any("총계 결합" in k for k in kinds)   # 자본과부채총계 = 부채+자본
    assert any("그룹" in k for k in kinds)        # 중첩 부모 (매출채권및기타 등)


def test_flat_detects_broken_leaf_via_reference_column():
    # 현금 500 → 700: 당기 열 분해 실패 → 전기 구조 적용 → 유동자산 불일치 검출
    doc = parse_contents(_flat_bs(cash="700"))
    report = run_checks(doc, only={"C2"})
    warns = [f for f in report.findings if f.level == "warn"]
    assert warns, "기준열 구조 적용으로 불일치를 잡아야 한다"
    assert any("유동자산" in f.where and "(200)" in f.message for f in warns)


def test_flat_shape_change_downgraded_not_false_warned():
    """기수 간 구성 변동(전기 0 → 당기 신규 하위 구조)은 기준열 구조 전이 시
    오탐(warn)이 아니라 '구조 이형 가능성' info로 강등 — 진탐 warn은 유지."""
    contents = factory.document(
        factory.p("표지"),
        factory.title("재무제표", 2),
        factory.statement_title_table("재 무 상 태 표"),
        factory.table(
            [["과목", "제 26 기 1분기말", "제 25 기말"]],
            [
                ["유동자산", "6,000", "5,800"],
                ["현금및현금성자산", "4,200", "3,800"],   # 훼손: 4,000 → 4,200
                ["재고자산", "2,000", "2,000"],
                ["비유동자산", "4,000", "4,000"],
                ["무형자산", "1,000", "0"],               # 전기 0 — 당기 신규 부모
                ["소프트웨어", "1,100", "0"],
                ["상각누계액", "(100)", "0"],
                ["유형자산", "3,000", "4,000"],
                ["자산총계", "10,000", "9,800"],
            ],
            unit="원",
        ),
    )
    report = run_checks(parse_contents(contents), only={"C2"})
    warns = [f for f in report.findings if f.level == "warn"]
    infos = [f for f in report.findings if f.level == "info"]
    # 진탐: 유동자산 (현금 훼손) — warn 유지
    assert any("유동자산" in f.where and "(200)" in f.message for f in warns)
    # 비유동자산은 warn이 아니라 구조 이형 info
    assert not any("비유동자산" in f.where for f in warns)
    assert any("비유동자산" in f.where and "구조 이형" in f.message for f in infos)


def test_flat_all_columns_broken_surfaces_info():
    # 양쪽 기수 모두 구조가 깨지면 침묵하지 않고 info 표면화
    doc = parse_contents(_flat_bs(cash="700", assets_total="9,999"))
    contents = factory.document(
        factory.p("표지"),
        factory.title("재무제표", 2),
        factory.statement_title_table("재 무 상 태 표"),
        factory.table(
            [["과목", "제 26 기 1분기말", "제 25 기말"]],
            [
                ["유동자산", "6,000", "5,800"],
                ["현금및현금성자산", "999", "999"],
                ["매출채권", "3,500", "3,200"],
                ["재고자산", "2,000", "2,000"],
                ["비유동자산", "4,000", "4,000"],
                ["유형자산", "3,999", "3,999"],
                ["자산총계", "10,000", "9,800"],
            ],
            unit="원",
        ),
    )
    report = run_checks(parse_contents(contents), only={"C2"})
    assert any("해석하지 못함" in f.message for f in report.findings)
