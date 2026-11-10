"""추가 검사 3그룹 회귀 테스트 — CF 앵커 캐스케이드·공정가치 등가·CE 이중 변형."""
from fixtures.foot_factory import build_dsd, table, _td, _amount_row

from dsdxlsx.foot import footing, sections
from dsdxlsx.foot.tableview import build_views
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.parse import parse_contents


def _analyze(tmp_path, body_xml: str):
    xml = ('<?xml version="1.0" encoding="UTF-8"?><BODY>' + body_xml + "</BODY>").encode()
    p = tmp_path / "t.dsd"
    p.write_bytes(build_dsd(xml))
    doc = parse_contents(read_dsd(str(p)).contents)
    plans = sections.build(doc)
    views, ctx = build_views(plans)
    return footing.analyze(views), ctx


def test_cf_anchor_cascade_no_roman(tmp_path):
    """로마자 힌트 없는 IFRS풍 CF: 순증가 = 영업+투자+재무, 활동합계 = 세부 합."""
    cf = table([
        [_td("과목"), _td("당기"), _td("전기")],
        _amount_row("영업활동으로 인한 현금흐름", ["100", "80"]),
        _amount_row("이자수취", ["60", "50"], 1),
        _amount_row("법인세납부", ["40", "30"], 1),
        _amount_row("투자활동으로 인한 현금흐름", ["(30)", "(20)"]),
        _amount_row("유형자산의 취득", ["(30)", "(20)"], 1),
        _amount_row("재무활동으로 인한 현금흐름", ["10", "5"]),
        _amount_row("차입금의 증가", ["10", "5"], 1),
        _amount_row("현금및현금성자산의 순증가(감소)", ["80", "65"]),
        _amount_row("기초의 현금및현금성자산", ["20", "(45)"]),
        _amount_row("기말의 현금및현금성자산", ["100", "20"]),
    ], 3)
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">현 금 흐 름 표</TITLE>' + cf)
    anchors = [s for s in specs if s.kind == "cf_anchor"]
    assert len([s for s in anchors if "순증가" in s.label]) == 2  # 당·전기
    assert all(s.ok for s in anchors)
    # 활동합계도 잡힘 (영업 = 이자수취+법인세납부)
    assert any("영업활동" in s.label for s in anchors)


def test_fv_equal_horizontal_and_vertical(tmp_path):
    note = (
        "<P>1. 금융자산&amp;cr;&amp;cr;보유 유가증권 내역입니다.</P>"
        + table([
            [_td("구분"), _td("취득원가"), _td("공정가치"), _td("장부가액")],
            _amount_row("종목가", ["50", "70", "70"]),
            _amount_row("종목나", ["30", "40", "40"]),
            _amount_row("종목다(원가측정)", ["20", "25", "20"]),
        ], 4)
        + "<P>수준별 공정가치 비교입니다.</P>"
        + table([
            [_td("구분"), _td("당기말"), _td("전기말")],
            _amount_row("장부금액", ["110", "90"]),
            _amount_row("공정가치", ["110", "90"]),
        ], 3)
    )
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    fv = [s for s in specs if s.kind == "fv_equal"]
    horiz = [s for s in fv if s.axis == "row"]
    vert = [s for s in fv if s.axis == "col"]
    assert len(horiz) == 2, "일치 종목 2건만 (원가측정 종목은 침묵)"
    assert all(s.ok for s in horiz)
    assert len(vert) == 2 and all(s.ok for s in vert)  # 당·전기말


def test_cf_anchor_kgaap_left_column(tmp_path):
    """구현 후보 ①: K-GAAP 2단 열 — 활동합계(C열) = '(n)' 소계(B열)의 합."""
    cf = table([
        [_td("과목"), _td("제 2(당)기", COLSPAN=2), _td("제 1(전)기", COLSPAN=2)],
        _amount_row("Ⅰ. 영업활동으로 인한 현금흐름", ["", "100", "", "80"]),
        _amount_row("(1) 당기순이익", ["70", "", "50", ""], 1),
        _amount_row("(2) 가산 항목", ["30", "", "30", ""], 1),
        _amount_row("Ⅱ. 투자활동으로 인한 현금흐름", ["", "(20)", "", "(10)"]),
        _amount_row("(1) 투자활동 현금유출액", ["(20)", "", "(10)", ""], 1),
    ], 5)
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">현 금 흐 름 표</TITLE>' + cf)
    anchors = [s for s in specs if s.kind == "cf_anchor" and s.ok]
    assert any("영업활동" in s.label for s in anchors), "좌열 소계 합산이 잡혀야 한다"


def test_shareholder_ratio_and_contract_rollforward(tmp_path):
    """지분율과 분양계약 총액·잔액의 명시적 헤더 관계."""
    note = (
        "<P>1. 회사의 개요&amp;cr;주주 현황입니다.</P>"
        + table([
            [_td("주주"), _td("주식수"), _td("금액"), _td("지분율")],
            _amount_row("갑", ["596", "2,980", "59.6"]),
            _amount_row("을", ["400", "2,000", "40.0"]),
            _amount_row("기타", ["4", "20", "0.4"]),
            _amount_row("합계", ["1,000", "5,000", "100.0"]),
        ], 4)
        + "<P>분양계약 내역입니다.</P>"
        + table([
            [_td("구분"), _td("기초분양계약"), _td("신규계약액"),
             _td("총분양금액"), _td("누적수익인식액"), _td("기말예약 매출잔액")],
            _amount_row("당기", ["100", "20", "120", "90", "30"]),
            _amount_row("전기", ["80", "20", "100", "70", "30"]),
        ], 6)
    )
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    ratios = [s for s in specs if s.kind == "ratio"]
    contracts = [s for s in specs if s.kind == "flow_h"
                 and ("총분양금액" in s.label or "매출잔액" in s.label)]
    assert len(ratios) == 4 and all(s.ok and s.factor == 100 for s in ratios)
    assert len(contracts) == 4 and all(s.ok for s in contracts)


def test_weighted_shares_and_eps_unit_rounding(tmp_path):
    """누적주식수·가중평균과 천원 순이익→원 단위 주당이익."""
    note = (
        "<P>1. 주당이익&amp;cr;기본주당이익 내역입니다.</P>"
        + table([
            [_td("구분"), _td("당기"), _td("전기")],
            _amount_row("보통주 지배기업소유주지분순손익", ["6,640,000", "16,720,000"]),
            _amount_row("가중평균 유통보통주식수", ["31,500,000", "31,500,000"]),
            _amount_row("기본주당이익(단위: 원)", ["211", "531"]),
        ], 3)
        + table([
            [_td("기간"), _td("항목"), _td("해당일수"), _td("유통주식수"),
             _td("누적유통보통주식수")],
            _amount_row("1.1 ~ 12.31", ["", "365", "31,500,000", "11,497,500,000"]),
            _amount_row("합계", ["", "", "", "11,497,500,000"]),
            _amount_row("가중평균유통보통주식수(=합계 / 365)",
                        ["", "", "", "31,500,000"]),
        ], 5)
    )
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    eps = [s for s in specs if s.op == "/" and s.factor == 1000]
    weighted = [s for s in specs if "가중평균유통" in s.label and s.op == "/"]
    products = [s for s in specs if s.op == "*" and "누적유통" in s.label]
    assert len(eps) == 2 and all(s.ok and s.digits == 0 for s in eps)
    assert len(weighted) == 1 and weighted[0].ok
    assert len(products) == 1 and products[0].ok


def test_variance_columns_and_embedded_pl(tmp_path):
    """별첨 분석표의 증감액과 주석 TITLE 안 손익계산서 단계손익을 검증한다."""
    note = (
        "<P>1. 재무제표 분석&amp;cr;요약표입니다.</P>"
        + table([
            [_td("구분"), _td("당기말"), _td("전기말"), _td("증(감)액"), _td("증(감)률")],
            _amount_row("유동자산", ["120", "100", "20", "20%"]),
            _amount_row("비유동자산", ["90", "60", "30", "50%"]),
            _amount_row("자산총계", ["210", "160", "50", "31.25%"]),
        ], 5)
        + table([
            [_td("구분"), _td("당기"), _td("전기")],
            _amount_row("매출액", ["100", "90"]),
            _amount_row("매출원가", ["60", "55"]),
            _amount_row("매출총이익", ["40", "35"]),
            _amount_row("판매비와관리비", ["10", "8"]),
            _amount_row("영업이익", ["30", "27"]),
            _amount_row("법인세비용", ["5", "4"]),
            _amount_row("당기순이익", ["25", "23"]),
        ], 3)
    )
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    variance = [s for s in specs if s.kind == "flow_h" and "증" in s.label]
    rates = [s for s in specs if s.kind == "ratio" and "증" in s.label]
    anchors = [s for s in specs if s.kind == "pl_anchor"]
    assert len(variance) == 3 and all(s.ok for s in variance)
    assert len(rates) == 3 and all(s.ok for s in rates)
    assert len(anchors) >= 6 and all(s.ok for s in anchors)


def test_flat_kgaap_outline_inside_appendix_note(tmp_path):
    """별첨의 무들여쓰기 2단 금액 열: (1) 소계와 로마자 상위합을 모두 검증."""
    note = (
        "<P>1. 별첨 재무제표&amp;cr;재무상태표입니다.</P>"
        + table([
            [_td("과목"), _td("당기 세부"), _td("당기 합계")],
            _amount_row("Ⅰ. 유동자산", ["", "100"]),
            _amount_row("(1) 당좌자산", ["", "70"]),
            _amount_row("현금", ["30", ""]),
            _amount_row("미수금", ["40", ""]),
            _amount_row("(2) 재고자산", ["", "30"]),
            _amount_row("상품", ["20", ""]),
            _amount_row("제품", ["10", ""]),
            _amount_row("Ⅱ. 비유동자산", ["", "50"]),
            _amount_row("(1) 유형자산", ["", "50"]),
            _amount_row("토지", ["20", ""]),
            _amount_row("건물", ["30", ""]),
        ], 3)
    )
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    groups = [s for s in specs if s.kind == "group"]
    labels = {s.label for s in groups if s.ok}
    assert any("유동자산" in x for x in labels)
    assert any("당좌자산" in x for x in labels)
    assert any("재고자산" in x for x in labels)
    assert any("유형자산" in x for x in labels)


def test_korean_hint_lines(tmp_path):
    """구현 후보 ②: 라벨 끝 '(가+나+다)' 합산식 힌트."""
    note = ("<P>1. 법인세비용&amp;cr;&amp;cr;법인세비용의 구성입니다.</P>"
            + table([
                [_td("구분"), _td("당기")],
                _amount_row("가. 당기 법인세부담액", ["100"]),
                _amount_row("나. 이연법인세 변동액", ["(30)"]),
                _amount_row("다. 자본에 직접 반영된 금액", ["10"]),
                _amount_row("라. 법인세비용 (가+나+다)", ["80"]),
            ], 2))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    hints = [s for s in specs if s.kind == "hint"]
    assert hints and all(s.ok for s in hints)


def test_bullet_subtotal_downward(tmp_path):
    """구현 후보 ③: 소계가 위, '- 세부항목'이 아래인 표 — 정확 일치만."""
    note = ("<P>1. 포괄손익&amp;cr;&amp;cr;포괄손익 내역입니다.</P>"
            + table([
                [_td("구분"), _td("세부", COLSPAN=1), _td("금액")],
                _amount_row("기타포괄손익", ["", "(180)"]),
                _amount_row("- 재평가손익", ["(60)", ""]),
                _amount_row("- 평가손익", ["(120)", ""]),
                _amount_row("총포괄손익", ["", "(180)"]),
            ], 3))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    down = [s for s in specs if s.kind == "total_row"
            and any(sign == 1 and (j, cc) != s.target for j, cc, sign in s.terms)
            and all(j > s.target[0] for j, cc, sign in s.terms)]
    assert down and all(s.ok for s in down), "하방 불릿 소계가 잡혀야 한다"


def test_summary_fin(tmp_path):
    """구현 후보 ④: 요약재무 자본 = 자산 − 부채 (일치만)."""
    note = ("<P>1. 종속기업투자&amp;cr;&amp;cr;요약재무정보입니다.</P>"
            + table([
                [_td("회사명"), _td("자산"), _td("부채"), _td("자본")],
                _amount_row("갑회사", ["1,000", "400", "600"]),
                _amount_row("을회사", ["500", "200", "300"]),
                _amount_row("병회사(비지배 있음)", ["900", "300", "550"]),
            ], 4))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    sf = [s for s in specs if s.kind == "summary_fin"]
    assert len(sf) == 2, "일치 2건만 (불일치 행은 침묵)"
    assert all(s.ok for s in sf)
    assert any(sign == -1 for s in sf for _, _, sign in s.terms), "부채는 차감 항"


def test_eq_carry_and_single_total(tmp_path):
    """구현 후보 ⑤: 무변동 이월(기초=기말)과 단일 구성항 소계 — 일치할 때만."""
    note = ("<P>1. 자본금 및 유가증권&amp;cr;&amp;cr;내역입니다.</P>"
            + table([
                [_td("구분"), _td("자본금"), _td("주식수")],
                _amount_row("전기초", ["500", "1,000"]),
                _amount_row("전기말", ["500", "1,000"]),
            ], 3)
            + "<P>유가증권 내역입니다.</P>"
            + table([
                [_td("구분"), _td("당기말")],
                _amount_row("시장성 없는 지분증권", ["50"]),
                _amount_row("소  계", ["50"]),
            ], 2)
            + "<P>불일치 이월은 침묵해야 합니다.</P>"
            + table([
                [_td("구분"), _td("금액")],
                _amount_row("기초", ["100"]),
                _amount_row("기말", ["120"]),
            ], 2))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    carry = [s for s in specs if s.kind == "eq_carry"]
    assert len(carry) == 2 and all(s.ok for s in carry), "동액 이월 2건(자본금·주식수)만"
    single = [s for s in specs if s.kind == "total_row" and len(s.terms) == 1]
    assert single and all(s.ok for s in single), "단일 구성항 소계가 잡혀야 한다"


def test_pl_roman_chain_medical(tmp_path):
    """FP1: 의료 PL — Ⅵ.법차전 = Ⅲ.의료이익+Ⅳ.의료외수익−Ⅴ.의료외비용 (Ⅱ 합계 없음)."""
    pl = table([
        [_td("과목"), _td("당기")],
        _amount_row("Ⅰ. 의료수익", ["1,000"]),
        _amount_row("Ⅱ. 의료비용", [""]),
        _amount_row("1. 인건비", ["600"], 1),
        _amount_row("Ⅲ. 의료이익", ["400"]),
        _amount_row("Ⅳ. 의료외수익", ["50"]),
        _amount_row("Ⅴ. 의료외비용", ["30"]),
        _amount_row("Ⅵ. 법인세비용차감전순이익", ["420"]),
    ], 2)
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">손 익 계 산 서</TITLE>' + pl)
    a = [s for s in specs if s.kind == "pl_anchor" and "법인세비용차감전순이익" in s.label]
    assert a and all(s.ok for s in a), "로마자 섹션 체인이 정확 일치해야 한다"
    assert any(sign == -1 for s in a for *_, sign in s.terms), "의료외비용은 차감 항"


def test_pl_roman_equal_carry(tmp_path):
    """FP1: ⅩⅡ.당기순이익 = ⅩⅠ.계속사업이익 (동액 단독항 — ⅩⅠ은 비앵커 라벨)."""
    pl = table([
        [_td("과목"), _td("당기")],
        _amount_row("Ⅸ. 법인세차감전계속사업손익", ["500"]),
        _amount_row("Ⅹ. 계속사업손익법인세비용", ["100"]),
        _amount_row("ⅩⅠ. 계속사업이익", ["400"]),
        _amount_row("ⅩⅡ. 당기순이익", ["400"]),
    ], 2)
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">손 익 계 산 서</TITLE>' + pl)
    a = [s for s in specs if s.kind == "pl_anchor" and "당기순이익" in s.label]
    assert a and all(s.ok for s in a), "동액 이월(ⅩⅡ=ⅩⅠ)이 잡혀야 한다"
    assert all(len(s.terms) == 1 for s in a), "단독항(ⅩⅠ)이어야 한다"


def test_total_after_named_subtotal(tmp_path):
    """FP4: '채권 합계'(명명 소계)+충당금 뒤의 '합 계' = 소계+충당금 (이중 합산 방지)."""
    note = ("<P>1. 기타채권&amp;cr;&amp;cr;채권 내역입니다.</P>"
            + table([
                [_td("구분"), _td("당기"), _td("전기")],
                _amount_row("대여금", ["100", "90"]),
                _amount_row("미수금", ["50", "40"]),
                _amount_row("채권 합계", ["150", "130"]),
                _amount_row("대손충당금", ["(30)", "(20)"]),
                _amount_row("합 계", ["120", "110"]),
            ], 3))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    tr = [s for s in specs if s.kind == "total_row"]
    assert tr and all(s.ok for s in tr), "명명 소계+충당금 후보가 정확 일치해야 한다"
    assert all(len(s.terms) == 2 for s in tr), "소계행+충당금행 2항 (세부 이중 합산 금지)"


def test_asset_cost_with_subsidy_column(tmp_path):
    """FP3: 장부금액 구성 — 기말 = 취득원가 + 상각누계액 + 국고보조금 (차감열 확장)."""
    note = ("<P>1. 유형자산&amp;cr;&amp;cr;유형자산 변동내역입니다.</P>"
            + table([
                [_td("구분"), _td("기초"), _td("취득"), _td("감가상각비"),
                 _td("국고보조금차감"), _td("기말"), _td("취득원가"),
                 _td("상각누계액"), _td("국고보조금")],
                _amount_row("차량운반구",
                            ["0", "0", "(45)", "45", "0", "1,639", "(1,624)", "(15)"]),
                _amount_row("비품",
                            ["174", "1,264", "(350)", "60", "1,148", "8,785",
                             "(7,617)", "(20)"]),
            ], 9))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    ac = [s for s in specs if s.kind == "asset_cost"]
    assert len(ac) == 2 and all(s.ok for s in ac), "보조금 열 포함 구성이 잡혀야 한다"
    assert all(len(s.terms) == 3 for s in ac), "취득원가+상각누계액+국고보조금 3항"


def test_inventory_rollforward_silent(tmp_path):
    """FP6: 매출원가 명세의 기초/기말완성건물재고액 — socie 개폐로 오인하지 않는다."""
    pl = table([
        [_td("과목"), _td("제 8(당) 기", COLSPAN=2), _td("제 7(전) 기", COLSPAN=2)],
        _amount_row("Ⅰ.매출액", ["", "900", "", "800"]),
        _amount_row("분양매출", ["900", "", "800", ""], 1),
        _amount_row("Ⅱ.매출원가", ["", "0", "", "700"]),
        _amount_row("기초완성건물재고액", ["0", "", "0", ""], 1),
        _amount_row("당기분양원가", ["0", "", "700", ""], 1),
        _amount_row("기말완성건물재고액", ["0", "", "0", ""], 1),
        _amount_row("Ⅲ.매출총이익", ["", "900", "", "100"]),
    ], 5)
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">손 익 계 산 서</TITLE>' + pl)
    inv = [s for s in specs
           if "재고액" in s.label and s.kind in ("socie", "total_row", "group")]
    assert not inv, "기말재고액 행은 기초+당기 소계가 아니다 (전기 열 오탐 방지)"


def test_bs_contra_flat_sum(tmp_path):
    """FP2: 총액/(충당금) 차감 표시 BS — 상위 소계는 평면합(충당금 포함), 총액행 무방출."""
    bs = table([
        [_td("과  목"), _td("제 5(당) 기", COLSPAN=2), _td("제 4(전) 기", COLSPAN=2)],
        _amount_row("유동자산", ["", "170", "", "150"]),
        _amount_row("현금및현금성자산", ["50", "", "40", ""], 1),
        _amount_row("외상매출금", ["150", "", "140", ""], 1),
        _amount_row("대손충당금", ["(30)", "", "(30)", ""], 2),
        _amount_row("자산총계", ["", "170", "", "150"]),
    ], 5)
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">재 무 상 태 표</TITLE>' + bs)
    grp = [s for s in specs if s.kind == "group"]
    top = [s for s in grp if "유동자산" in s.label]
    assert top and all(s.ok for s in top), "평면합(충당금 포함)이 정확 일치해야 한다"
    assert any(len(s.terms) == 3 for s in top), "현금+외상매출금+충당금 3항"
    assert not any("외상매출금" in s.label for s in grp), "총액=충당금 오인 방출 금지"


def test_deferred_tax_net_pair(tmp_path):
    """FP5: 이연법인세자산(부채)순액 = 자산 계 − 부채 계 (비인접 짝, 정확 일치만)."""
    note = ("<P>1. 법인세비용&amp;cr;&amp;cr;이연법인세 내역입니다.</P>"
            + table([
                [_td("구분"), _td("당기"), _td("전기")],
                _amount_row("차감할 일시적차이", ["300", "250"]),
                _amount_row("이연법인세자산 소계", ["300", "250"]),
                _amount_row("가산할 일시적차이", ["(100)", "(80)"]),
                _amount_row("이연법인세부채 소계", ["(100)", "(80)"]),
                _amount_row("이연법인세자산(부채)순액", ["200", "170"]),
            ], 3)
            + "<P>부채 양수 표기 변형입니다.</P>"
            + table([
                [_td("구분"), _td("당기")],
                _amount_row("이연법인세자산 계", ["500"]),
                _amount_row("이연법인세부채 계", ["120"]),
                _amount_row("이연법인세자산(부채)순액", ["380"]),
            ], 2))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    npair = [s for s in specs if s.kind == "net_pair"]
    assert len(npair) == 3 and all(s.ok for s in npair), "당·전기+양수변형 3건"
    assert all(len(s.terms) == 2 for s in npair), "자산 계·부채 계 2항"
    bad = [s for s in specs if not s.ok and s.kind == "total_row"]
    assert not bad, "순액 셀의 total_row 최근접 오답은 net_pair가 대체해야 한다"


def test_flow_columns_prefixed_headers(tmp_path):
    """연결 갭①: '당기초|…|당기말' 헤더 흐름표 — 접두 허용 (이연법인세 명세 실측)."""
    note = ("<P>1. 법인세&amp;cr;&amp;cr;일시적차이 변동입니다.</P>"
            + table([
                [_td("구분"), _td("당기초"), _td("연결범위 변동"), _td("포괄손익계산서"),
                 _td("당기말")],
                _amount_row("대손충당금", ["230", "0", "980", "1,210"]),
                _amount_row("재고자산", ["1,520", "0", "(56)", "1,464"]),
            ], 5))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    fl = [s for s in specs if s.kind == "flow_h"]
    assert len(fl) == 2 and all(s.ok for s in fl), "당기초→당기말 흐름이 잡혀야 한다"


def test_asset_cost_unamortized_balance(tmp_path):
    """연결 갭②: 공단형 '기말 장부가액 = 기말 미상각잔액 + 취득보조금'."""
    note = ("<P>1. 유형자산&amp;cr;&amp;cr;자기유형자산 내역입니다.</P>"
            + table([
                [_td("구분"), _td("기초미상각잔액"), _td("취득"), _td("처분"),
                 _td("기말미상각잔액"), _td("취득보조금"), _td("기말장부가액")],
                _amount_row("건물", ["1,000", "200", "0", "1,200", "(100)", "1,100"]),
                _amount_row("구축물", ["500", "0", "(50)", "450", "(30)", "420"]),
            ], 7))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    ac = [s for s in specs if s.kind == "asset_cost"]
    assert len(ac) == 2 and all(s.ok for s in ac), "미상각잔액+보조금 구성이 잡혀야 한다"
    fl = [s for s in specs if s.kind == "flow_h"]
    assert fl and all(s.ok for s in fl), "기초→기말 미상각잔액 흐름도 함께"


def test_grand_total_column_nested(tmp_path):
    """연결 갭③: CE '총자본 = 지배지분 합계 + 비지배지분' 중첩 가로합 ."""
    ce = table([
        [_td("구분", ROWSPAN=2), _td("지배기업 소유주지분", COLSPAN=3),
         _td("비지배지분", ROWSPAN=2), _td("총자본", ROWSPAN=2)],
        [_td("자본금"), _td("이익잉여금"), _td("합  계")],
        _amount_row("Ⅰ.2025.01.01(당기초)", ["100", "50", "150", "30", "180"]),
        _amount_row("당기순이익", ["0", "20", "20", "5", "25"]),
        _amount_row("연차배당", ["0", "(10)", "(10)", "0", "(10)"]),
        _amount_row("Ⅱ.2025.12.31(당기말)", ["100", "60", "160", "35", "195"]),
    ], 6, head_rows=2)
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">연 결 자 본 변 동 표</TITLE>' + ce)
    tc = [s for s in specs if s.kind == "total_col" and s.ok]
    nested = [s for s in tc if len(s.terms) == 2]
    assert len(nested) >= 4, "전 행에 총자본=합계+비지배 2항 중첩합이 잡혀야 한다"
    flat = [s for s in tc if len(s.terms) == 2 and s.target[1] != nested[0].target[1]] \
        or [s for s in tc if len(s.terms) >= 2]
    assert any(len(s.terms) == 2 for s in tc), "지배 합계열 가로합도 유지"


def test_cascade_signed(tmp_path):
    """연결 갭④: 부문정보 '당기순손익 = 영업이익 + 금융손익 − 법인세' 부호 캐스케이드."""
    note = ("<P>1. 부문정보&amp;cr;&amp;cr;부문별 손익입니다.</P>"
            + table([
                [_td("구분"), _td("당기")],
                _amount_row("보고부문 영업이익", ["100"]),
                _amount_row("금융손익 등", ["(20)"]),
                _amount_row("법인세", ["30"]),
                _amount_row("당기순손익", ["50"]),
            ], 2))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    cs = [s for s in specs if s.kind == "cascade" and s.ok]
    assert cs and any(any(sg == -1 for *_, sg in s.terms) for s in cs), \
        "법인세 차감 부호 캐스케이드가 잡혀야 한다"


def test_running_totals_discovery(tmp_path):
    """연결 갭⑤: 라벨 표식 없는 소계 체인 — 계상액 = 계 + 조정 + 미인식계 (스택 발견)."""
    note = ("<P>1. 이연법인세&amp;cr;&amp;cr;일시적차이 내역입니다.</P>"
            + table([
                [_td("구분"), _td("당기")],
                _amount_row("차감할항목A", ["123,457"]),
                _amount_row("차감할항목B", ["54,321"]),
                _amount_row("이연법인세자산(부채) 계", ["177,778"]),
                _amount_row("연결조정항목", ["(31,113)"]),
                _amount_row("미인식항목A", ["(11,117)"]),
                _amount_row("미인식항목B", ["(22,229)"]),
                _amount_row("미인식된 이연법인세부채", ["(33,346)"]),
                _amount_row("이연법인세자산(부채)계상액", ["113,319"]),
            ], 2))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    cs = [s for s in specs if s.kind == "cascade" and s.ok]
    three = [s for s in cs if len(s.terms) == 3]
    assert three, "계상액 = 계+조정+미인식계 3항(접힌 소계) 체인이 잡혀야 한다"


def test_tax_ratio_no_tags(tmp_path):
    """연결 갭⑥: 태그 힌트 없는 유효세율 = 법인세비용 ÷ 차감전순이익 (나눗셈)."""
    note = ("<P>1. 법인세비용&amp;cr;&amp;cr;유효세율 내역입니다.</P>"
            + table([
                [_td("구분"), _td("당기")],
                _amount_row("법인세비용차감전순이익", ["1,000"]),
                _amount_row("법인세비용", ["250"]),
                _amount_row("평균유효세율", ["0.25"]),
            ], 2))
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + note)
    ratio = [s for s in specs if s.kind == "combo" and s.op == "/"]
    assert ratio and all(s.ok for s in ratio), "유효세율 나눗셈이 잡혀야 한다"


_RATIO_BODY = (
    "<P>1. 대출채권&amp;cr;&amp;cr;구성내역입니다.</P>"
    + table([
        [_td("구분"), _td("금액"), _td("구성비율(%)")],
        _amount_row("기업", ["81,711", "83.6"]),
        _amount_row("가계", ["8,102", "8.3"]),
        _amount_row("공공", ["7,925", "8.1"]),
        _amount_row("계", ["97,738", "100.0"]),
    ], 3)
    + "<P>접미 % 표기 변형입니다.</P>"
    + table([
        [_td("구분"), _td("금액"), _td("구성비")],
        _amount_row("A", ["600", "60.0%"]),
        _amount_row("B", ["400", "40.0%"]),
        _amount_row("합계", ["1,000", "100.0%"]),
    ], 3)
    + "<P>불일치 비율은 빨간불로 방출합니다.</P>"
    + table([
        [_td("구분"), _td("금액"), _td("구성비율(%)")],
        _amount_row("갑", ["500", "70.0"]),
        _amount_row("을", ["500", "30.0"]),
        _amount_row("계", ["1,000", "100.0"]),
    ], 3))


def test_ratio_columns_both_scales(tmp_path):
    """비율 검증: 구성비 = 금액÷합계행 — 퍼센트 수치(factor 100)·분수(%) 두 스케일,
    표시 자릿수 반올림은 통과, 실제 어긋난 비율은 빨간불로 방출."""
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">주석</TITLE>' + _RATIO_BODY)
    ratio = [s for s in specs if s.kind == "ratio"]
    assert len(ratio) == 10, "표1 4행 + 표2 3행 + 표3 3행"
    assert {s.factor for s in ratio} == {1, 100}, "두 스케일 모두 인식"
    bad = [s for s in ratio if not s.ok]
    assert len(bad) == 2, "갑(70≠50)·을(30≠50)만 빨간불"
    assert all(s.ok for s in ratio if s.factor == 1), "분수 저장 표는 전부 일치"


def test_ratio_selfaudit_roundtrip(tmp_path):
    """비율 수식(A/B·A/B*100 + ROUND(d,n))이 자체감사 파서와 호환되는지 왕복."""
    import warnings

    import openpyxl

    from dsdxlsx.foot import pipeline
    from dsdxlsx.foot.selfaudit import audit
    from test_foot_selfaudit import _rebuild_checkset

    xml = ('<?xml version="1.0" encoding="UTF-8"?><BODY>'
           '<TITLE ATOC="Y">주석</TITLE>' + _RATIO_BODY + "</BODY>").encode()
    p = tmp_path / "ratio.dsd"
    p.write_bytes(build_dsd(xml))
    out = tmp_path / "ratio.xlsx"
    pipeline.process_file(p, out)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(str(out))
    problems = audit(wb, _rebuild_checkset(p))
    assert problems == [], problems[:4]


def test_socie_dual_variant(tmp_path):
    """자본변동표 잔액행: 아래(소계합)와 오른쪽(항목합) 이중 방출."""
    ce = table([
        [_td("과목"), _td("자본금"), _td("이익잉여금")],
        _amount_row("기초잔액", ["500", "100"]),
        _amount_row("총포괄손익", ["", ""]),
        _amount_row("당기순이익", ["0", "70"], 1),
        _amount_row("총포괄손익 합계", ["0", "70"]),
        _amount_row("기말잔액", ["500", "170"]),
    ], 3)
    specs, _ = _analyze(tmp_path, '<TITLE ATOC="Y">자 본 변 동 표</TITLE>' + ce)
    socie = [s for s in specs if s.kind == "socie"]
    end_axis = {(s.target, s.axis) for s in socie}
    # 같은 기말 셀에 col(아래)·row(오른쪽) 두 변형이 공존
    cols = {t for t, a in end_axis if a == "col"}
    rows = {t for t, a in end_axis if a == "row"}
    assert cols & rows, "이중 변형이 최소 한 셀에는 있어야 한다"


def test_price_qty_multiply(tmp_path):
    """곱셈형: 자본금 주석 — 금액 = 1주당 금액 × 주식수 (정확 일치만)."""
    t = table([
        [_td("구분"), _td("1주당 금액"), _td("발행주식수"), _td("금액")],
        _amount_row("보통주", ["5,000", "10,000", "50,000,000"]),
        _amount_row("우선주", ["5,000", "2,000", "10,000,000"]),
        _amount_row("어긋남", ["5,000", "3,000", "15,000,001"]),  # 불일치 — 방출 안 함
    ], 4)
    body = '<TITLE ATOC="Y">주석</TITLE><P>1. 자본금&amp;cr;내역.</P>' + t
    specs, _ = _analyze(tmp_path, body)
    mult = [s for s in specs if s.kind == "combo" and s.op == "*"]
    assert len(mult) == 2, "정확 일치 2행만 (어긋난 행은 추측형이라 침묵)"
    assert all(s.ok for s in mult)


def test_relaxed_gate_small_rounding(tmp_path):
    """relaxed: 소액(|기재|<1000) ±1 절사도 후보로 승격 — AI검토 전제."""
    t = table([
        [_td("구분"), _td("금액")],
        _amount_row("가", ["401"]),
        _amount_row("나", ["302"]),
        _amount_row("장부금액", ["704"]),   # 401+302=703 — ±1, 소액 (라벨 무표식 소계)
    ], 2)
    body = '<TITLE ATOC="Y">주석</TITLE><P>1. 소액표&amp;cr;내역.</P>' + t
    xml = ('<?xml version="1.0" encoding="UTF-8"?><BODY>' + body + "</BODY>").encode()
    p = tmp_path / "r.dsd"
    p.write_bytes(build_dsd(xml))
    doc = parse_contents(read_dsd(str(p)).contents)
    plans = sections.build(doc)
    views, _ctx = build_views(plans)
    strict_hit = any(s.kind == "cascade" for s in footing.analyze(views))
    relaxed_hit = any(s.kind == "cascade"
                      for s in footing.analyze(views, relaxed=True))
    assert not strict_hit, "기본 모드는 소액 ±1 게이트로 침묵"
    assert relaxed_hit, "relaxed면 소액 ±1도 절사 검증으로 방출"
    assert footing.RELAXED is False, "analyze 종료 후 플래그 원복"
