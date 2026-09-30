"""공시사항점검표 자동 판정 — 합성 review.json 으로 규칙 동작만 확인(한공회 원본 불필요)."""
from dsdxlsx.disclosure import RULES_KGAAP, RULES_KIFRS, Evidence, Item, judge, roll_up


def _review(notes, fs=(), paras=()):
    return {"document": {
        "notes": [{"no": n, "title": t} for n, t in notes],
        "fs_rows": [{"section": sec, "label": lab, "amounts": {"당기": amt}, "refs": refs}
                    for sec, lab, amt, refs in fs],
        "paragraphs": [{"section": sec, "text": t} for sec, t in paras],
        "tables": [],
    }}


def _judge(rules, code, ev):
    it = Item(code=code, title="", row=0, level=code.count("."))
    judge(it, rules[code], ev)
    return it


def test_note_title_and_balance_give_y():
    ev = Evidence(_review([(1, "일반사항"), (2, "중요한 회계정책"), (9, "재고자산")],
                          fs=[("재무상태표", "재고자산", 100, [9])],
                          paras=[("표지", "한국채택국제회계기준에 따라 작성")]))
    it = _judge(RULES_KIFRS, "DN2.4.1", ev)
    assert it.verdict == "Y" and it.notes == [9] and not it.warn


def test_balance_without_note_warns():
    ev = Evidence(_review([(2, "중요한 회계정책")], fs=[("재무상태표", "재고자산", 100, [])],
                          paras=[("주석2", "재고자산은 총평균법으로 평가합니다")]))
    it = _judge(RULES_KGAAP, "DN2.1.3", ev)
    assert it.verdict == "Y" and it.warn.startswith("주석 누락 후보") and "회계정책" in it.warn


def test_policy_note_only_mention_is_na():
    ev = Evidence(_review([(2, "중요한 회계정책")], paras=[("주석2", "사업결합은 취득법으로 회계처리합니다")]))
    it = _judge(RULES_KIFRS, "DN2.7", ev)
    assert it.verdict == "NA" and "회계정책" in it.basis


def test_income_statement_line_is_not_balance_evidence():
    ev = Evidence(_review([(2, "중요한 회계정책")], fs=[("손익계산서", "지분법이익", 50, [])]))
    it = _judge(RULES_KGAAP, "DN2.1.5", ev)
    assert it.verdict == "NA"


def test_strong_rule_text_only_needs_check():
    ev = Evidence(_review([(31, "특수관계자 거래")],
                          paras=[("주석31", "관계기업 및 기타 특수관계자"), ("주석31", "관계기업 거래 없음")]))
    it = _judge(RULES_KIFRS, "DN2.4.8.3", ev)
    assert it.verdict == "확인 필요"


def test_listed_only_items_na_for_private_company():
    ev = Evidence(_review([(2, "중요한 회계정책")]))
    ev.listed = False
    assert _judge(RULES_KIFRS, "DN4.2", ev).verdict == "NA"


def test_roll_up_parent():
    items = [Item("DN2", "", 0, 1), Item("DN2.1", "", 0, 1, verdict="NA"), Item("DN2.2", "", 0, 1, verdict="Y", notes=[3])]
    roll_up(items)
    assert items[0].verdict == "Y" and items[0].notes == [3]
