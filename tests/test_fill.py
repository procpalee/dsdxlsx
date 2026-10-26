"""DSD 당기 칸 채우기(fill) — 밀기(roll-interim) 뒤 골격 × 합성 조서, 라벨 · 단위 · 합계 · 치환 예외."""
from __future__ import annotations

import openpyxl

from dsdxlsx import fill as F
from dsdxlsx import period as P
from dsdxlsx import recon as R
from dsdxlsx import splice
from dsdxlsx.parse import parse_contents

from test_rollinterim import _doc_1q_2026, _doc_2q_2025, _row, _run


def _skeleton():
    doc2, _roll, _d = _run(_doc_1q_2026(), _doc_2q_2025())
    return doc2


def _wp(tmp_path):
    """2026 반기 조서(8110 모양): 반기말 · 전기말 / 반기 3개월 · 누적 · 전기 반기."""
    wb = openpyxl.Workbook()
    bs = wb.active
    bs.title = "BS"
    bs.append(["계정과목", None, "2026.06.30", "2025.12.31"])
    bs.append(["현금및현금성자산", None, 1500, 1000])
    bs.append(["대손충당금", None, -120, -90])
    bs.append(["자산총계", None, 1500, 1000])
    pl = wb.create_sheet("IS")
    pl.append(["계정과목", None, "2026.06.30", None, "2025.06.30", None])
    pl.append([None, None, "최종3개월", "누적", "최종3개월", "누적"])
    pl.append(["매출액", None, 2100, 4100, 1900, 3700])
    pl.append(["매출원가", None, -600, -1500, -500, -900])
    pl.append(["반기순이익", None, -50, 150, 80, 200])
    cf = wb.create_sheet("CF")
    cf.append(["구분", None, "2026.06.30"])
    cf.append(["영업활동현금흐름", None, 350])
    cf.append(["유형자산의 취득", None, 60])
    cf.append(["기말 현금및현금성자산", None, 1500])
    nt = wb.create_sheet("주석")
    nt.append([None, "1. 임대수익"])
    nt.append([None, "구분", "당반기", "전반기"])
    nt.append([None, "임대수익", 61, 55])
    nt.append([None, "2. 보증금"])
    nt.append([None, "구분", "당반기말", "전기말"])
    nt.append([None, "보증금", 12, 9])
    path = tmp_path / "8110.xlsx"
    wb.save(path)
    return str(path)


def _filled(tmp_path):
    doc = _skeleton()
    ops, rep = F.fill_ops(doc, R.load_workpaper(_wp(tmp_path)))
    return parse_contents(splice.apply(doc.data, ops)), rep


def test_fill_statements_and_notes(tmp_path):
    doc, rep = _filled(tmp_path)
    assert _row(doc, "현금및현금성자산")[1:] == ["1,500", "1,000"]
    assert _row(doc, "대손충당금")[1:] == ["(120)", "(90)"]          # 앵커 = 전기말 값 일치
    assert _row(doc, "매출액")[1:] == ["2,100", "4,100", "1,900", "3,700"]
    assert _row(doc, "영업활동현금흐름")[1:] == ["350", "300"]     # 전기 = 전년 동기 DSD
    # 현금흐름표 유출 행: 조서 양수 · DSD 괄호(전기 앵커가 부호 관례를 정함)
    assert _row(doc, "유형자산의 취득")[1] in ("(60)", "60")
    assert _row(doc, "임대수익")[1:] == ["61", "55"]
    assert _row(doc, "보증금")[1:] == ["12", "9"]
    assert rep.count("채움") >= 10


def test_pl_result_label_regex():
    rx = F._PL_RESULT_RE
    for label, want in (("V. 영업손실", "V. 영업이익(손실)"), ("Ⅲ. 매출총이익", "Ⅲ. 매출총이익(손실)"),
                        ("Ⅹ. 분기순손실", "Ⅹ. 분기순이익(손실)")):
        m = rx.search(label)
        assert m and label[:m.start()] + m.group(1) + "이익(손실)" + label[m.end():] == want
    assert rx.search("반기순이익(손실)") is None          # 이미 '(손실)'


def test_fill_then_recon_has_no_mismatch(tmp_path):
    doc, _rep = _filled(tmp_path)
    rep = R.reconcile(doc, R.load_workpaper(_wp(tmp_path)))
    bad = [x for x in rep.items if x.verdict in ("불일치", "부호 다름") and x.period == "당기"]
    assert not bad, [(x.sheet, x.label, x.col, x.dsd, x.wp) for x in bad]


def test_vocab_protects_general_words():
    r = P.vocab_rules(2, 3)
    assert P.apply_vocab("일반기업회계기준에 따른 반기재무제표", r) == "일반기업회계기준에 따른 분기재무제표"
    r = P.vocab_rules(1, 2)
    assert P.apply_vocab("구분기준에 따른 분기순이익", r) == "구분기준에 따른 반기순이익"


def test_pick_same_two_unit_sides():
    class Row:
        def __init__(self, n):
            self.row, self.sheet = n, "주석"
    opts = [(Row(1), 3, 6334709.0, None), (Row(1), 14, 6334709642.0, None)]
    got = F._pick_same(opts, 1000, 50000.0)
    assert got and got[0] == 6334709.0
