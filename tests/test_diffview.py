"""diff - 두 DSD 보이는 차이 보고서(합성 문서)."""
from __future__ import annotations

from dsdxlsx import diffview as D
from dsdxlsx.cli import main
from dsdxlsx.parse import parse_contents

from fixtures import factory as F


def _doc(*, cash="1,000", label="보통예금", blank="-", note_text="담보 금액은 12,345,678천원입니다.",
         extra_row=False, spaced="대체 된 금액", merged=False, rate="운영자금"):
    rows = [[label, cash, "900"], ["외화예금", "200", blank]]
    if extra_row:
        rows.append(["정기예금", "50", "-"])
    loans = [["산업은행", rate, "5.25"], ["", "시설자금", "4.82"]] if not merged else \
        [["산업은행", rate, "5.25"], ["시설자금", "4.82"]]
    return parse_contents(F.document(
        F.title("재무제표"),
        F.p("6. 현금및현금성자산"),
        F.table([["구분", "당기", "전기"]], rows, unit="원"),
        F.p(note_text),
        F.p("17. 차입금"),
        F.table([["차입처", "구분", "이자율"]], loans),
        F.p(f"(*) 건설중인자산에서 {spaced}입니다."),
    ))


def test_kinds():
    a = _doc()
    b = _doc(cash="1,200", blank="", note_text="담보 금액은 23,456,789천원입니다.", extra_row=True,
             spaced="대체된 금액", rate="시설자금", merged=True)
    ch = D.compare(a, b)
    kinds = {(c.note, c.kind) for c in ch}
    assert ("6. 현금및현금성자산", "number") in kinds         # 1,000 → 1,200 · 문단 숫자
    assert ("6. 현금및현금성자산", "dash") in kinds           # '-' → 빈 칸
    assert ("6. 현금및현금성자산", "added") in kinds          # 정기예금 줄
    assert ("17. 차입금", "text") in kinds                   # 운영자금 → 시설자금
    assert ("17. 차입금", "layout") in kinds                 # 빈 칸이 빠진 병합 줄
    assert ("17. 차입금", "space") in kinds                  # '대체 된' → '대체된'
    s = D.summary(ch)
    assert s["total"] == len(ch) and s["by_note"]["6. 현금및현금성자산"] >= 3


def test_same_doc_no_change():
    assert D.compare(_doc(), _doc()) == []


def test_base_marks_inherited_and_changed():
    base = _doc(cash="700")                  # 직전 분기: 숫자는 다르지만 모양은 같음, '-'·운영자금은 그대로
    a = _doc()
    b = _doc(blank="", rate="시설자금")
    ch = {c.kind: c for c in D.compare(a, b, base)}
    assert ch["text"].base == "a"            # 운영자금 - 기준에도 있음(이어받은 오류)
    assert ch["dash"].base == "a"


def test_base_period_words_and_split_paragraph():
    # 직전 반기 문단은 '전반기말' 이고 다음 주석 제목까지 한 문단 - 그래도 이어받은 숫자로 잡혀야 함
    base = parse_contents(F.document(F.p("10. 유형자산"), F.p("(*) 전반기말 담보 12,345,678천원입니다.\n11. 무형자산")))
    a = parse_contents(F.document(F.p("10. 유형자산"), F.p("(*) 전분기말 담보 12,345,678천원입니다.")))
    b = parse_contents(F.document(F.p("10. 유형자산"), F.p("(*) 전분기말 담보 23,456,789천원입니다.")))
    (c,) = D.compare(a, b, base)
    assert c.kind == "number" and c.base == "a"


def test_html_and_cli(tmp_path):
    a, b = tmp_path / "a.dsd", tmp_path / "b.dsd"
    F.write_dsd(a, F.document(F.p("6. 현금"), F.table([["구분", "당기"]], [["예금", "1,000"]])))
    F.write_dsd(b, F.document(F.p("6. 현금"), F.table([["구분", "당기"]], [["예금", "1,200"]])))
    out = tmp_path / "r.html"
    assert main(["diff", str(a), str(b), "--labels", "우리,감사인", "-o", str(out)]) == 0
    h = out.read_text(encoding="utf-8")
    assert "<!doctype html>" in h and "우리" in h and "감사인" in h
    assert "chg-a" in h and "chg-b" in h and "1,200" in h and "6. 현금" in h
