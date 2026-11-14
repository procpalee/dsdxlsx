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


# ── 나란히 보기 ─────────────────────────────────────

def _side(d, side):
    """나란히 보기 줄에서 한쪽에 실제로 있는 줄(빈 자리 뺌)."""
    return [(e.x if side == "a" else e.y) for e in d.entries if (e.x if side == "a" else e.y) is not None]


def test_entries_cover_both_docs_in_order():
    a = _doc()
    b = _doc(extra_row=True, cash="1,200", merged=True)
    d = D.diff(a, b)
    # 양쪽 문서의 모든 문단 · 표 행(빈 행 포함)이 순서대로 한 번씩 나온다
    assert [i.ref for i in _side(d, "a")] == [i.ref for i in D.flatten(a, keep_empty=True)]
    assert [i.ref for i in _side(d, "b")] == [i.ref for i in D.flatten(b, keep_empty=True)]
    # 바뀐 곳은 compare 와 같고, 추가 줄은 왼쪽이 빈 자리
    assert [c.kind for c in d.changes] == [c.kind for c in D.compare(a, b)]
    (add,) = [e for e in d.entries if e.c and e.c.kind == "added"]
    assert add.x is None and add.y.cells[0] == "정기예금"


def test_empty_rows_kept_but_not_counted():
    rows = [["예금", "1,000"], ["", ""], ["합계", "1,000"]]
    a = parse_contents(F.document(F.p("6. 현금"), F.table([["구분", "당기"]], rows)))
    b = parse_contents(F.document(F.p("6. 현금"), F.table([["구분", "당기"]], rows[:1] + rows[2:])))
    d = D.diff(a, b)
    assert d.changes == [] and D.compare(a, a) == []
    (bl,) = [e for e in d.entries if e.x is not None and e.x.empty]
    assert bl.blank and bl.y is None and bl.c is None


def test_side_table_rowspan_stretches_over_gap():
    # 왼쪽: 2줄 병합 '차입처' / 오른쪽: 가운데 줄이 하나 더 있어 3줄 병합
    a = parse_contents(F.document(F.p("17. 차입금"), F.table([["차입처", "구분", "금액"]], [
        ["||ROWSPAN=2||산업은행", "운영자금", "100"], ["시설자금", "200"]])))
    b = parse_contents(F.document(F.p("17. 차입금"), F.table([["차입처", "구분", "금액"]], [
        ["||ROWSPAN=3||산업은행", "운영자금", "100"], ["일반자금", "150"], ["시설자금", "200"]])))
    d = D.diff(a, b)
    h = D.render_html(d, ("A", "B", None), ("a.dsd", "b.dsd", None))
    assert '<td class="a bd" rowspan="3" style="text-align:left">산업은행</td>' in h   # 왼쪽 병합이 빈 자리까지 덮음
    assert '<td class="a ph" colspan="2"></td>' in h               # 빈 자리는 병합 칸 옆 두 칸만
    assert '<td class="b bd" rowspan="3" style="text-align:left">산업은행</td>' in h
    assert h.count('<table class="pair">') == 1                     # 한 표로 묶임(줄 높이 맞춤)


def test_side_whole_table_added():
    a = parse_contents(F.document(F.p("6. 현금"), F.p("끝")))
    b = parse_contents(F.document(F.p("6. 현금"), F.table([["구분", "당기"]], [["예금", "1"], ["합계", "1"]]), F.p("끝")))
    h = D.render_html(D.diff(a, b), ("A", "B", None), ("a", "b", None))
    assert '<td class="a ph" rowspan="3"></td>' in h               # 머리글 포함 3줄을 빗금 한 칸으로


def test_side_marks_changed_cells_and_chars():
    a = _doc()
    b = _doc(cash="1,200", rate="시설자금")
    h = D.render_html(D.diff(a, b, base=_doc()), ("우리", "감사인", "기준"), ("a", "b", "c"))
    assert 'class="d k-number b-a"' in h and 'class="d k-text b-a"' in h
    assert "1,<del>0</del>00" in h and "1,<ins>2</ins>00" in h       # 바뀐 글자만, 숫자는 오른쪽부터 자리 맞춤
    assert "<del>운영</del>자금" in h and "<ins>시설</ins>자금" in h
    assert 'class="a bd n x"' in h and 'class="b bd n x"' in h      # 바뀐 숫자 칸
    assert 'class="eq"' in h                                        # 같은 줄(흐리게)
    assert "id='sidev'" in h and "data-view='side'" in h and "id='cards'" in h
    assert "KeyN" in h and "KeyP" in h                              # n/p 이동


def test_cli_view_option(tmp_path):
    a, b = tmp_path / "a.dsd", tmp_path / "b.dsd"
    F.write_dsd(a, F.document(F.p("6. 현금"), F.table([["구분", "당기"]], [["예금", "1,000"]])))
    F.write_dsd(b, F.document(F.p("6. 현금"), F.table([["구분", "당기"]], [["예금", "1,200"]])))
    out = tmp_path / "r.html"
    assert main(["diff", str(a), str(b), "--view", "cards", "-o", str(out)]) == 0
    h = out.read_text(encoding="utf-8")
    assert "<body data-view='cards'>" in h and "나란히 보기에서 보기" in h


def test_cli_json_lists_changed_cells(tmp_path, capsys):
    import json

    a, b = tmp_path / "a.dsd", tmp_path / "b.dsd"
    F.write_dsd(a, F.document(F.p("6. 현금"), F.table([["구분", "당기", "전기"]], [["예금", "1,000", "900"]]),
                              F.p("담보는 없습니다.")))
    F.write_dsd(b, F.document(F.p("6. 현금"), F.table([["구분", "당기", "전기"]], [["예금", "1,200", "900"]]),
                              F.p("담보 100 이 있습니다.")))
    assert main(["diff", str(a), str(b), "-o", str(tmp_path / "r.html"), "--json"]) == 0
    env = json.loads(capsys.readouterr().out)
    rows = env["changes"]
    assert env["total"] == len(rows) == 2 and "changes_truncated" not in env
    row = next(r for r in rows if r["after"]["type"] == "row")
    assert row["kind"] == "숫자" and row["section"] == "6. 현금" and row["after"]["label"] == "예금"
    assert row["cols"] == [{"col": 1, "before": "1,000", "after": "1,200"}]
    para = next(r for r in rows if r["after"]["type"] == "para")
    assert para["before"]["text"] == "담보는 없습니다." and "cols" not in para


def test_changes_json_puts_cosmetic_last_and_cuts():
    a = _doc()
    b = _doc(cash="1,200", spaced="대체된 금액", blank="")
    rows, cut = D.changes_json(D.compare(a, b))
    kinds = [r["kind"] for r in rows]
    assert kinds[0] == "숫자" and kinds.index("공백") > kinds.index("숫자") and not cut
    rows, cut = D.changes_json(D.compare(a, b), limit=1)
    assert len(rows) == 1 and cut
