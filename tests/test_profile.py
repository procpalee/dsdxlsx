"""프로파일 인프라 — 기본값=v1 하드코딩 동일, 회사 override 병합, suppress."""
import textwrap

from dsdxlsx.check import run_checks
from dsdxlsx.parse import parse_contents
from dsdxlsx.profile import (
    default_profile, load_profile, regname_of, resolve_profile,
)

from fixtures import factory


def test_default_matches_v1_hardcoded_values():
    p = default_profile()
    assert p.name == "_default"
    assert p.footing_tol == 1
    assert p.global_search_min == 1000
    assert p.total_re.match("합계") and p.total_re.match("소  계")
    assert p.grand_total_re.search("자산총계")
    assert p.deduct_total_re.search("차감후금액")
    assert p.unit_mult["천원"] == 1_000
    assert p.anchor_re.search("영업이익")
    assert "합계" in p.generic_labels
    assert not p.suppress


def _write_company(tmp_path, name, body):
    d = tmp_path / "profiles" / name
    d.mkdir(parents=True)
    (d / "profile.toml").write_text(textwrap.dedent(body), encoding="utf-8")


def test_company_override_merges_partially(tmp_path):
    _write_company(tmp_path, "테스트사", """
        [tolerance]
        footing = 0

        [[suppress]]
        check = "C2"
        where_contains = "주석3"
        reason = "전기 심리완료 — 들여쓰기 관례 차이"
    """)
    p = load_profile("테스트사", base_dir=tmp_path)
    assert p.name == "테스트사"
    assert p.footing_tol == 0                       # override 반영
    assert p.global_search_min == 1000              # 미지정 키는 기본값
    assert p.total_re.match("합계")                 # 기본 라벨 유지
    assert p.suppress[0].check == "C2" and "심리완료" in p.suppress[0].reason


def test_company_name_normalization(tmp_path):
    _write_company(tmp_path, "합성테스트", "[tolerance]\nfooting = 2\n")
    # '(주)'·공백 차이는 정규화 매치
    p = load_profile("합성테스트(주)", base_dir=tmp_path)
    assert p.footing_tol == 2


def test_regname_and_resolve(tmp_path):
    meta = factory.meta(regname="합성테스트(주)")
    assert regname_of(meta) == "합성테스트(주)"
    # 회사 프로파일이 없으면 기본으로 폴백
    p = resolve_profile(meta, None, base_dir=tmp_path)
    assert p.name == "_default"
    _write_company(tmp_path, "합성테스트(주)", "[tolerance]\nfooting = 3\n")
    p = resolve_profile(meta, None, base_dir=tmp_path)
    assert p.footing_tol == 3


def test_synthetic_doc_passes_all_checks():
    doc = parse_contents(factory.standard_contents())
    report = run_checks(doc)
    problems = [f for f in report.findings if f.level in ("warn", "error")]
    assert problems == []
    # 검사가 실제로 수행됐는지 (침묵 통과 방지)
    assert report.counts.get("C2", 0) >= 1   # Ⅰ.유동자산 그룹
    assert report.counts.get("C3", 0) >= 3   # 자산총계·주석4 합계·주석5 차감후
    assert report.counts.get("C4", 0) >= 2   # 현금·매출채권 대사


def test_period_bucket_gisu_fallback():
    """(당)/(전) 표기 없는 '제 N 기' 헤더 — 최대 기수=당기, 둘째=전기."""
    from dsdxlsx.check import build_views

    contents = factory.document(
        factory.p("표지"),
        factory.title("재무제표", 2),
        factory.statement_title_table("재 무 상 태 표"),
        factory.table(
            [["과  목", "제 26 기 1분기말", "제 25 기말"]],
            [["현금", "1,000", "900"], ["자산총계", "1,000", "900"]],
            unit="원",
        ),
    )
    views, ctx = build_views(parse_contents(contents))
    tv = next(v for v in ctx["statements"] if v.ncols == 3)  # 제목표 제외
    assert tv.period_bucket(1) == "당기"
    assert tv.period_bucket(2) == "전기"
    assert tv.period_bucket(0) is None


def test_coverage_reports_unchecked_tables():
    """주석7(합계 없는 단일행 표)은 재무 검산 미적용 — 커버리지에 드러나야 한다."""
    doc = parse_contents(factory.standard_contents())
    report = run_checks(doc)
    assert report.coverage
    uncovered = [c for c in report.coverage if not c["skipped"] and not c["checks"]]
    assert any(c["context"].startswith("주석7") for c in uncovered)
    covered = [c for c in report.coverage if c["checks"]]
    assert covered  # 검산 적용 표도 존재


def test_injected_error_detected_then_suppressed(tmp_path):
    # 주석4 합계를 일부러 틀리게 → C3 warn
    doc = parse_contents(factory.standard_contents(cash_note_total="1,100"))
    report = run_checks(doc)
    c3 = [f for f in report.findings if f.check == "C3" and f.level == "warn"]
    assert len(c3) >= 1 and any("주석4" in f.where for f in c3)

    # 같은 오류를 회사 프로파일 suppress 규칙으로 억제 — findings에서 빠지고 집계에 남음
    _write_company(tmp_path, "합성테스트(주)", """
        [[suppress]]
        check = "C3"
        where_contains = "주석4"
        reason = "검토 완료 — 의도된 반올림 차이"
    """)
    profile = load_profile("합성테스트(주)", base_dir=tmp_path)
    report2 = run_checks(doc, profile=profile)
    assert not [f for f in report2.findings if f.check == "C3" and "주석4" in f.where]
    assert any("주석4" in f.where for f, _ in report2.suppressed)
    assert report2.profile_name == "합성테스트(주)"
