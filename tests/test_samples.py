"""다회사 샘플 회귀: sample/ 폴더의 실제 DSD(IFRS·K-GAAP)에서
①왕복 바이트 동일 ②합계검산 클린이 유지돼야 한다.

sample/은 고객 데이터라 git 제외 — 신규 클론에서는 샘플이 없어 전부 skip된다.
회사별 허용 예외(실측 확인된 문서 이상)는 gitignore된 tests/sample_exceptions.py의
ALLOWED(파일명 키워드 → {check: 최대 warn 수})에 두어 저장소에 회사 정보를 남기지 않는다.
"""
import glob
import hashlib
import os

import pytest

from dsdxlsx.check import run_checks
from dsdxlsx.dsd_io import read_dsd
from dsdxlsx.export_xlsx import export_xlsx
from dsdxlsx.import_xlsx import import_xlsx
from dsdxlsx.parse import parse_contents

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLES = sorted(glob.glob(os.path.join(PROJECT_ROOT, "sample", "*.dsd")))

try:
    from sample_exceptions import ALLOWED  # 로컬 전용(gitignore)
except ImportError:
    ALLOWED = {}


@pytest.mark.parametrize("path", SAMPLES, ids=[os.path.basename(p)[:24] for p in SAMPLES])
def test_sample_roundtrip_and_sums(path, tmp_path):
    if not SAMPLES:
        pytest.skip("sample 폴더 없음")
    pkg = read_dsd(path)
    doc = parse_contents(pkg.contents)

    # 무수정 왕복 바이트 동일
    xlsx = str(tmp_path / "t.xlsx")
    with open(path, "rb") as f:
        export_xlsx(doc, f.read(), path, xlsx)
    _, new_contents, rep = import_xlsx(xlsx)
    assert not rep.changed
    assert hashlib.sha256(new_contents).digest() == hashlib.sha256(pkg.contents).digest()

    # 합계검산(C1·C2·C3·C7) — 허용 예외 외에는 전수 통과
    report = run_checks(doc, only={"C1", "C2", "C3", "C7"})
    allowed = next((v for k, v in ALLOWED.items() if k in os.path.basename(path)), {})
    by_check = {}
    for f in report.findings:
        by_check.setdefault(f.check, []).append(f)
    for check, findings in by_check.items():
        limit = allowed.get(check, 0)
        assert len(findings) <= limit, (
            f"{os.path.basename(path)} {check} {len(findings)}건 (허용 {limit}):\n"
            + "\n".join(f"  {f.where}: {f.message}" for f in findings[:8])
        )