"""풋팅 시험 준비물 - 합성 문서, 로컬 실제 샘플(있을 때만)."""
import os
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]


def discover_samples() -> list[Path]:
    """실제 샘플 .dsd - 저장소의 sample/ 과 환경변수 DSDXLSX_FOOT_SAMPLES(폴더, ; 로 여러 개). git 제외."""
    dirs = [ROOT / "sample"] + [Path(d) for d in os.environ.get("DSDXLSX_FOOT_SAMPLES", "").split(";") if d]
    return [p for d in dirs if d.is_dir() for p in sorted(d.rglob("*.dsd"))]


def pytest_generate_tests(metafunc):
    """foot_sample_dsd 인자를 쓰는 시험을 샘플 파일별로 돌린다(샘플이 없으면 수집되지 않음)."""
    if "foot_sample_dsd" in metafunc.fixturenames:
        paths = discover_samples()
        metafunc.parametrize("foot_sample_dsd", paths, ids=[f"s{i:02d}" for i in range(len(paths))])


@pytest.fixture()
def synth_doc(tmp_path):
    """합성 DSD → (package, doc)."""
    from fixtures.foot_factory import write_dsd

    from dsdxlsx.dsd_io import read_dsd
    from dsdxlsx.parse import parse_contents

    p = tmp_path / "synth.dsd"
    write_dsd(p)
    package = read_dsd(str(p))
    return package, parse_contents(package.contents)
