import glob
import os
import sys

import pytest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # tests/ 로컬 모듈 임포트


def find_sample_dsd() -> str | None:
    hits = sorted(glob.glob(os.path.join(PROJECT_ROOT, "*.dsd")))
    hits += sorted(glob.glob(os.path.join(PROJECT_ROOT, "samples", "*.dsd")))
    # 도구 산출물(수정본 등)이 아닌 원본을 우선 — 검산 테스트는 심리 완료 원본 기준
    originals = [h for h in hits if "수정본" not in h and "왕복" not in h]
    hits = originals or hits
    return hits[0] if hits else None


@pytest.fixture(scope="session")
def sample_dsd() -> str:
    path = find_sample_dsd()
    if path is None:
        pytest.skip("샘플 .dsd 파일 없음 (gitignore 대상이라 로컬 전용)")
    return path


@pytest.fixture(autouse=True)
def _dsdxlsx_state(tmp_path_factory, monkeypatch):
    """출력 기록(outputs.json)이 시험마다 따로, 실제 PC 폴더를 건드리지 않게."""
    monkeypatch.setenv("DSDXLSX_STATE", str(tmp_path_factory.mktemp("state")))
