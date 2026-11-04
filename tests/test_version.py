"""버전 표기 - pyproject 와 패키지 __version__ 이 같고 --version 이 그 값을 낸다."""
import os
import re

import pytest

from dsdxlsx import __version__
from dsdxlsx.cli import main

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def test_version_matches_pyproject():
    with open(os.path.join(ROOT, "pyproject.toml"), encoding="utf-8") as f:
        m = re.search(r'^version = "([^"]+)"', f.read(), re.M)
    assert m and m.group(1) == __version__


def test_cli_version(capsys):
    with pytest.raises(SystemExit) as e:
        main(["--version"])
    assert e.value.code == 0
    assert capsys.readouterr().out.strip() == f"dsdxlsx {__version__}"
