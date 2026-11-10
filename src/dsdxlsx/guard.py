"""입력 파일 점검과 출력 덮어쓰기 보호.

입력: 없는 파일, DSD 자리에 엑셀(또는 그 반대), 옛 엑셀 형식(.xls), 깨진 파일을
      명령이 돌기 전에 걸러 한 줄로 알린다(InputError → 종료코드 2).

출력: 사람이 고치는 결과물(편집용 엑셀, 정산표, 결과 DSD)은 이미 있으면 그냥 덮어쓰지 않는다.
      dsdxlsx 가 만든 그대로면 덮어쓰고, 만든 뒤에 바뀌었거나 만든 기록이 없으면 멈춘다
      (Stopped → 종료코드 3). 기록은 파일 옆이 아니라 PC 의 dsdxlsx 폴더(outputs.json)에 둔다.
"""
from __future__ import annotations

import hashlib
import json
import os
import zipfile
from pathlib import Path


class InputError(Exception):
    """입력 파일이 없거나 형식이 다르다."""


class Stopped(Exception):
    """안전장치로 멈춤 - 아무것도 쓰지 않았다."""


# ── 입력 ─────────────────────────────────────────────────────────────

_WANT = {
    "dsd": "DSD 파일(.dsd)",
    "xlsx": "dsdxlsx export 로 만든 엑셀(.xlsx)",
    "either": "DSD 파일(.dsd)이나 export 로 만든 엑셀(.xlsx)",
    "workbook": "엑셀 파일(.xlsx, .xlsm)",
}


def kind_of(path: str) -> str:
    """'dsd' | 'xlsx' | 'xls' | 'zip' | 'other'. 없는 파일·폴더는 InputError."""
    if not path or not os.path.exists(path):
        raise InputError(f"파일이 없다: {path}")
    if os.path.isdir(path):
        raise InputError(f"폴더다 - 파일을 지정해야 한다: {path}")
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except PermissionError:
        raise InputError(f"파일을 읽을 수 없다(다른 프로그램이 잠갔거나 권한이 없음): {path}") from None
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return "xls"
    if not head.startswith(b"PK"):
        return "other"
    try:
        with zipfile.ZipFile(path) as zf:
            names = [n.lower() for n in zf.namelist()]
    except zipfile.BadZipFile:
        raise InputError(f"파일이 깨졌다(압축을 풀 수 없음): {path}") from None
    if "contents.xml" in names:
        return "dsd"
    if any(n.startswith("xl/") for n in names):
        return "xlsx"
    return "zip"


def need(path: str, want: str) -> str:
    """path 가 want('dsd'·'xlsx'·'either'·'workbook') 형식인지 보고 실제 종류를 돌려준다."""
    kind = kind_of(path)
    ok = {"dsd": ("dsd",), "xlsx": ("xlsx",), "either": ("dsd", "xlsx"),
          "workbook": ("xlsx",)}[want]
    if kind in ok:
        return kind
    if kind == "xls":
        got = "옛 엑셀 형식(.xls)이거나 암호가 걸린 엑셀이다 - 암호 없는 .xlsx 로 다시 저장해야 한다"
    elif kind == "dsd":
        got = "DSD 파일이다"
        if want in ("xlsx",):
            got += " - 엑셀로 바꾸려면 dsdxlsx export"
    elif kind == "xlsx":
        got = "엑셀 파일이다"
        if want == "dsd":
            got += " - DSD 로 바꾸려면 dsdxlsx import"
    else:
        got = "DSD 도 엑셀도 아니다"
    raise InputError(f"{path}: {got}. 이 자리에는 {_WANT[want]}이 와야 한다")


# ── 출력 ─────────────────────────────────────────────────────────────

def state_dir() -> Path:
    base = (os.environ.get("DSDXLSX_STATE") or os.environ.get("LOCALAPPDATA")
            or os.environ.get("XDG_STATE_HOME") or str(Path.home()))
    d = Path(base) / "dsdxlsx"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _store() -> Path:
    return state_dir() / "outputs.json"


def _key(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def _sha(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _load() -> dict:
    try:
        return json.loads(_store().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def remember(*paths: str) -> None:
    """방금 쓴 출력 파일을 기록한다(다음에 덮어써도 되는지 판단할 기준)."""
    try:
        data = _load()
        for p in paths:
            if p and os.path.isfile(p):
                data[_key(p)] = _sha(p)
        if len(data) > 3000:                      # 지워진 파일 기록 정리
            data = {k: v for k, v in data.items() if os.path.exists(k)}
        _store().write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def untouched(path: str) -> bool | None:
    """dsdxlsx 가 만든 그대로면 True, 만든 뒤 바뀌었으면 False, 기록이 없으면 None."""
    want = _load().get(_key(path))
    if want is None:
        return None
    try:
        return _sha(path) == want
    except OSError:
        return False


def check_out(outs, *, force: bool = False, inputs=()) -> None:
    """출력 경로들을 써도 되는지 본다. 안 되면 Stopped."""
    ins = {_key(i) for i in inputs if i}
    for o in outs:
        if not o:
            continue
        if _key(o) in ins:
            raise Stopped(f"입력 파일 자리에 쓰지 않는다: {o} - 다른 -o 를 주어야 한다")
    if force:
        return
    for o in outs:
        if not o or not os.path.exists(o):
            continue
        st = untouched(o)
        if st is True:
            continue
        why = ("dsdxlsx 가 만든 뒤에 고쳐진 파일이다" if st is False
               else "이미 있는 파일이고 이 PC 에서 dsdxlsx 가 만든 기록이 없다")
        raise Stopped(f"{o}: {why} - 덮어쓰지 않았다. 다른 -o 를 주거나, 덮어써도 되면 --force")
