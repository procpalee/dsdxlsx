"""DSD ZIP 컨테이너 입출력.

DSD = ZIP(deflate): contents.xml + meta.xml. 재패킹 시 엔트리 순서·이름·타임스탬프를
원본 그대로 유지한다. meta.xml의 FILE-INFO size는 contents.xml 크기와 연동 → 갱신.
"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass


@dataclass
class DsdPackage:
    entries: list[tuple[zipfile.ZipInfo, bytes]]  # 원본 순서 유지

    @property
    def contents(self) -> bytes:
        return self._get("contents.xml")

    @property
    def meta(self) -> bytes:
        return self._get("meta.xml")

    def _get(self, name: str) -> bytes:
        for info, data in self.entries:
            if info.filename.lower() == name:
                return data
        raise FileNotFoundError(f"DSD 안에 {name}이 없음")


def read_dsd(path: str) -> DsdPackage:
    with zipfile.ZipFile(path) as zf:
        entries = [(zf.getinfo(name), zf.read(name)) for name in zf.namelist()]
    return DsdPackage(entries=entries)


def patch_meta_size(meta: bytes, new_size: int) -> bytes:
    """FILE-INFO의 size 속성만 새 contents.xml 크기로 치환."""
    def repl(m: re.Match) -> bytes:
        return m.group(1) + str(new_size).encode() + m.group(3)

    patched, n = re.subn(
        rb'(<FILE-INFO\b[^>]*?size=")(\d+)(")', repl, meta, count=1
    )
    if n != 1:
        raise ValueError("meta.xml에서 FILE-INFO size 속성을 찾지 못함")
    return patched


def write_dsd(path: str, package: DsdPackage, new_contents: bytes) -> None:
    """수정된 contents.xml로 DSD 재패킹. 무수정이면 meta도 원본 바이트 유지."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for info, data in package.entries:
            name = info.filename.lower()
            if name == "contents.xml":
                data = new_contents
            elif name == "meta.xml" and new_contents != package.contents:
                data = patch_meta_size(data, len(new_contents))
            out = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            out.compress_type = zipfile.ZIP_DEFLATED
            out.external_attr = info.external_attr
            zf.writestr(out, data)
