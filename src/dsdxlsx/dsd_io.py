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
    """FILE-INFO의 size 속성만 새 contents.xml 크기로 치환.

    회사·편집기 버전에 따라 size=""(빈 값)이거나 FILE-INFO 자체가 없는
    meta.xml이 실존한다(K-GAAP 샘플 실측). 그 경우 편집기가 쓴 형태를
    그대로 유지한다 — 빈 값은 빈 값대로, 없으면 없는 대로.
    """
    def repl(m: re.Match) -> bytes:
        if not m.group(2):  # size="" → 원본 관례 유지
            return m.group(0)
        return m.group(1) + str(new_size).encode() + m.group(3)

    patched, _ = re.subn(
        rb'(<FILE-INFO\b[^>]*?size=")(\d*)(")', repl, meta, count=1
    )
    return patched


class _RawNameInfo(zipfile.ZipInfo):
    """원본 zip의 파일명 바이트·인코딩 플래그를 그대로 보존.

    Python zipfile은 비ASCII 파일명을 무조건 UTF-8(+flag 0x800)로 다시 쓰지만,
    DART 편집기가 만든 zip은 첨부 이미지의 한글 이름을 cp949(플래그 없음)로
    저장한다. 재패킹에서 이름 바이트가 바뀌면 meta.xml의 IMG 참조와 어긋나
    편집기가 파일 열기를 거부한다 (실측 — 날인 이미지 동봉 문서)."""

    def _encodeFilenameFlags(self):
        raw = getattr(self, "_raw_name", None)
        if raw is not None:
            return raw, self.flag_bits | self._raw_encoding_flag
        return super()._encodeFilenameFlags()


def _raw_name_bytes(info: zipfile.ZipInfo) -> bytes:
    """읽을 때 디코드된 파일명 → 원본 헤더 바이트 복원.
    zipfile은 flag 0x800이면 UTF-8, 아니면 cp437로 디코드했다 — 둘 다 무손실."""
    if info.flag_bits & 0x800:
        return info.filename.encode("utf-8")
    return info.filename.encode("cp437")


def write_dsd(path: str, package: DsdPackage, new_contents: bytes) -> None:
    """수정된 contents.xml로 DSD 재패킹. 무수정 엔트리(첨부 이미지 포함)는
    이름 바이트·플래그까지 원본 그대로 유지."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for info, data in package.entries:
            name = info.filename.lower()
            if name == "contents.xml":
                data = new_contents
            elif name == "meta.xml" and new_contents != package.contents:
                data = patch_meta_size(data, len(new_contents))
            out = _RawNameInfo(info.filename, date_time=info.date_time)
            out.compress_type = zipfile.ZIP_DEFLATED
            out.external_attr = info.external_attr
            out.internal_attr = info.internal_attr
            out.create_system = info.create_system
            out._raw_name = _raw_name_bytes(info)
            out._raw_encoding_flag = info.flag_bits & 0x800
            zf.writestr(out, data)
