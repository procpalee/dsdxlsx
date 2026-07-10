"""Phase 0 스파이크 — 편집기 재저장본 diff 분석.

변형.dsd(도구 산출) vs 변형_저장본.dsd(DART 편집기 재저장)를 비교해
편집기가 재작성하는 것(속성·공백·WIDTH·meta 등)을 관찰한다.
이 관찰이 compose/blockops의 산술 규칙(docs/DART호환성체크리스트.md)을 확정한다.

사용법: PYTHONPATH=src uv run --no-editable python spike/analyze_resave.py [spike폴더]
"""
from __future__ import annotations

import difflib
import hashlib
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from dsdxlsx.parse import parse_contents


def sha(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()[:12]


def zip_layout(path: Path) -> list[tuple]:
    with zipfile.ZipFile(path) as zf:
        return [(i.filename, i.compress_type, i.date_time, i.file_size)
                for i in zf.infolist()]


def entry(path: Path, name: str) -> bytes:
    with zipfile.ZipFile(path) as zf:
        for n in zf.namelist():
            if n.lower() == name:
                return zf.read(n)
    return b""


def line_diff(a: bytes, b: bytes, label: str, limit: int = 30) -> list[str]:
    """\\r\\n 라인 기준 unified diff (표시용 — 변경 패턴 관찰)."""
    al = a.decode("utf-8", errors="replace").splitlines()
    bl = b.decode("utf-8", errors="replace").splitlines()
    out = []
    for line in difflib.unified_diff(al, bl, lineterm="", n=0,
                                     fromfile=f"{label}(도구)", tofile=f"{label}(편집기)"):
        if line.startswith(("---", "+++", "@@")):
            if line.startswith("@@"):
                out.append(line)
            continue
        out.append(line[:220])
        if len(out) >= limit:
            out.append(f"… (diff 라인 {limit}개 초과 — 이하 생략)")
            break
    return out


def classify_patterns(a: bytes, b: bytes) -> list[str]:
    """자주 의심되는 재작성 패턴 자동 감지."""
    notes = []
    for pat, name in [
        (rb'editver="[\d.]+"', "GENERATOR editver"),
        (rb'<FILE-INFO[^>]*>', "FILE-INFO (date/size)"),
        (rb'WIDTH="\d+"', "WIDTH 속성"),
        (rb'HEIGHT="\d+"', "HEIGHT 속성"),
    ]:
        va = re.findall(pat, a)
        vb = re.findall(pat, b)
        if va != vb:
            diffs = [(x, y) for x, y in zip(va, vb) if x != y][:4]
            extra = f" (개수 {len(va)}→{len(vb)})" if len(va) != len(vb) else ""
            notes.append(f"{name} 변경{extra}: " + "; ".join(
                f"{x.decode()}→{y.decode()}" for x, y in diffs) if diffs or extra
                else f"{name} 변경")
    # 공백류 정규화 감지: 공백 제거 후 동일하면 순수 whitespace 차이
    if a != b and re.sub(rb"\s+", b"", a) == re.sub(rb"\s+", b"", b):
        notes.append("차이가 전부 공백/개행 (내용 동일)")
    return notes


def structural_equal(a: bytes, b: bytes) -> list[str]:
    """모델 수준 비교: 표/행/셀 수와 전체 셀·문단 텍스트."""
    da, db = parse_contents(a), parse_contents(b)
    notes = []
    if len(da.tables) != len(db.tables):
        notes.append(f"표 수 {len(da.tables)} → {len(db.tables)}")
    ta = [(t.ncols, len(t.rows)) for t in da.tables.values()]
    tb = [(t.ncols, len(t.rows)) for t in db.tables.values()]
    if ta != tb:
        notes.append("표 열/행 구성 상이")
    xa = [n.text for n in da.nodes.values()]
    xb = [n.text for n in db.nodes.values()]
    if xa != xb:
        changed = sum(1 for p, q in zip(xa, xb) if p != q)
        notes.append(f"텍스트 노드 차이 {changed}건 (+수량차 {len(xb)-len(xa):+d})")
    if not notes:
        notes.append("구조·텍스트 완전 동일")
    return notes


def main() -> int:
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent
    pairs = []
    for v in sorted(folder.glob("V*_*.dsd")):
        if v.stem.endswith("_저장본"):
            continue
        saved = folder / f"{v.stem}_저장본.dsd"
        if saved.exists():
            pairs.append((v, saved))
    if not pairs:
        print("저장본(*_저장본.dsd)이 없습니다.")
        return 2

    for v, s in pairs:
        print(f"\n{'=' * 72}\n■ {v.stem}  ({v.stat().st_size:,}B → 저장본 {s.stat().st_size:,}B)")
        la, lb = zip_layout(v), zip_layout(s)
        if [x[0] for x in la] != [x[0] for x in lb]:
            print(f"  ZIP 엔트리 변경: {[x[0] for x in la]} → {[x[0] for x in lb]}")
        for name in ("contents.xml", "meta.xml"):
            a, b = entry(v, name), entry(s, name)
            if a == b:
                print(f"  {name}: ✅ 바이트 동일 (sha {sha(a)})")
                continue
            print(f"  {name}: ✏️ 재작성됨 ({len(a):,}B → {len(b):,}B, {len(b)-len(a):+,}B)")
            for n in classify_patterns(a, b):
                print(f"     - {n}")
            if name == "contents.xml":
                for n in structural_equal(a, b):
                    print(f"     - [구조] {n}")
            for line in line_diff(a, b, name):
                print(f"     {line}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
