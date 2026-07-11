"""dsd_io 단위 테스트 — meta.xml 실측 편차(size="", FILE-INFO 부재) 관용."""
from dsdxlsx.dsd_io import patch_meta_size


META_NUM = '<?xml version="1.0"?>\n<METAINFO><FILE-INFO date="20260331" size="423853"/></METAINFO>'.encode()
META_EMPTY = '<?xml version="1.0"?>\n<METAINFO><FILE-INFO date="20260331" size=""/></METAINFO>'.encode()
META_ABSENT = '<?xml version="1.0"?>\n<METAINFO><DOCUMENT-INFO docver="1"/></METAINFO>'.encode()


def test_numeric_size_is_replaced():
    out = patch_meta_size(META_NUM, 999)
    assert b'size="999"' in out
    assert b'size="423853"' not in out


def test_empty_size_is_kept_empty():
    # K-GAAP 실측: size="" — 편집기 관례를 그대로 유지해야 한다
    assert patch_meta_size(META_EMPTY, 999) == META_EMPTY


def test_missing_fileinfo_is_tolerated():
    assert patch_meta_size(META_ABSENT, 999) == META_ABSENT


def test_repack_preserves_cp949_attachment_names(tmp_path):
    """DART 편집기 zip은 한글 첨부 이름을 cp949(플래그 없음)로 저장 —
    재패킹이 UTF-8로 바꾸면 meta의 IMG 참조와 어긋나 편집기가 열지 못한다."""
    import zipfile

    from dsdxlsx.dsd_io import _RawNameInfo, read_dsd, write_dsd

    raw_name = "0. 날인 이미지_1.jpg".encode("cp949")
    src = tmp_path / "원본.dsd"
    with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("contents.xml", b"<DOCUMENT><P>x</P></DOCUMENT>")
        zf.writestr("meta.xml", META_EMPTY)
        # 편집기 방식 재현: cp949 이름 바이트를 플래그 없이 기록
        # (표준 ZipInfo는 비ASCII를 UTF-8로 강제하므로 보존 클래스로 생성)
        info = _RawNameInfo(raw_name.decode("cp437"))
        info._raw_name = raw_name
        info._raw_encoding_flag = 0
        zf.writestr(info, b"JPEGDATA")
    assert raw_name in src.read_bytes()  # 전제: 원본에 cp949 이름 바이트

    pkg = read_dsd(str(src))
    out = tmp_path / "재패킹.dsd"
    write_dsd(str(out), pkg, pkg.contents)  # 무수정 재패킹
    blob = out.read_bytes()
    assert raw_name in blob, "cp949 파일명 바이트가 보존돼야 함"
    assert raw_name.decode("cp949").encode("utf-8") not in blob  # UTF-8 변조 없음
    with zipfile.ZipFile(out) as zf:  # 데이터도 온전
        assert zf.read(raw_name.decode("cp437")) == b"JPEGDATA"
