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
