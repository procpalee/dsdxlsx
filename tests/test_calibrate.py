"""calibrate — 새 회사 온보딩: 전수 통과 확인 + suppress 초안 스캐폴딩."""
from dsdxlsx.cli import main

from fixtures import factory


def _dsd(tmp_path, name="전기.dsd", **kw):
    p = tmp_path / name
    p.write_bytes(factory.dsd_bytes(factory.standard_contents(**kw)))
    return str(p)


def test_calibrate_clean_passes(tmp_path, capsys, monkeypatch):
    monkeypatch.chdir(tmp_path)
    rc = main(["calibrate", _dsd(tmp_path)])
    out = capsys.readouterr().out
    assert rc == 0 and "전수 통과" in out


def test_calibrate_scaffolds_suppress_draft(tmp_path, capsys, monkeypatch):
    monkeypatch.chdir(tmp_path)
    # 주석4 합계 오류 주입 → warn 1건 → 초안 출력 + --write로 주석 상태 추가
    rc = main(["calibrate", _dsd(tmp_path, cash_note_total="1,100"), "--write"])
    out = capsys.readouterr().out
    assert rc == 1
    assert "# [[suppress]]" in out and 'check = "C3"' in out

    ptoml = tmp_path / "profiles" / "합성테스트(주)" / "profile.toml"
    assert ptoml.is_file()
    text = ptoml.read_text(encoding="utf-8")
    assert '[meta]' in text and "# [[suppress]]" in text

    # 주석 상태(#)라 아직 비활성: 같은 문서 check는 여전히 warn — 로드도 정상
    from dsdxlsx.profile import load_profile
    prof = load_profile("합성테스트(주)", base_dir=tmp_path)
    assert not prof.suppress

    # 사람이 검토 후 주석 해제하면 활성화
    active = text.replace("# [[suppress]]", "[[suppress]]") \
                 .replace('# check =', 'check =') \
                 .replace('# where_contains =', 'where_contains =') \
                 .replace('# reason = "TODO: 검토 사유 기입"', 'reason = "검토 완료"')
    ptoml.write_text(active, encoding="utf-8")
    prof2 = load_profile("합성테스트(주)", base_dir=tmp_path)
    assert prof2.suppress and prof2.suppress[0].check == "C3"
