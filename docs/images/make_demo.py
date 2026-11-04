"""README 화면 예시용 합성 DSD 두 판을 만든다(실제 회사 자료 아님).

    uv run --no-editable python docs/images/make_demo.py <출력 폴더>

- 샘플_초안.dsd: 주석 4 현금 표의 '보통예금'을 1,100 으로 잘못 적어 합계(1,200)와 안 맞는 판
- 샘플_수정본.dsd: 그 칸을 1,000 으로 고치고 재고자산 구성과 주석 문장 하나를 바꾼 판

화면 예시는 이 두 파일로 찍었다.
- check: dsdxlsx export 샘플_초안.dsd → dsdxlsx check 샘플_초안.xlsx → 엑셀에서 '4' 시트
- diff:  dsdxlsx diff 샘플_초안.dsd 샘플_수정본.dsd -o diff.html → 브라우저
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tests" / "fixtures"))
import factory as f  # noqa: E402

HEAD = [["과  목", "주  석", "제 2(당) 기말", "제 1(전) 기말"]]
NOTE_HEAD = [["구  분", "당기말", "전기말"]]


def contents(*, deposit: str, goods: str, materials: str, pledge_text: str) -> bytes:
    bs = f.table(HEAD, [
        ["Ⅰ.유동자산", "", "8,500", "7,000"],
        ["   현금및현금성자산", "||ALIGN=CENTER||4", "1,200", "900"],
        ["   매출채권", "||ALIGN=CENTER||5", "5,000", "4,000"],
        ["   재고자산", "||ALIGN=CENTER||6", "2,300", "2,100"],
        ["Ⅱ.비유동자산", "", "12,000", "12,500"],
        ["   유형자산", "||ALIGN=CENTER||7", "12,000", "12,500"],
        ["자산총계", "", "20,500", "19,500"],
    ], unit="백만원")
    note4 = f.table(NOTE_HEAD, [
        ["보유현금", "200", "150"],
        ["보통예금", deposit, "750"],
        ["합  계", "1,200", "900"],
    ], unit="백만원")
    note5 = f.table(NOTE_HEAD, [
        ["매출채권", "5,500", "4,300"],
        ["대손충당금", "(500)", "(300)"],
        ["차감후 장부금액", "5,000", "4,000"],
    ], unit="백만원")
    note6 = f.table(NOTE_HEAD, [
        ["제품", goods, "1,400"],
        ["원재료", materials, "700"],
        ["합  계", "2,300", "2,100"],
    ], unit="백만원")
    note7 = f.table(NOTE_HEAD, [
        ["토지", "5,000", "5,000"],
        ["건물", "7,000", "7,500"],
        ["합  계", "12,000", "12,500"],
    ], unit="백만원")
    return f.document(
        f.p("샘플전자 주식회사 (합성 예시 문서)"),
        f.title("재무제표", 2),
        f.statement_title_table("재 무 상 태 표"),
        bs,
        f.title("주석", 3),
        f.p("1. 일반사항\n샘플전자 주식회사는 README 화면 예시를 위해 만든 가상의 회사입니다."),
        f.p("4. 현금및현금성자산\n당기말과 전기말 현재 현금및현금성자산의 내역은 다음과 같습니다."),
        note4,
        f.p("5. 매출채권\n당기말과 전기말 현재 매출채권의 내역은 다음과 같습니다."),
        note5,
        f.p("6. 재고자산\n당기말과 전기말 현재 재고자산의 내역은 다음과 같습니다."),
        note6,
        f.p("7. 유형자산\n" + pledge_text),
        note7,
    )


def main(out_dir: str) -> None:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    meta = f.meta(regname="샘플전자(주)")
    f.write_dsd(out / "샘플_초안.dsd", contents(
        deposit="1,100", goods="1,500", materials="800",
        pledge_text="당기말과 전기말 현재 유형자산의 내역은 다음과 같습니다."), meta)
    f.write_dsd(out / "샘플_수정본.dsd", contents(
        deposit="1,000", goods="1,480", materials="820",
        pledge_text="당기말과 전기말 현재 유형자산의 내역은 다음과 같으며, "
                    "토지와 건물은 차입금 담보로 제공되어 있습니다."), meta)
    print(f"만듦: {out / '샘플_초안.dsd'}, {out / '샘플_수정본.dsd'}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
