"""서식 상수 — excel-theme 스킬 `default` 테마 (모노톤 라이트 그레이·삼선표).

색·폰트는 ~/.claude/skills/excel-theme PALETTES["default"] 레지스트리 값
(맑은 고딕 11, primary 404040, band F2F2F2, border_in BFBFBF, border_out 000000).
자체 룩.
검증 수식의 숫자서식([RED]오류/[BLACK]"TRUE")은 기능 요소라 유지한다.
"""
from __future__ import annotations

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ── 테마 토큰 (excel-theme default) ───────────────────────────────────
_FONT_NAME = "맑은 고딕"
_PRIMARY = "404040"        # 진회색 — 헤더 채움·강조
_PRIMARY_DARK = "262626"   # 소계·합계·검증 라벨 글자
_BAND = "F2F2F2"           # 헤더·라벨 배경 (라이트 그레이)
_BORDER_IN = "BFBFBF"      # 내부 얇은 선
_BORDER_OUT = "000000"     # 표 상/하 medium 선

FONT = Font(name=_FONT_NAME, size=11)
FONT_BOLD = Font(name=_FONT_NAME, size=11, bold=True, color=_PRIMARY_DARK)
FONT_HEAD = Font(name=_FONT_NAME, size=11, bold=True)          # 표 머리행
FONT_LINK = Font(name=_FONT_NAME, size=11, color="0563C1", underline="single")
FONT_HEAD_INV = Font(name=_FONT_NAME, size=11, bold=True, color="FFFFFF")
FONT_SMALL_DIM = Font(name=_FONT_NAME, size=9, color="AAAAAA")  # 총괄표 예비 열

_THIN = Side(style="thin", color=_BORDER_IN)
_MED = Side(style="medium", color=_BORDER_OUT)
_DOTTED = Side(style="dotted", color=_BORDER_IN)

# 삼선표(frame): 좌/우 외곽선 없음, 상/하 medium, 내부 thin — 위치별 4변 조합
def table_border(top: bool, bottom: bool, left: bool, right: bool) -> Border:
    return Border(
        left=None if left else _THIN, right=None if right else _THIN,
        top=_MED if top else _THIN, bottom=_MED if bottom else _THIN)


BORDER_CHECK = Border(left=_DOTTED, right=_DOTTED, top=_DOTTED, bottom=_DOTTED)

_BAND2 = "DDE3E8"          # 세컨더리 헤더(연청회색) — excel_header2_fill

FILL_HEAD = PatternFill(fill_type="solid", fgColor=_BAND)      # 표 머리행
FILL_LABEL = PatternFill(fill_type="solid", fgColor=_BAND)     # 합계검증(풋팅) 라벨
FILL_XREF = PatternFill(fill_type="solid", fgColor=_BAND2)     # 대사검증(크로스) 라벨
_BAND_AI = "E8E1F0"        # AI 제안 대사 라벨(연보라) — 결정론 대사와 시각 구분
FILL_AI = PatternFill(fill_type="solid", fgColor=_BAND_AI)
_BAND_DART = "E2F0D9"      # DART 대사 라벨(연녹) — 결정론·AI 대사와 시각 구분
FILL_DART = PatternFill(fill_type="solid", fgColor=_BAND_DART)
FILL_DARK = PatternFill(fill_type="solid", fgColor=_PRIMARY)   # 총괄표 헤더
# 오류 강조(엑셀 표준 연빨강). start/end 둘 다 지정 — 조건부서식(dxf)은 bgColor를 쓴다
FILL_BAD = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FONT_BAD = Font(name=_FONT_NAME, size=11, bold=True, color="9C0006")

# 금액: 양수 콤마, 음수 빨간 괄호, 0은 '-'
NUMFMT_INT = "#,##0_);[Red](#,##0);-"
# 검증 수식 셀: 차이금액(양·음수) 빨강, 0(및 TRUE 불리언)은 검정 — 조건부서식 불필요
NUMFMT_CHECK = '[RED]#,##0;[RED]-#,##0;[BLACK]"TRUE"'

COL_WIDTH = 25.0
GUTTER_WIDTH = 2.5         # A열(빈 거터) 고정 너비
CHECK_NUM_WIDTH = 14.0     # 검증 수식(숫자) 열 너비
FREEZE = "B6"              # 1~5행(헤더 바·링크)과 A열 거터 고정

ALIGN_MAP = {"LEFT": "left", "RIGHT": "right", "CENTER": "center", "JUSTIFY": "justify"}


def number_format(text: str, decimal_places: int, suffix: str) -> str:
    """원본 표기 기반 숫자서식 (소수 자릿수·접미 단위 유지)."""
    lit = f'"{suffix}"' if suffix else ""
    if decimal_places:
        frac = "0" * decimal_places
        return f"#,##0.{frac}{lit}_);[Red](#,##0.{frac}{lit});-"
    if suffix:
        return f"#,##0{lit}_);[Red](#,##0{lit});-"
    return NUMFMT_INT


def para_alignment(text: str) -> Alignment:
    # 자동 줄바꿈 없음 — 행 높이가 세로로 늘어나지 않게 (사용자 지정)
    return Alignment(vertical="top", wrap_text=False)


# ── Office 테마 XML (excel-theme 스킬 MODERN_OFFICE_THEME 이식) ──────
# openpyxl 기본 저장은 구형 'Office 2007-2010' 테마를 심는다. 최신 Office 테마
# (맑은 고딕 major/minor)를 강제해 scheme=minor 기본 폰트가 맑은 고딕으로 렌더된다.
# write_workbook에서 wb.loaded_theme = OFFICE_THEME 로 주입.
OFFICE_THEME = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="맑은 고딕"/><a:ea typeface="맑은 고딕"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="맑은 고딕"/><a:ea typeface="맑은 고딕"/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
