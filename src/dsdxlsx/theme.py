"""excel-theme 스킬의 default 테마를 로드해 export 스타일의 단일 소스로 쓴다.

스킬(~/.claude/skills/excel-theme/excel_theme.py)의 PALETTES["default"]를 읽고,
없는 환경(다른 PC 등)에서는 동일 값의 내장 폴백을 쓴다. 스킬 모듈 import 시
openpyxl 저장 테마도 최신 Office 테마로 교체된다(스킬의 설계).
"""
from __future__ import annotations

import importlib.util
import os

# excel_theme.py PALETTES["default"]와 동일 값 (스킬 부재 시 폴백)
_FALLBACK = {
    "font": "맑은 고딕",
    "size_body": 11,
    "size_header": 11,
    "size_title": 12,
    "colors": {
        "primary": "404040",
        "primary_dark": "262626",
        "border_in": "BFBFBF",
        "border_out": "000000",
        "header_fill": "F2F2F2",
        "header_text": "000000",
        "header2_fill": "DDE3E8",  # 서브 헤더 (연청회색)
        "header2_text": "000000",
        "linked": "F2F2F2",   # 참조/자동 셀 — 연회색
        "note": "FFFBEF",     # 입력 셀 — 크림
        "accent": "C00000",
        "title_color": "000000",
        "hlink": "0563C1",
    },
    "modernize": None,
}

_SKILL_PATH = os.path.join(
    os.path.expanduser("~"), ".claude", "skills", "excel-theme", "excel_theme.py"
)


def load_theme() -> dict:
    theme = {**_FALLBACK, "colors": dict(_FALLBACK["colors"])}
    try:
        spec = importlib.util.spec_from_file_location("_excel_theme_skill", _SKILL_PATH)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        pal = mod.PALETTES["default"]
        c = pal.get("colors", {})
        theme["font"] = pal.get("font_excel") or pal.get("font_name") or theme["font"]
        theme["size_body"] = pal.get("excel_size_body", theme["size_body"])
        theme["size_header"] = pal.get("excel_size_header", theme["size_header"])
        theme["size_title"] = pal.get("excel_title_size", theme["size_title"])
        colors = theme["colors"]
        colors["primary"] = c.get("primary", colors["primary"])
        colors["primary_dark"] = c.get("primary_dark", colors["primary_dark"])
        colors["border_in"] = c.get("border_in", colors["border_in"])
        colors["border_out"] = c.get("border_out", colors["border_out"])
        colors["header_fill"] = c.get("excel_header_fill", colors["header_fill"])
        colors["header_text"] = c.get("excel_header_text", colors["header_text"])
        colors["header2_fill"] = c.get("excel_header2_fill", colors["header2_fill"])
        colors["header2_text"] = c.get("excel_header2_text", colors["header2_text"])
        colors["linked"] = c.get("linked", colors["linked"])
        colors["note"] = c.get("note", colors["note"])
        colors["accent"] = c.get("accent", colors["accent"])
        colors["title_color"] = c.get("title_color", colors["title_color"])
        theme["modernize"] = getattr(mod, "modernize_theme", None)
    except Exception:
        pass  # 스킬 미설치 환경 — 폴백 사용
    return theme
