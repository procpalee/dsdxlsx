@echo off
rem 전기(직전 보고서) .dsd 파일을 이 스크립트에 드래그&드롭
rem -> 당기 기초 DSD(_차기기초.dsd) + 표준 정산표(_차기기초.xlsx) 생성
rem    기수+1, 당기값->전기 열 이동, 당기 열 비움, '롤포워드' 리포트 시트 포함
cd /d "%~dp0"
if "%~1"=="" (
    echo 사용법: 전기 .dsd 파일을 이 스크립트에 드래그 앤 드롭하세요.
    pause
    exit /b 1
)
uv run --no-editable dsdxlsx rollforward "%~1"
pause
