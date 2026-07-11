@echo off
rem 편집한 .xlsx(또는 .dsd)를 이 스크립트에 드래그&드롭
rem -> 합계검산·재무제표<->주석 대사·대사맵·표기 검증
rem    xlsx면 그 파일 안에 '검증요약'·'대사맵' 시트와 살아있는 검산 수식을 기록
rem    (엑셀에서 파일을 닫은 상태로 실행하세요)
cd /d "%~dp0"
if "%~1"=="" (
    echo 사용법: 편집한 .xlsx 또는 .dsd 파일을 이 스크립트에 드래그 앤 드롭하세요.
    pause
    exit /b 1
)
uv run --no-editable dsdxlsx check "%~1"
pause
