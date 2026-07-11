@echo off
rem 결산 수치를 입력한 .xlsx를 이 스크립트에 드래그&드롭
rem -> 상위 합계(유동자산·자산총계 등) 재계산 제안을 미리보기로 표시
rem    (실제 기입은 아래 안내대로 --apply 로 한 번 더 실행)
cd /d "%~dp0"
if "%~1"=="" (
    echo 사용법: 결산 수치를 입력한 .xlsx 파일을 드래그 앤 드롭하세요.
    pause
    exit /b 1
)
uv run --no-editable dsdxlsx refoot "%~1"
echo.
echo [실제 기입하려면] 아래 명령을 복사해 실행하세요:
echo    uv run --no-editable dsdxlsx refoot "%~1" --apply
pause
