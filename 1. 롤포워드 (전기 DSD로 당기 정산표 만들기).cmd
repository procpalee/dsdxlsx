@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" goto :usage
uv run --no-editable dsdxlsx rollforward "%~1"
echo.
pause
exit /b 0
:usage
echo.
echo  [rollforward] Drag a prior-period .dsd file onto this script.
echo  Creates: ^<name^>_cha-gi-gi-cho.dsd + .xlsx (next-period worksheet)
echo.
pause
