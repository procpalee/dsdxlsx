@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" goto :usage
uv run --no-editable dsdxlsx export "%~1"
echo.
pause
exit /b 0
:usage
echo.
echo  [export] Drag a .dsd file onto this script -^> creates .xlsx next to it.
echo.
pause
