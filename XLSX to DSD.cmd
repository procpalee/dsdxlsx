@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" goto :usage
uv run --no-editable dsdxlsx import "%~1"
echo.
pause
exit /b 0
:usage
echo.
echo  [import] Drag an edited .xlsx onto this script -^> creates a new .dsd.
echo  (The original DSD is embedded in the xlsx, so only the xlsx is needed.)
echo.
pause
