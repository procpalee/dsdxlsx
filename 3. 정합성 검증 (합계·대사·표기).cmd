@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" goto :usage
uv run --no-editable dsdxlsx check "%~1"
echo.
pause
exit /b 0
:usage
echo.
echo  [check] Drag an edited .xlsx (or a .dsd) onto this script.
echo  Writes verification sheets and live formulas into the .xlsx.
echo  Close the file in Excel before running.
echo.
pause
