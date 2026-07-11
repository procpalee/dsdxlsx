@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" goto :usage
uv run --no-editable dsdxlsx refoot "%~1"
echo.
echo  --------------------------------------------------------------
echo   To actually write the values, run again with --apply :
echo     uv run --no-editable dsdxlsx refoot "%~1" --apply
echo  --------------------------------------------------------------
echo.
pause
exit /b 0
:usage
echo.
echo  [refoot] Drag an edited .xlsx (with current-period figures) here.
echo  Previews recalculated subtotals/totals. Add --apply to write them.
echo.
pause
