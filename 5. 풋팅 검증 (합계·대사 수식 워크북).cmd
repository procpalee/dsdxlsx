@echo off
chcp 65001 >nul
cd /d "%~dp0"
where uv >nul 2>nul || goto :nouv
if "%~1"=="" goto :usage
uv run --no-editable dsdxlsx foot %*
echo.
pause
exit /b 0
:usage
echo.
echo  [foot] Drag one or more .dsd files onto this script.
echo  Writes a footing workbook (.xlsx) next to each file: summary sheet, statements, notes,
echo  with live footing and cross-reference formulas.
echo.
pause
exit /b 0

:nouv
echo.
echo  uv is not installed. Install it once in PowerShell, then open a new window:
echo    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
echo  See README: https://github.com/procpalee/dsdxlsx
echo.
pause
exit /b 1
