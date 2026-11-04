@echo off
chcp 65001 >nul
cd /d "%~dp0"
where uv >nul 2>nul || goto :nouv
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
exit /b 0

:nouv
echo.
echo  uv is not installed. Install it once in PowerShell, then open a new window:
echo    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
echo  See README: https://github.com/procpalee/dsdxlsx
echo.
pause
exit /b 1
