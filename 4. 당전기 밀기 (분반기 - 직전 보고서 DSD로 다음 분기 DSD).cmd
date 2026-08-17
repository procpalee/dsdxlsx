@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" goto :usage
set "PRIOR_SAME="
set "PRIOR_ANNUAL="
echo.
echo  [roll-interim] base = "%~1"
echo  Optional: prior-year same-quarter DSD (source of prior-period flow values).
set /p PRIOR_SAME= "   prior-year same-quarter .dsd path (Enter to skip): "
echo  Optional: prior-year audit report DSD (only for 3Q -> next 1Q; source of prior year-end balances).
set /p PRIOR_ANNUAL= "   prior audit report .dsd path (Enter to skip): "
set "ARGS="
if not "%PRIOR_SAME%"=="" set ARGS=%ARGS% --prior-same "%PRIOR_SAME:"=%"
if not "%PRIOR_ANNUAL%"=="" set ARGS=%ARGS% --prior-annual "%PRIOR_ANNUAL:"=%"
uv run --no-editable dsdxlsx roll-interim "%~1" %ARGS%
echo.
pause
exit /b 0
:usage
echo.
echo  [roll-interim] Drag the most recent interim-report .dsd (1Q for 2Q work, 2Q for 3Q, prior 3Q for 1Q) onto this script.
echo  Creates: ^<name^>_dang-jeon-gi-mil-gi.dsd + .xlsx (next-quarter DSD: current blank, prior filled from prior-year files/DART)
echo  OPENDART_API_KEY in .env enables DART comparison (listed companies).
echo.
pause
