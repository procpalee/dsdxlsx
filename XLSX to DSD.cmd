@echo off
rem Drag & drop an edited .xlsx (made by dsdxlsx export) onto this script
rem -> creates a new .dsd next to it (original DSD is embedded in the xlsx)
cd /d "%~dp0"
if "%~1"=="" (
    echo Usage: drag and drop an .xlsx file onto this script.
    pause
    exit /b 1
)
uv run --no-editable dsdxlsx import "%~1"
pause
