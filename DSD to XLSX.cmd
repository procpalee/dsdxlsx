@echo off
rem Drag & drop a .dsd file onto this script -> creates .xlsx next to it
cd /d "%~dp0"
if "%~1"=="" (
    echo Usage: drag and drop a .dsd file onto this script.
    pause
    exit /b 1
)
uv run --no-editable dsdxlsx export "%~1"
pause
