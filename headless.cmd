@echo off
REM Run the headless CLI: headless.cmd <workbook.xlsx> [--conversion conv.xlsx] --out result.xlsx
setlocal
cd /d "%~dp0"
if not exist .venv (
    echo Virtual environment not found. Run run.cmd first to set things up.
    exit /b 1
)
call .venv\Scripts\activate.bat
set PYTHONPATH=%cd%\src
python -m recon.cli %*
endlocal
