@echo off
REM Run the test suite. Sets up the venv if missing, then invokes pytest.
setlocal
cd /d "%~dp0"
if not exist .venv (
    echo Virtual environment not found. Run run.cmd first to set things up.
    exit /b 1
)
call .venv\Scripts\activate.bat
set PYTHONPATH=%cd%\src
python -m pytest tests -v
endlocal
