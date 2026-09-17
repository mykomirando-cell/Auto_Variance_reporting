@echo off
REM Run the Streamlit app with the right env. Double-click to launch.
setlocal
cd /d "%~dp0"
if not exist .venv (
    echo Creating virtual environment...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    echo Installing dependencies...
    pip install -e ".[dev]" >nul
) else (
    call .venv\Scripts\activate.bat
)
set PYTHONPATH=%cd%\src
echo Launching Inventory Reconciliation App...
streamlit run app.py
endlocal
