@echo off
REM Virtualization Sizing Calculator - Windows launcher
REM First run: installs dependencies into a private .venv folder. Every run: launches the app.
cd /d "%~dp0"

set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY (where python >nul 2>&1 && set "PY=python")
if not defined PY (
    echo Python 3 not found. Install it from https://www.python.org/downloads/
    echo During install, tick "Add python.exe to PATH". Then re-run this file.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo [setup] Creating Python environment ^(first run only^)...
    %PY% -m venv .venv || (echo Failed to create virtual environment. & pause & exit /b 1)
)

fc /b requirements.txt ".venv\requirements.installed" >nul 2>&1
if errorlevel 1 (
    echo [setup] Installing libraries ^(Streamlit, Pandas, OpenPyXL^)...
    ".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
    ".venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt || (echo Library install failed. Check network/proxy. & pause & exit /b 1)
    copy /y requirements.txt ".venv\requirements.installed" >nul
)

echo Launching... your browser will open shortly. Close this window to stop the app.
REM Install the AHEAD theme (.streamlit\config.toml)
".venv\Scripts\python.exe" sizing_app.py --setup >nul 2>&1
".venv\Scripts\python.exe" -m streamlit run sizing_app.py
pause
