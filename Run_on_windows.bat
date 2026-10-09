@echo off
REM ==========================================================================
REM  Host Sizer - Windows launcher. Double-click this file.
REM  First run: finds Python (or offers to install it), creates a private
REM  .venv and installs the libraries. Every run: starts the app and opens
REM  your browser.
REM ==========================================================================
setlocal EnableExtensions EnableDelayedExpansion
title Host Sizer
cd /d "%~dp0"
echo ===========================================
echo    Host Sizer  -  Powered by AHEAD
echo ===========================================
echo.

REM --------------------------------------------------------------------------
REM 1. Python 3.9+  (ignores the Microsoft Store "python" placeholder)
REM --------------------------------------------------------------------------
echo [1/4] Checking for Python...
set "PY="
call :find_python
if not defined PY call :install_python
if not defined PY goto :no_python
<nul set /p "=      Using "
!PY! --version

REM --------------------------------------------------------------------------
REM 2. Private environment for this tool (rebuilt automatically if it breaks)
REM --------------------------------------------------------------------------
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys" >nul 2>&1
    if errorlevel 1 (
        echo [2/4] The tool's Python environment is damaged - rebuilding it...
        rmdir /s /q ".venv"
    )
)
if not exist ".venv\Scripts\python.exe" (
    echo [2/4] Creating the tool's Python environment ^(first run only^)...
    !PY! -m venv .venv
    if errorlevel 1 goto :venv_failed
) else (
    echo [2/4] Python environment ready.
)

REM --------------------------------------------------------------------------
REM 3. Libraries - installed on first run and whenever requirements.txt changes
REM --------------------------------------------------------------------------
fc /b requirements.txt ".venv\requirements.installed" >nul 2>&1
if errorlevel 1 (
    echo [3/4] Installing libraries ^(Streamlit, Pandas, OpenPyXL^). First run takes a few minutes...
    ".venv\Scripts\python.exe" -m pip install --quiet --disable-pip-version-check --upgrade pip >nul 2>&1
    ".venv\Scripts\python.exe" -m pip install --quiet --disable-pip-version-check -r requirements.txt
    if errorlevel 1 goto :pip_failed
    copy /y requirements.txt ".venv\requirements.installed" >nul
) else (
    echo [3/4] Libraries up to date.
)

REM --------------------------------------------------------------------------
REM 4. Launch
REM --------------------------------------------------------------------------
echo [4/4] Starting Host Sizer - your browser will open shortly.
echo       Keep this window open while you use the tool. Close it to stop the tool.
echo.
".venv\Scripts\python.exe" sizing_app.py --setup >nul 2>&1
".venv\Scripts\python.exe" -m streamlit run sizing_app.py
pause
exit /b 0


REM ==========================================================================
REM  Helpers
REM ==========================================================================
:find_python
REM Commands on PATH. The Store placeholder fails this check, so it is skipped.
for %%C in ("py -3" "python" "python3") do (
    if not defined PY (
        %%~C -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1 && set "PY=%%~C"
    )
)
if defined PY exit /b 0
REM Standard install folders (covers a Python installed moments ago, before PATH refreshes).
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*" "%ProgramFiles%\Python3*") do (
    if not defined PY (
        if exist "%%~D\python.exe" (
            "%%~D\python.exe" -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1 && set PY="%%~D\python.exe"
        )
    )
)
exit /b 0

:install_python
echo.
echo       Python is not installed on this PC.
where winget >nul 2>&1
if errorlevel 1 goto :open_python_site
choice /c YN /m "      Install Python 3.12 now using Microsoft's winget installer"
if errorlevel 2 goto :open_python_site
echo       Installing Python 3.12 for your account. This can take a few minutes...
winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
call :find_python
if not defined PY echo       winget finished but Python was not found. Opening python.org instead.
if not defined PY goto :open_python_site
exit /b 0

:open_python_site
echo       Opening python.org. Install Python 3.12 or newer and tick "Add python.exe to PATH",
echo       then double-click this launcher again.
start "" "https://www.python.org/downloads/windows/"
exit /b 0

:no_python
echo.
echo Python 3.9 or newer is required and was not found.
echo After installing Python, double-click this launcher again.
pause
exit /b 1

:venv_failed
echo.
echo Could not create the Python environment in:
echo   %CD%
echo Move the tool folder somewhere you can write to ^(e.g. Documents^) and try again.
pause
exit /b 1

:pip_failed
echo.
echo The libraries could not be downloaded from the Python Package Index ^(pypi.org^).
echo   - Check that you are connected to the internet ^(and VPN, if required^).
echo   - If your network uses a proxy, it may be blocking pypi.org.
echo     Ask your eTech team to allow pypi.org and files.pythonhosted.org.
echo Then double-click this launcher again.
pause
exit /b 1
