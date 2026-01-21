@echo off
cd /d "%~dp0"
echo Checking dependencies...
pip install -r requirements.txt
echo Launching Sizing Tool...
streamlit run sizing_app.py
pause