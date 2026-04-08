@echo off
setlocal

cd /d "%~dp0"

set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
set "STREAMLIT_SERVER_HEADLESS=true"
set "DASHBOARD_URL=http://localhost:8501"

start "" pythonw -m streamlit run dashboard_app.py

timeout /t 4 /nobreak >nul
start "" "%DASHBOARD_URL%"
exit /b
