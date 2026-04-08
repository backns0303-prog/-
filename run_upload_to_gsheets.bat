@echo off
setlocal

cd /d "%~dp0"

set "CREDENTIALS=streamlit-sheets-upload-34b193fd0a59.json"
set "SPREADSHEET_ID=1Jy1DFHveJYFEw2lVg_pUGeE7HCcFmYaeUb6FwSrZGJM"
set "PATTERN=*.xls"

echo [1/2] Starting upload...
python upload_xls_to_gsheets.py --credentials "%CREDENTIALS%" --spreadsheet-id "%SPREADSHEET_ID%" --pattern "%PATTERN%" --cleanup-daily --cleanup-apply

if errorlevel 1 (
    echo.
    echo Upload failed.
    pause
    exit /b 1
)

echo.
echo Upload completed successfully.
pause
