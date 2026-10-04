@echo off
rem Starts the CCI Alumni Tracer Study site and opens it in the browser.
rem Double-click this file; close the window (or press Ctrl+C) to stop the site.
title CCI Alumni Tracer Study
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found. Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
    pause
    exit /b 1
)

rem Install the needed packages the first time
python -c "import flask, openpyxl, PIL" 2>nul
if errorlevel 1 (
    echo Installing required packages...
    python -m pip install flask openpyxl pillow
    if errorlevel 1 (
        echo Could not install the required packages.
        pause
        exit /b 1
    )
)

rem Open the browser a few seconds from now, once the site is up
start "" /min cmd /c "timeout /t 3 /nobreak >nul & start http://127.0.0.1:5000"

echo Site address: http://127.0.0.1:5000
python db.py
pause
