@echo off
setlocal
cd /d "%~dp0"

py -3 --version >nul 2>&1
if errorlevel 1 (
    echo Python 3 was not found. Install Python 3.11 or newer and ensure the "py" launcher is available.
    pause
    exit /b 1
)

py -3 -c "import PySide6" >nul 2>&1
if errorlevel 1 (
    echo PySide6 is not installed for Python 3.
    echo Install the app requirements with:
    echo py -3 -m pip install --user -r requirements.txt
    pause
    exit /b 1
)

where pyw >nul 2>&1
if errorlevel 1 (
    echo The Python windowed launcher "pyw" was not found. Repair or reinstall Python with the launcher.
    pause
    exit /b 1
)

start "" pyw -3 "%~dp0app.py"
