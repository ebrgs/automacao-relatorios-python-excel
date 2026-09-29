@echo off
cd /d "%~dp0"
python -c "import openpyxl, tkinter" >nul 2>&1
if errorlevel 1 (
    echo Python ou dependencias ausentes.
    echo Instale Python 3.10+ e execute: pip install -r requirements.txt
    pause
    exit /b 1
)
start "" pythonw "%~dp0interface.py"
