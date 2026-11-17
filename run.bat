@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "VENV=%~dp0.venv"
set "PY=%VENV%\Scripts\python.exe"
set "PYW=%VENV%\Scripts\pythonw.exe"

if not exist "%PY%" (
    echo [INFO] Creating the project virtual environment...
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3 -m venv "%VENV%"
    ) else (
        where python >nul 2>nul
        if errorlevel 1 (
            echo [ERROR] Python 3 was not found. Install Python 3.10+ and try again.
            pause
            exit /b 1
        )
        python -m venv "%VENV%"
    )
    if errorlevel 1 (
        echo [ERROR] Failed to create the virtual environment.
        pause
        exit /b 1
    )
)

if not exist "%PY%" (
    echo [ERROR] The virtual environment is incomplete: "%PY%"
    pause
    exit /b 1
)

echo [INFO] Checking project dependencies...
"%PY%" -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo [ERROR] Dependency installation failed. Check Python and network access.
    pause
    exit /b 1
)

if not exist "%PYW%" set "PYW=%PY%"
start "" "%PYW%" "%~dp0pdf_duplex_print.py"
endlocal
