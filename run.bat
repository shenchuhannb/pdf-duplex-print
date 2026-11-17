@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "ROOT=%~dp0"
set "ENV_CONFIG=%ROOT%.pdf-duplex-env"
set "VENV="

if /I "%~1"=="--reset-env" (
    if exist "%ENV_CONFIG%" del /q "%ENV_CONFIG%"
    shift
)
if not "%~1"=="" set "VENV=%~1"
if not defined VENV if exist "%ENV_CONFIG%" set /p "VENV="<"%ENV_CONFIG%"

if not defined VENV (
    echo Choose the virtual environment directory.
    echo Press Enter to use the project default:
    echo %ROOT%.venv
    set /p "VENV=Environment path: "
    if not defined VENV set "VENV=%ROOT%.venv"
    >"%ENV_CONFIG%" echo %VENV%
)

set "PY=%VENV%\Scripts\python.exe"
set "PYW=%VENV%\Scripts\pythonw.exe"

if not exist "%PY%" (
    echo [INFO] Creating the virtual environment:
    echo %VENV%
    if defined PDF_DUPLEX_PYTHON (
        if not exist "%PDF_DUPLEX_PYTHON%" (
            echo [ERROR] PDF_DUPLEX_PYTHON does not exist:
            echo %PDF_DUPLEX_PYTHON%
            pause
            exit /b 1
        )
        "%PDF_DUPLEX_PYTHON%" -m venv "%VENV%"
    ) else (
        where py >nul 2>nul
        if not errorlevel 1 (
            py -3 -m venv "%VENV%"
        ) else (
            where python >nul 2>nul
            if errorlevel 1 (
                echo [ERROR] Python 3 was not found. Install Python 3.10+ first.
                pause
                exit /b 1
            )
            python -m venv "%VENV%"
        )
    )
    if errorlevel 1 (
        echo [ERROR] Failed to create the virtual environment.
        pause
        exit /b 1
    )
)

if not exist "%PY%" (
    echo [ERROR] The selected directory is not a valid Python venv:
    echo %VENV%
    echo Use run.bat --reset-env to choose another directory.
    pause
    exit /b 1
)

echo [INFO] Checking project dependencies...
"%PY%" -m pip install -r "%ROOT%requirements.txt"
if errorlevel 1 (
    echo [ERROR] Dependency installation failed. Check Python and network access.
    pause
    exit /b 1
)

if not exist "%PYW%" set "PYW=%PY%"
start "" "%PYW%" "%ROOT%pdf_duplex_print.py"
endlocal
