@echo off
setlocal EnableExtensions
cd /d "%~dp0.."

set "VENV="
if not "%~1"=="" set "VENV=%~1"
if not defined VENV if exist ".pdf-duplex-env" set /p "VENV="<".pdf-duplex-env"
if not defined VENV set "VENV=.venv"
set "PY=%VENV%\Scripts\python.exe"

if not exist "%PY%" (
    echo [ERROR] Selected environment was not found:
    echo %VENV%
    echo Run run.bat first or pass the environment path as the first argument.
    pause
    exit /b 1
)

"%PY%" -m pip install pyinstaller
if errorlevel 1 (
    echo [ERROR] Failed to install PyInstaller.
    pause
    exit /b 1
)

"%PY%" -m PyInstaller --noconfirm --clean --windowed ^
    --name "PDFDuplexPrint" --icon "app_icon.ico" ^
    --collect-all tkinterdnd2 --add-data "app_icon.ico;." ^
    pdf_duplex_print.py
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo [OK] Build completed: dist\PDFDuplexPrint\
endlocal
