@echo off
setlocal EnableExtensions
cd /d "%~dp0.."

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] Run run.bat first to create the virtual environment.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -m pip install pyinstaller
if errorlevel 1 (
    echo [ERROR] Failed to install PyInstaller.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --windowed ^
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
