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
    echo "%VENV%"
    echo Run run.bat first or pass the environment path as the first argument.
    pause
    exit /b 1
)

"%PY%" -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    "%PY%" -m pip --version >nul 2>&1
    if errorlevel 1 "%PY%" -m ensurepip --upgrade
    "%PY%" -m pip install -r "%ROOT%requirements-build.txt"
    if errorlevel 1 (
        echo [ERROR] Failed to install the pinned PyInstaller.
        pause
        exit /b 1
    )
)

"%PY%" -m PyInstaller --noconfirm --clean --windowed ^
    --name "PDFDuplexPrint" --icon "app_icon.ico" ^
    --version-file "tools\version_info.txt" ^
    --collect-all tkinterdnd2 --add-data "app_icon.ico;." ^
    pdf_duplex_print.py
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

set "PDF_DUPLEX_SELF_CHECK_FILE=%ROOT%build\frozen-self-check.txt"
if exist "%PDF_DUPLEX_SELF_CHECK_FILE%" del /q "%PDF_DUPLEX_SELF_CHECK_FILE%"
start "" /wait "%ROOT%dist\PDFDuplexPrint\PDFDuplexPrint.exe" --self-check
if not exist "%PDF_DUPLEX_SELF_CHECK_FILE%" (
    echo [ERROR] Frozen self-check did not produce a result.
    pause
    exit /b 1
)
findstr /c:"ok=True" "%PDF_DUPLEX_SELF_CHECK_FILE%" >nul
if errorlevel 1 (
    echo [ERROR] Frozen self-check failed.
    pause
    exit /b 1
)
set "PDF_DUPLEX_SELF_CHECK_FILE="
"%PY%" tools\write_checksums.py "dist\PDFDuplexPrint"
if errorlevel 1 (
    echo [ERROR] Failed to write SHA-256 checksums.
    pause
    exit /b 1
)

echo [OK] Build and frozen self-check completed: dist\PDFDuplexPrint\
endlocal
