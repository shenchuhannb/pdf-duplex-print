@echo off
rem 用项目 venv 的 PyInstaller 打包为单目录 exe（产物在 dist\PDF双面打印\）
rem 依赖：.venv\Scripts\python.exe -m pip install pyinstaller
setlocal
cd /d "%~dp0.."
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --windowed ^
    --name "PDF双面打印" --icon "app_icon.ico" ^
    --collect-all tkinterdnd2 --add-data "app_icon.ico;." ^
    pdf_duplex_print.py
endlocal
