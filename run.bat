@echo off
rem 启动本项目（使用项目内独立 venv 的 pythonw，不依赖系统 Python 环境）
setlocal
cd /d "%~dp0"
if not exist "%~dp0.venv\Scripts\pythonw.exe" (
    echo [错误] 未找到项目虚拟环境：%~dp0.venv\Scripts\pythonw.exe
    echo 请先执行：python.exe -m venv "%~dp0.venv"
    echo 然后执行： "%~dp0.venv\Scripts\python.exe" -m pip install -r "%~dp0requirements.txt"
    pause
    exit /b 1
)
start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0pdf_duplex_print.py"
endlocal
