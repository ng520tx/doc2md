@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem 优先使用 Python 3.13 虚拟环境（MinerU 需要 ≤3.13）
if exist ".venv313\Scripts\python.exe" (
    set "PYTHON=.venv313\Scripts\python.exe"
) else (
    set "PYTHON=python"
)

%PYTHON% gui.py
if errorlevel 1 (
    echo.
    echo [Error] Failed to start GUI. Installing dependencies...
    %PYTHON% -m pip install -r requirements.txt
    echo.
    echo Retrying...
    %PYTHON% gui.py
)
pause
