@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   文档转 Markdown 工具 - 环境安装
echo ============================================
echo.

rem 检查 Python 3.13 虚拟环境
if exist ".venv313\Scripts\python.exe" (
    echo [OK] 已检测到 Python 3.13 虚拟环境
    set "PYTHON=.venv313\Scripts\python.exe"
    set "PIP=.venv313\Scripts\pip.exe"
    goto :install_deps
)

rem 检查 py launcher 是否有 3.13
py -3.13 --version >nul 2>&1
if not errorlevel 1 (
    echo [1/3] 创建 Python 3.13 虚拟环境...
    py -3.13 -m venv .venv313
    set "PYTHON=.venv313\Scripts\python.exe"
    set "PIP=.venv313\Scripts\pip.exe"
    goto :install_deps
)

rem 降级到系统 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.10-3.13
    echo 下载地址: https://www.python.org/downloads/
    echo 注意: MinerU 不支持 Python 3.14+
    pause
    exit /b 1
)

echo [提示] 未找到 Python 3.13，使用系统 Python（MinerU 可能不可用）
set "PYTHON=python"
set "PIP=pip"

:install_deps
echo.
echo [2/3] 正在安装 Python 依赖...
%PIP% install -r "%~dp0requirements.txt"

if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败，请检查网络连接
    pause
    exit /b 1
)

echo.
echo [3/3] 检查 .doc 转换环境...
%PYTHON% -c "import win32com.client; w=win32com.client.Dispatch('Word.Application'); print('[OK] Word/WPS:', w.Version); w.Quit()" 2>nul
if errorlevel 1 (
    where soffice >nul 2>&1
    if errorlevel 1 (
        if exist "C:\Program Files\LibreOffice\program\soffice.exe" (
            echo [OK] 检测到 LibreOffice
        ) else (
            echo [提示] 未检测到 Word/WPS/LibreOffice，.doc 格式转换将不可用
        )
    ) else (
        echo [OK] 检测到 LibreOffice
    )
)

echo.
echo ============================================
echo   安装完成！
echo.
echo   启动 GUI:  start_gui.bat
echo   命令行:    .venv313\Scripts\python.exe doc2md.py 文档.pdf
echo ============================================
pause
