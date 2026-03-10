@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   文档转 Markdown 工具
echo ============================================
echo.

rem 优先使用 Python 3.13 虚拟环境
if exist ".venv313\Scripts\python.exe" (
    set "PYTHON=.venv313\Scripts\python.exe"
) else (
    set "PYTHON=python"
)

if "%~1"=="" (
    echo 用法: 将文件拖拽到此 .bat 文件上即可转换
    echo 或者: convert.bat 文件路径.docx
    echo.
    echo 支持格式: .doc .docx .pdf
    echo.
    set /p INPUT_FILE="请输入文件路径（或拖拽文件到此窗口）: "
) else (
    set "INPUT_FILE=%~1"
)

if "%INPUT_FILE%"=="" (
    echo [错误] 未指定文件
    pause
    exit /b 1
)

echo 正在转换: %INPUT_FILE%
echo.

%PYTHON% "%~dp0doc2md.py" "%INPUT_FILE%"

echo.
echo ============================================
pause
