@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem 检测 Python 环境
set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY ( where py >nul 2>nul && set "PY=py" )
if not defined PY (
    echo ============================================
    echo   菜品选择系统 - 一键启动
    echo ============================================
    echo.
    echo 未检测到 Python 环境。
    echo 请先安装 Python（https://www.python.org/downloads/）后重试。
    echo.
    pause
    exit /b 1
)

echo ============================================
echo   菜品选择系统 - 一键启动
echo ============================================
echo.
echo [1/2] 同步本地选菜记录到 Excel ...
%PY% sync_records.py --no-pause
echo.
echo [2/2] 启动菜品系统服务（将自动打开浏览器）...
echo        关闭本窗口即停止服务
echo.
%PY% 菜品系统服务.py
echo.
pause
