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
echo [1/3] 同步本地选菜记录到 Excel ...
%PY% sync_records.py --no-pause
echo.
echo [2/3] 比对更新 HTML 菜品数据（Excel 最新数据同步到页面）...
%PY% update_html.py --no-pause
echo.
echo [3/3] 启动菜品系统服务（后台运行，用于新增菜品/选菜记录自动写 Excel）...
start "菜品系统服务" %PY% 菜品系统服务.py --no-browser
timeout /t 2 /nobreak >nul
echo        打开最新页面 ...
start "" "菜品选择.html"
echo.
echo 如页面显示"未连接"，请稍等 2~3 秒后刷新页面。
echo.
pause
