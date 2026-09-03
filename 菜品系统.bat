@echo off
setlocal
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
%PY% "%~dp0sync_records.py" --no-pause
echo.
echo [2/3] 比对更新 HTML 菜品数据 ...
%PY% "%~dp0update_html.py" --no-pause
echo.
echo [3/3] 启动菜品系统服务（后台运行）...
start "菜品系统服务" %PY% "%~dp0菜品系统服务.py" --no-browser
timeout /t 2 /nobreak >nul
echo       打开最新页面 ...

rem 探测浏览器打开页面（避免 .html 关联异常导致打不开）
set "BROWSER="
if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" set "BROWSER=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if not defined BROWSER if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" set "BROWSER=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
if not defined BROWSER if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "BROWSER=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not defined BROWSER if exist "%ProgramFiles%\Mozilla Firefox\firefox.exe" set "BROWSER=%ProgramFiles%\Mozilla Firefox\firefox.exe"
if defined BROWSER (
    start "" "%BROWSER%" "%~dp0index.html"
) else (
    start "" "%~dp0index.html"
)
echo.
echo 如页面显示"未连接"，请稍等 2~3 秒后刷新页面。
echo.
pause
