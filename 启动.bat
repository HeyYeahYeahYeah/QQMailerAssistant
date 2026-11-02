@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

if exist "dist\QQ批量发送助手\QQ批量发送助手.exe" (
    start "" "dist\QQ批量发送助手\QQ批量发送助手.exe"
    goto :eof
)

if exist "dist\QQ批量发送助手.exe" (
    start "" "dist\QQ批量发送助手.exe"
    goto :eof
)

where pythonw >nul 2>nul
if not errorlevel 1 (
    start "" pythonw "app.py"
    goto :eof
)

where pyw >nul 2>nul
if not errorlevel 1 (
    start "" pyw -3 "app.py"
    goto :eof
)

where python >nul 2>nul
if not errorlevel 1 (
    echo 正在用 Python 启动（这个窗口不要关，关掉就退出程序）。
    python "app.py"
    goto :eof
)

echo.
echo 没有找到可用的程序。
echo 如果你拿到的是 exe 文件夹，请直接双击里面的「QQ批量发送助手.exe」。
echo 如果你想用源码运行，请先安装 Python 3.10 以上版本（安装时勾选 Add Python to PATH）。
echo.
pause
