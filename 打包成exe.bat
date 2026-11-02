@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo 没有找到 Python，无法打包。请先安装 Python 3.10 以上版本。
    pause
    exit /b 1
)

echo 正在检查 PyInstaller...
python -m pip install --upgrade pyinstaller
if errorlevel 1 (
    echo.
    echo PyInstaller 安装失败，请检查网络后重试，或手动执行：
    echo   python -m pip install pyinstaller
    pause
    exit /b 1
)

echo.
echo 正在打包（绿色版文件夹）...
python -m PyInstaller --noconfirm --clean --onedir --windowed ^
    --name "QQ批量发送助手" ^
    --icon "%CD%\assets\app.ico" ^
    --hidden-import openpyxl ^
    app.py
if errorlevel 1 (
    echo.
    echo 打包失败，请把上面的报错发给我。
    pause
    exit /b 1
)

echo.
echo 打包完成：dist\QQ批量发送助手\QQ批量发送助手.exe
echo 把 dist\QQ批量发送助手 整个文件夹复制给别人，对方不用装 Python。
pause
