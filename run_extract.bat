@echo off
echo Excel图片提取工具启动中...
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误: 未检测到Python安装。请安装Python 3.6或更高版本。
    echo 您可以从 https://www.python.org/downloads/ 下载Python。
    pause
    exit /b
)

echo 正在安装所需依赖...
pip install -r requirements.txt
echo.

echo 正在启动Excel图片提取工具...
python extract_images.py
echo.

echo 程序已关闭。如果您已提取图片，它们将保存在您选择的输出文件夹中。
echo 按任意键退出...
pause