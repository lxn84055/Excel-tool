@echo off
chcp 65001 >nul
echo ========================================
echo    Excel/Word 数据处理工具 - 打包脚本
echo ========================================
echo.

:: 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python，请先安装Python 3.7+
    pause
    exit /b 1
)

:: 检查PyInstaller
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [信息] 正在安装PyInstaller...
    pip install pyinstaller
)

:: 检查依赖库
echo [信息] 检查依赖库...
pip install pandas openpyxl python-docx reportlab python-pptx numpy

:: 清理旧文件
echo [信息] 清理旧文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

:: 打包
echo [信息] 开始打包...
pyinstaller --noconfirm --clean --onefile --windowed ^
    --name "数据处理工具" ^
    --hidden-import "pandas" ^
    --hidden-import "openpyxl" ^
    --hidden-import "docx" ^
    --hidden-import "reportlab" ^
    --hidden-import "pptx" ^
    --hidden-import "numpy" ^
    --hidden-import "csv" ^
    --hidden-import "threading" ^
    --hidden-import "datetime" ^
    --hidden-import "re" ^
    --hidden-import "os" ^
    --hidden-import "tkinter" ^
    --hidden-import "scrolledtext" ^
    --hidden-import "filedialog" ^
    --hidden-import "messagebox" ^
    data_tool_gui.py

if errorlevel 1 (
    echo [错误] 打包失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo    打包完成！
echo    输出文件: dist\数据处理工具.exe
echo ========================================
echo.
echo 提示：可以将exe复制到任何Windows电脑使用
echo.
pause
