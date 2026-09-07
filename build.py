import subprocess
import sys
import os
import shutil

def main():
    print("=" * 60)
    print("  Excel/Word 数据处理工具 - 打包脚本")
    print("  模式：--onefile（单文件）+ 优化启动速度")
    print("=" * 60)
    print()
    
    # 检查Python版本
    print(f"Python版本: {sys.version}")
    print()
    
    # 安装依赖
    print("[1/4] 检查并安装依赖库...")
    subprocess.run([
        sys.executable, "-m", "pip", "install",
        "pyinstaller", "pandas", "openpyxl", "python-docx",
        "reportlab", "python-pptx", "numpy"
    ], check=True)
    print("依赖库安装完成")
    print()
    
    # 清理旧文件
    print("[2/4] 清理旧文件...")
    for folder in ["build", "dist"]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
    if os.path.exists("DataTool.spec"):
        os.remove("DataTool.spec")
    print("清理完成")
    print()
    
    # 打包
    print("[3/4] 开始打包（可能需要3-5分钟）...")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",           # 单文件模式
        "--windowed",          # 无控制台窗口
        "--name", "DataTool",  # 输出名称
        "--noupx",             # 禁用UPX压缩（加速启动）
        # 排除不需要的库（减小体积）
        "--exclude-module", "matplotlib",
        "--exclude-module", "scipy",
        "--exclude-module", "sklearn",
        "--exclude-module", "PIL",
        "--exclude-module", "tkinter.test",
        "--exclude-module", "unittest",
        "--exclude-module", "pydoc",
        "--exclude-module", "doctest",
        # 隐藏导入必要库
        "--hidden-import", "pandas",
        "--hidden-import", "openpyxl",
        "--hidden-import", "docx",
        "--hidden-import", "reportlab",
        "--hidden-import", "pptx",
        "--hidden-import", "numpy",
        "data_tool_gui.py"
    ]
    
    result = subprocess.run(cmd)
    
    if result.returncode != 0:
        print()
        print("[错误] 打包失败！")
        return
    
    print()
    print("[4/4] 打包完成！")
    print()
    print("=" * 60)
    print(f"  输出文件: dist{os.sep}DataTool.exe")
    print(f"  文件大小: {os.path.getsize(f'dist{os.sep}DataTool.exe') / 1024 / 1024:.1f} MB")
    print("=" * 60)
    print()
    print("提示：")
    print("  - 只需要分发 DataTool.exe 一个文件")
    print("  - 首次运行可能需要3-5秒加载")
    print("  - 目标电脑需要Windows 7及以上系统")
    
if __name__ == "__main__":
    main()
