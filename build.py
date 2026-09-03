import subprocess
import sys
import os

def main():
    print("=" * 50)
    print("  Excel/Word 数据处理工具 - 打包脚本")
    print("=" * 50)
    print()
    
    # 检查Python版本
    print(f"Python版本: {sys.version}")
    print()
    
    # 安装依赖
    print("[1/3] 安装依赖库...")
    subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller", "pandas", "openpyxl", "python-docx", "reportlab", "python-pptx", "numpy"], check=True)
    print("依赖库安装完成")
    print()
    
    # 清理旧文件
    print("[2/3] 清理旧文件...")
    for folder in ["build", "dist"]:
        if os.path.exists(folder):
            import shutil
            shutil.rmtree(folder)
    print("清理完成")
    print()
    
    # 打包
    print("[3/3] 开始打包...")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean", "--onefile", "--windowed",
        "--name", "DataTool",
        "--hidden-import", "pandas",
        "--hidden-import", "openpyxl",
        "--hidden-import", "docx",
        "--hidden-import", "reportlab",
        "--hidden-import", "pptx",
        "--hidden-import", "numpy",
        "--hidden-import", "csv",
        "--hidden-import", "threading",
        "--hidden-import", "datetime",
        "--hidden-import", "re",
        "--hidden-import", "os",
        "--hidden-import", "tkinter",
        "data_tool_gui.py"
    ]
    
    result = subprocess.run(cmd, check=True)
    
    print()
    print("=" * 50)
    print("  打包完成！")
    print(f"  输出文件: dist{os.sep}DataTool.exe")
    print("=" * 50)
    print()
    print("提示：可以将exe复制到任何Windows电脑使用")
    
if __name__ == "__main__":
    main()
