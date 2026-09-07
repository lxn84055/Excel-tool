import subprocess
import sys
import os
import shutil

def main():
    print("=" * 50)
    print("  使用 --onedir 模式打包")
    print("=" * 50)
    print()
    
    # 清理旧文件
    print("[1/3] 清理旧文件...")
    for folder in ["build", "dist"]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
    print("清理完成")
    print()
    
    # 打包
    print("[2/3] 开始打包...")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--onedir",          # 使用onedir模式
        "--windowed",        # 无控制台窗口
        "--name", "DataTool",
        "data_tool_gui.py"
    ]
    
    result = subprocess.run(cmd)
    
    if result.returncode != 0:
        print("打包失败！")
        return
    
    print()
    print("[3/3] 打包完成！")
    print()
    print("输出位置: dist/DataTool/")
    print("主程序: dist/DataTool/DataTool.exe")
    print()
    print("注意：需要将整个 DataTool 文件夹一起分发")
    print("不能只复制 DataTool.exe 一个文件")
    
if __name__ == "__main__":
    main()
