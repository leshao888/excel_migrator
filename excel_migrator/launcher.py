"""
Excel 数据迁移工具 - 打包启动器 (tkinter + ttkbootstrap)
用于 PyInstaller 独立打包
"""
import os
import sys

def main():
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    os.chdir(base_dir)

    if base_dir not in sys.path:
        sys.path.insert(0, base_dir)

    print("=" * 50)
    print("Excel 数据迁移工具 v1.0")
    print("=" * 50)
    print()

    from ui.tkinter_app import main as tk_main
    tk_main()

if __name__ == "__main__":
    main()
