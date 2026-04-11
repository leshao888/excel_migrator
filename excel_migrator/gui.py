"""
Excel 数据迁移工具 - GUI 启动器 (tkinter + ttkbootstrap)
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.tkinter_app import main

if __name__ == "__main__":
    main()
