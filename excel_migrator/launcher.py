"""
Excel 数据迁移工具 - 打包启动器
用于 PyInstaller 独立打包
"""
import os
import sys

def main():
    # 获取打包后的资源目录
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    # 切换到应用目录
    os.chdir(base_dir)

    # 添加 base_dir 到 Python 路径
    if base_dir not in sys.path:
        sys.path.insert(0, base_dir)

    # 打印启动信息
    print("=" * 50)
    print("Excel 数据迁移工具 v1.0")
    print("=" * 50)
    print()
    print("正在启动应用...")
    print("启动后请在浏览器中访问: http://localhost:8501")
    print()
    print("按 Ctrl+C 可停止服务")
    print("=" * 50)
    print()

    # 设置 streamlit 配置参数
    sys.argv = [
        "streamlit", "run",
        os.path.join(base_dir, "main.py"),
        "--server.port", "8501",
        "--server.address", "localhost",
        "--server.headless", "true"
    ]

    # 直接调用 streamlit CLI
    from streamlit.web import cli as stcli
    stcli.main()

if __name__ == "__main__":
    main()
