"""
Excel 数据迁移工具 - GUI 启动器
使用此脚本启动应用（替代直接运行 main.py）
"""
import subprocess
import sys
import os

def main():
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 启动 Streamlit
    app_path = os.path.join(current_dir, "main.py")

    print("=" * 50)
    print("Excel 数据迁移工具 v1.0")
    print("=" * 50)
    print()
    print("正在启动应用...")
    print("启动后请在浏览器中访问: http://localhost:8501")
    print()
    print("按 Ctrl+C 可停止服务")
    print("=" * 50)

    # 使用 subprocess 启动 Streamlit
    cmd = [sys.executable, "-m", "streamlit", "run", app_path,
           "--server.port", "8501",
           "--server.address", "localhost"]

    try:
        subprocess.run(cmd, check=True)
    except KeyboardInterrupt:
        print("\n应用已停止")
    except subprocess.CalledProcessError as e:
        print(f"启动失败: {e}")
        input("按回车键退出...")

if __name__ == "__main__":
    main()
