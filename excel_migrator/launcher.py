"""
Excel 数据迁移工具 - 打包启动器 (tkinter + ttkbootstrap)
用于 PyInstaller 独立打包
"""
import os
import sys
import tempfile

def main():
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
        app_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        app_dir = base_dir

    os.chdir(app_dir)

    if base_dir not in sys.path:
        sys.path.insert(0, base_dir)

    print("=" * 50)
    print("Excel 数据迁移工具 v1.3")
    print("=" * 50)
    print()

    # 检查更新（仅在打包环境下）
    if getattr(sys, 'frozen', False):
        try:
            from updater import check_for_updates, download_update, launch_updater, get_current_version

            print("[*] 正在检查更新...")
            has_update, latest_version, download_url, release_notes = check_for_updates()

            if has_update:
                print(f"[*] 发现新版本: v{latest_version}")
                print("[*] 提示用户更新...")

                # 使用 tkinter 显示更新对话框
                import tkinter as tk
                from tkinter import messagebox

                root = tk.Tk()
                root.withdraw()  # 隐藏主窗口
                root.attributes('-topmost', True)  # 置顶

                response = messagebox.askyesno(
                    "发现新版本",
                    f"发现新版本 v{latest_version}！\n\n是否下载并更新？\n\n更新说明：\n{release_notes[:200]}..."
                )

                if response and download_url:
                    print("[*] 开始下载新版本...")
                    messagebox.showinfo("正在下载", "正在下载新版本，请稍候...")

                    new_exe = download_update(download_url)
                    if new_exe:
                        print("[*] 下载完成，正在更新...")
                        launch_updater(sys.executable, new_exe)
                        print("[*] 正在重启程序...")
                        root.destroy()
                        sys.exit(0)
                    else:
                        messagebox.showerror("下载失败", "下载新版本失败，请手动更新。")
                        print("[!] 下载失败")

                root.destroy()
            else:
                print("[*] 已是最新版本")

        except Exception as e:
            print(f"[!] 检查更新时出错: {e}")
            import traceback
            traceback.print_exc()

    from ui.tkinter_app import main as tk_main
    tk_main()

if __name__ == "__main__":
    main()
