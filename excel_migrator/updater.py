"""
Excel 数据迁移工具 - 自动更新模块
"""
import json
import os
import subprocess
import sys
import tempfile
import urllib.request
import urllib.error
from pathlib import Path


# 当前版本
CURRENT_VERSION = "1.3"

# GitHub 仓库
GITHUB_REPO = "leshao888/excel_migrator"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
GITHUB_DOWNLOAD_URL = f"https://github.com/{GITHUB_REPO}/releases/download"


def get_current_version():
    """获取当前版本"""
    return CURRENT_VERSION


def compare_versions(current: str, latest: str) -> int:
    """
    比较版本号
    返回: 1 最新版更新, 0 版本相同, -1 当前版本更新
    """
    def normalize(v):
        # 移除 v 前缀
        v = v.lstrip('v')
        # 分割并转为整数
        parts = []
        for p in v.split('.'):
            try:
                parts.append(int(p))
            except ValueError:
                parts.append(0)
        return parts

    cur = normalize(current)
    lat = normalize(latest)

    # 补齐长度
    while len(cur) < len(lat):
        cur.append(0)
    while len(lat) < len(cur):
        lat.append(0)

    for c, l in zip(cur, lat):
        if c > l:
            return -1
        elif c < l:
            return 1
    return 0


def check_for_updates():
    """
    检查 GitHub 是否有新版本
    返回: (has_update, latest_version, download_url, release_notes)
    """
    try:
        req = urllib.request.Request(
            GITHUB_API_URL,
            headers={
                'User-Agent': 'ExcelMigrator',
                'Accept': 'application/vnd.github.v3+json'
            }
        )

        with urllib.request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read().decode('utf-8'))

        latest_version = data.get('tag_name', CURRENT_VERSION)
        release_notes = data.get('body', '')

        # 获取 exe 下载地址
        download_url = None
        for asset in data.get('assets', []):
            if asset.get('name', '').lower() == 'excelmigrator.exe':
                download_url = asset.get('browser_download_url')
                break

        # 如果没有找到精确匹配的，从 tag 构建下载链接
        if not download_url:
            download_url = f"{GITHUB_DOWNLOAD_URL}/{latest_version}/ExcelMigrator.exe"

        has_update = compare_versions(CURRENT_VERSION, latest_version) == 1

        return (has_update, latest_version, download_url, release_notes)

    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as e:
        print(f"[Update] 检查更新失败: {e}")
        return (False, CURRENT_VERSION, None, "")
    except Exception as e:
        print(f"[Update] 检查更新异常: {e}")
        return (False, CURRENT_VERSION, None, "")


def download_update(url: str, progress_callback=None) -> str:
    """
    下载新版本 exe
    返回: 下载的 exe 文件路径
    """
    temp_dir = tempfile.gettempdir()
    dest_path = os.path.join(temp_dir, "ExcelMigrator_new.exe")

    try:
        req = urllib.request.Request(
            url,
            headers={
                'User-Agent': 'ExcelMigrator',
                'Accept': 'application/octet-stream'
            }
        )

        with urllib.request.urlopen(req, timeout=60) as response:
            total_size = int(response.headers.get('Content-Length', 0))
            downloaded = 0
            chunk_size = 8192

            with open(dest_path, 'wb') as f:
                while True:
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total_size > 0:
                        progress_callback(downloaded, total_size)

        return dest_path

    except Exception as e:
        print(f"[Update] 下载失败: {e}")
        if os.path.exists(dest_path):
            os.remove(dest_path)
        return None


def launch_updater(current_exe: str, new_exe: str):
    """
    启动更新程序，替换当前 exe
    """
    # 创建批处理脚本
    temp_dir = tempfile.gettempdir()
    batch_path = os.path.join(temp_dir, "update_excelmigrator.bat")

    # 批处理内容：等待原程序退出，替换exe，启动新程序
    batch_content = f'''@echo off
timeout /t 2 /nobreak >nul
:wait
tasklist /FI "IMAGENAME eq ExcelMigrator.exe" 2>nul | find /I "ExcelMigrator.exe" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait
)
copy /Y "{new_exe}" "{current_exe}"
del "{new_exe}"
start "" "{current_exe}"
del "%~f0"
'''

    with open(batch_path, 'w', encoding='gbk') as f:
        f.write(batch_content)

    # 启动批处理程序（以隐藏窗口方式）
    subprocess.Popen(
        ['cmd', '/c', 'start', '', '/min', batch_path],
        shell=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        stdin=subprocess.DEVNULL
    )


if __name__ == "__main__":
    # 测试用
    has_update, latest, url, notes = check_for_updates()
    print(f"当前版本: {CURRENT_VERSION}")
    print(f"最新版本: {latest}")
    print(f"有更新: {has_update}")
    if url:
        print(f"下载链接: {url}")
