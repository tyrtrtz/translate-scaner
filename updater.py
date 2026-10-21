"""Public GitHub releases and the Windows executable replacement helper."""

import hashlib
from http.client import IncompleteRead
import json
import os
import re
import shutil
import ssl
import struct
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

VERSION = "1.1.7"
REPOSITORY = "tyrtrtz/translate-scaner"
ASSET_NAME = "HeaderChecker.exe"
BACKUP_ROOT = f"https://raw.githubusercontent.com/{REPOSITORY}/updates"


def version_numbers(value):
    if not isinstance(value, str) or not re.fullmatch(r"v?\d+\.\d+\.\d+", value):
        raise ValueError("发布版本号必须为 v主版本.次版本.修订版本")
    return tuple(map(int, value.removeprefix("v").split(".")))


def fetch_from_channels(urls, read):
    """Retry the complete read, then use our own repository's alternate host."""
    errors = []
    for url in urls:
        for attempt in range(2):
            try:
                return read(Request(url, headers={"User-Agent": f"HeaderChecker/{VERSION}"}))
            except (URLError, ConnectionError, TimeoutError, IncompleteRead, ssl.SSLError) as error:
                errors.append(error)
                if isinstance(error, HTTPError) and error.code != 429 and error.code < 500:
                    break
                if attempt == 0:
                    time.sleep(0.5)
    if all(isinstance(error, HTTPError) and error.code == 404 for error in errors):
        raise errors[-1]
    raise ConnectionError("GitHub 发布通道和仓库备用通道均未连接成功。"
                          "能打开网页不代表能下载文件，请更换网络后重试。"
                          f"\n最后错误：{errors[-1]}") from errors[-1]


def validate_release(release):
    if not isinstance(release, dict):
        raise ValueError("发布信息格式无效")
    version = release.get("version")
    version_numbers(version)
    if version.startswith("v"):
        raise ValueError("发布信息版本号不能带 v 前缀")
    url = f"https://github.com/{REPOSITORY}/releases/download/v{version}/{ASSET_NAME}"
    if (release.get("url") != url or type(release.get("size")) is not int
            or not 0 < release["size"] <= 200 * 1024 * 1024
            or not re.fullmatch(r"[0-9a-f]{64}", release.get("sha256") or "")):
        raise ValueError("发布文件的地址、大小或校验信息无效")
    return dict(version=version, url=url, size=release["size"], sha256=release["sha256"])


def check_for_update(current=VERSION):
    def read(request):
        with urlopen(request, timeout=10) as response:
            return validate_release(json.loads(response.read(1024 * 1024)))

    try:
        release = fetch_from_channels([
            f"https://github.com/{REPOSITORY}/releases/latest/download/update.json",
            f"{BACKUP_ROOT}/update.json"], read)
    except HTTPError as error:
        if error.code == 404:
            return None  # Neither channel has a published release yet.
        raise
    return release if version_numbers(release["version"]) > version_numbers(current) else None


def download_update(release, target, progress=lambda percent: None):
    release = validate_release(release)
    target = Path(target).resolve()
    # Keep staging and backup on the same volume as the executable; also fail
    # before downloading if its folder cannot be written by the current user.
    directory = Path(tempfile.mkdtemp(prefix="HeaderChecker-update-", dir=target.parent))
    destination = directory / ASSET_NAME
    def read(request):
        digest, size, last_percent = hashlib.sha256(), 0, -1
        progress(0)
        with urlopen(request, timeout=30) as response, destination.open("wb") as output:
            while chunk := response.read(128 * 1024):
                size += len(chunk)
                if size > release["size"]:
                    raise ValueError("下载文件超过发布大小，已取消更新")
                output.write(chunk)
                digest.update(chunk)
                percent = size * 100 // release["size"]
                if percent != last_percent:
                    progress(percent)
                    last_percent = percent
        if size != release["size"] or digest.hexdigest() != release["sha256"]:
            raise ValueError("下载不完整或 SHA256 校验失败，原程序未修改")

    try:
        fetch_from_channels([release["url"], f"{BACKUP_ROOT}/v{release['version']}/{ASSET_NAME}"], read)
        with destination.open("rb") as executable:
            header = executable.read(64)
            if len(header) != 64 or header[:2] != b"MZ":
                raise ValueError("更新文件不是 Windows 程序")
            executable.seek(struct.unpack_from("<I", header, 60)[0])
            if executable.read(6) != b"PE\x00\x00\x64\x86":
                raise ValueError("更新文件不是 Windows x64 程序")
        return destination
    except Exception:
        shutil.rmtree(directory, ignore_errors=True)
        raise


def start_installation(download, target, process_id=None):
    """Wait for this app to close, atomically replace it, and restart it."""
    download, target = Path(download).resolve(), Path(target).resolve()
    process_id = os.getpid() if process_id is None else int(process_id)
    quote = lambda value: "'" + str(value).replace("'", "''") + "'"
    script = download.parent / "install.ps1"
    script.write_text(f"""$ErrorActionPreference = 'Stop'
$targetPath = {quote(target)}
$downloadPath = {quote(download)}
$backupPath = {quote(download.parent / 'previous.exe')}
$stagingPath = {quote(download.parent)}
$oldProcessId = {process_id}
function Start-UpdatedApp {{
    $info = New-Object System.Diagnostics.ProcessStartInfo
    $info.FileName = $targetPath
    $info.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetPath)
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    # A restarted onefile app needs its own extracted runtime, not the old one.
    $info.EnvironmentVariables['PYINSTALLER_RESET_ENVIRONMENT'] = '1'
    [System.Diagnostics.Process]::Start($info) | Out-Null
}}
try {{
    $deadline = (Get-Date).AddSeconds(60)
    while (Get-Process -Id $oldProcessId -ErrorAction SilentlyContinue) {{
        if ((Get-Date) -gt $deadline) {{ throw '原程序尚未退出，请关闭后重新更新。' }}
        Start-Sleep -Milliseconds 250
    }}
    # PyInstaller's parent launcher can retain the exe briefly after Tk closes.
    for ($attempt = 0; $attempt -lt 60; $attempt++) {{
        try {{ [System.IO.File]::Move($targetPath, $backupPath); break }}
        catch {{ if ($attempt -eq 59) {{ throw }}; Start-Sleep -Milliseconds 500 }}
    }}
    [System.IO.File]::Move($downloadPath, $targetPath)
    Start-UpdatedApp
    Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction SilentlyContinue
}} catch {{
    $failure = $_.Exception.Message
    if (Test-Path -LiteralPath $backupPath) {{
        try {{
            if (Test-Path -LiteralPath $targetPath) {{ Remove-Item -LiteralPath $targetPath -Force }}
            [System.IO.File]::Move($backupPath, $targetPath)
            Start-UpdatedApp
        }} catch {{ $failure += "`n恢复文件位置：$backupPath`n" + $_.Exception.Message }}
    }}
    $failure | Set-Content -LiteralPath (Join-Path $stagingPath 'update-error.txt') -Encoding UTF8
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("更新失败。原程序或备份已保留。`n$failure", '更新失败') | Out-Null
    exit 1
}}
""", encoding="utf-8-sig")
    return subprocess.Popen(["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                             "-File", str(script)], creationflags=subprocess.CREATE_NO_WINDOW,
                            close_fds=True)


def can_install_updates():
    return sys.platform == "win32" and bool(getattr(sys, "frozen", False))
