"""Publish verified release files to our repository's alternate download branch."""

import hashlib
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from updater import ASSET_NAME, REPOSITORY, VERSION, validate_release, version_numbers


def publish(build=Path("dist")):
    release = validate_release(json.loads((build / "update.json").read_text(encoding="utf-8")))
    executable = build / ASSET_NAME
    if (release["version"] != VERSION or executable.stat().st_size != release["size"]
            or hashlib.sha256(executable.read_bytes()).hexdigest() != release["sha256"]):
        raise ValueError("发布文件与版本或校验信息不一致")
    with tempfile.TemporaryDirectory(prefix="headerchecker-publish-") as directory:
        root = Path(directory)

        def git(*args, check=True):
            return subprocess.run(["git", *args], cwd=root, check=check, capture_output=True,
                                  text=True, env=dict(os.environ, GIT_TERMINAL_PROMPT="0"))

        git("init", "-b", "updates")
        git("config", "user.name", "github-actions[bot]")
        git("config", "user.email", "41898282+github-actions[bot]@users.noreply.github.com")
        git("config", "credential.helper", "!gh auth git-credential")
        git("remote", "add", "origin", f"https://github.com/{REPOSITORY}.git")
        exists = git("ls-remote", "--exit-code", "--heads", "origin", "updates", check=False)
        if exists.returncode == 0:
            git("fetch", "--depth=1", "origin", "updates")
            git("checkout", "-B", "updates", "FETCH_HEAD")
        elif exists.returncode != 2:
            raise RuntimeError("无法读取备用发布分支：" + exists.stderr)
        version_dir = root / f"v{VERSION}"
        version_dir.mkdir(exist_ok=True)
        target = version_dir / ASSET_NAME
        if target.exists() and hashlib.sha256(target.read_bytes()).hexdigest() != release["sha256"]:
            raise ValueError("同版本的备用文件已存在且内容不同，禁止覆盖")
        shutil.copyfile(executable, target)
        latest = root / "update.json"
        if not latest.exists() or version_numbers(json.loads(latest.read_text())["version"]) <= version_numbers(VERSION):
            shutil.copyfile(build / "update.json", latest)
        git("add", "--", f"v{VERSION}/{ASSET_NAME}", "update.json")
        if git("diff", "--cached", "--quiet", check=False).returncode:
            git("commit", "-m", f"Publish alternate download v{VERSION}")
            git("push", "origin", "updates")
    print(f"Published alternate download v{VERSION}")


if __name__ == "__main__":
    publish()
