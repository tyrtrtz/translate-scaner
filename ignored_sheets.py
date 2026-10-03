"""Local, per-user sheet exclusions. No third-party dependencies."""

import json
import os
import sys
import tempfile
import unicodedata
from pathlib import Path


def default_ignore_file():
    if sys.platform == "win32":
        base = Path(os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
                    or Path.home() / "AppData" / "Local")
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return base / "HeaderChecker" / "ignored_sheets.json"


def sheet_key(path, sheet):
    filename = unicodedata.normalize("NFC", str(Path(path).expanduser().resolve()))
    return os.path.normcase(filename), unicodedata.normalize("NFC", sheet)


def load_ignored(path=None):
    path = Path(path) if path is not None else default_ignore_file()
    try:
        content = path.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        return set()
    try:
        data = json.loads(content)
        if not isinstance(data, dict) or data.get("version") != 1 or not isinstance(data.get("sheets"), list):
            raise ValueError
        records = set()
        for item in data["sheets"]:
            if (not isinstance(item, dict) or not isinstance(item.get("file"), str)
                    or not isinstance(item.get("sheet"), str) or not item["sheet"]
                    or not Path(item["file"]).is_absolute()):
                raise ValueError
            records.add(sheet_key(item["file"], item["sheet"]))
        return records
    except (ValueError, UnicodeError) as error:
        raise ValueError(f"忽略记录格式无效，请修复此文件后重启：{path}") from error


def save_ignored(records, path=None):
    path = Path(path) if path is not None else default_ignore_file()
    path.parent.mkdir(parents=True, exist_ok=True)
    data = {"version": 1, "sheets": [
        {"file": filename, "sheet": sheet} for filename, sheet in sorted(records)
    ]}
    temporary = None
    try:
        with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=path.parent,
                                         prefix="ignored_", suffix=".tmp", delete=False) as stream:
            temporary = Path(stream.name)
            json.dump(data, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
