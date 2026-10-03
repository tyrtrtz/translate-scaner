"""Validated per-user preferences, saved atomically beside sheet exclusions."""

import json
from pathlib import Path

from ignored_sheets import default_ignore_file, write_json
from scanner import CHINESE_HEADERS, KOREAN_HEADERS, normalize


def default_settings():
    return dict(version=1, source_folder="", baseline_folder="", output_folder="", rows=10,
                chinese_headers=list(CHINESE_HEADERS), korean_headers=list(KOREAN_HEADERS),
                recursive=True, include_hidden_sheets=True, include_hidden_rows=True)


def validate_settings(data):
    if not isinstance(data, dict) or type(data.get("version")) is not int or data["version"] != 1:
        raise ValueError("设置版本或格式无效")
    result = default_settings()
    result.update({key: value for key, value in data.items() if key in result})
    for key in ("source_folder", "baseline_folder", "output_folder"):
        if not isinstance(result[key], str):
            raise ValueError("文件夹设置必须是文本")
    if type(result["rows"]) is not int or result["rows"] < 1:
        raise ValueError("检查行数必须是大于等于 1 的整数")
    for key in ("recursive", "include_hidden_sheets", "include_hidden_rows"):
        if type(result[key]) is not bool:
            raise ValueError("扫描开关设置必须为布尔值")
    for key in ("chinese_headers", "korean_headers"):
        values = result[key]
        if not isinstance(values, list) or not values or any(not isinstance(v, str) or not normalize(v) for v in values):
            raise ValueError("中韩表头均需至少一个非空名称")
    if {normalize(v) for v in result["chinese_headers"]} & {normalize(v) for v in result["korean_headers"]}:
        raise ValueError("中文和韩语表头不能使用同一个名称")
    return result


def load_settings(path=None):
    path = Path(path) if path is not None else default_ignore_file().with_name("settings.json")
    try:
        content = path.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        return default_settings()
    try:
        return validate_settings(json.loads(content))
    except (ValueError, UnicodeError) as error:
        raise ValueError(f"本地设置格式无效，请修复后重启：{path}\n{error}") from error


def save_settings(data, path=None):
    path = Path(path) if path is not None else default_ignore_file().with_name("settings.json")
    write_json(validate_settings(data), path)
