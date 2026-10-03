import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from settings import default_settings, load_settings, save_settings


class SavedSettings(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "配置/settings.json"

    def test_defaults_and_custom_round_trip(self):
        self.assertEqual(load_settings(self.path), default_settings())
        values = default_settings()
        values.update(source_folder="C:/客户 한국어/源文件", baseline_folder="基准", output_folder="结果",
                      rows=1, recursive=False, include_hidden_sheets=False, include_hidden_rows=False,
                      chinese_headers=["自定义中文"], korean_headers=["번역"])
        save_settings(values, self.path)
        self.assertEqual(load_settings(self.path), values)
        self.assertIn("客户 한국어", self.path.read_text())

    def test_invalid_values_never_replace_saved_settings(self):
        values = default_settings()
        save_settings(values, self.path)
        before = self.path.read_bytes()
        for changes in ({"rows":0}, {"rows":True}, {"recursive":1}, {"chinese_headers":[]},
                        {"korean_headers":["对白文本"]}, {"output_folder":None}, {"version":2}):
            with self.subTest(changes=changes), self.assertRaises(ValueError):
                save_settings(values | changes, self.path)
            self.assertEqual(self.path.read_bytes(), before)
        with patch("ignored_sheets.os.replace", side_effect=OSError("磁盘写入失败")), self.assertRaises(OSError):
            save_settings(values | {"rows":2}, self.path)
        self.assertEqual(self.path.read_bytes(), before)
        self.assertFalse(list(self.path.parent.glob("*.tmp")))

    def test_corrupt_settings_are_reported_and_not_overwritten(self):
        self.path.parent.mkdir()
        for data in ("{broken", json.dumps({"version":1,"rows":"10"}), "[]"):
            self.path.write_text(data)
            with self.assertRaisesRegex(ValueError, "本地设置格式无效"):
                load_settings(self.path)
            self.assertEqual(self.path.read_text(), data)


if __name__ == "__main__":
    unittest.main()
