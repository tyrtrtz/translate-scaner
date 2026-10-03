import hashlib
import tempfile
import unittest
import re
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from ignored_sheets import sheet_key
from matcher import match_folder
from scanner import scan_folder


class TranslationMatches(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.source = self.root / "源한국어"
        self.baseline = self.root / "基准"
        self.output = self.root / "结果"
        self.source.mkdir()
        self.baseline.mkdir()

    def write_source(self, rows, name="客户.xlsx", title="正文"):
        book = Workbook()
        sheet = book.active
        sheet.title = title
        sheet.append(["台词", "한국어"])
        for row in rows:
            sheet.append(row)
        path = self.source / name
        book.save(path)
        book.close()
        return path

    def write_baseline(self, texts, name="story.xlsx"):
        book = Workbook()
        sheet = book.active
        sheet.cell(1, 7, "Text")
        sheet.cell(1, 1, "Key")
        sheet.cell(2, 1, "=1+1")
        sheet.cell(2, 1).font = Font(bold=True)
        sheet.column_dimensions["G"].width = 40
        for number, text in enumerate(texts, 2):
            sheet.cell(number, 7, text)
        path = self.baseline / name
        book.save(path)
        book.close()
        return path

    def test_counts_details_stats_unicode_and_originals_unchanged(self):
        rows = [["一条", "번역1"], ["  两条  ", "번역2"], ["两条", "번역3"],
                *[["三条", f"번역{i}"] for i in range(3)],
                *[["四条", f"번역{i}"] for i in range(4)],
                ["无基准", "외부"], [None, "空中文"], ["   ", "空白"], [" 空译文 ", None], ["内部 空格", "구별"]]
        source = self.write_source(rows)
        baseline = self.write_baseline([" 一条 ", "两条", "三条", "四条", "未找到", "空译文", "内部空格"])
        book = load_workbook(source)
        book.active.row_dimensions[2].hidden = True
        book.active.auto_filter.ref = "A1:B15"
        hidden = book.create_sheet("隐藏")
        hidden.append(["台词", "韩文"])
        hidden.append(["独有隐藏", "숨김"])
        hidden.sheet_state = "hidden"
        book.save(source)
        book.close()
        self.write_baseline(["独有隐藏", "一条"], name="char.xlsx")
        original = {p: hashlib.sha256(p.read_bytes()).digest() for p in (source, *self.baseline.iterdir())}
        result = match_folder(scan_folder(self.source, 1), self.baseline, self.output)
        self.assertEqual((result.baseline_rows, result.matched_rows), (9, 7))
        book = load_workbook(result.destination / "基准匹配结果/story.xlsx")
        sheet = book.active
        self.assertEqual([sheet.cell(i, 27).value for i in range(2, 9)], [1, 2, 3, 4, 0, 1, 0])
        self.assertEqual(sheet["R2"].value, "一条")
        self.assertEqual(sheet["S2"].value, "번역1")
        self.assertIn("行号：2", sheet["T2"].value)
        self.assertIn("Sheet：正文", sheet["T2"].value)
        self.assertEqual(sheet["R3"].value, "  两条  ")
        self.assertEqual(sheet["U3"].value, "两条")
        self.assertEqual(sheet["X4"].value, "三条")
        self.assertTrue(all(sheet.cell(5, c).value is None for c in range(18, 27)))
        self.assertIsNone(sheet["S7"].value)
        self.assertEqual(sheet["A2"].value, "=1+1")
        self.assertTrue(sheet["A2"].font.bold)
        self.assertEqual(sheet.column_dimensions["G"].width, 40)
        book.close()
        stats = load_workbook(result.destination / "源文件匹配统计.xlsx")
        row = list(stats["文件统计"].values)[1]
        self.assertEqual(row[2:7], (2, 0, 14, 12, 2))
        self.assertAlmostEqual(row[7], 12 / 14)
        self.assertEqual(stats["文件统计"]["H2"].number_format, "0.00%")
        stats.close()
        self.assertEqual(original, {p: hashlib.sha256(p.read_bytes()).digest() for p in original})
        self.assertFalse(any(self.output.glob("匹配临时_*")))
        self.assertFalse(any(result.destination.glob("*.sqlite3")))

    def test_ignored_and_preexisting_abnormal_sheets_never_count(self):
        path = self.write_source([["命中", "有效"]])
        book = load_workbook(path)
        abnormal = book.create_sheet("异常")
        abnormal.append(["台词"])
        abnormal.append(["命中", "不应输出"])
        ignored = book.create_sheet("忽略")
        ignored.append(["台词", "韩文"])
        ignored.append(["命中", "不应输出"])
        book.save(path)
        book.close()
        record = {sheet_key(path, "忽略")}
        report = scan_folder(self.source, 1, ignored_sheets=record)
        # A sheet fixed after checking remains excluded until the user rechecks.
        book = load_workbook(path)
        book["异常"]["B1"] = "韩文"
        book.save(path)
        book.close()
        self.write_baseline(["命中"])
        result = match_folder(report, self.baseline, self.output, ignored_sheets=record)
        self.assertEqual((result.source_rows, result.source_matched, result.skipped), (1, 1, 2))
        self.assertEqual([r.sheet for r in result.check_report.exceptions], ["异常"])
        stats = load_workbook(result.destination / "源文件匹配统计.xlsx")
        self.assertEqual({r[2] for r in list(stats["跳过清单"].values)[1:]}, {"异常", "忽略"})
        stats.close()

    def test_changed_header_excluded_and_body_failure_rolls_back(self):
        path = self.write_source([["命中", "错误"]], title="变更")
        book = load_workbook(path)
        for title, rows in (("失败", [["命中", "半途数据"], ["命中", "=1+1"]]), ("合法", [["命中", "正确"]])):
            sheet = book.create_sheet(title)
            sheet.append(["台词", "韩文"])
            for row in rows:
                sheet.append(row)
        book.save(path)
        book.close()
        report = scan_folder(self.source, 1)
        book = load_workbook(path)
        book["变更"]["B1"] = "新未知表头"
        book.save(path)
        book.close()
        self.write_baseline(["命中"])
        result = match_folder(report, self.baseline, self.output)
        self.assertEqual((result.source_rows, result.source_matched, result.skipped), (1, 1, 2))
        book = load_workbook(result.destination / "基准匹配结果/story.xlsx")
        self.assertEqual(book.active["AA2"].value, 1)
        self.assertEqual(book.active["S2"].value, "正确")
        book.close()

    def test_literal_formula_like_text_not_executed(self):
        path = self.write_source([["=文字", "=번역"]])
        book = load_workbook(path)
        book.active["A2"].data_type = book.active["B2"].data_type = "s"
        book.save(path)
        book.close()
        baseline = self.write_baseline(["=文字"])
        book = load_workbook(baseline)
        book.active["G2"].data_type = "s"
        book.save(baseline)
        book.close()
        result = match_folder(scan_folder(self.source, 1), self.baseline, self.output)
        book = load_workbook(result.destination / "基准匹配结果/story.xlsx")
        self.assertEqual(book.active["S2"].value, "=번역")
        self.assertEqual(book.active["S2"].data_type, "s")
        book.close()

    def test_scan_switches_control_matching_scope_and_denominator(self):
        path = self.write_source([["命中", "可见"], ["命中", "隐藏行"], ["未匹配", "可见"]])
        book = load_workbook(path)
        book.active.row_dimensions[3].hidden = True
        hidden = book.create_sheet("隐藏表")
        hidden.append(["台词", "韩文"])
        hidden.append(["命中", "隐藏表译文"])
        hidden.sheet_state = "veryHidden"
        book.save(path)
        book.close()
        nested = self.source / "子目录"
        nested.mkdir()
        self.write_source([["命中", "子目录译文"]], name="子目录/嵌套.xlsx")
        self.write_baseline(["命中"])
        report = scan_folder(self.source, 1, recursive=False, include_hidden_sheets=False, include_hidden_rows=False)
        self.assertEqual(report.files, 1)
        self.assertEqual(report.passed_sheets, 1)
        self.assertEqual(len(report.skipped_sheets), 1)
        result = match_folder(report, self.baseline, self.output)
        self.assertEqual((result.source_rows, result.source_matched, result.skipped), (2, 1, 1))
        book = load_workbook(result.destination / "基准匹配结果/story.xlsx")
        self.assertEqual(book.active["AA2"].value, 1)
        self.assertEqual(book.active["S2"].value, "可见")
        book.close()
        stats = load_workbook(result.destination / "源文件匹配统计.xlsx")
        self.assertEqual(stats["文件统计"]["H2"].value, .5)
        stats.close()

    def test_hidden_header_row_is_not_used_when_disabled(self):
        path = self.write_source([["中文", "번역"]])
        book = load_workbook(path)
        book.active.row_dimensions[1].hidden = True
        book.save(path)
        book.close()
        self.assertEqual(scan_folder(self.source, 1).passed_sheets, 1)
        report = scan_folder(self.source, 1, include_hidden_rows=False)
        self.assertEqual(report.passed_sheets, 0)
        self.assertEqual(len(report.exceptions), 1)
        self.assertIn("排除隐藏行：1", report.exceptions[0].details)

    def test_validation_existing_output_columns_and_repeat_outputs(self):
        self.write_source([["中文", "한국어"]])
        baseline = self.write_baseline(["中文"])
        report = scan_folder(self.source, 1)
        for source, output in ((self.source, self.output), (self.baseline, self.source / "输出")):
            with self.assertRaises(ValueError):
                match_folder(report, source, output)
        first = match_folder(report, self.baseline, self.output)
        second = match_folder(report, self.baseline, self.output)
        self.assertNotEqual(first.destination, second.destination)
        book = load_workbook(baseline)
        book.active["R2"] = "原有内容"
        book.save(baseline)
        book.close()
        before = sorted(self.output.iterdir())
        with self.assertRaisesRegex(ValueError, "已有内容"):
            match_folder(report, self.baseline, self.output)
        self.assertEqual(sorted(self.output.iterdir()), before)

    def test_streaming_preserves_package_parts_prefixes_and_bad_dimensions(self):
        self.write_source([["中文", "한국어"]])
        baseline = self.write_baseline(["中文"])
        book = load_workbook(baseline)
        book.active["R2"].font = Font(italic=True)  # Empty styled cell must not duplicate the new R2.
        book.save(baseline)
        book.close()
        with ZipFile(baseline) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        xml = parts['xl/worksheets/sheet1.xml'].decode()
        xml = re.sub(r'<dimension\b[^>]*/>', '<dimension ref="A1"/>', xml)
        xml = xml.replace('xmlns=', 'xmlns:m=', 1)
        xml = re.sub(r'<(/?)([A-Za-z][A-Za-z0-9]*)(?=[\s/>])', r'<\1m:\2', xml)
        xml = xml.replace('<m:worksheet ', '<m:worksheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac" ', 1)
        parts['xl/worksheets/sheet1.xml'] = xml.encode()
        with ZipFile(baseline, 'w') as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        result = match_folder(scan_folder(self.source, 1), self.baseline, self.output)
        target = result.destination / '基准匹配结果/story.xlsx'
        with ZipFile(target) as archive:
            for name, data in parts.items():
                if name != 'xl/worksheets/sheet1.xml':
                    self.assertEqual(archive.read(name), data)
            xml = archive.read('xl/worksheets/sheet1.xml')
            self.assertIn(b'mc:Ignorable="x14ac"', xml)
            ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            self.assertEqual(len(ET.fromstring(xml).findall('.//m:c[@r="R2"]', ns)), 1)
        book = load_workbook(target, read_only=True)
        self.assertEqual(book.active['R2'].value, '中文')
        self.assertEqual(book.active['AA2'].value, 1)
        self.assertEqual(book.active['A2'].value, '=1+1')
        book.close()


if __name__ == "__main__":
    unittest.main()
