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
from matcher import align_boundary_newlines, match_folder, write_detail_sheet
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

    def test_output_boundary_newlines_follow_each_baseline_row(self):
        def preserve_cell_text(path, cells):
            # Without lxml, openpyxl's fixture writer emits literal CR. XML
            # readers normalize it, so encode CR as Excel's character reference.
            with ZipFile(path) as archive:
                parts = [(item, archive.read(item)) for item in archive.infolist()]
            with ZipFile(path, "w") as archive:
                for item, data in parts:
                    if item.filename == "xl/worksheets/sheet1.xml":
                        root = ET.fromstring(data)
                        for cell in root.findall(".//{*}c"):
                            if cell.get("r") in cells:
                                cell.find("{*}is/{*}t").text = cells[cell.get("r")]
                        data = ET.tostring(root, encoding="utf-8").replace(b"\r", b"&#13;")
                    archive.writestr(item, data)

        chinese = " \n\n\t甲\n乙 \t\n\n "
        korean = " \n\n\t번역\n둘 \t\n\n "
        source = self.write_source([[chinese, korean], ["无译文", None],
                                    ["\n三组\n", "\n단일\n"], ["\r三组\r", "\r단일\r"], ["\r\n三组\r\n", "\r\n단일\r\n"]])
        baselines = ["甲\n乙", "\r\n甲\n乙", "甲\n乙\r\n", " \r\n甲\n乙\n ", "无译文", "三组"]
        baseline = self.write_baseline(baselines)
        preserve_cell_text(source, {"A5": "\r三组\r", "B5": "\r단일\r",
                                    "A6": "\r\n三组\r\n", "B6": "\r\n단일\r\n"})
        preserve_cell_text(baseline, {f"G{row}": text for row, text in enumerate(baselines, 2)})
        # A fourth source occurrence suppresses R:Z, but AD must still be cleaned.
        many = self.write_source([["\r\n多条\r\n", "\r\n여러\r\n"]] * 4, name="多条.xlsx")
        preserve_cell_text(many, {f"{column}{row}": text for row in range(2, 6)
                                 for column, text in (("A", "\r\n多条\r\n"), ("B", "\r\n여러\r\n"))})
        other = self.write_baseline(["多条", "\r\n多条\r\n"], name="另一个.xlsx")
        preserve_cell_text(other, {"G3": "\r\n多条\r\n"})
        original = {p: hashlib.sha256(p.read_bytes()).digest() for p in (source, baseline, other)}
        result = match_folder(scan_folder(self.source, 1), self.baseline, self.output)
        self.assertEqual((result.baseline_rows, result.matched_rows, result.recommended_rows, result.source_rows,
                          result.source_matched, result.match_pairs, result.conflict_texts), (8, 8, 7, 9, 9, 16, 0))
        book = load_workbook(result.destination / "基准匹配结果/story.xlsx")
        clean_chinese, clean_korean = " \t甲\n乙 \t ", " \t번역\n둘 \t "
        expected = [(clean_chinese, clean_korean),
                    (" \n\n\t甲\n乙 \t ", " \n\n\t번역\n둘 \t "),
                    (" \t甲\n乙 \t\n\n ", " \t번역\n둘 \t\n\n "),
                    (chinese, korean)]
        for row, (cn, ko) in enumerate(expected, 2):
            self.assertEqual(book.active.cell(row, 7).value, baselines[row - 2])
            self.assertEqual(book.active.cell(row, 18).value, cn)
            self.assertEqual(book.active.cell(row, 19).value, ko)
            self.assertEqual(book.active.cell(row, 30).value, ko)
            self.assertEqual(book.active.cell(row, 33).value, "唯一译文")
        self.assertIsNone(book.active["S6"].value)
        self.assertIsNone(book.active["AD6"].value)
        self.assertEqual([book.active.cell(7,c).value for c in (18,21,24)], ["三组"] * 3)
        self.assertEqual([book.active.cell(7,c).value for c in (19,22,25,30)], ["단일"] * 4)
        book.close()
        book = load_workbook(result.destination / "基准匹配结果/另一个.xlsx")
        self.assertIsNone(book.active["R2"].value)
        self.assertEqual(book.active["AD2"].value, "여러")
        self.assertEqual(book.active["AA2"].value, 4)
        self.assertEqual(book.active["G3"].value, "\r\n多条\r\n")
        self.assertEqual(book.active["AD3"].value, "\r\n여러\r\n")
        book.close()
        # Evidence remains raw, and none of the original workbooks is saved.
        book = load_workbook(result.destination / "匹配明细.xlsx")
        self.assertTrue(all(r[8:10] == (chinese, korean) for r in list(book["匹配明细"].values)[1:] if r[4] == source.name and r[7] == 2))
        book.close()
        self.assertEqual(original, {p: hashlib.sha256(p.read_bytes()).digest() for p in original})
        self.assertEqual(align_boundary_newlines("\n=文字\n", False, False), "=文字")
        self.assertEqual(align_boundary_newlines(" \r\n\t ", False, False), " \t ")
        self.assertEqual(align_boundary_newlines("原文无换行", True, True), "\n原文无换行\n")
        self.assertEqual(align_boundary_newlines(" \t正文\t ", True, False), "\n \t正文\t ")
        self.assertEqual(align_boundary_newlines(" \r\n正文\n\n ", True, True), " \r\n正文\n\n ")
        self.assertEqual(align_boundary_newlines("正文\n\n", True, False), "\n正文")
        self.assertEqual(align_boundary_newlines("\n正文", False, True), "正文\n")
        self.assertEqual(align_boundary_newlines(None, True, True), None)
        self.assertEqual(align_boundary_newlines(" \t ", True, True), " \t ")

    def test_missing_boundary_newlines_are_added_per_baseline(self):
        self.write_source([["中文\n正文", "번역\n본문"]] * 3 + [["缺译", None], ["未推荐", "甲译文"], ["未推荐", "乙译文"]])
        self.write_source([["多条", "여러"]] * 4, name="多条.xlsx")
        baselines = ["中文\n正文", "\n中文\n正文", "中文\n正文\n", "\n\n中文\n正文\n\n", "\n缺译\n", "\n未命中\n", "\n未推荐\n", "\n多条\n"]
        self.write_baseline(baselines)
        result = match_folder(scan_folder(self.source, 1), self.baseline, self.output)
        self.assertEqual((result.baseline_rows, result.matched_rows, result.recommended_rows), (8, 7, 5))
        book = load_workbook(result.destination / "基准匹配结果/story.xlsx")
        expected = [("中文\n正文", "번역\n본문"), ("\n中文\n正文", "\n번역\n본문"),
                    ("中文\n正文\n", "번역\n본문\n"), ("\n中文\n正文\n", "\n번역\n본문\n")]
        for row, (chinese, korean) in enumerate(expected, 2):
            self.assertEqual(book.active.cell(row, 7).value, baselines[row - 2])
            self.assertEqual([book.active.cell(row, c).value for c in (18, 21, 24)], [chinese] * 3)
            self.assertEqual([book.active.cell(row, c).value for c in (19, 22, 25, 30)], [korean] * 4)
        for row in (6, 7, 8):
            self.assertIsNone(book.active.cell(row, 30).value)
        self.assertIsNone(book.active['S6'].value)
        self.assertEqual(book.active['AG8'].value, '待确认')
        self.assertIsNone(book.active['R9'].value)
        self.assertEqual(book.active['AD9'].value, '\n여러\n')
        self.assertEqual(book.active['AA9'].value, 4)
        book.close()
        book = load_workbook(result.destination / "匹配明细.xlsx")
        self.assertEqual([(row[8], row[9]) for row in list(book['匹配明细'].values)[1:] if row[8] == '中文\n正文'],
                         [('中文\n正文', '번역\n본문')] * 12)
        book.close()

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

    def test_complete_details_missing_and_conflicts_keep_all_candidates(self):
        path = self.write_source([
            ["重复","韩语甲"], ["重复",None], ["首尾一致"," 번역 "], ["首尾一致","번역"],
            ["单一缺译"," \t "], ["四条","韩1"], ["四条","韩1"], ["四条","韩2"], ["四条",None],
            ["不在基准",None], ["不在基准","例1"], ["不在基准","例2"],
        ], name="A.xlsx")
        self.write_source([["重复","韩语乙"]], name="B.xlsx")
        book = load_workbook(path)
        ignored = book.create_sheet("忽略")
        ignored.append(["台词","韩文"])
        ignored.append(["重复","不应算冲突"])
        ignored.append(["单一缺译",None])
        invalid = book.create_sheet("异常")
        invalid.append(["台词"])
        invalid.append(["重复","不应算冲突"])
        book.save(path)
        book.close()
        exclusions = {sheet_key(path, "忽略")}
        self.write_baseline(["重复","首尾一致","单一缺译","四条","未找到"])
        self.write_baseline(["重复"], name="char.xlsx")
        report = scan_folder(self.source, 1, ignored_sheets=exclusions)
        result = match_folder(report, self.baseline, self.output, ignored_sheets=exclusions)
        self.assertEqual((result.source_rows,result.source_matched,result.missing_translations,result.conflict_texts,result.match_pairs), (13,10,4,3,13))
        book = load_workbook(result.destination / "基准匹配结果/story.xlsx")
        sheet = book.active
        self.assertEqual([(sheet.cell(r,27).value,sheet.cell(r,28).value,sheet.cell(r,29).value) for r in range(2,7)],
                         [(3,1,"是"),(2,0,"否"),(1,1,"否"),(4,1,"是"),(0,0,"未匹配")])
        self.assertTrue(all(sheet.cell(5,c).value is None for c in range(18,27)))
        book.close()
        book = load_workbook(result.destination / "匹配明细.xlsx")
        matched = list(book["匹配明细"].values)[1:]
        self.assertEqual(len(matched),13)
        repeat = [r for r in matched if r[8]=="重复"]
        self.assertEqual(len(repeat),6)
        self.assertEqual({r[0] for r in repeat},{"story.xlsx","char.xlsx"})
        self.assertTrue(all(r[11]=="是" for r in repeat))
        self.assertFalse(any(r[6] in ("忽略","异常") for r in matched))
        unmatched = list(book["未匹配明细"].values)[1:]
        self.assertEqual(len(unmatched),3)
        self.assertEqual({r[4] for r in unmatched},{"不在基准"})
        self.assertEqual({r[3] for r in unmatched},{11,12,13})
        missing = list(book["缺译明细"].values)[1:]
        self.assertEqual(len(missing),4)
        self.assertEqual(sum(r[6]=="否" for r in missing),1)
        conflicts = list(book["译文冲突"].values)[1:]
        self.assertEqual({r[4] for r in conflicts},{"重复","四条","不在基准"})
        self.assertTrue(all(r[8]==2 for r in conflicts))
        self.assertEqual(len(conflicts),10)
        self.assertEqual(book["匹配明细"].freeze_panes,"A2")
        book.close()
        book = load_workbook(result.destination / "源文件匹配统计.xlsx")
        rows = list(book["文件统计"].values)[1:]
        self.assertEqual([(r[0],r[9],r[10]) for r in rows],[("A.xlsx",4,3),("B.xlsx",0,1)])
        book.close()

    def test_details_split_at_sheet_limit_without_losing_rows(self):
        book = Workbook(write_only=True)
        count = write_detail_sheet(book,"明细",["中文"],[("=原文",),("第二行",),("第三行",),("第四行",),("第五行",)],[30],row_limit=3)
        path = self.root / "分页.xlsx"
        book.save(path)
        book.close()
        self.assertEqual(count,5)
        book = load_workbook(path)
        self.assertEqual(book.sheetnames,["明细","明细_2","明细_3"])
        self.assertEqual([row[0] for sheet in book for row in list(sheet.values)[1:]], ["=原文","第二行","第三行","第四行","第五行"])
        self.assertEqual(book["明细"]["A2"].data_type,"s")
        self.assertTrue(all(s.max_row<=3 for s in book))
        book.close()

    def test_context_recommends_per_baseline_and_keeps_full_evidence(self):
        a = ["甲前一", "甲前二", "甲前三", "共同句", "甲后一", "甲后二", "甲后三"]
        b = ["乙前一", "乙前二", "乙前三", "共同句", "乙后一", "乙后二", "乙后三"]
        self.write_source([[t, "韩甲" if t == "共同句" else "译文"] for t in a], name="A.xlsx")
        self.write_source([[t, "韩乙" if t == "共同句" else "译文"] for t in b], name="B.xlsx")
        # One-sided candidate must lose to the candidate matching both sides.
        self.write_source([[t, "干扰译文" if t == "共同句" else "译文"] for t in a[:4] + b[4:]], name="C.xlsx")
        # Empty Korean never wins, even with the best context.
        self.write_source([[t, None if t == "共同句" else "译文"] for t in a], name="D.xlsx")
        self.write_baseline(a, name="A.xlsx")
        (self.baseline / "子目录").mkdir()
        self.write_baseline(b, name="子目录/B.xlsx")
        result = match_folder(scan_folder(self.source, 1), self.baseline, self.output)
        self.assertEqual((result.baseline_files, result.baseline_rows, result.matched_rows, result.recommended_rows), (2, 14, 14, 14))
        for name, expected in (("A.xlsx", "韩甲"), ("子目录/B.xlsx", "韩乙")):
            book = load_workbook(result.destination / "基准匹配结果" / name)
            self.assertEqual(book.active["AD5"].value, expected)
            self.assertEqual(book.active["AG5"].value, "上下文推荐")
            self.assertIn("前3/3句一致，后3/3句一致", book.active["AF5"].value)
            self.assertEqual(book.active["AC5"].value, "是")
            book.close()
        book = load_workbook(result.destination / "匹配明细.xlsx")
        evidence = [dict(zip(next(book["匹配明细"].values), row)) for row in list(book["匹配明细"].values)[1:] if row[8] == "共同句"]
        selected = [r for r in evidence if r["推荐来源"] == "是"]
        self.assertEqual(len(selected), 2)
        self.assertEqual({r["韩语"] for r in selected}, {"韩甲", "韩乙"})
        self.assertTrue(all(r["完整一致侧数"] == 2 and r["一致位置总数"] == 6 for r in selected))
        self.assertIn("前1句 [行4] 甲前三", selected[0]["基准前3句"])
        book.close()
        book = load_workbook(result.destination / "基准文件匹配统计.xlsx")
        for row in list(book["文件统计"].values)[1:]:
            self.assertEqual(row[2:11], (7, 7, 7, 7, 0, 0, 1, 1, 0))
        self.assertEqual(list(book["总体统计"].values)[1][2:11], (14, 14, 14, 14, 0, 0, 2, 2, 0))
        self.assertEqual(book["文件统计"]["L2"].number_format, "0.00%")
        book.close()

    def test_context_ties_boundaries_single_side_and_hidden_rows(self):
        texts = ["前1", "前2", "前3", "争议", "后1", "后2", "后3"]
        self.write_source([[t, "韩甲" if t == "争议" else "同译"] for t in texts], name="A.xlsx")
        self.write_source([[t, "韩乙" if t == "争议" else "同译"] for t in texts], name="B.xlsx")
        self.write_baseline(texts)
        result = match_folder(scan_folder(self.source, 1), self.baseline, self.output)
        book = load_workbook(result.destination / "基准匹配结果/story.xlsx")
        self.assertIsNone(book.active["AD5"].value)
        self.assertEqual(book.active["AG5"].value, "待确认")
        self.assertIn("最高上下文得分", book.active["AF5"].value)
        book.close()
        # At a boundary, even two identical neighbours are insufficient for a conflict.
        self.write_baseline(["争议", "后1", "后2"])
        result = match_folder(scan_folder(self.source, 1), self.baseline, self.output)
        book = load_workbook(result.destination / "基准匹配结果/story.xlsx")
        self.assertEqual(book.active["AG2"].value, "待确认")
        self.assertIn("完整前3句或完整后3句", book.active["AF2"].value)
        book.close()
        # Skip a hidden distractor; the next three nonempty records then match.
        path = self.write_source([["争议", "韩甲"], ["隐藏干扰", "不可参与"], ["后1", "同译"],
                                  [None, "空中文"], ["后2", "同译"], ["后3", "同译"]], name="A.xlsx")
        book = load_workbook(path)
        book.active.row_dimensions[3].hidden = True
        book.save(path)
        book.close()
        self.write_source([["争议", "韩乙"], ["不同1", "同译"], ["不同2", "同译"], ["不同3", "同译"]], name="B.xlsx")
        self.write_baseline(["争议", "后1", "后2", "后3"])
        result = match_folder(scan_folder(self.source, 1, include_hidden_rows=False), self.baseline, self.output)
        book = load_workbook(result.destination / "基准匹配结果/story.xlsx")
        self.assertEqual(book.active["AD2"].value, "韩甲")
        self.assertIn("前0/3句一致，后3/3句一致", book.active["AF2"].value)
        book.close()

    def test_many_baselines_summary_empty_missing_and_source_deduplication(self):
        self.write_source([["唯一", " 번역 "], ["唯一", "번역"], ["缺译", None]])
        for i in range(100):
            self.write_baseline(["唯一", "缺译", "无匹配"], name=f"{i:03}.xlsx")
        empty = self.write_baseline([], name="空.xlsx")
        result = match_folder(scan_folder(self.source, 1), self.baseline, self.output)
        self.assertEqual((result.baseline_files, result.baseline_rows, result.matched_rows, result.recommended_rows,
                          result.source_rows, result.source_matched, result.match_pairs), (101, 300, 200, 100, 3, 3, 300))
        book = load_workbook(result.destination / "基准文件匹配统计.xlsx")
        rows = {r[0]: r for r in list(book["文件统计"].values)[1:]}
        self.assertEqual(rows["000.xlsx"][2:11], (3, 2, 1, 1, 1, 1, 0, 0, 0))
        self.assertEqual(rows["空.xlsx"][2:11], (0,) * 9)
        self.assertEqual(rows["空.xlsx"][11:], (None, None, None))
        total = list(book["总体统计"].values)[1]
        self.assertEqual(total[:11], (101, 101, 300, 200, 100, 100, 100, 100, 0, 0, 0))
        self.assertAlmostEqual(total[11], 2/3)
        book.close()
        book = load_workbook(result.destination / "基准匹配结果/000.xlsx")
        self.assertEqual(book.active["AG2"].value, "唯一译文")
        self.assertEqual(book.active["AG3"].value, "全部缺译")
        self.assertEqual(book.active["AG4"].value, "未匹配")
        book.close()
        # New recommendation columns must also be protected from overwrite.
        book = load_workbook(empty)
        book.active["AG2"] = "已有判断"
        book.save(empty)
        book.close()
        with self.assertRaisesRegex(ValueError, "R～AG 列已有内容"):
            match_folder(scan_folder(self.source, 1), self.baseline, self.output)



if __name__ == "__main__":
    unittest.main()
