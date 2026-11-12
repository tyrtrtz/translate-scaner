"""Package-level regression checks for pure-backfill string storage."""

import hashlib
import sqlite3
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from backfill import write_copy
from matcher import match_folder
from scanner import scan_folder
from structure_scan import scan_structure


SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
TYPE_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
NS = {'s': SHEET_NS}


def shared_fixture(path, *, part='xl/sharedStrings.xml', counts=True, rich_text=False, extension=False):
    """Independently turn a small test package into the customer's SST format."""
    with ZipFile(path) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    sst = ET.Element(f'{{{SHEET_NS}}}sst')
    indexes, reference_count = {}, 0
    for name, data in list(entries.items()):
        if not name.startswith('xl/worksheets/') or not name.endswith('.xml'):
            continue
        root = ET.fromstring(data)
        for cell in root.findall('.//s:c[@t="inlineStr"]', NS):
            inline = cell.find('s:is', NS)
            if inline is None:
                continue
            item = ET.Element(f'{{{SHEET_NS}}}si')
            if rich_text and ''.join(inline.itertext()) == '富文本':
                item.extend(ET.fromstring(f'<si xmlns="{SHEET_NS}"><r><rPr><b/></rPr><t>富</t></r><r><t>文本</t></r></si>'))
            else:
                item.extend(list(inline))
            serialized = ET.tostring(item)
            if serialized not in indexes:
                indexes[serialized] = len(sst)
                sst.append(item)
            cell.remove(inline)
            cell.set('t', 's')
            ET.SubElement(cell, f'{{{SHEET_NS}}}v').text = str(indexes[serialized])
            reference_count += 1
        # ElementTree's explicit namespace prefix exercises prefixed worksheets.
        entries[name] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    if counts:
        sst.set('count', str(reference_count))
        sst.set('uniqueCount', str(len(sst)))
    if extension:
        tail = ET.SubElement(sst, f'{{{SHEET_NS}}}extLst')
        ET.SubElement(tail, f'{{{SHEET_NS}}}ext', {'uri': 'urn:customer:preserve'})
    entries[part] = ET.tostring(sst, encoding='utf-8', xml_declaration=True)
    types = ET.fromstring(entries['[Content_Types].xml'])
    ET.SubElement(types, f'{{{TYPE_NS}}}Override', {
        'PartName': '/' + part,
        'ContentType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'})
    entries['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8')
    relationships = ET.fromstring(entries['xl/_rels/workbook.xml.rels'])
    ET.SubElement(relationships, f'{{{REL_NS}}}Relationship', {
        'Id': 'rIdSharedFixture',
        'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
        'Target': part.removeprefix('xl/')})
    entries['xl/_rels/workbook.xml.rels'] = ET.tostring(relationships, encoding='utf-8')
    with ZipFile(path, 'w', compression=ZIP_DEFLATED) as archive:
        for name, data in entries.items():
            archive.writestr(name, data)


class BackfillStorageTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)

    def database(self, baseline, sheets, values):
        db = sqlite3.connect(':memory:')
        self.addCleanup(db.close)
        db.execute('CREATE TABLE fill_sheets (path TEXT,sheet TEXT,column_number INTEGER,header_row INTEGER,new_header TEXT)')
        db.execute('CREATE TABLE fill_values (path TEXT,sheet TEXT,row INTEGER,value TEXT)')
        db.executemany('INSERT INTO fill_sheets VALUES (?,?,?,?,?)', [(str(baseline), *row) for row in sheets])
        db.executemany('INSERT INTO fill_values VALUES (?,?,?,?)', [(str(baseline), *row) for row in values])
        return db

    def test_shared_strings_reuse_and_append_across_sheets_preserving_package(self):
        path = self.root / '原件.xlsx'
        book = Workbook()
        sheet = book.active
        sheet.title = '正文'
        for row in [('Text', 'TextTrans', '备注'), ('甲', '旧译文', '=1+1'), ('乙', '待替换', 7),
                    ('未匹配', '保留原译文', '富文本'), ('公式不回填', '=2+2', True),
                    ('富文本译文', '旧译文')]:
            sheet.append(row)
        sheet['B2'].font = Font(bold=True, color='FF112233')
        sheet['B2'].fill = PatternFill('solid', fgColor='FFFFFF00')
        sheet.column_dimensions['B'].width = 44
        extra = book.create_sheet('新增译文列')
        for row in [('Text',), ('丙',), ('丁',)]:
            extra.append(row)
        untouched = book.create_sheet('未处理')
        untouched.append(['原文', '富文本', '旧译文'])
        book.save(path); book.close()
        part = 'xl/stringData/customer-strings.xml'
        shared_fixture(path, part=part, rich_text=True, extension=True)
        before_hash = hashlib.sha256(path.read_bytes()).digest()
        value = ' 신규 & <color=red>번역</color>\r\n끝 _x0041_ '
        db = self.database(path, [('正文', 2, 1, ''), ('新增译文列', 2, 1, 'TextTrans')],
                           [('正文', 2, value), ('正文', 3, '旧译文'), ('正文', 6, '富文本'),
                            ('新增译文列', 2, value), ('新增译文列', 3, '=literal')])
        output = self.root / '纯回填.xlsx'
        parts = {'xl/worksheets/sheet1.xml': ('正文',), 'xl/worksheets/sheet2.xml': ('新增译文列',)}
        write_copy(path, output, db, parts)
        self.assertEqual(before_hash, hashlib.sha256(path.read_bytes()).digest())
        with ZipFile(path) as original, ZipFile(output) as changed:
            self.assertEqual(set(original.namelist()), set(changed.namelist()))
            original_sst = ET.fromstring(original.read(part))
            new_sst = ET.fromstring(changed.read(part))
            old_items = original_sst.findall('s:si', NS)
            new_items = new_sst.findall('s:si', NS)
            # Appending new strings must not change any existing index/rich text.
            self.assertEqual([ET.tostring(e) for e in old_items],
                             [ET.tostring(e) for e in new_items[:len(old_items)]])
            self.assertEqual(new_sst[-1].tag, f'{{{SHEET_NS}}}extLst')
            self.assertEqual(ET.tostring(original_sst[-1]), ET.tostring(new_sst[-1]))
            self.assertIn(b'_x005F_x0041_', changed.read(part))
            for name in original.namelist():
                if name not in (*parts, part):
                    self.assertEqual(original.read(name), changed.read(name), name)
            roots = [ET.fromstring(changed.read(f'xl/worksheets/sheet{i}.xml')) for i in range(1, 4)]
            self.assertFalse(any(root.findall('.//s:c[@t="inlineStr"]', NS) for root in roots))
            shared_cells = [cell for root in roots for cell in root.findall('.//s:c[@t="s"]', NS)]
            self.assertEqual(int(new_sst.get('count')), len(shared_cells))
            self.assertEqual(int(new_sst.get('uniqueCount')), len(new_items))
            self.assertTrue(all(0 <= int(c.find('s:v', NS).text) < len(new_items) for c in shared_cells))
            def index(root, coordinate):
                return int(root.find(f'.//s:c[@r="{coordinate}"]/s:v', NS).text)
            self.assertEqual(index(roots[0], 'B2'), index(roots[1], 'B2'))
            self.assertEqual(index(roots[0], 'B3'), index(ET.fromstring(original.read('xl/worksheets/sheet1.xml')), 'B2'))
            self.assertEqual(index(roots[1], 'B1'), index(roots[0], 'B1'))
            plain_rich_index = index(roots[0], 'B6')
            self.assertGreaterEqual(plain_rich_index, len(old_items))
            self.assertIsNone(new_items[plain_rich_index].find('s:r', NS))
            self.assertEqual(new_items[plain_rich_index].find('s:t', NS).text, '富文本')
            self.assertEqual(len(new_items), len(old_items) + 3)
        book = load_workbook(output)
        self.assertEqual(book['正文']['B2'].value, value)
        self.assertEqual(book['正文']['B3'].value, '旧译文')
        self.assertEqual(book['正文']['B4'].value, '保留原译文')
        self.assertEqual(book['正文']['B5'].value, '=2+2')
        self.assertEqual(book['正文']['B6'].value, '富文本')
        self.assertEqual(book['正文']['C2'].value, '=1+1')
        self.assertEqual(book['正文']['C4'].value, '富文本')
        self.assertTrue(book['正文']['B2'].font.bold)
        self.assertEqual(book['正文']['B2'].fill.fgColor.rgb, 'FFFFFF00')
        self.assertEqual(book['正文'].column_dimensions['B'].width, 44)
        self.assertEqual(book['新增译文列']['B1'].value, 'TextTrans')
        self.assertEqual(book['新增译文列']['B2'].value, value)
        self.assertEqual(book['新增译文列']['B3'].value, '=literal')
        self.assertEqual(book['新增译文列']['B3'].data_type, 's')
        book.close()

    def test_optional_sst_counts_and_original_inline_storage(self):
        for shared in (False, True):
            with self.subTest(shared=shared):
                path = self.root / f'{shared}.xlsx'
                book = Workbook(); book.active.title = '正文'
                book.active.append(['Text', 'TextTrans']); book.active.append(['甲', '原值'])
                book.save(path); book.close()
                if shared:
                    shared_fixture(path, counts=False)
                db = self.database(path, [('正文', 2, 1, '')], [('正文', 2, '번역')])
                output = self.root / f'output-{shared}.xlsx'
                write_copy(path, output, db, {'xl/worksheets/sheet1.xml': ('正文',)})
                with ZipFile(output) as archive:
                    root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
                    cell = root.find('.//s:c[@r="B2"]', NS)
                    self.assertEqual(cell.get('t'), 's' if shared else 'inlineStr')
                    if shared:
                        sst = ET.fromstring(archive.read('xl/sharedStrings.xml'))
                        self.assertEqual(('count' in sst.attrib), ('uniqueCount' in sst.attrib))
                    else:
                        self.assertNotIn('xl/sharedStrings.xml', archive.namelist())
                book = load_workbook(output)
                self.assertEqual(book.active['B2'].value, '번역'); book.close()

    def test_normal_and_structural_matching_keep_shared_string_backfills(self):
        source, baseline, output = [self.root / name for name in ('源', '基准', '结果')]
        source.mkdir(); baseline.mkdir()
        book = Workbook(); book.active.append(['台词', '한국어'])
        book.active.append(['完整中文', '완전한 번역'])
        book.active.append(['1、甲选项\n2、乙选项', '1. 첫째\n2. 둘째'])
        book.save(source / 'source.xlsx'); book.close()
        book = Workbook(); book.active.append(['Text', 'TextTrans'])
        for text in ('完整中文', '甲选项', '乙选项'):
            book.active.append([text, '旧译文'])
        path = baseline / 'story.xlsx'; book.save(path); book.close()
        shared_fixture(path)
        original_hash = hashlib.sha256(path.read_bytes()).digest()
        report = scan_folder(source)
        normal = match_folder(report, baseline, output)
        structural = scan_structure(report, baseline, output).match_report
        for result, expected in ((normal, ['완전한 번역', '旧译文', '旧译文']),
                                 (structural, ['완전한 번역', '첫째', '둘째'])):
            with self.subTest(result=result.destination):
                pure = result.destination / '纯回填结果/story.xlsx'
                with ZipFile(pure) as archive:
                    root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
                    self.assertFalse(root.findall('.//s:c[@t="inlineStr"]', NS))
                book = load_workbook(pure)
                self.assertEqual([book.active.cell(r, 2).value for r in (2, 3, 4)], expected)
                book.close()
        self.assertEqual(original_hash, hashlib.sha256(path.read_bytes()).digest())


if __name__ == '__main__':
    unittest.main()
