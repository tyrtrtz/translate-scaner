import hashlib
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from key_matching import business_key, content_key, metadata_columns, similarity, speaker_compare, option_parts
from matcher import match_folder
from scanner import scan_folder
from compare_backfills import compare_backfills


class CustomerRules(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.source, self.baseline, self.output = [self.root / s for s in ('源', '基准', '输出')]
        self.source.mkdir(); self.baseline.mkdir()

    def workbook(self, folder, rows, name='table.xlsx'):
        book = Workbook()
        book.active.title = '正文'
        for row in rows:
            book.active.append(row)
        path = folder / name
        book.save(path); book.close()
        return path

    def run_match(self):
        originals = {p: hashlib.sha256(p.read_bytes()).digest() for root in (self.source, self.baseline) for p in root.glob('*.xlsx')}
        result = match_folder(scan_folder(self.source), self.baseline, self.output)
        self.assertEqual(originals, {p: hashlib.sha256(p.read_bytes()).digest() for p in originals})
        return result

    def values(self, result, folder='基准匹配结果', name='table.xlsx'):
        book = load_workbook(result.destination / folder / name)
        try:
            return list(book.active.values)
        finally:
            book.close()

    def test_id_and_offset_resolve_conflicts_without_overwriting_legacy_success(self):
        self.workbook(self.source, [['台词', '한국어', '对话ID', '说话角色名'],
                                   ['嗯', '첫 번째', 99, '米砂'], ['嗯', '두 번째', 200, '米砂'],
                                   ['原来就能得到唯一译文', '旧成功结果', 999, '甲']])
        self.workbook(self.baseline, [['translationKey', 'chs', 'translate', 'speaker'],
            ['story_text_100_Dialog_Detail_v0', '嗯', '旧一', '米砂'],
            ['story_text_200_Dialog_Detail_v0', '嗯', '旧二', '米砂'],
            ['story_text_999_Dialog_Detail_v0', '原来就能得到唯一译文', '旧三', '乙']])
        result = self.run_match()
        rows = self.values(result)
        self.assertEqual([r[29] for r in rows[1:]], ['첫 번째', '두 번째', '旧成功结果'])
        self.assertEqual([r[32] for r in rows[1:]], ['键定位推荐·待复核', '键定位推荐·待复核', '唯一译文'])
        self.assertIn('ID ±5', rows[1][31])
        self.assertIn('v1.1.9规则优先', rows[3][31])
        self.assertEqual([r[26] for r in rows[1:]], [2, 2, 1])  # Old conflicting candidates remain.
        filled = self.values(result, '纯回填结果')
        self.assertEqual(len(filled[0]), 4)
        self.assertEqual([r[2] for r in filled[1:]], [r[29] for r in rows[1:]])
        self.assertEqual(result.backfilled_rows, 3)

    def test_plot_ids_and_role_conflicts(self):
        self.workbook(self.source, [['台词', '한국어', '剧情ID', '说话角色名'],
                                   ['好了', '첫째', 123, '甲'], ['好了', '둘째', None, '甲'],
                                   ['不能串角色', '错一', 124, '乙'], ['不能串角色', '错二', None, '乙']])
        self.workbook(self.baseline, [['translationKey', 'chs', 'translate', 'speaker'],
            ['story_convenient_12301_Dialog_Detail_v0', '好了', '原一', '甲'],
            ['story_convenient_12302_Dialog_Detail_v0', '好了', '原二', '甲'],
            ['story_convenient_12401_Dialog_Detail_v0', '不能串角色', '保留', '甲']])
        result = self.run_match()
        rows = self.values(result)
        self.assertEqual([r[29] for r in rows[1:]], ['첫째', '둘째', None])
        self.assertIn('剧情ID', rows[1][31])
        self.assertEqual(self.values(result, '纯回填结果')[3][2], '保留')

    def test_comment_business_keys_drive_id_matching_logs_and_customer_comparison(self):
        self.workbook(self.source, [['台词', '한국어', '对话ID', '说话角色名'],
            ['嗯', '첫 번째', 100, '甲'], ['嗯', '두 번째', 200, '甲']])
        headers = ['TransComment', 'Key', 'GameKey', 'Ver', 'Status', 'Category', 'Feature',
                   'Text', 'TransText', 'Comment', 'CategoryDesc', 'Speaker', 'SkillName',
                   'AffiliationPrefab', 'AffiliatioinPrefabPath', 'TextTrans']
        keys = [f'story_text_{n}_Dialog_Detail_v0' for n in (100, 200)]
        self.workbook(self.baseline, [headers,
            ['', 701, 'story_900000001', '', '', '', '', '嗯', '参考原译1', keys[0], '', '甲', '', '', '', '旧一'],
            ['', 702, 'story_900000002', '', '', '', '', '嗯', '参考原译2', keys[1], '', '甲', '', '', '', '旧二']])
        result = self.run_match()
        rows = self.values(result)
        self.assertEqual([r[29] for r in rows[1:]], ['첫 번째', '두 번째'])
        self.assertTrue(all(r[32] == '键定位推荐·待复核' and '对话ID与中文一致' in r[31] for r in rows[1:]))
        pure = self.values(result, '纯回填结果')
        self.assertEqual([r[15] for r in pure[1:]], ['첫 번째', '두 번째'])
        self.assertEqual([r[8] for r in pure[1:]], ['参考原译1', '参考原译2'])
        book = load_workbook(result.destination / '回填日志.xlsx')
        self.assertEqual([book.active.cell(i, 4).value for i in (2, 3)], keys)
        book.close()

        customer = self.root / '客户'; customer.mkdir()
        customer_header = list(headers)
        customer_header[15] = 'html回填结果'
        # Different filename, order and numeric record IDs: only Comment is stable.
        customer_rows = [list(pure[2]), list(pure[1])]
        for i, r in enumerate(customer_rows):
            r[1] = 900 + i
        exported = self.workbook(customer, [customer_header, *customer_rows], name='客户导出.xlsx')
        _, counts = compare_backfills(result.destination, exported)
        self.assertEqual(counts, {'一致可定稿': 2})

        # v1.1.16 logs used numeric Key; keep valid old logs usable after upgrading.
        log = result.destination / '回填日志.xlsx'
        book = load_workbook(log); book.active['D2'] = '701'; book.save(log); book.close()
        _, counts = compare_backfills(result.destination, exported)
        self.assertEqual(counts, {'一致可定稿': 2})
        book = load_workbook(log); book.active['D2'] = 'unrelated-key'; book.save(log); book.close()
        _, counts = compare_backfills(result.destination, exported)
        self.assertEqual(counts, {'待确认': 1, '一致可定稿': 1})

        # A named customer result also wins over an original TextTrans column.
        book = load_workbook(exported); book.active['Q1'] = 'TextTrans'; book.active['Q2'] = '旧值'; book.save(exported); book.close()
        _, counts = compare_backfills(result.destination, exported)
        self.assertEqual(counts, {'待确认': 1, '一致可定稿': 1})
        book = load_workbook(exported); book.active['Q1'] = 'html回填结果'; book.save(exported); book.close()
        with self.assertRaisesRegex(ValueError, '表头缺失／不唯一'):
            compare_backfills(result.destination, exported)

    def test_comment_is_only_used_for_a_recognizable_business_key(self):
        book = Workbook(); self.addCleanup(book.close)
        sheet = book.active
        sheet.append(['Key', 'Comment'])
        key = 'story_text_100_Dialog_Detail_v0'
        for numeric, comment, expected in [
            (12, key, key), (12.0, key, key), (None, key, key),
            (12, '请参考旧台本', '12'), (12, 'story_900000001', '12'),
            (12, key + ' 请确认', '12'), (12, '=A1', '12'),
            ('explicit-key', key, 'explicit-key'),
            (key, 'story_text_200_Dialog_Detail_v0', key),
        ]:
            with self.subTest(numeric=numeric, comment=comment):
                sheet.append([numeric, comment])
                row = next(sheet.iter_rows(min_row=sheet.max_row, max_row=sheet.max_row))
                self.assertEqual(business_key(row, metadata_columns(sheet[1])), expected)
        sheet['C1'] = 'Comment'; sheet['C2'] = key
        self.assertEqual(business_key(sheet[2], metadata_columns(sheet[1])), '12')

    def test_zero_based_options_keep_segment_order_and_reject_gaps_or_duplicates(self):
        self.workbook(self.source, [['台词', '한국어', '对话ID', '说话角色名'],
            ['这是选项前的一句话', '앞 문장', 100, '玩家'],
            ['1. 点头。\n2. 摇头。', '1. 끄덕인다\n2. 젓는다', None, '玩家']])
        header = ['Key', 'Text', 'TextTrans', 'Comment', 'Speaker']
        def baseline(numbers):
            self.workbook(self.baseline, [header,
                [1, '这是选项前的一句话', '旧', 'story_text_100_Dialog_Detail_v0', '{playerName}'],
                [2, '点头。', '旧一', f'story_text_100_Option{numbers[0]}_v0', '{playerName}'],
                [3, '摇头。', '旧二', f'story_text_100_Option{numbers[1]}_v0', '{playerName}']])
        baseline((0, 1))
        result = self.run_match()
        rows = self.values(result)
        self.assertEqual([r[29] for r in rows[1:]], ['앞 문장', '끄덕인다', '젓는다'])
        self.assertTrue(all(r[32] == '选项推荐·待复核' and 'Option0～Option1' in r[31] for r in rows[2:]))
        self.assertIn('选项1', rows[2][30])  # First source segment stays part 1, not whole-line part 0.
        self.assertIn('选项2', rows[3][30])
        self.assertEqual([r[2] for r in self.values(result, '纯回填结果')[1:]], ['앞 문장', '끄덕인다', '젓는다'])
        book = load_workbook(result.destination / '匹配明细.xlsx')
        self.assertEqual(list(book['未推荐明细'].values)[1:], [])
        book.close()
        for numbers in ((0, 2), (0, 0), (2, 3)):
            with self.subTest(numbers=numbers):
                baseline(numbers)
                result = self.run_match()
                self.assertTrue(all(r[29] is None for r in self.values(result)[2:]))
                self.assertEqual([r[2] for r in self.values(result, '纯回填结果')[2:]], ['旧一', '旧二'])
        baseline((0, 1))
        source = self.source / 'table.xlsx'
        book = load_workbook(source); book.active['B3'] = '1. 끄덕인다'; book.save(source); book.close()
        result = self.run_match()
        self.assertTrue(all(r[29] is None for r in self.values(result)[2:]))
        self.assertEqual([r[2] for r in self.values(result, '纯回填结果')[2:]], ['旧一', '旧二'])

    def test_option_keys_write_individual_translations_and_log_full_source_coverage(self):
        self.workbook(self.source, [['台词', '한국어', '对话ID', '说话角色名'],
                                   ['这是选项前的一句话', '앞 문장', 100, '玩家'],
                                   ['1：你是谁\n2；我在哪', '1. 누구야\n2. 어디야', None, '玩家']])
        self.workbook(self.baseline, [['translationKey', 'chs', 'translate', 'speaker'],
            ['story_text_100_Dialog_Detail_v0', '这是选项前的一句话', '', '{playerName}'],
            ['story_text_100_Option1_v0', '你是谁', '旧一', '{playerName}'],
            ['story_text_100_Option2_v0', '我在哪', '旧二', '{playerName}']])
        result = self.run_match()
        rows = self.values(result)
        self.assertEqual([r[29] for r in rows[1:]], ['앞 문장', '누구야', '어디야'])
        self.assertEqual([r[32] for r in rows[2:]], ['选项推荐·待复核'] * 2)
        self.assertEqual([r[2] for r in self.values(result, '纯回填结果')[1:]], ['앞 문장', '누구야', '어디야'])
        book = load_workbook(result.destination / '匹配明细.xlsx')
        self.assertEqual(list(book['未推荐明细'].values)[1:], [])
        book.close()

    def test_ordered_short_lines_and_same_key_translation_conflicts(self):
        self.workbook(self.source, [['台词', '한국어', '对话ID', '说话角色名'],
            ['开场！', '시작', 100, '甲'], ['嗯', '앞 음', None, '甲'], ['好的', '좋아', None, '甲'],
            ['嗯', '뒤 음', None, '甲'], ['收场！', '끝', 104, '甲'],
            ['重复的冲突句', '충돌1', 500, '甲'], ['重复的冲突句', '충돌2', 500, '甲']])
        self.workbook(self.baseline, [['translationKey', 'chs', 'translate', 'speaker'],
            *[[f'story_text_{n}_Dialog_Detail_v0', z, '原值', '甲'] for n, z in
              [(100, '开场。'), (101, '嗯'), (102, '好的'), (103, '嗯'), (104, '收场。'),
               (500, '重复的冲突句'), (600, '重复的冲突句')]]])
        result = self.run_match()
        rows = self.values(result)
        self.assertEqual(rows[2][29], '앞 음')
        self.assertEqual(rows[4][29], '뒤 음')
        self.assertIn('顺序', rows[2][31])
        self.assertIsNone(rows[6][29]); self.assertIsNone(rows[7][29])

    def test_full_text_normalization_and_global_similarity(self):
        target = '她在树林深处遇到了一位正在寻找朋友的小妖精'
        self.workbook(self.source, [['台词', '한국어'],
            [' '.join('这是诶呀一段特别的原文') + '（轻声）？！', '본문'],
            [' '.join(target.replace('她', '他')), '비슷한 문장']])
        self.workbook(self.baseline, [['Text', 'TextTrans'], ['这是欸呀一段特别的原文', '旧文'], [target, '旧译']])
        result = self.run_match()
        rows = self.values(result)
        self.assertEqual([r[29] for r in rows[1:]], ['본문', '비슷한 문장'])
        self.assertTrue(all(r[32] == '键定位推荐·待复核' for r in rows[1:]))
        self.assertIn('全表相似原文', rows[2][31])
        self.assertEqual(content_key('{playerName}<color=red>诶！</color>（轻声）', parentheses=True), '欸')
        self.assertEqual(content_key('1<2>3\\r\\n'), '123')
        self.assertEqual(speaker_compare('？？？（米砂）', '米砂'), 1)
        self.assertEqual(speaker_compare('村长', '村长夫人'), -1)
        self.assertEqual(option_parts('1；甲\n2：乙'), ['甲', '乙'])
        self.assertEqual(option_parts('1；甲\n3：乙'), [])
        self.assertLess(similarity('甲' * 400 + '乙' * 100, '甲' * 400 + '丙' * 100), .9)

    def test_pure_backfill_preserves_styles_formulas_blank_and_original_package_parts(self):
        self.workbook(self.source, [['台词', '한국어'], ['替换句', '=literal'], ['公式句', '新译文'], ['缺译句', None]])
        book = load_workbook(self.source / 'table.xlsx'); book.active['B2'].data_type = 's'; book.save(self.source / 'table.xlsx'); book.close()
        baseline = self.workbook(self.baseline, [['Key', 'TextTrans', 'Text', '其它'],
            ['k1', '旧译文', '替换句', '=1+1'], ['k2', '=2+2', '公式句', '不动'],
            ['k3', '保留缺译值', '缺译句'], ['k4', '保留未匹配值', '完全不相关']])
        book = load_workbook(baseline)
        book.active['B2'].font = Font(bold=True, color='FF112233')
        book.active['B2'].fill = PatternFill('solid', fgColor='FFFFFF00')
        book.active.column_dimensions['B'].width = 44
        book.save(baseline); book.close()
        result = self.run_match()
        pure = result.destination / '纯回填结果/table.xlsx'
        book = load_workbook(pure)
        self.assertEqual(book.active['B2'].value, '=literal'); self.assertEqual(book.active['B2'].data_type, 's')
        self.assertEqual(book.active['B3'].value, '=2+2'); self.assertEqual(book.active['B3'].data_type, 'f')
        self.assertEqual(book.active['B4'].value, '保留缺译值')
        self.assertEqual(book.active['B5'].value, '保留未匹配值')
        self.assertEqual(book.active['D2'].value, '=1+1')
        self.assertTrue(book.active['B2'].font.bold)
        self.assertEqual(book.active.column_dimensions['B'].width, 44)
        self.assertEqual(book.active.max_column, 4)
        book.close()
        with ZipFile(baseline) as original, ZipFile(pure) as changed:
            for name in original.namelist():
                if not name.startswith('xl/worksheets/sheet'):
                    self.assertEqual(original.read(name), changed.read(name))
        self.assertEqual(result.backfilled_rows, 1)
        book = load_workbook(result.destination / '回填日志.xlsx')
        self.assertIn('公式', book.active['H3'].value)
        self.assertEqual(book.active['H4'].value, '未推荐，保留原译文')
        book.close()

    def test_missing_translation_header_is_added_but_unlabelled_data_is_protected(self):
        self.workbook(self.source, [['台词', '한국어'], ['中文', '번역']])
        baseline = self.workbook(self.baseline, [['Text'], ['中文']])
        result = self.run_match()
        self.assertEqual(self.values(result, '纯回填结果'), [('Text', 'TextTrans'), ('中文', '번역')])
        book = load_workbook(baseline); book.active['B2'] = '未命名原值'; book.save(baseline); book.close()
        before = set(self.output.iterdir())
        with self.assertRaisesRegex(ValueError, '已有内容'):
            self.run_match()
        self.assertEqual(set(self.output.iterdir()), before)

    def test_character_file_cannot_supplement_other_characters_voice_keys(self):
        self.workbook(self.source, [['台词', '한국어'], [' '.join('只属于米砂的一句长台词'), '미사']], name='米砂_角色文档.xlsx')
        self.workbook(self.baseline, [['translationKey', 'chs', 'translate', 'speaker'],
            ['char_voice_words_idle_100001_1', '只属于米砂的一句长台词', '', '米砂'],
            ['char_voice_words_idle_100002_1', '只属于米砂的一句长台词', '保留其他角色', '别的角色']])
        result = self.run_match()
        rows = self.values(result)
        self.assertEqual(rows[1][29], '미사')
        self.assertIsNone(rows[2][29])
        self.assertEqual(self.values(result, '纯回填结果')[2][2], '保留其他角色')

    def test_compare_exact_difference_unchanged_and_missing(self):
        self.workbook(self.source, [['台词', '한국어'], ['第一句', '새 번역'], ['第二句', '<b>둘</b>'], ['第三句', '同原值']])
        self.workbook(self.baseline, [['translationKey', 'chs', 'translate'],
            ['k1', '第一句', '旧一'], ['k2', '第二句', '旧二'], ['k3', '第三句', '同原值'],
            ['k4', '未匹配原文', '保留'], ['k5', '客户漏掉句', '原译']])
        result = self.run_match()
        customer = self.root / '客户'; customer.mkdir()
        exported = self.workbook(customer, [['translationKey', 'chs', 'translate'],
            ['k3', '第三句', '同原值'], ['k1', '第一句', '새 번역'], ['k2', '第二句', '둘'],
            ['k4', '未匹配原文', '保留'], ['extra', '客户额外句', '새로']], name='回填后.xlsx')
        before = exported.read_bytes()
        path, counts = compare_backfills(result.destination, customer)
        self.assertEqual(counts, {'一致可定稿': 1, '译文不一致': 1, '待确认': 2, '无法对应': 1})
        self.assertEqual(exported.read_bytes(), before)
        book = load_workbook(path)
        self.assertEqual(book['一致可定稿']['F2'].value, 'k1')
        self.assertIn('标签', book['译文不一致']['B2'].value)
        self.assertEqual(book['客户未对应']['D2'].value, 'extra')
        book.close()
        with self.assertRaisesRegex(ValueError, '自身对比'):
            compare_backfills(result.destination, result.destination / '纯回填结果')

    def test_compare_duplicate_keys_changed_original_and_formula_never_finalize(self):
        self.workbook(self.source, [['台词', '한국어'], ['句一', '하나'], ['句二', '둘'], ['句三', '셋']])
        self.workbook(self.baseline, [['translationKey', 'chs', 'translate'],
            ['k1', '句一', '旧一'], ['k2', '句二', '旧二'], ['k3', '句三', '旧三']])
        result = self.run_match()
        customer = self.root / '客户'; customer.mkdir()
        self.workbook(customer, [['translationKey', 'chs', 'translate'],
            ['k1', '句一', '하나'], ['k1', '句一', '하나'], ['k2', '改过的原文', '둘'],
            ['k3', '句三', '=1+1']], name='回填后.xlsx')
        path, counts = compare_backfills(result.destination, customer)
        self.assertEqual(counts, {'无法对应': 2, '待确认': 1})
        book = load_workbook(path)
        self.assertIn('不唯一', book['无法对应']['B2'].value)
        self.assertIn('原文不同', book['无法对应']['B3'].value)
        self.assertIn('公式', book['待确认']['B2'].value)
        book.close()

    def test_compare_no_key_uses_position_and_checks_changed_local_copy(self):
        self.workbook(self.source, [['台词', '한국어'], ['中文', '한국어']])
        self.workbook(self.baseline, [['Text', 'TextTrans'], ['中文', '原译']])
        result = self.run_match()
        customer = self.root / '客户'; customer.mkdir()
        exported = self.workbook(customer, [['Text', 'TextTrans'], ['中文', '한국어']])
        _, counts = compare_backfills(result.destination / '纯回填结果', exported)
        self.assertEqual(counts, {'一致可定稿': 1})
        pure = result.destination / '纯回填结果/table.xlsx'
        for path in (pure, exported):
            book = load_workbook(path); book.active['B2'] = '手工改后'; book.save(path); book.close()
        _, counts = compare_backfills(result.destination, exported)
        self.assertEqual(counts, {'待确认': 1})


if __name__ == '__main__':
    unittest.main()
