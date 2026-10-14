"""build_report 的測試（不需要 GUI）。執行：python -m unittest discover -s tests"""
import os
import tempfile
import unittest

from PIL import Image
from docx import Document

from _common import TEMPLATE
import report_generator_gui as g


def unique_cells(doc):
    for t in doc.tables:
        for r in t.rows:
            seen = {}
            for c in r.cells:
                seen[id(c._tc)] = c
            yield from seen.values()


class BuildReportTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        d = cls.tmp.name
        cls.p1 = os.path.join(d, 'a.jpg')
        Image.new('RGB', (400, 200), 'red').save(cls.p1)
        cls.p2 = os.path.join(d, 'b.jpg')
        im = Image.new('RGB', (400, 200), 'blue')
        ex = im.getexif()
        ex[0x0112] = 6
        im.save(cls.p2, exif=ex)
        cls.p3 = os.path.join(d, 'c.png')
        Image.new('RGBA', (100, 100), 'green').save(cls.p3)
        cls.p4 = os.path.join(d, 'bad.jpg')
        with open(cls.p4, 'wb') as f:
            f.write(b'not an image')

        data = {k: f'內容_{k}\n● 第二行_{k}' for k in g.FIELD_NAMES}
        data.update(year='115', month='10', day='12', personnel='王小明', start_year='115',
                    start_month='10', start_day='10', end_year='115', end_month='10', end_day='11',
                    location='排雲山莊')
        photos = {c: [] for c in g.PHOTO_CATEGORIES}
        photos['道路、步道狀況及維護情形'] = [cls.p1, cls.p2, cls.p3]
        photos['其他'] = [cls.p4]
        caps = {'道路、步道狀況及維護情形': {cls.p1: '自訂說明A'}}
        out = os.path.join(d, 'out.docx')
        cls.warnings = g.build_report(TEMPLATE, data, photos, caps, out)
        cls.doc = Document(out)
        cls.text = '\n'.join(c.text for c in unique_cells(cls.doc))

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_each_field_filled_exactly_once(self):
        for key in g.FIELD_NAMES:
            with self.subTest(key=key):
                self.assertEqual(self.text.count(f'內容_{key}'), 1)

    def test_dates(self):
        self.assertIn('115年10月12日', self.text)
        self.assertIn('115年10月10日至115年10月11日', self.text)

    def test_no_double_bullet(self):
        self.assertNotIn('● ●', self.text)

    def test_captions(self):
        self.assertIn('自訂說明A', self.text)
        self.assertNotIn('[照片說明]', self.text)  # 預設說明字樣現在是空白

    def test_images_and_exif(self):
        self.assertEqual(len(self.doc.inline_shapes), 3)
        # 3 張照片中只有 EXIF 旋轉過的 b.jpg 是直式（橫式排前、直式排後，所以是最後一張）
        portrait = [s for s in self.doc.inline_shapes if s.height > s.width]
        self.assertEqual(len(portrait), 1)
        self.assertIs(portrait[0]._inline, self.doc.inline_shapes[2]._inline)

    def test_bad_image_warns_not_crashes(self):
        self.assertTrue(any('bad.jpg' in w for w in self.warnings))

    def test_single_photo_heading(self):
        n = sum('九、服勤照片' in p.text for p in self.doc.paragraphs) + self.text.count('九、服勤照片')
        self.assertEqual(n, 1)

    def test_missing_field_warns(self):
        with tempfile.TemporaryDirectory() as d:
            tpl = os.path.join(d, 't.docx')
            doc = Document(TEMPLATE)
            for c in unique_cells(doc):
                for p in c.paragraphs:
                    for r in p.runs:
                        r.text = r.text.replace('六、登山遊客反應意見：', '')
            doc.save(tpl)
            data = {k: '' for k in g.FIELD_NAMES}
            data.update(year='1', month='1', day='1', start_year='1', start_month='1', start_day='1',
                        end_year='1', end_month='1', end_day='1', personnel='', location='', feedback='x')
            warns = g.build_report(tpl, data, {}, {}, os.path.join(d, 'o.docx'))
            self.assertTrue(any('六、登山遊客反應意見' in w for w in warns))


class PrepareImageTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def make(self, name, size, exif_orientation=None, mode='RGB'):
        path = os.path.join(self.tmp.name, name)
        im = Image.effect_noise(size, 80).convert(mode)  # 雜訊圖：不易壓縮，能反映檔案大小差異
        if exif_orientation:
            ex = im.getexif()
            ex[0x0112] = exif_orientation
            im.save(path, exif=ex)
        else:
            im.save(path)
        return path

    def test_small_image_returned_untouched(self):
        p = self.make('s.jpg', (800, 600))
        src, size = g.prepare_image(p)
        self.assertEqual(src, p)
        self.assertEqual(size, (800, 600))

    def test_large_image_downscaled_keeps_aspect(self):
        p = self.make('l.jpg', (4000, 3000))
        out, size = g.prepare_image(p)
        self.assertEqual(size, (g.MAX_IMAGE_WIDTH_PX, round(3000 * g.MAX_IMAGE_WIDTH_PX / 4000)))
        with Image.open(out) as im:
            self.assertEqual(im.width, g.MAX_IMAGE_WIDTH_PX)
            self.assertEqual(im.height, round(3000 * g.MAX_IMAGE_WIDTH_PX / 4000))
        out.seek(0, 2)
        self.assertLess(out.tell(), os.path.getsize(p))

    def test_large_rotated_image_is_rotated_then_scaled(self):
        p = self.make('r.jpg', (4000, 3000), exif_orientation=6)  # 轉正後 3000x4000（直式）
        with Image.open(g.prepare_image(p)[0]) as im:
            self.assertEqual(im.width, g.MAX_IMAGE_WIDTH_PX)
            self.assertGreater(im.height, im.width)

    def test_large_png_with_alpha_stays_png(self):
        p = os.path.join(self.tmp.name, 'a.png')
        Image.new('RGBA', (3000, 1000), (0, 255, 0, 128)).save(p)
        with Image.open(g.prepare_image(p)[0]) as im:
            self.assertEqual(im.format, 'PNG')
            self.assertEqual(im.width, g.MAX_IMAGE_WIDTH_PX)

    def test_docx_size_much_smaller_for_large_photos(self):
        p = self.make('big.jpg', (4000, 3000))
        data = {k: '' for k in g.FIELD_NAMES}
        data.update(year='1', month='1', day='1', start_year='1', start_month='1', start_day='1',
                    end_year='1', end_month='1', end_day='1', personnel='', location='')
        photos = {c: [] for c in g.PHOTO_CATEGORIES}
        photos['其他'] = [p, p]
        out = os.path.join(self.tmp.name, 'o.docx')
        g.build_report(TEMPLATE, data, photos, {}, out)
        self.assertLess(os.path.getsize(out), os.path.getsize(p))


def make_item(w, h, caption='c', fit=3.4, max_h=4.7):
    item = g.PhotoItem(path='x', caption=caption, px_w=w, px_h=h)
    g.size_photo(item, fit, max_h)
    return item


class LayoutTest(unittest.TestCase):
    LINE_CHARS = 20
    PAGE_H = 10.4

    def test_landscape_fits_cell_width(self):
        item = make_item(4000, 3000)
        self.assertAlmostEqual(item.img_w, 3.4)
        self.assertAlmostEqual(item.img_h, 3.4 * 0.75)

    def test_portrait_fits_width_but_is_capped_by_height(self):
        normal = make_item(3000, 4000)  # 3:4 -> 高 4.53 < 4.7
        self.assertAlmostEqual(normal.img_w, 3.4)
        tall = make_item(1000, 3000)  # 1:3 -> 以高度上限縮小
        self.assertAlmostEqual(tall.img_h, 4.7)
        self.assertAlmostEqual(tall.img_w, 4.7 / 3)

    def test_groups_by_orientation_two_per_row(self):
        items = [make_item(4, 3), make_item(3, 4), make_item(4, 3), make_item(3, 4), make_item(4, 3)]
        plan = [r for _, r in g.plan_layout({'a': items}, self.PAGE_H * 2, 0, self.LINE_CHARS)]
        for row in plan:
            self.assertLessEqual(len(row), 2)
            self.assertEqual(len({p.landscape for p in row}), 1)  # 同一列同方向
        kinds = [row[0].landscape for row in plan]
        self.assertEqual(kinds, sorted(kinds, reverse=True))  # 橫式在前、直式在後
        self.assertEqual(sum(len(r) for r in plan), 5)

    def test_categories_are_never_interleaved(self):
        # a 的第 3 列在 12 吋頁高放不下，換頁後才排；就算這頁剩下的空間放得下 b 的列，也不能插進去
        a = [make_item(3, 4) for _ in range(6)]
        b = [make_item(2, 1) for _ in range(2)]
        seq = g.plan_layout({'a': a, 'b': b}, 12.0, 0, self.LINE_CHARS)
        self.assertEqual([c for c, _ in seq], ['a', 'a', 'a', 'b'])

    def test_each_category_is_contiguous_for_many_shapes(self):
        import random
        rnd = random.Random(7)
        for _ in range(40):
            cats = {name: [make_item(*rnd.choice([(4, 3), (3, 4), (2, 1), (1, 3)]), caption='c' * rnd.randint(0, 40))
                           for _ in range(rnd.randint(0, 7))] for name in 'abcde'}
            seq = g.plan_layout(cats, rnd.choice([9.0, 10.4, 12.0]), rnd.choice([0, 0.45]), self.LINE_CHARS)
            order = [c for c, _ in seq]
            seen = []
            for c in order:
                if not seen or seen[-1] != c:
                    seen.append(c)
            self.assertEqual(seen, sorted(set(order)), order)  # 每個分類只出現一段，且依原順序
            self.assertEqual(sum(len(r) for _, r in seq), sum(len(v) for v in cats.values()))

    def test_gap_fill_still_works_inside_one_category(self):
        # 同一分類內：直式列放不下的空隙，可由後面的橫式列補上
        port = [make_item(3, 4) for _ in range(6)]   # 3 列直式，每列約 4.8
        land = [make_item(2, 1) for _ in range(2)]   # 1 列扁的橫式列
        seq = g.plan_layout({'a': port + land}, 12.0, 0, self.LINE_CHARS)
        rows = [r for _, r in seq]
        self.assertEqual([r[0].landscape for r in rows], [True, False, False, False])  # 橫式在前（分類內分組）
        self.assertEqual(len(rows), 4)

    def test_mixed_row_for_leftover_landscape_and_portrait(self):
        items = ([make_item(4, 3) for _ in range(3)] + [make_item(3, 4) for _ in range(3)])
        rows = [r for _, r in g.plan_layout({'a': items}, self.PAGE_H * 3, 0, self.LINE_CHARS)]
        self.assertEqual([len(r) for r in rows], [2, 2, 2])
        mixed = [r for r in rows if len({p.landscape for p in r}) == 2]
        self.assertEqual(len(mixed), 1)
        self.assertTrue(mixed[0][0].landscape and not mixed[0][1].landscape)

    def test_no_mixed_row_when_only_one_side_is_odd(self):
        items = ([make_item(4, 3) for _ in range(3)] + [make_item(3, 4) for _ in range(2)])
        rows = [r for _, r in g.plan_layout({'a': items}, self.PAGE_H * 3, 0, self.LINE_CHARS)]
        self.assertEqual([len(r) for r in rows], [2, 1, 2])
        self.assertTrue(all(len({p.landscape for p in r}) == 1 for r in rows))

    def test_shrink_to_fill_page_not_below_min(self):
        # 4 列橫式約需 4*(2.55+0.32)+0.32 = 11.8 > 10.4，縮到 80% 可放入 -> 全部縮小且總高不超過頁高
        items = [make_item(4, 3) for _ in range(8)]
        plan = [r for _, r in g.plan_layout({'a': items}, 10.4, 0, self.LINE_CHARS)]
        self.assertEqual(len(plan), 4)
        scale = items[0].img_w / 3.4
        self.assertLess(scale, 1.0)
        self.assertGreaterEqual(scale, g.MIN_PHOTO_SCALE - 1e-9)
        total = g.HEADER_H + sum(g.row_height(r, self.LINE_CHARS) for r in plan)
        self.assertLessEqual(total, 10.4 + 1e-6)

    def test_no_shrink_when_it_cannot_help(self):
        items = [make_item(4, 3) for _ in range(10)]  # 5 列：縮 80% 也放不下第 5 列 -> 換頁
        plan = [r for _, r in g.plan_layout({'a': items}, 10.4, 0, self.LINE_CHARS)]
        self.assertEqual(len(plan), 5)
        self.assertTrue(all(p.img_w >= 3.4 * g.MIN_PHOTO_SCALE - 1e-9 for p in items))

    def test_caption_height_grows_with_text(self):
        self.assertLess(g.caption_height('短', 20), g.caption_height('長' * 45, 20))


class MoveItemTest(unittest.TestCase):
    def test_move_down_up_and_same(self):
        seq = ['a', 'b', 'c', 'd']
        g.move_item(seq, 0, 2)
        self.assertEqual(seq, ['b', 'c', 'a', 'd'])
        g.move_item(seq, 3, 0)
        self.assertEqual(seq, ['d', 'b', 'c', 'a'])
        g.move_item(seq, 1, 1)
        self.assertEqual(seq, ['d', 'b', 'c', 'a'])


class OrderKeptTest(unittest.TestCase):
    def test_user_order_is_kept_within_orientation_group(self):
        items = [make_item(4, 3, caption=f'L{i}') for i in range(4)]
        rows = [r for _, r in g.plan_layout({'a': items}, 30, 0, 20)]
        self.assertEqual([[p.caption for p in r] for r in rows], [['L0', 'L1'], ['L2', 'L3']])
        items.reverse()  # 使用者把順序調換 -> 版面跟著變
        rows = [r for _, r in g.plan_layout({'a': items}, 30, 0, 20)]
        self.assertEqual([[p.caption for p in r] for r in rows], [['L3', 'L2'], ['L1', 'L0']])


class DocumentStructureTest(BuildReportTest):
    """沿用 BuildReportTest 產生的文件，檢查模板結構被正確整理。"""

    def test_heading_page_break_and_photos_follow_heading(self):
        body = list(self.doc.element.body)
        heading = next(e for e in body if e.tag.endswith('}p') and '九、服勤照片' in ''.join(t.text or '' for t in e.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')))
        idx = body.index(heading)
        self.assertTrue(heading.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pageBreakBefore') is not None)
        self.assertTrue(body[idx + 1].tag.endswith('}tbl'))  # 照片表格緊接在標題後
        self.assertTrue(body[idx - 1].tag.endswith('}tbl'))  # 標題前的空白填充段落已清掉

    def test_no_long_trailing_empty_paragraphs(self):
        body = [e for e in self.doc.element.body if not e.tag.endswith('}sectPr')]
        trailing = 0
        for e in reversed(body):
            if e.tag.endswith('}p') and not any((t.text or '').strip() for t in e.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')):
                trailing += 1
            else:
                break
        self.assertLessEqual(trailing, 1)

    def test_rows_do_not_split_across_pages(self):
        w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        photo_table = self.doc.tables[-1]
        self.assertTrue(all(r._tr.find(f'.//{w}cantSplit') is not None for r in photo_table.rows))
        self.assertIsNotNone(photo_table.rows[0]._tr.find(f'.//{w}tblHeader'))


class ContentLinesTest(unittest.TestCase):
    def test_lines_are_kept_as_typed(self):
        # 項目符號/編號由使用者在編輯區自己加，填表時不再自動補符號
        self.assertEqual(g.content_lines(''), [''])
        self.assertEqual(g.content_lines(' a \n\n● b\n1. c'), ['a', '', '● b', '1. c'])


class ListToolsTest(unittest.TestCase):
    def test_toggle_bullets_add_and_remove(self):
        lines = ['a', '', 'b']
        added = g.toggle_bullets(lines)
        self.assertEqual(added, ['● a', '', '● b'])
        self.assertEqual(g.toggle_bullets(added), ['a', '', 'b'])

    def test_toggle_bullets_mixed_selection_adds_to_all(self):
        self.assertEqual(g.toggle_bullets(['● a', 'b']), ['● a', '● b'])

    def test_toggle_numbers_sequence_skips_blank_lines(self):
        numbered = g.toggle_numbers(['a', '', 'b', 'c'])
        self.assertEqual(numbered, ['1. a', '', '2. b', '3. c'])
        self.assertEqual(g.toggle_numbers(numbered), ['a', '', 'b', 'c'])

    def test_default_bullet_is_solid_circle_and_dot_is_gone_from_menu(self):
        self.assertEqual(g.BULLET_CHOICES[0], ('●', '實心圓形'))
        self.assertEqual([name for _, name in g.BULLET_CHOICES], ['實心圓形', '實心菱形', '實心方塊'])
        self.assertNotIn('•', g.BULLET_SYMBOLS)
        self.assertEqual(g.toggle_bullets(['a']), ['● a'])

    def test_legacy_small_dot_is_still_recognised(self):
        self.assertTrue(g.LIST_MARKER_RE.match('• 舊的項目'))
        self.assertEqual(g.strip_list_marker('• 舊的項目'), '舊的項目')
        self.assertEqual(g.toggle_bullets(['• a', 'b']), ['● a', '● b'])  # 改選實心圓形會把舊的小圓點換掉
        self.assertEqual(g.continue_list('• 舊的項目'), ('continue', '• '))
        ind = g._build_content_paragraph('• 舊的項目').find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr').find(
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind')
        self.assertEqual(ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hanging'), '240')

    def test_bullet_with_chosen_symbol(self):
        added = g.toggle_bullets(['a', '', 'b'], '◆')
        self.assertEqual(added, ['◆ a', '', '◆ b'])
        self.assertEqual(g.toggle_bullets(added, '◆'), ['a', '', 'b'])  # 同一個符號再選一次 = 取消

    def test_choosing_another_symbol_replaces_the_old_one(self):
        self.assertEqual(g.toggle_bullets(['◆ a', '● b', '1. c'], '■'), ['■ a', '■ b', '■ c'])

    def test_remove_list_markers(self):
        self.assertEqual(g.remove_list_markers(['◆ a', '', '2. b', '- c', '一般']), ['a', '', 'b', 'c', '一般'])

    def test_every_dropdown_symbol_is_recognised_as_a_marker(self):
        for symbol, _ in g.BULLET_CHOICES:
            line = f'{symbol} 內容'
            with self.subTest(symbol=symbol):
                self.assertTrue(g.LIST_MARKER_RE.match(line))
                self.assertEqual(g.strip_list_marker(line), '內容')
                self.assertEqual(g.continue_list(line), ('continue', f'{symbol} '))
                self.assertEqual(g.continue_list(f'{symbol} '), ('end', ''))

    def test_hyphen_needs_space_to_count_as_marker(self):
        self.assertIsNone(g.LIST_MARKER_RE.match('-3 度'))
        self.assertEqual(g.strip_list_marker('-3 度'), '-3 度')
        self.assertEqual(g.continue_list('-3 度'), (None, ''))
        self.assertTrue(g.LIST_MARKER_RE.match('- 項目'))

    def test_switch_between_bullets_and_numbers(self):
        self.assertEqual(g.toggle_numbers(['● a', '● b']), ['1. a', '2. b'])
        self.assertEqual(g.toggle_bullets(['1. a', '2. b']), ['● a', '● b'])

    def test_continue_list(self):
        self.assertEqual(g.continue_list('● 山屋正常'), ('continue', '● '))
        self.assertEqual(g.continue_list('2. 水塔滿水'), ('continue', '3. '))
        self.assertEqual(g.continue_list('● '), ('end', ''))
        self.assertEqual(g.continue_list('3. '), ('end', ''))
        self.assertEqual(g.continue_list('一般文字'), (None, ''))


class OutputNameTest(unittest.TestCase):
    BASE = dict(start_year='115', start_month='9', start_day='25',
                end_year='115', end_month='9', end_day='28', location='排雲山莊')

    def test_roc_date_range(self):
        d = dict(self.BASE)
        self.assertEqual(g.roc_date_range(d), '1150925-28')
        d.update(end_day='25')
        self.assertEqual(g.roc_date_range(d), '1150925')
        d.update(end_month='10', end_day='2')
        self.assertEqual(g.roc_date_range(d), '1150925-1002')
        d.update(end_year='116', end_month='1', end_day='3')
        self.assertEqual(g.roc_date_range(d), '1150925-1160103')

    def test_default_output_name(self):
        self.assertEqual(g.default_output_name(self.BASE), '1150925-28_排雲山莊協勤回報單')

    def test_invalid_filename_chars_removed_and_empty_location(self):
        d = dict(self.BASE, location='排雲/山:莊?')
        self.assertEqual(g.default_output_name(d), '1150925-28_排雲山莊協勤回報單')
        d = dict(self.BASE, location='')
        self.assertEqual(g.default_output_name(d), '1150925-28_協勤回報單')


class FormattingTest(BuildReportTest):
    """填入內容後，模板原有格式要保留；新增內容依參考回報單格式（標楷體 12pt、固定行高 22pt）。"""

    W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    def _paragraphs(self, startswith):
        for c in unique_cells(self.doc):
            for p in c.paragraphs:
                if p.text.startswith(startswith):
                    return c, p
        self.fail(startswith)

    def test_template_headings_keep_bold_and_14pt(self):
        _, heading = self._paragraphs('三、自然資源狀況')
        self.assertTrue(all(r.font.bold for r in heading.runs))
        self.assertEqual(heading.runs[0].font.size.pt, 14)

    def test_fill_date_is_centered_and_service_date_is_left_aligned_in_word(self):
        _, p = self._paragraphs('115年10月12日')
        jc = p._p.pPr.find(self.W + 'jc')
        self.assertIsNotNone(jc)
        self.assertEqual(jc.get(self.W + 'val'), 'center')  # 填表日期置中
        _, p = self._paragraphs('115年10月10日至115年10月11日')
        jc = p._p.pPr.find(self.W + 'jc')
        self.assertIsNotNone(jc)
        self.assertEqual(jc.get(self.W + 'val'), 'left')    # 協勤日期靠左

    def test_question_placeholder_is_removed(self):
        self.assertNotIn('因為', self.text)
        self.assertNotIn('1.相關疑問', self.text)
        self.assertIn('八、相關疑問與解答', self.text)
        self.assertEqual(self.text.count('內容_questions'), 1)

    def test_question_placeholder_is_removed_even_when_questions_are_empty(self):
        d = tempfile.mkdtemp()
        data = {k: '' for k in g.FIELD_NAMES}
        data.update(year='115', month='10', day='12', personnel='王', start_year='115', start_month='10',
                    start_day='10', end_year='115', end_month='10', end_day='11', location='排雲山莊')
        out = os.path.join(d, 'empty.docx')
        g.build_report(TEMPLATE, data, {c: [] for c in g.PHOTO_CATEGORIES}, {}, out)
        text = chr(10).join(c.text for c in unique_cells(Document(out)))
        self.assertNotIn('因為', text)
        self.assertNotIn('1.相關疑問', text)
        self.assertIn('八、相關疑問與解答', text)

    def test_is_question_placeholder_matches_both_ellipsis_styles(self):
        self.assertTrue(g.is_question_placeholder('1.相關疑問？答：因為……。'))
        self.assertTrue(g.is_question_placeholder('1.相關疑問？答：因為......。'))
        self.assertFalse(g.is_question_placeholder('1.雲海什麼時候最漂亮？答：清晨。'))
        self.assertFalse(g.is_question_placeholder('八、相關疑問與解答：'))

    def test_personnel_and_location_are_not_changed_to_center(self):
        _, p = self._paragraphs('王小明')
        jc = p._p.pPr.find(self.W + 'jc') if p._p.pPr is not None else None
        self.assertTrue(jc is None or jc.get(self.W + 'val') != 'center')

    def test_title_and_date_alignment_kept(self):
        _, title = self._paragraphs('玉山國家公園高山保育志工協勤回報單')
        self.assertTrue(title.runs[0].font.bold)
        _, date = self._paragraphs('115年10月12日')
        self.assertIsNotNone(date._p.pPr.find(self.W + 'jc'))  # 右對齊沒有被洗掉

    def test_header_values_use_kai_and_keep_template_size(self):
        _, name = self._paragraphs('王小明')
        run = name.runs[0]
        rfonts = run._r.rPr.rFonts
        for attr in ('ascii', 'hAnsi', 'eastAsia'):
            self.assertEqual(rfonts.get(self.W + attr), '標楷體')
        self.assertEqual(run.font.size.pt, 14)

    def test_content_paragraph_format(self):
        _, p = self._paragraphs('● 第二行_road_status')  # 使用者自己加的項目符號
        run = p.runs[0]
        self.assertEqual(run.font.size.pt, 12)
        self.assertEqual(run._r.rPr.rFonts.get(self.W + 'eastAsia'), '標楷體')
        spacing = p._p.pPr.find(self.W + 'spacing')
        self.assertEqual(spacing.get(self.W + 'line'), '440')
        self.assertEqual(spacing.get(self.W + 'lineRule'), 'exact')
        ind = p._p.pPr.find(self.W + 'ind')
        self.assertEqual(ind.get(self.W + 'hanging'), '240')

    def test_plain_line_has_no_hanging_indent_and_no_auto_bullet(self):
        _, p = self._paragraphs('內容_road_status')  # 使用者沒加符號 -> 不會被自動加上「● 」
        self.assertFalse(p.text.startswith('●'))
        ind = p._p.pPr.find(self.W + 'ind')
        self.assertEqual(ind.get(self.W + 'left'), '480')
        self.assertIsNone(ind.get(self.W + 'hanging'))

    def test_any_dropdown_symbol_gets_hanging_indent_in_report(self):
        for symbol, _ in g.BULLET_CHOICES:
            ind = g._build_content_paragraph(f'{symbol} 內容').find(self.W + 'pPr').find(self.W + 'ind')
            with self.subTest(symbol=symbol):
                self.assertEqual(ind.get(self.W + 'hanging'), '240')

    def test_numbered_line_hanging_indent(self):
        doc_p = g._build_content_paragraph('1. 項目')
        ind = doc_p.find(self.W + 'pPr').find(self.W + 'ind')
        self.assertEqual(ind.get(self.W + 'hanging'), '360')

    def test_content_is_inserted_right_after_label(self):
        _, label = self._paragraphs('1、道路、步道狀況及維護情形：')
        nxt = label._p.getnext()
        self.assertIn('內容_road_status', ''.join(t.text or '' for t in nxt.iter(self.W + 't')))

    def test_cell_is_not_flattened(self):
        # 舊作法用 cell.text 重組整格，標題的粗體會消失；現在標題與內容是不同段落
        cell, _ = self._paragraphs('三、自然資源狀況')
        self.assertGreater(len(cell.paragraphs), 3)


if __name__ == '__main__':
    unittest.main()


class LabelColumnTest(unittest.TestCase):
    def test_label_column_widened_and_table_width_unchanged(self):
        doc = Document(TEMPLATE)
        table = doc.tables[0]
        grid_before = [int(c.get(g.qn('w:w'))) for c in table._tbl.tblGrid.findall(g.qn('w:gridCol'))]
        g.widen_label_column(table)
        grid = [int(c.get(g.qn('w:w'))) for c in table._tbl.tblGrid.findall(g.qn('w:gridCol'))]
        self.assertEqual(grid[0], g.LABEL_COLUMN_TWIPS)
        self.assertEqual(sum(grid), sum(grid_before))
        # 每一列的儲存格寬度加起來仍等於整個表格寬度
        for tr in table._tbl.findall(g.qn('w:tr')):
            total = sum(int(tc.find(g.qn('w:tcPr')).find(g.qn('w:tcW')).get(g.qn('w:w')))
                        for tc in tr.findall(g.qn('w:tc')))
            self.assertEqual(total, sum(grid))
        # 「協勤人員姓名：」14pt 共 7 個全形字，要放得下一行（寬度 = 欄寬 - 左右邊距各 108 twips）
        self.assertGreaterEqual(grid[0] - 216, 7 * 14 * 20)

    def test_is_idempotent(self):
        doc = Document(TEMPLATE)
        g.widen_label_column(doc.tables[0])
        once = [c.get(g.qn('w:w')) for c in doc.tables[0]._tbl.tblGrid.findall(g.qn('w:gridCol'))]
        g.widen_label_column(doc.tables[0])
        twice = [c.get(g.qn('w:w')) for c in doc.tables[0]._tbl.tblGrid.findall(g.qn('w:gridCol'))]
        self.assertEqual(once, twice)
