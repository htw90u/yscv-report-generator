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

        data = {k: g.add_bullet_points(f'內容_{k}\n• 第二行_{k}') for k in g.FIELD_NAMES}
        data.update(year='115', month='10', day='12', personnel='王小明', start_year='115',
                    start_month='10', start_day='10', end_year='115', end_month='10', end_day='11',
                    location='排雲山莊')
        photos = {c: [] for c in g.PHOTO_CATEGORIES_1 + g.PHOTO_CATEGORIES_2}
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
        self.assertNotIn('• •', self.text)

    def test_captions(self):
        self.assertIn('自訂說明A', self.text)
        self.assertIn('[照片說明]', self.text)

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
        photos = {c: [] for c in g.PHOTO_CATEGORIES_1 + g.PHOTO_CATEGORIES_2}
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
        plan = g.plan_layout({'a': items}, self.PAGE_H * 2, 0, self.LINE_CHARS)['a']
        for row in plan:
            self.assertLessEqual(len(row), 2)
            self.assertEqual(len({p.landscape for p in row}), 1)  # 同一列同方向
        kinds = [row[0].landscape for row in plan]
        self.assertEqual(kinds, sorted(kinds, reverse=True))  # 橫式在前、直式在後
        self.assertEqual(sum(len(r) for r in plan), 5)

    def test_gap_fill_pulls_row_that_fits(self):
        # 這頁剩下的空間放不下直式列，但放得下橫式列 -> 橫式列要補進來，而不是留白換頁
        port = [make_item(3, 4) for _ in range(4)]   # 2 列直式，每列約 4.8
        land = [make_item(4, 3) for _ in range(2)]   # 1 列橫式，約 2.8
        plan = g.plan_layout({'a': land[:0] + port + land}, 10.4, 0, self.LINE_CHARS)['a']
        # 橫式列排在第一列（橫式在前），其後兩列直式放不下就換頁；不應有任何一列被拆開
        self.assertEqual([len(r) for r in plan], [2, 2, 2])

    def test_shrink_to_fill_page_not_below_min(self):
        # 4 列橫式約需 4*(2.55+0.32)+0.32 = 11.8 > 10.4，縮到 80% 可放入 -> 全部縮小且總高不超過頁高
        items = [make_item(4, 3) for _ in range(8)]
        plan = g.plan_layout({'a': items}, 10.4, 0, self.LINE_CHARS)['a']
        self.assertEqual(len(plan), 4)
        scale = items[0].img_w / 3.4
        self.assertLess(scale, 1.0)
        self.assertGreaterEqual(scale, g.MIN_PHOTO_SCALE - 1e-9)
        total = g.HEADER_H + sum(g.row_height(r, self.LINE_CHARS) for r in plan)
        self.assertLessEqual(total, 10.4 + 1e-6)

    def test_no_shrink_when_it_cannot_help(self):
        items = [make_item(4, 3) for _ in range(10)]  # 5 列：縮 80% 也放不下第 5 列 -> 換頁
        plan = g.plan_layout({'a': items}, 10.4, 0, self.LINE_CHARS)['a']
        self.assertEqual(len(plan), 5)
        self.assertTrue(all(p.img_w >= 3.4 * g.MIN_PHOTO_SCALE - 1e-9 for p in items))

    def test_caption_height_grows_with_text(self):
        self.assertLess(g.caption_height('短', 20), g.caption_height('長' * 45, 20))


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


class BulletTest(unittest.TestCase):
    def test_bullets(self):
        self.assertEqual(g.add_bullet_points(''), '')
        self.assertEqual(g.add_bullet_points('a\n\n• b'), '    • a\n\n    • b')


if __name__ == '__main__':
    unittest.main()
