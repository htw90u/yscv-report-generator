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
        self.assertIn('[照片2說明]', self.text)

    def test_images_and_exif(self):
        self.assertEqual(len(self.doc.inline_shapes), 3)
        shp = self.doc.inline_shapes[1]
        self.assertGreater(shp.height, shp.width)  # EXIF 轉正後為直式

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
        self.assertEqual(g.prepare_image(p), p)

    def test_large_image_downscaled_keeps_aspect(self):
        p = self.make('l.jpg', (4000, 3000))
        out = g.prepare_image(p)
        with Image.open(out) as im:
            self.assertEqual(im.width, g.MAX_IMAGE_WIDTH_PX)
            self.assertEqual(im.height, round(3000 * g.MAX_IMAGE_WIDTH_PX / 4000))
        out.seek(0, 2)
        self.assertLess(out.tell(), os.path.getsize(p))

    def test_large_rotated_image_is_rotated_then_scaled(self):
        p = self.make('r.jpg', (4000, 3000), exif_orientation=6)  # 轉正後 3000x4000（直式）
        with Image.open(g.prepare_image(p)) as im:
            self.assertEqual(im.width, g.MAX_IMAGE_WIDTH_PX)
            self.assertGreater(im.height, im.width)

    def test_large_png_with_alpha_stays_png(self):
        p = os.path.join(self.tmp.name, 'a.png')
        Image.new('RGBA', (3000, 1000), (0, 255, 0, 128)).save(p)
        with Image.open(g.prepare_image(p)) as im:
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


class BulletTest(unittest.TestCase):
    def test_bullets(self):
        self.assertEqual(g.add_bullet_points(''), '')
        self.assertEqual(g.add_bullet_points('a\n\n• b'), '    • a\n\n    • b')


if __name__ == '__main__':
    unittest.main()
