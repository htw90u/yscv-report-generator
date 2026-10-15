"""塔塔加保育宣導回報單、生態廁所回報單：Word 填寫與畫面流程測試。執行：python -m unittest discover -s tests"""
import gc
import json
import os
import tempfile
import time
import tkinter as tk
import unittest
from tkinter import filedialog
from unittest import mock

from PIL import Image
from docx import Document
from docx.oxml.ns import qn

from _common import TATAKA_TEMPLATE, ECO_TEMPLATE
import yscv_report_generator as g
import yscv_reports as r

messagebox = g.messagebox


def all_text(doc):
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.append(cell.text)
    return '\n'.join(parts)


def cell_text(table, row, col):
    return r._cell(table, row, col).text


def checked_symbols(paragraph_el):
    """段落裡所有方框符號的狀態，依序：True＝打勾。"""
    return [sym.get(qn('w:char')) == r.CHECKED_SYMBOL for sym in paragraph_el.iter(qn('w:sym'))]


nl = chr(10)


TATAKA_DATA = dict(
    year='115', month='10', day='15', volunteer='王小明', time_text='09:00～16:00', weather='陰',
    guide_booked='1', guide_walkin='', center_on='', center_text='櫃檯諮詢', other_on='', other_text='',
    incident='有', incident_text='遊客扭傷腳踝', facility='● 木棧板鬆動\n● 衛生紙不足', visitors='指標不清楚',
    reflection='收穫很多', questions='1. 問：海拔多高？答：約2600公尺。')

ECO_DATA = dict(
    year='115', month='10', day='15', personnel='王小明', start_year='115', start_month='10', start_day='15',
    end_year='115', end_month='10', end_day='17', location='玉山主峰線', nature='● 玉山杜鵑', environment='帶下山垃圾 2 袋',
    reflection='狀況良好')
for _prefix, _ in r.TOILETS:
    for _spec, _value in zip(r.EXCRETA_FIELDS + r.COMPOST_FIELDS,
                             ['12.5', '晴', '有', '1.2', '3', '4', f'{_prefix}備註一', '正常', '55%', '18℃', '0.5', '是',
                              f'{_prefix}備註二']):
        ECO_DATA[f'{_prefix}_{_spec[0]}'] = _value


class ReportBuildTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def photos(self, count, size=(400, 200)):
        paths = []
        for i in range(count):
            p = os.path.join(self.tmp.name, f'p{i}.png')
            Image.new('RGB', size if i % 2 == 0 else size[::-1], 'red').save(p)
            paths.append(p)
        return paths

    def tataka(self, data=None, count=0, captions=None):
        out = os.path.join(self.tmp.name, 'tataka.docx')
        paths = self.photos(count)
        caps = {r.TATAKA_PHOTO_CATEGORY: captions or {p: f'名稱{i}\n地點{i}' for i, p in enumerate(paths)}}
        warnings = r.build_tataka_report(TATAKA_TEMPLATE, data or TATAKA_DATA, {r.TATAKA_PHOTO_CATEGORY: paths},
                                         caps, out)
        return Document(out), warnings

    def eco(self, data=None, count=0):
        out = os.path.join(self.tmp.name, 'eco.docx')
        paths = self.photos(count)
        caps = {r.ECO_PHOTO_CATEGORY: {p: f'照片說明{i}' for i, p in enumerate(paths)}}
        warnings = r.build_ecotoilet_report(ECO_TEMPLATE, data or ECO_DATA, {r.ECO_PHOTO_CATEGORY: paths}, caps, out)
        return Document(out), warnings

    # ---- 塔塔加
    def test_tataka_header_cells_and_weather_box(self):
        doc, warnings = self.tataka()
        self.assertEqual(warnings, [])
        table = doc.tables[0]
        self.assertEqual(cell_text(table, 1, 0), '王小明')
        self.assertEqual(cell_text(table, 1, 1).replace('\n', ''), '115年10月15日')
        self.assertEqual(cell_text(table, 1, 2), '09:00～16:00')
        weather = r._cell(table, 1, 3)._tc.findall(qn('w:p'))[0]
        self.assertEqual(checked_symbols(weather), [False, True, False])   # 晴 陰 雨：只有「陰」打勾
        self.assertEqual(r._cell(table, 1, 3).text, '晴陰雨')

    def test_split_names_detects_common_separators(self):
        expected = ['王小明', '李小華', '陳大文']
        for text in ('王小明 李小華 陳大文', '王小明、李小華、陳大文', '王小明,李小華,陳大文', '王小明，李小華；陳大文',
                     '王小明' + chr(0x3000) + '李小華/陳大文', '  王小明 ,  、李小華　　陳大文 ', '王小明' + chr(10) + '李小華' + chr(9) + '陳大文'):
            self.assertEqual(r.split_names(text), expected, repr(text))
        self.assertEqual(r.split_names('王小明'), ['王小明'])
        # 外文姓名：中間的空白不拆；中文姓名之間的空白照常拆
        self.assertEqual(r.split_names('John Smith'), ['John Smith'])
        self.assertEqual(r.split_names('John Smith, Mary Jane Watson'), ['John Smith', 'Mary Jane Watson'])
        self.assertEqual(r.split_names('王小明 John Smith 李小華'), ['王小明', 'John Smith', '李小華'])
        self.assertEqual(r.split_names('王小明 李小華 John Smith'), ['王小明', '李小華', 'John Smith'])
        self.assertEqual(r.split_names('John Smith、王小明' + chr(0x3000) + 'Mary Jane'), ['John Smith', '王小明', 'Mary Jane'])
        self.assertEqual(r.split_names('김 철수 田中 太郎'), ['김 철수', '田中', '太郎'])   # 韓文不拆，漢字姓名照拆
        self.assertEqual(r.split_names('約翰·史密斯 王小明'), ['約翰·史密斯', '王小明'])
        self.assertEqual(r.split_names('John  Smith' + chr(10) + 'Mary Jane'), ['John Smith', 'Mary Jane'])   # 換行一定分隔
        self.assertEqual(r.split_names('  ， 、 '), [])
        self.assertEqual(r.split_names(''), [])

    def test_tataka_each_volunteer_on_its_own_centered_line(self):
        doc, _ = self.tataka(dict(TATAKA_DATA, volunteer='王小明、李小華 陳大文,林小美'))
        cell = r._cell(doc.tables[0], 1, 0)
        paragraphs = cell._tc.findall(qn('w:p'))
        self.assertEqual([g._paragraph_text(p) for p in paragraphs], ['王小明', '李小華', '陳大文', '林小美'])
        for p in paragraphs:
            self.assertEqual(p.find(qn('w:pPr')).find(qn('w:jc')).get(qn('w:val')), 'center')
            self.assertEqual([c.tag.split('}')[1] for c in p.find(qn('w:pPr'))][-1], 'rPr')   # rPr 放最後（Word 的順序）
        height = doc.tables[0]._tbl.tr_lst[1].trPr.find(qn('w:trHeight'))
        self.assertNotEqual(height.get(qn('w:hRule')), 'exact')            # 人多時列會長高，不會被切掉

    def test_tataka_foreign_name_stays_on_one_line(self):
        doc, _ = self.tataka(dict(TATAKA_DATA, volunteer='王小明 John Smith 李小華'))
        paragraphs = r._cell(doc.tables[0], 1, 0)._tc.findall(qn('w:p'))
        self.assertEqual([g._paragraph_text(p) for p in paragraphs], ['王小明', 'John Smith', '李小華'])

    def test_tataka_service_items_and_incident(self):
        doc, _ = self.tataka()
        table = doc.tables[0]
        paragraphs = r._cell(table, 2, 1)._tc.findall(qn('w:p'))
        self.assertEqual(checked_symbols(paragraphs[0]), [True, False])    # 預約打勾、散客沒有
        self.assertEqual(checked_symbols(paragraphs[2]), [True])           # 有填遊客中心的文字，自動打勾
        self.assertIn('櫃檯諮詢', g._paragraph_text(paragraphs[2]))
        self.assertEqual(checked_symbols(paragraphs[3]), [False])          # 其他：沒填
        incident = r._cell(table, 3, 1)._tc.findall(qn('w:p'))[0]
        self.assertEqual(checked_symbols(incident), [False, True])         # 無／有：選「有」
        self.assertIn('遊客扭傷腳踝', g._paragraph_text(incident))

    def test_tataka_several_incidents_are_numbered_under_the_yes_box(self):
        records = nl.join(['遊客扭傷腳踝', '山屋停水', '有人違規露營'])
        doc, _ = self.tataka(dict(TATAKA_DATA, incident_text=records))
        paragraphs = r._cell(doc.tables[0], 3, 1)._tc.findall(qn('w:p'))
        self.assertEqual(len(paragraphs), 4)                              # 「有：」那行＋三筆
        self.assertEqual(checked_symbols(paragraphs[0]), [False, True])
        self.assertNotIn('遊客扭傷腳踝', g._paragraph_text(paragraphs[0]))
        texts = [g._paragraph_text(p).replace(' ', '').replace(chr(0x3000), '') for p in paragraphs[1:]]
        self.assertEqual(texts, ['1.遊客扭傷腳踝', '2.山屋停水', '3.有人違規露營'])
        height = doc.tables[0]._tbl.tr_lst[3].trPr.find(qn('w:trHeight'))
        self.assertNotEqual(height.get(qn('w:hRule')), 'exact')

    def test_tataka_incident_records_ignored_unless_yes(self):
        doc, _ = self.tataka(dict(TATAKA_DATA, incident='無', incident_text=nl.join(['甲', '乙'])))
        paragraphs = r._cell(doc.tables[0], 3, 1)._tc.findall(qn('w:p'))
        self.assertEqual(len(paragraphs), 1)
        self.assertNotIn('甲', g._paragraph_text(paragraphs[0]))

    def test_tataka_no_incident_checks_none_box(self):
        doc, _ = self.tataka(dict(TATAKA_DATA, incident='無', incident_text=''))
        incident = r._cell(doc.tables[0], 3, 1)._tc.findall(qn('w:p'))[0]
        self.assertEqual(checked_symbols(incident), [True, False])

    def test_tataka_text_rows_are_filled_and_not_clipped(self):
        doc, _ = self.tataka()
        table = doc.tables[0]
        self.assertIn('木棧板鬆動', cell_text(table, 4, 1))
        self.assertIn('衛生紙不足', cell_text(table, 4, 1))
        self.assertEqual(cell_text(table, 5, 1), '指標不清楚')
        self.assertEqual(cell_text(table, 6, 1), '收穫很多')
        for tr in table._tbl.tr_lst[1:5]:
            height = tr.trPr.find(qn('w:trHeight'))
            self.assertNotEqual(height.get(qn('w:hRule')), 'exact')        # 固定高度會切掉長文字

    def test_tataka_question_sample_line_removed(self):
        doc, _ = self.tataka()
        text = cell_text(doc.tables[0], 7, 1)
        self.assertNotIn('因為', text)
        self.assertIn('海拔多高', text)
        self.assertIn('將服勤期間相關疑問', text)             # 說明文字保留
        doc, _ = self.tataka(dict(TATAKA_DATA, questions=''))
        self.assertNotIn('因為', cell_text(doc.tables[0], 7, 1))   # 沒填也要拿掉範例

    def test_tataka_one_row_per_photo_with_name_and_place(self):
        doc, _ = self.tataka(count=3)
        table = doc.tables[0]
        self.assertEqual(len(table.rows), 9 + 3)                  # 範本 10 列：最後的空白列換成 3 張照片
        self.assertEqual(len(doc.inline_shapes), 3)
        self.assertEqual(cell_text(table, 9, 0), '名稱0\n地點0')
        self.assertEqual(cell_text(table, 11, 0), '名稱2\n地點2')
        for shape in doc.inline_shapes:                           # 照片放得進儲存格，不超過高度上限
            self.assertLessEqual(shape.height / 914400, r.PHOTO_MAX_HEIGHT_IN + 0.01)
            self.assertLess(shape.width / 914400, 4.5)

    def test_tataka_name_only_caption_and_no_photos(self):
        paths = self.photos(1)
        doc, _ = self.tataka(count=1, captions={paths[0]: '只有名稱'})
        self.assertEqual(cell_text(doc.tables[0], 9, 0), '只有名稱')
        doc, _ = self.tataka(count=0)
        self.assertEqual(len(doc.tables[0].rows), 10)             # 沒照片：維持範本的空白列
        self.assertEqual(len(doc.inline_shapes), 0)

    def test_tataka_bad_template_is_reported(self):
        out = os.path.join(self.tmp.name, 'bad.docx')
        with self.assertRaises(ValueError):
            r.build_tataka_report(ECO_TEMPLATE, TATAKA_DATA, {}, {}, out)

    # ---- 生態廁所
    def test_eco_header_and_dates(self):
        doc, warnings = self.eco()
        self.assertEqual(warnings, [])
        header = doc.tables[0]
        self.assertIn('115 年 10 月 15 日', cell_text(header, 1, 0))
        self.assertEqual(cell_text(header, 2, 1), '王小明')
        self.assertEqual(cell_text(header, 3, 1), '115年10月15日　至　115年10月17日')
        self.assertEqual(cell_text(header, 4, 1), '玉山主峰線')

    def test_eco_records_go_into_the_right_toilet_row(self):
        doc, _ = self.eco()
        excreta, compost = doc.tables[1], doc.tables[2]
        for row, prefix in ((3, 'm'), (7, 'b')):
            values = [c.text for c in excreta.rows[row].cells]
            self.assertEqual(values, ['12.5', '晴', '有', '1.2', '3', '4', f'{prefix}備註一'])
            values = [c.text for c in compost.rows[row].cells]
            self.assertEqual(values, ['正常', '55%', '18℃', '0.5', '是', f'{prefix}備註二'])
        self.assertEqual(cell_text(excreta, 2, 0), 'kg')         # 單位列維持範本原樣

    def test_eco_sections_fill_after_headings(self):
        doc, _ = self.eco()
        texts = [p.text.strip() for p in doc.paragraphs]
        for heading, content in (('二、自然資源狀況', '● 玉山杜鵑'), ('三、環境清理及遊客宣導情形', '帶下山垃圾 2 袋'),
                                 ('四、協勤心得與建議', '狀況良好')):
            index = next(i for i, t in enumerate(texts) if t.startswith(heading))
            following = [t for t in texts[index + 1:index + 4] if t]
            self.assertIn(content, following, heading)
        # 二、三的標題有兩行：內容放在第二行之後，不能插在標題中間
        nature = next(i for i, t in enumerate(texts) if t.startswith('二、自然資源狀況'))
        self.assertTrue(texts[nature + 1].endswith('等）'))
        self.assertEqual(texts[nature + 2], '● 玉山杜鵑')

    def test_eco_headings_stay_with_tables(self):
        doc, _ = self.eco()
        for p in doc.paragraphs:
            if p.text.strip().startswith(('（一）', '（二）', '（三）')):
                self.assertTrue(p.paragraph_format.keep_with_next, p.text)

    def test_eco_photo_grid_grows_and_shrinks(self):
        for count, rows in ((0, 2), (1, 2), (4, 4), (5, 6), (8, 8)):
            doc, _ = self.eco(count=count)
            grid = doc.tables[3]
            self.assertEqual(len(grid.rows), rows, count)
            self.assertEqual(len(doc.inline_shapes), count)
            self.assertNotIn('照片圖檔', all_text(doc))
            self.assertNotIn('照片文字', all_text(doc))
        doc, _ = self.eco(count=3)
        grid = doc.tables[3]
        self.assertEqual([cell_text(grid, 1, 0), cell_text(grid, 1, 1)], ['照片說明0', '照片說明1'])
        self.assertEqual([cell_text(grid, 3, 0), cell_text(grid, 3, 1)], ['照片說明2', ''])   # 單數張：最後一格空白

    def test_eco_photos_fit_the_cell(self):
        doc, _ = self.eco(count=4)
        for shape in doc.inline_shapes:
            self.assertLessEqual(shape.width / 914400, 3.4)
            self.assertLessEqual(shape.height / 914400, r.PHOTO_MAX_HEIGHT_IN + 0.01)

    def test_eco_bad_template_is_reported(self):
        out = os.path.join(self.tmp.name, 'bad.docx')
        with self.assertRaises(ValueError):
            r.build_ecotoilet_report(TATAKA_TEMPLATE, ECO_DATA, {}, {}, out)


class FormGuiTest(unittest.TestCase):
    app_class = None

    def setUp(self):
        if self.app_class is None:
            self.skipTest('基底類別')
        self.tmp = tempfile.TemporaryDirectory()
        self.root = tk.Tk()
        self.autosave = os.path.join(self.tmp.name, 'autosave.json')
        self.app = self.app_class(self.root, autosave_path=self.autosave, restore_on_start=False,
                                  settings_path=os.path.join(self.tmp.name, 'settings.json'))
        self.shown, self.pdf_calls, self.confirm_calls = [], [], []

        def fake_confirm(title, message, **kwargs):
            self.confirm_calls.append(message)
            return True

        def fake_pdf(docx, pdf):
            self.pdf_calls.append((docx, pdf))
            with open(pdf, 'wb') as f:
                f.write(b'%PDF-fake')

        for p in (mock.patch.object(messagebox, 'showerror', lambda t, m: self.shown.append(('err', m))),
                  mock.patch.object(messagebox, 'showinfo', lambda t, m: self.shown.append(('info', m))),
                  mock.patch.object(g, 'export_pdf', fake_pdf),
                  mock.patch.object(g, 'open_file', lambda path: None),
                  mock.patch.object(messagebox, 'askyesno', fake_confirm)):
            p.start()
            self.addCleanup(p.stop)

    def tearDown(self):
        if self.app_class is None:
            return
        self.root.destroy()
        self.app = self.root = None
        gc.collect()
        self.tmp.cleanup()

    def make_photos(self, names=('a.jpg', 'b.jpg')):
        category = self.app.PHOTO_CATEGORIES[0]
        paths = []
        for n in names:
            p = os.path.join(self.tmp.name, n)
            Image.new('RGB', (400, 200), 'red').save(p)
            paths.append(p)
        self.app._add_photo_files(category, paths)
        return category, self.app.listboxes[category], paths

    def wait_done(self, timeout=20):
        t0 = time.time()
        while time.time() - t0 < timeout and str(self.app.generate_btn['state']) == 'disabled':
            self.root.update()
            time.sleep(0.05)
        self.root.update()

    def generate(self):
        out = os.path.join(self.tmp.name, 'out.docx')
        captured = {}

        def ask(**kw):
            captured.update(kw)
            return out

        with mock.patch.object(filedialog, 'asksaveasfilename', ask):
            self.app.generate_report()
            self.wait_done()
        return out, captured

    def fill(self, name, value):
        g.widget_set(getattr(self.app, name), value)

    def set_text(self, key, value):
        self.app.text_fields[key].delete('1.0', tk.END)
        self.app.text_fields[key].insert('1.0', value)


class TatakaGuiTest(FormGuiTest):
    app_class = r.TatakaReportGUI

    def fill_all(self):
        values = dict(year='115', month='10', day='15', volunteer='王小明', time_text='09:00～16:00', weather='晴',
                      guide_booked='1', center_text='櫃檯諮詢', incident='無')
        for name, value in values.items():
            self.fill(name, value)
        for key in ('facility', 'visitors', 'reflection', 'questions'):
            self.set_text(key, f'{key}內容')

    def test_window_title_names_the_report(self):
        self.assertEqual(self.root.title(), f'{g.APP_TITLE}－塔塔加保育宣導回報單　{g.version_label()}')

    def test_tabs_and_default_date(self):
        titles = [self.app.notebook.tab(t, 'text').replace(' ' + g.MARK_MISSING, '') for t in self.app.notebook.tabs()]
        self.assertEqual(titles, ['基本資料', '服務項目', '設施與遊客意見', '心得與疑問解答', '相片記錄'])
        self.assertEqual(self.app.year.get(), g.roc_date_parts()[0])      # 值勤日期預設今天

    def test_blank_fields_and_tab_marks(self):
        blanks = self.app.find_blank_fields({})
        self.assertEqual(list(blanks), ['基本資料', '服務項目', '設施與遊客意見', '心得與疑問解答', '相片記錄'])
        self.assertIn('志工姓名', blanks['基本資料'])
        self.assertIn('服務項目（至少選一項）', blanks['服務項目'])
        self.fill_all()
        self.assertEqual(self.app.find_blank_fields({}), {'相片記錄': ['相片記錄（尚未加入照片）']})
        self.make_photos(('a.jpg',))
        self.assertEqual(self.app.find_blank_fields({}), {})
        self.assertIn('相片記錄', self.app._tabs_with_blanks())          # 照片還沒填資料：分頁仍標 ●
        self.fill('incident', '有')
        self.assertIn('偶發事件的說明', self.app.find_blank_fields({})['服務項目'])

    def test_two_caption_fields_are_joined_and_split(self):
        category, table, (a, b) = self.make_photos()
        table.set_selection([0])
        self.app._on_photo_select(category, table)
        self.app.caption_vars[0].set(' 塔塔加遊客中心 ')
        self.app.caption_vars[1].set('入口大廳')
        self.app._commit_caption()
        self.assertEqual(self.app.photo_captions[category][a], '塔塔加遊客中心\n入口大廳')
        self.assertEqual(table.item('0', 'values')[1], '塔塔加遊客中心／入口大廳')
        table.set_selection([1])
        self.app._on_photo_select(category, table)
        self.assertEqual([v.get() for v in self.app.caption_vars], ['', ''])
        table.set_selection([0])
        self.app._on_photo_select(category, table)
        self.assertEqual([v.get() for v in self.app.caption_vars], ['塔塔加遊客中心', '入口大廳'])
        self.app.caption_vars[0].set('')
        self.app.caption_vars[1].set('')
        self.app._commit_caption()
        self.assertNotIn(a, self.app.photo_captions[category])         # 兩欄都空＝沒填

    def test_enter_moves_to_next_field_then_next_photo(self):
        category, table, (a, b) = self.make_photos()
        self.app._select_photo(category, 0)
        self.app.caption_vars[0].set('甲')
        self.assertEqual(self.app._caption_return(None, 0), 'break')
        self.assertEqual(self.app.photo_captions[category].get(a), None)   # 還在第一張，還沒儲存
        self.app.caption_vars[1].set('地點')
        self.app._caption_return(None, 1)
        self.assertEqual(self.app.photo_captions[category][a], '甲\n地點')
        self.assertEqual(list(table.selected_indexes()), [1])           # 跳到下一張

    def test_collect_data_validates_date(self):
        self.fill_all()
        self.fill('month', '2')
        self.fill('day', '30')
        data, error = self.app.collect_data()
        self.assertIsNone(data)
        self.assertIn('值勤日期', error)

    def test_draft_roundtrip_keeps_checkboxes_and_choices(self):
        self.fill_all()
        self.fill('guide_walkin', '1')
        category, table, (a, b) = self.make_photos()
        self.app.photo_captions[category][a] = '名稱\n地點'
        draft = json.loads(json.dumps(self.app.build_draft()))
        self.assertEqual(draft['report'], 'tataka')
        self.assertEqual(draft['fields']['guide_booked'], '1')
        self.app._clear_widgets()
        self.assertEqual(self.app.guide_booked.get(), '')
        self.assertEqual(self.app.weather.get(), '')
        self.assertEqual(self.app.apply_draft(draft), [])
        self.assertEqual((self.app.guide_booked.get(), self.app.guide_walkin.get(), self.app.weather.get(),
                          self.app.incident.get()), ('1', '1', '晴', '無'))
        self.assertEqual(self.app.photo_captions[category][a], '名稱\n地點')
        self.assertEqual(self.app.text_fields['facility'].get('1.0', tk.END).strip(), 'facility內容')

    def test_plus_button_adds_incident_records(self):
        box = self.app.incident_text
        self.assertEqual(len(box.rows), 1)
        self.assertTrue(box.rows[0][2].instate(['disabled']))             # 只有一筆時不能刪
        box.add_button.invoke()
        box.add_button.invoke()
        self.assertEqual(len(box.rows), 3)
        self.assertFalse(box.rows[0][2].instate(['disabled']))
        for (frame, entry, button), text in zip(box.rows, ('甲', '', '丙')):
            entry.insert(0, text)
        self.assertEqual(box.get(), '甲' + nl + '丙')                       # 空白筆不算
        before = self.app._signature()
        box.rows[1][2].invoke()                                           # 刪掉中間（空白）那筆
        self.assertEqual(len(box.rows), 2)
        box.rows[0][2].invoke()
        self.assertEqual(box.get(), '丙')
        self.assertNotEqual(self.app._signature(), before)
        self.assertTrue(box.rows[0][2].instate(['disabled']))             # 只剩一筆：刪除鈕停用
        box.remove_row(box.rows[0])                                       # 硬要刪就只是清空文字
        self.assertEqual((len(box.rows), box.get()), (1, ''))

    def test_incident_records_survive_draft_and_mark_the_tab(self):
        self.fill('incident', '有')
        self.assertIn('偶發事件的說明', self.app.find_blank_fields({})['服務項目'])
        self.fill('incident_text', '甲' + nl + '乙')
        self.assertEqual(len(self.app.incident_text.rows), 2)
        self.assertNotIn('偶發事件的說明', self.app.find_blank_fields({}).get('服務項目', []))
        draft = json.loads(json.dumps(self.app.build_draft()))
        self.assertEqual(draft['fields']['incident_text'], '甲' + nl + '乙')
        self.app._clear_widgets()
        self.assertEqual((len(self.app.incident_text.rows), self.app.incident_text.get()), (1, ''))
        self.app.apply_draft(draft)
        self.assertEqual([e.get() for _, e, _ in self.app.incident_text.rows], ['甲', '乙'])
        data, _ = self.app.collect_data()
        self.assertEqual(data['incident_text'], '甲' + nl + '乙')

    def test_enter_in_incident_entry_adds_next_record(self):
        box = self.app.incident_text
        box.rows[0][1].insert(0, '甲')
        self.app.notebook.select(self.app._tab_by_title(self.app.TAB_SERVICE))
        self.root.deiconify()
        self.root.update()
        box.rows[0][1].focus_force()
        self.root.update()
        box.rows[0][1].event_generate('<Return>')
        self.root.update()
        self.assertEqual(len(box.rows), 2)

    def test_checkbox_change_marks_the_tab_and_dirty(self):
        before = self.app._signature()
        self.assertFalse(self.app.is_dirty())
        self.app.guide_booked.invoke()
        self.assertEqual(self.app.guide_booked.get(), '1')
        self.assertNotEqual(self.app._signature(), before)
        self.assertTrue(self.app.is_dirty())

    def test_draft_of_another_report_is_rejected(self):
        patrol_draft = {'version': g.DRAFT_VERSION, 'fields': {}}              # 舊草稿沒有種類欄位＝保育巡查回報單
        with self.assertRaises(ValueError) as caught:
            self.app.apply_draft(patrol_draft)
        self.assertIn('保育巡查回報單', str(caught.exception))
        with self.assertRaises(ValueError):
            self.app.apply_draft(dict(patrol_draft, report='ecotoilet'))
        self.app.apply_draft(dict(patrol_draft, report='tataka'))

    def test_file_names(self):
        self.fill_all()
        data, _ = self.app.collect_data()
        self.assertEqual(self.app.output_name(data), '1151015_塔塔加保育宣導回報單_王小明')
        data['volunteer'] = '王小明 李小華,John Smith'
        self.assertEqual(self.app.output_name(data), '1151015_塔塔加保育宣導回報單_王小明、李小華、John Smith')
        data['volunteer'] = '王小明 李小華,陳大文'
        self.assertEqual(self.app.output_name(data), '1151015_塔塔加保育宣導回報單_王小明、李小華、陳大文')
        name = self.app.draft_name({'year': '115', 'month': '10', 'day': '15', 'volunteer': '王小:明'},
                                   g.datetime(2026, 10, 15, 9, 30))
        self.assertEqual(name, '1151015_塔塔加保育宣導回報單草稿_王小明_20261015_0930')
        self.assertEqual(self.app.draft_name({}, g.datetime(2026, 10, 15, 9, 30)),
                         '塔塔加保育宣導回報單草稿_20261015_0930')

    def test_generate_report_end_to_end(self):
        self.fill_all()
        category, table, (a, b) = self.make_photos()
        self.app.photo_captions[category] = {a: '甲\n地點甲', b: '乙'}
        out, captured = self.generate()
        self.assertEqual(captured['initialfile'], '1151015_塔塔加保育宣導回報單_王小明.docx')
        self.assertEqual(self.confirm_calls, [])                        # 都填了：不會問「有欄位未填」
        self.assertEqual(self.shown, [])
        self.assertEqual(self.pdf_calls, [(out, out[:-5] + '.pdf')])
        doc = Document(out)
        self.assertEqual(len(doc.inline_shapes), 2)
        self.assertEqual(cell_text(doc.tables[0], 9, 0), '甲\n地點甲')

    def test_generate_asks_when_something_is_blank(self):
        self.fill('volunteer', '王小明')
        self.generate()
        self.assertEqual(len(self.confirm_calls), 1)
        self.assertIn('服務項目', self.confirm_calls[0])
        self.assertNotIn('志工姓名', self.confirm_calls[0])      # 姓名已經填了


class EcoToiletGuiTest(FormGuiTest):
    app_class = r.EcoToiletReportGUI

    def fill_all(self):
        for name, value in dict(year='115', month='10', day='15', personnel='王小明', start_year='115',
                                start_month='10', start_day='15', end_year='115', end_month='10', end_day='17',
                                location='玉山主峰線').items():
            self.fill(name, value)
        for key in r.TOILET_FIELD_NAMES:
            self.fill(key, ECO_DATA[key])
        for key in ('nature', 'environment', 'reflection'):
            self.set_text(key, f'{key}內容')

    def test_window_title_and_tabs(self):
        self.assertEqual(self.root.title(), f'{g.APP_TITLE}－生態廁所回報單　{g.version_label()}')
        titles = [self.app.notebook.tab(t, 'text').replace(' ' + g.MARK_MISSING, '') for t in self.app.notebook.tabs()]
        self.assertEqual(titles, ['基本資料', '孟祿亭', '白木林', '自然資源/環境清理', '協勤心得與建議', '監測及維護照片'])

    @unittest.skipUnless(os.path.exists(g.EMOJI_FONT), '需要 Windows 的彩色表情符號字型')
    def test_toilet_tabs_have_a_toilet_icon(self):
        nb = self.app.notebook
        for _, name in r.TOILETS:
            self.assertTrue(str(nb.tab(self.app._tab_by_title(name), 'image')), name)
        for title in ('基本資料', self.app.TAB_NATURE, self.app.TAB_REFLECT, self.app.TAB_PHOTO):
            self.assertEqual(str(nb.tab(self.app._tab_by_title(title), 'image')), '', title)
        # 標籤文字（含未填的 ●）不受圖示影響，仍能用名稱找到分頁
        self.assertEqual(nb.tab(self.app._tab_by_title('孟祿亭'), 'text'), '孟祿亭 ' + g.MARK_MISSING)

    def test_toilet_tabs_still_work_without_the_emoji_font(self):
        self.root.destroy()
        self.root = tk.Tk()
        with mock.patch.object(g, 'EMOJI_FONT', os.path.join(self.tmp.name, 'missing.ttf')):
            app = r.EcoToiletReportGUI(self.root, autosave_path=self.autosave, restore_on_start=False,
                                       settings_path=os.path.join(self.tmp.name, 's2.json'))
        self.assertEqual(str(app.notebook.tab(app._tab_by_title('白木林'), 'image')), '')
        self.assertEqual(app._tab_icons, [])

    def test_blank_fields_are_grouped_by_toilet_and_remarks_are_optional(self):
        blanks = self.app.find_blank_fields({})
        self.assertIn('本次收集桶總量（kg）', blanks['孟祿亭'])
        self.assertIn('本次收集桶總量（kg）', blanks['白木林'])
        self.assertNotIn('備註', blanks['孟祿亭'])
        self.fill_all()
        self.make_photos(('a.jpg',))
        self.app.photo_captions[self.app.PHOTO_CATEGORIES[0]] = {self.app.photo_files[self.app.PHOTO_CATEGORIES[0]][0]: 'x'}
        self.assertEqual(self.app.find_blank_fields({}), {})
        self.fill('b_temp', '')
        self.assertEqual(self.app.find_blank_fields({}), {'白木林': ['堆積溫度（範本：常溫）']})
        self.assertIn('白木林', self.app._tabs_with_blanks())

    def test_collect_data_includes_toilet_fields_and_validates_dates(self):
        self.fill_all()
        data, error = self.app.collect_data()
        self.assertIsNone(error)
        self.assertEqual(data['m_total'], '12.5')
        self.assertEqual(data['b_remark2'], 'b備註二')
        self.assertEqual(data['nature'], 'nature內容')
        self.fill('end_day', '1')
        data, error = self.app.collect_data()
        self.assertIn('起日', error)

    def test_draft_roundtrip_and_report_kind(self):
        self.fill_all()
        draft = json.loads(json.dumps(self.app.build_draft()))
        self.assertEqual(draft['report'], 'ecotoilet')
        self.app._clear_widgets()
        self.assertEqual(self.app.m_weather.get(), '')
        self.app.apply_draft(draft)
        self.assertEqual((self.app.m_weather.get(), self.app.b_turned.get(), self.app.m_total.get()),
                         ('晴', '是', '12.5'))
        with self.assertRaises(ValueError):
            self.app.apply_draft(dict(draft, report='tataka'))

    def test_generate_report_end_to_end(self):
        self.fill_all()
        category, table, (a, b, c) = self.make_photos(('a.jpg', 'b.jpg', 'c.jpg'))
        self.app.photo_captions[category] = {a: '甲', b: '乙', c: '丙'}
        out, captured = self.generate()
        self.assertEqual(captured['initialfile'], '1151015-17_玉山主峰線生態廁所回報單.docx')
        self.assertEqual(self.confirm_calls, [])
        self.assertEqual(self.shown, [])
        doc = Document(out)
        self.assertEqual(len(doc.inline_shapes), 3)
        self.assertEqual(cell_text(doc.tables[1], 3, 0), '12.5')
        self.assertEqual(cell_text(doc.tables[3], 1, 1), '乙')

    def test_draft_file_name(self):
        name = self.app.draft_name({'start_year': '115', 'start_month': '10', 'start_day': '15', 'end_year': '115',
                                    'end_month': '10', 'end_day': '15', 'location': '玉山主峰線'},
                                   g.datetime(2026, 10, 15, 9, 30))
        self.assertEqual(name, '1151015_玉山主峰線生態廁所回報單草稿_20261015_0930')


class LauncherTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = tk.Tk()

    def tearDown(self):
        try:
            self.root.destroy()
        except tk.TclError:
            pass
        gc.collect()

    def launcher(self):
        return g.ReportLauncher(self.root, settings_path=os.path.join(self.tmp.name, 's.json'))

    def test_lists_the_three_reports_in_order(self):
        launcher = self.launcher()
        self.assertEqual(list(launcher.buttons), ['patrol', 'tataka', 'ecotoilet'])
        self.assertEqual([c[1] for c in g.REPORT_CHOICES], ['保育巡查回報單', '塔塔加保育宣導回報單', '生態廁所回報單'])
        self.assertIn(g.version_label(), self.root.title())
        self.assertTrue(self.root.title().startswith(g.APP_TITLE))

    def test_clicking_a_button_chooses_and_closes(self):
        launcher = self.launcher()
        launcher.buttons['tataka'].invoke()
        self.assertEqual(launcher.choice, 'tataka')
        with self.assertRaises(tk.TclError):                      # 選完就關掉選單，換成回報單視窗
            self.root.winfo_exists()

    def test_card_names_are_bold_and_descriptions_are_not(self):
        import tkinter.font as tkfont
        card = self.launcher().buttons['tataka']
        _, name, description = card.parts
        self.assertEqual(name.cget('text'), '塔塔加保育宣導回報單')
        self.assertEqual(tkfont.Font(root=self.root, font=name.cget('font')).actual('weight'), 'bold')
        self.assertEqual(tkfont.Font(root=self.root, font=description.cget('font')).actual('weight'), 'normal')
        size = lambda w: tkfont.Font(root=self.root, font=w.cget('font')).actual('size')
        self.assertGreater(size(name), size(description))              # 名稱比說明大
        self.assertNotEqual(name.cget('foreground'), description.cget('foreground'))

    @unittest.skipUnless(os.path.exists(g.EMOJI_FONT), '需要 Windows 的彩色表情符號字型')
    def test_each_card_has_a_color_icon(self):
        launcher = self.launcher()
        self.assertEqual(set(g.REPORT_ICONS), {'patrol', 'tataka', 'ecotoilet'})
        for report_id, card in launcher.buttons.items():
            self.assertIsNotNone(card.icon_label, report_id)
        icon = g.emoji_image(g.REPORT_ICONS['ecotoilet'], 44)
        self.assertEqual(max(icon.size), 44)
        colors = {pixel[:3] for pixel in icon.getdata() if pixel[3] > 200}
        self.assertGreater(len(colors), 3)                               # 是彩色圖，不是單色字形

    def test_launcher_still_works_without_the_emoji_font(self):
        with mock.patch.object(g, 'EMOJI_FONT', os.path.join(self.tmp.name, 'missing.ttf')):
            launcher = self.launcher()
        for card in launcher.buttons.values():
            self.assertIsNone(card.icon_label)
        launcher.buttons['tataka'].invoke()
        self.assertEqual(launcher.choice, 'tataka')
        self.assertIsNone(g.emoji_image('', 44))

    def test_enter_on_the_focused_card_chooses_it(self):
        launcher = self.launcher()
        self.root.update()
        launcher.buttons['patrol'].focus_force()
        self.root.update()
        launcher.buttons['patrol'].event_generate('<Return>')
        self.root.update()
        self.assertEqual(launcher.choice, 'patrol')

    def test_clicking_a_card_label_chooses_it(self):
        launcher = self.launcher()
        self.root.update()
        launcher.buttons['ecotoilet'].parts[1].event_generate('<Button-1>')
        self.root.update()
        self.assertEqual(launcher.choice, 'ecotoilet')

    def test_number_keys_choose(self):
        launcher = self.launcher()
        self.root.update()
        self.root.focus_force()
        self.root.update()
        self.root.event_generate('<Key-3>')
        self.root.update()
        self.assertEqual(launcher.choice, 'ecotoilet')

    def test_escape_closes_without_choice(self):
        launcher = self.launcher()
        self.root.update()
        self.root.event_generate('<Escape>')
        self.root.update()
        self.assertIsNone(launcher.choice)

    def test_create_app_builds_each_report(self):
        self.root.destroy()
        for report_id, cls_name in (('patrol', 'ReportGeneratorGUI'), ('tataka', 'TatakaReportGUI'),
                                    ('ecotoilet', 'EcoToiletReportGUI')):
            root = tk.Tk()
            try:
                app = g.create_app(report_id, root, autosave_path=os.path.join(self.tmp.name, 'a.json'),
                                   restore_on_start=False, settings_path=os.path.join(self.tmp.name, 's.json'))
                self.assertEqual(type(app).__name__, cls_name)
                self.assertEqual(app.REPORT_ID, report_id)
                self.assertTrue(os.path.exists(app.template_path), app.template_path)
            finally:
                root.destroy()
                gc.collect()
        self.root = tk.Tk()

    def test_each_report_has_its_own_autosave_file(self):
        names = {cls.DEFAULT_AUTOSAVE_PATH for cls in (g.ReportGeneratorGUI, r.TatakaReportGUI, r.EcoToiletReportGUI)}
        self.assertEqual(len(names), 3)


if __name__ == '__main__':
    unittest.main()
