"""GUI 流程測試（需要可用的顯示環境）。執行：python -m unittest discover -s tests"""
import os
import tempfile
import time
import tkinter as tk
import unittest
from tkinter import ttk, filedialog, messagebox, simpledialog
from unittest import mock

from PIL import Image
from docx import Document

from _common import TEMPLATE
import report_generator_gui as g
from test_report import unique_cells


class GuiTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = tk.Tk()
        self.app = g.ReportGeneratorGUI(self.root)
        self.app.template_path = TEMPLATE
        self.shown = []
        for p in (
            mock.patch.object(messagebox, 'showerror', lambda t, m: self.shown.append(('err', m))),
            mock.patch.object(messagebox, 'showinfo', lambda t, m: self.shown.append(('info', m))),
        ):
            p.start()
            self.addCleanup(p.stop)

    def tearDown(self):
        self.root.destroy()
        self.tmp.cleanup()

    def fill_dates(self, month='10', day='12', start_day='1', end_day='2'):
        for e, v in [(self.app.year, '115'), (self.app.month, month), (self.app.day, day),
                     (self.app.start_year, '115'), (self.app.start_month, '10'), (self.app.start_day, start_day),
                     (self.app.end_year, '115'), (self.app.end_month, '10'), (self.app.end_day, end_day)]:
            e.delete(0, tk.END)
            e.insert(0, v)

    def test_date_validation(self):
        self.app.generate_report()
        self.assertIn('日期', self.shown[-1][1])
        self.fill_dates(month='2', day='30')
        self.app.generate_report()
        self.assertIn('年月日', self.shown[-1][1])
        self.fill_dates(start_day='5', end_day='1')
        self.app.generate_report()
        self.assertIn('起日', self.shown[-1][1])

    def test_photo_list_and_caption(self):
        p = os.path.join(self.tmp.name, 'x.jpg')
        Image.new('RGB', (30, 20)).save(p)
        cat = g.PHOTO_CATEGORIES_1[0]
        lb = self.app.listboxes[cat]
        with mock.patch.object(filedialog, 'askopenfilenames', lambda **k: (p,)):
            self.app.select_photos(cat, lb)
            self.app.select_photos(cat, lb)
        self.assertEqual(lb.size(), 1)
        lb.selection_set(0)
        with mock.patch.object(simpledialog, 'askstring', lambda *a, **k: '測試說明'):
            self.app.edit_caption(cat, lb)
        self.assertIn('測試說明', lb.get(0))
        with mock.patch.object(messagebox, 'askyesno', lambda *a, **k: True):
            self.app.clear_all()
        self.assertEqual(lb.size(), 0)
        self.assertFalse(self.app.photo_files[cat] or self.app.photo_captions[cat])

    def test_generate_in_thread(self):
        self.fill_dates()
        self.app.text_fields['nature'].insert('1.0', '看到山羌\n• 已有bullet')
        out = os.path.join(self.tmp.name, 'o.docx')
        with mock.patch.object(filedialog, 'asksaveasfilename', lambda **k: out):
            self.app.generate_report()
        self.assertEqual(str(self.app.generate_btn['state']), 'disabled')
        t0 = time.time()
        while time.time() - t0 < 15 and not any(k == 'info' for k, _ in self.shown):
            self.root.update()
            time.sleep(0.05)
        self.root.update()
        self.assertTrue(any(k == 'info' for k, _ in self.shown), self.shown)
        self.assertEqual(str(self.app.generate_btn['state']), 'normal')
        text = '\n'.join(c.text for c in unique_cells(Document(out)))
        self.assertEqual(text.count('看到山羌'), 1)
        self.assertNotIn('• •', text)

    def _add_photo(self, name='d.jpg', caption='說明X'):
        p = os.path.join(self.tmp.name, name)
        Image.new('RGB', (30, 20)).save(p)
        cat = g.PHOTO_CATEGORIES_2[0]
        self.app.photo_files[cat].append(p)
        self.app.photo_captions[cat][p] = caption
        self.app.listboxes[cat].insert(tk.END, self.app._photo_label(cat, p))
        return cat, p

    def _save_draft_to(self, path):
        with mock.patch.object(filedialog, 'asksaveasfilename', lambda **k: path):
            self.app.save_draft()

    def _load_draft_from(self, path, confirm=True):
        with mock.patch.object(filedialog, 'askopenfilename', lambda **k: path), \
                mock.patch.object(messagebox, 'askyesno', lambda *a, **k: confirm):
            self.app.load_draft()

    def test_draft_roundtrip(self):
        self.fill_dates()
        self.app.personnel.insert(0, '王小明')
        self.app.text_fields['nature'].insert('1.0', '看到山羌\n第二行')
        cat, p = self._add_photo()
        draft = os.path.join(self.tmp.name, 'd.json')
        self._save_draft_to(draft)
        self.assertTrue(os.path.exists(draft))

        with mock.patch.object(messagebox, 'askyesno', lambda *a, **k: True):
            self.app.clear_all()
        self.assertEqual(self.app.personnel.get(), '')
        self._load_draft_from(draft)

        self.assertEqual(self.app.personnel.get(), '王小明')
        self.assertEqual(self.app.year.get(), '115')
        self.assertEqual(self.app.text_fields['nature'].get('1.0', tk.END).strip(), '看到山羌\n第二行')
        self.assertEqual(self.app.photo_files[cat], [p])
        self.assertEqual(self.app.photo_captions[cat][p], '說明X')
        self.assertIn('說明X', self.app.listboxes[cat].get(0))
        self.assertEqual(self.app.template_path, TEMPLATE)

    def test_draft_load_missing_photo_warns(self):
        cat, p = self._add_photo()
        draft = os.path.join(self.tmp.name, 'd.json')
        self._save_draft_to(draft)
        os.remove(p)
        self._load_draft_from(draft)
        self.assertEqual(self.app.photo_files[cat], [])
        self.assertEqual(self.app.listboxes[cat].size(), 0)
        self.assertTrue(any(k == 'info' and p in m for k, m in self.shown))

    def test_draft_invalid_file_rejected_and_data_kept(self):
        self.app.personnel.insert(0, '保留我')
        bad = os.path.join(self.tmp.name, 'bad.json')
        with open(bad, 'w', encoding='utf-8') as f:
            f.write('{ not json')
        self._load_draft_from(bad)
        self.assertEqual(self.shown[-1][0], 'err')
        wrong = os.path.join(self.tmp.name, 'wrong.json')
        with open(wrong, 'w', encoding='utf-8') as f:
            f.write('{"version": 99}')
        self._load_draft_from(wrong)
        self.assertEqual(self.shown[-1][0], 'err')
        self.assertEqual(self.app.personnel.get(), '保留我')

    def test_draft_load_declined_keeps_current(self):
        draft = os.path.join(self.tmp.name, 'd.json')
        self._save_draft_to(draft)
        self.app.personnel.insert(0, '現有內容')
        self._load_draft_from(draft, confirm=False)
        self.assertEqual(self.app.personnel.get(), '現有內容')

    def test_mousewheel_scrolls_tab_from_child_widget(self):
        self.root.geometry('900x300')
        self.root.update()
        nb = next(w for w in self.root.winfo_children()[0].winfo_children() if isinstance(w, ttk.Notebook))
        nb.select(1)  # 道路步道
        self.root.update()
        tab = nb.nametowidget(nb.tabs()[1])
        canvas = next(w for w in tab.winfo_children() if isinstance(w, tk.Canvas))
        inner = canvas.nametowidget(canvas.winfo_children()[0])
        label = next(w for w in inner.winfo_children() if isinstance(w, ttk.Label))
        self.assertEqual(canvas.yview()[0], 0.0)
        label.event_generate('<MouseWheel>', delta=-120)
        self.root.update()
        self.assertGreater(canvas.yview()[0], 0.0)


if __name__ == '__main__':
    unittest.main()
