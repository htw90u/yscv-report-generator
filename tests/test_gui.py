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
