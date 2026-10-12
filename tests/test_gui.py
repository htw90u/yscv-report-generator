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
        self.autosave = os.path.join(self.tmp.name, 'autosave.json')
        self.app = g.ReportGeneratorGUI(self.root, autosave_path=self.autosave, restore_on_start=False)
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

    # ---- 未存內容提示 / 自動存檔
    def _close_with(self, answer, save_ok=True):
        """模擬按下關閉，回傳 (是否關閉視窗, 提示是否出現)。"""
        asked = []
        destroyed = []
        self.app.root.destroy = lambda: destroyed.append(True)
        self.app.save_draft = lambda: save_ok
        def ask(*a, **k):
            asked.append(True)
            return answer
        with mock.patch.object(messagebox, 'askyesnocancel', ask):
            self.app.on_close()
        return bool(destroyed), bool(asked)

    def test_close_clean_no_prompt(self):
        closed, asked = self._close_with(None)
        self.assertTrue(closed)
        self.assertFalse(asked)

    def test_template_only_change_is_not_dirty(self):
        self.app.template_path = None
        self.assertFalse(self.app.is_dirty())

    def test_close_dirty_cancel_stays(self):
        self.app.personnel.insert(0, 'x')
        self.assertTrue(self.app.is_dirty())
        closed, asked = self._close_with(None)
        self.assertTrue(asked)
        self.assertFalse(closed)

    def test_close_dirty_discard_closes_and_removes_autosave(self):
        self.app.personnel.insert(0, 'x')
        self.assertTrue(self.app.autosave_now())
        self.assertTrue(os.path.exists(self.autosave))
        closed, _ = self._close_with(False)
        self.assertTrue(closed)
        self.assertFalse(os.path.exists(self.autosave))

    def test_close_dirty_save_ok_closes_but_save_cancelled_stays(self):
        self.app.personnel.insert(0, 'x')
        closed, _ = self._close_with(True, save_ok=True)
        self.assertTrue(closed)
        self.app.personnel.insert(0, 'y')
        closed, _ = self._close_with(True, save_ok=False)
        self.assertFalse(closed)

    def test_save_draft_marks_clean_and_removes_autosave(self):
        self.app.personnel.insert(0, 'x')
        self.app.autosave_now()
        self._save_draft_to(os.path.join(self.tmp.name, 'd.json'))
        self.assertFalse(self.app.is_dirty())
        self.assertFalse(os.path.exists(self.autosave))
        self.app.personnel.insert(0, 'z')
        self.assertTrue(self.app.is_dirty())

    def test_autosave_skips_blank_clean_and_unchanged(self):
        self.assertFalse(self.app.autosave_now())  # 空白
        self.app.personnel.insert(0, 'x')
        self.assertTrue(self.app.autosave_now())
        self.assertFalse(self.app.autosave_now())  # 沒有新變動
        self.app.personnel.insert(0, 'y')
        self.assertTrue(self.app.autosave_now())
        self.assertFalse(os.path.exists(self.autosave + '.tmp'))

    def test_autosave_restore_roundtrip(self):
        self.app.personnel.insert(0, '王小明')
        self.app.text_fields['nature'].insert('1.0', '山羌')
        self.app.autosave_now()
        # 模擬程式當機後重開：新的 app 實例
        self.root.destroy()
        self.root = tk.Tk()
        app2 = g.ReportGeneratorGUI(self.root, autosave_path=self.autosave, restore_on_start=False)
        with mock.patch.object(messagebox, 'askyesno', lambda *a, **k: True):
            app2.offer_restore()
        self.assertEqual(app2.personnel.get(), '王小明')
        self.assertEqual(app2.text_fields['nature'].get('1.0', tk.END).strip(), '山羌')
        self.assertTrue(app2.is_dirty())  # 還原後仍視為未存
        self.app = app2

    def test_autosave_restore_declined_deletes_file(self):
        self.app.personnel.insert(0, 'x')
        self.app.autosave_now()
        with mock.patch.object(messagebox, 'askyesno', lambda *a, **k: False):
            self.app.offer_restore()
        self.assertFalse(os.path.exists(self.autosave))
        self.assertEqual(self.app.personnel.get(), 'x')

    def test_corrupt_autosave_is_discarded_silently(self):
        with open(self.autosave, 'w', encoding='utf-8') as f:
            f.write('{ broken')
        self.app.offer_restore()
        self.assertFalse(os.path.exists(self.autosave))
        self.assertEqual(self.shown, [])

    def test_autosave_tick_reschedules(self):
        self.app.personnel.insert(0, 'x')
        with mock.patch.object(self.root, 'after') as after:
            self.app._autosave_tick()
        self.assertTrue(os.path.exists(self.autosave))
        after.assert_called_once()

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
