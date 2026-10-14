"""GUI 流程測試（需要可用的顯示環境）。執行：python -m unittest discover -s tests"""
import gc
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


def style_lookup(root, style_name, option, state=None):
    return ttk.Style(root).lookup(style_name, option, state or [])


class GuiTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = tk.Tk()
        self.autosave = os.path.join(self.tmp.name, 'autosave.json')
        self.settings = os.path.join(self.tmp.name, 'settings.json')
        self.app = g.ReportGeneratorGUI(self.root, autosave_path=self.autosave, restore_on_start=False,
                                      settings_path=self.settings)
        self.app.template_path = TEMPLATE
        self.shown = []
        self.pdf_calls = []
        self.confirm_calls = []  # 產生前「有欄位未填」的確認視窗；測試預設按「是」

        def fake_confirm(title, message, **kwargs):
            self.confirm_calls.append((title, message, kwargs))
            return True


        def fake_pdf(docx, pdf):
            self.pdf_calls.append((docx, pdf))
            with open(pdf, 'wb') as f:
                f.write(b'%PDF-fake')

        for p in (
            mock.patch.object(messagebox, 'showerror', lambda t, m: self.shown.append(('err', m))),
            mock.patch.object(messagebox, 'showinfo', lambda t, m: self.shown.append(('info', m))),
            mock.patch.object(g, 'export_pdf', fake_pdf),
            mock.patch.object(messagebox, 'askyesno', fake_confirm),
        ):
            p.start()
            self.addCleanup(p.stop)

    def tearDown(self):
        self.root.destroy()
        self.app = self.root = None
        # 在主執行緒先回收 Tk 物件；否則可能在背景執行緒被 GC，造成 Tcl_AsyncDelete 當機
        gc.collect()
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
        self.assertIn('填表日期', self.shown[-1][1])
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

    def _wait_until_done(self, timeout=20):
        t0 = time.time()
        while time.time() - t0 < timeout and str(self.app.generate_btn['state']) == 'disabled':
            self.root.update()
            time.sleep(0.05)
        self.root.update()

    def test_default_filename_and_pdf_generated(self):
        self.fill_dates()
        self.app.location.insert(0, '排雲山莊')
        out = os.path.join(self.tmp.name, 'o.docx')
        captured = {}

        def ask(**kw):
            captured.update(kw)
            return out

        with mock.patch.object(filedialog, 'asksaveasfilename', ask), \
                mock.patch.object(g, 'open_file', lambda path: None):
            self.app.generate_report()
            self._wait_until_done()
        self.assertEqual(captured['initialfile'], '1151001-02_排雲山莊協勤回報單.docx')
        self.assertEqual(self.pdf_calls, [(out, os.path.join(self.tmp.name, 'o.pdf'))])
        self.assertTrue(os.path.exists(os.path.join(self.tmp.name, 'o.pdf')))

    def test_pdf_failure_is_reported_but_docx_kept(self):
        self.fill_dates()
        out = os.path.join(self.tmp.name, 'o.docx')

        def boom(docx, pdf):
            raise RuntimeError('沒有 Word')

        with mock.patch.object(filedialog, 'asksaveasfilename', lambda **k: out), \
                mock.patch.object(g, 'export_pdf', boom), \
                mock.patch.object(g, 'open_file', lambda path: None):
            self.app.generate_report()
            self._wait_until_done()
        self.assertTrue(os.path.exists(out))
        self.assertTrue(any(k == 'info' and 'PDF 未產生' in m for k, m in self.shown))

    def test_auto_open_is_always_on_and_has_no_option(self):
        self.assertFalse(hasattr(self.app, 'auto_open'))
        texts = [w.cget('text') for w in self._all_widgets(self.root) if isinstance(w, (ttk.Checkbutton, ttk.Button))]
        self.assertNotIn('產生後自動開啟', texts)
        self.fill_dates()
        out = os.path.join(self.tmp.name, 'o.docx')
        opened = []
        with mock.patch.object(filedialog, 'asksaveasfilename', lambda **k: out), \
                mock.patch.object(g, 'open_file', lambda path: opened.append(path)):
            self.app.generate_report()
            self._wait_until_done()
        self.assertEqual(opened, [out])

    def test_button_order_and_styles(self):
        frame = self.app.generate_btn.master
        labels = [w.cget('text') for w in frame.winfo_children() if isinstance(w, ttk.Button)]
        self.assertEqual(labels, ['載入草稿', '儲存草稿', '產生報告', '清除', '離開'])
        self.assertEqual(str(self.app.clear_btn.cget('style')), 'Danger.TButton')
        self.assertEqual(str(self.app.generate_btn.cget('style')), 'Primary.TButton')
        palette = g.THEMES[g.DEFAULT_THEME]
        for style_name, key in (('Primary.TButton', 'primary'), ('Danger.TButton', 'danger')):
            self.assertEqual(style_lookup(self.root, style_name, 'background'), palette[key])
            self.assertEqual(style_lookup(self.root, style_name, 'foreground'), palette[f'{key}_fg'])
            self.assertEqual(style_lookup(self.root, style_name, 'background', ['active']), palette[f'{key}_active'])
        # 離開維持預設灰色，不上色
        leave = [w for w in frame.winfo_children() if isinstance(w, ttk.Button) and w.cget('text') == '離開'][0]
        self.assertIn(str(leave.cget('style')), ('', 'TButton'))

    def test_open_file_failure_does_not_raise(self):
        with mock.patch.object(os, 'startfile', side_effect=OSError('boom'), create=True):
            g.open_file(os.path.join(self.tmp.name, 'nope.docx'))

    def test_generate_in_thread(self):
        self.fill_dates()
        self.app.text_fields['nature'].insert('1.0', '看到山羌\n• 已有bullet')
        out = os.path.join(self.tmp.name, 'o.docx')
        opened = []
        with mock.patch.object(filedialog, 'asksaveasfilename', lambda **k: out), \
                mock.patch.object(g, 'open_file', lambda path: opened.append(path)):
            self.app.generate_report()
            self.assertEqual(str(self.app.generate_btn['state']), 'disabled')
            self._wait_until_done()
        self.assertEqual(str(self.app.generate_btn['state']), 'normal')
        self.assertEqual(opened, [out])  # 產生後自動開啟
        self.assertTrue(self.root.winfo_exists())  # 工具沒有被關掉
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

        def ask(*a, **k):
            asked.append(True)
            return answer

        with mock.patch.object(self.app.root, 'destroy', lambda: destroyed.append(True)),                 mock.patch.object(self.app, 'save_draft', lambda: save_ok),                 mock.patch.object(messagebox, 'askyesnocancel', ask):
            self.app.on_close()
        return bool(destroyed), bool(asked)

    def test_close_clean_no_prompt(self):
        closed, asked = self._close_with(None)
        self.assertTrue(closed)
        self.assertFalse(asked)

    def test_close_dirty_cancel_stays(self):
        self.app.personnel.insert(0, 'x')
        self.assertTrue(self.app.is_dirty())
        closed, asked = self._close_with(None)
        self.assertTrue(asked)
        self.assertFalse(closed)

    def test_close_dirty_discard_keeps_autosave_with_latest_content(self):
        self.app.personnel.insert(0, 'x')
        self.assertTrue(self.app.autosave_now())
        self.app.personnel.insert(tk.END, 'y')  # 自動存檔之後又改了內容
        closed, _ = self._close_with(False)
        self.assertTrue(closed)
        self.assertTrue(os.path.exists(self.autosave))  # 保留
        import json
        with open(self.autosave, encoding='utf-8') as f:
            self.assertEqual(json.load(f)['fields']['personnel'], 'xy')  # 離開前補存了最新內容

    def test_close_dirty_discard_with_blank_screen_removes_stale_autosave(self):
        self.app.personnel.insert(0, 'x')
        self.app.autosave_now()
        self.app._saved_signature = 'something else'  # 假裝上次存檔的內容和現在（空白）不同 -> 算有改動
        self.app.personnel.delete(0, tk.END)
        self.assertTrue(self.app.is_dirty())
        closed, _ = self._close_with(False)
        self.assertTrue(closed)
        self.assertFalse(os.path.exists(self.autosave))

    def test_close_clean_removes_autosave(self):
        with open(self.autosave, 'w', encoding='utf-8') as f:
            f.write('{}')
        closed, asked = self._close_with(None)
        self.assertTrue(closed)
        self.assertFalse(asked)
        self.assertFalse(os.path.exists(self.autosave))

    def test_autosave_interval_is_30_seconds(self):
        self.assertEqual(g.AUTOSAVE_INTERVAL_MS, 30_000)

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
        app2 = g.ReportGeneratorGUI(self.root, autosave_path=self.autosave, restore_on_start=False,
                                      settings_path=self.settings)
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

    def test_draft_does_not_store_template_and_ignores_old_template_path(self):
        draft = os.path.join(self.tmp.name, 'd.json')
        self._save_draft_to(draft)
        import json
        with open(draft, encoding='utf-8') as f:
            data = json.load(f)
        self.assertNotIn('template_path', data)
        data['template_path'] = os.path.join(self.tmp.name, 'old.docx')  # 舊版草稿帶有這個欄位
        with open(draft, 'w', encoding='utf-8') as f:
            json.dump(data, f)
        self._load_draft_from(draft)
        self.assertEqual(self.app.template_path, TEMPLATE)

    def test_template_is_fixed_no_selection_ui(self):
        self.assertEqual(g.TEMPLATE_PATH, TEMPLATE)
        self.assertFalse(hasattr(self.app, 'select_template'))
        self.assertFalse(hasattr(self.app, 'template_label'))
        self.assertTrue(os.path.exists(g.TEMPLATE_PATH))

    def test_missing_template_blocks_generation(self):
        self.app.template_path = os.path.join(self.tmp.name, 'missing.docx')
        self.fill_dates()
        self.app.generate_report()
        self.assertEqual(self.shown[-1][0], 'err')
        self.assertIn('找不到模板', self.shown[-1][1])

    def test_output_same_as_template_rejected(self):
        import shutil
        tpl = os.path.join(self.tmp.name, 'tpl.docx')
        shutil.copy(TEMPLATE, tpl)
        self.app.template_path = tpl
        self.fill_dates()
        with mock.patch.object(filedialog, 'asksaveasfilename', lambda **k: tpl):
            self.app.generate_report()
        self.assertEqual(self.shown[-1][0], 'err')
        self.assertIn('模板', self.shown[-1][1])
        self.assertEqual(str(self.app.generate_btn['state']), 'normal')

    # ---- 編輯區項目符號 / 編號、字體
    def test_bullet_and_number_buttons_on_selection(self):
        w = self.app.text_fields['nature']
        w.insert('1.0', 'a\nb\nc')
        w.tag_add(tk.SEL, '1.0', '2.end')
        self.app.apply_list(w, g.toggle_bullets)
        self.assertEqual(w.get('1.0', 'end-1c'), '• a\n• b\nc')
        self.app.apply_list(w, g.toggle_numbers)
        self.assertEqual(w.get('1.0', 'end-1c'), '1. a\n2. b\nc')
        self.app.apply_list(w, g.toggle_numbers)  # 再按一次取消
        self.assertEqual(w.get('1.0', 'end-1c'), 'a\nb\nc')

    def test_buttons_apply_to_current_line_without_selection(self):
        w = self.app.text_fields['nature']
        w.insert('1.0', 'a\nb')
        w.mark_set(tk.INSERT, '2.1')
        self.app.apply_list(w, g.toggle_bullets)
        self.assertEqual(w.get('1.0', 'end-1c'), 'a\n• b')

    def _bullet_menu(self, widget):
        for w in self._all_widgets(self.root):
            if isinstance(w, ttk.Menubutton) and getattr(w, 'menu', None) is not None:
                # 同一個欄位的選單：以兄弟元件裡的文字區比對
                if widget.master in w.master.master.winfo_children():  # ScrolledText 的文字區外面還包了一層 Frame
                    return w.menu
        self.fail('找不到項目符號選單')

    def _all_widgets(self, parent):
        for child in parent.winfo_children():
            yield child
            yield from self._all_widgets(child)

    def test_bullet_dropdown_lists_all_symbols_and_applies_choice(self):
        w = self.app.text_fields['nature']
        menu = self._bullet_menu(w)
        labels = [menu.entrycget(i, 'label') for i in range(menu.index('end') + 1)
                  if menu.type(i) == 'command']
        self.assertEqual(len(labels), len(g.BULLET_CHOICES) + 2)  # 全部符號 + 編號 + 取消
        self.assertTrue(any('編號' in lbl and lbl.startswith('1.') for lbl in labels))
        self.assertTrue(any(lbl.startswith('◆') for lbl in labels))
        w.insert('1.0', 'a\nb')
        w.tag_add(tk.SEL, '1.0', 'end-1c')
        menu.invoke([i for i in range(menu.index('end') + 1) if menu.type(i) == 'command'
                     and menu.entrycget(i, 'label').startswith('◆')][0])
        self.assertEqual(w.get('1.0', 'end-1c'), '◆ a\n◆ b')
        menu.invoke(menu.index('end'))  # 最後一項：取消項目符號
        self.assertEqual(w.get('1.0', 'end-1c'), 'a\nb')

    def test_numbering_is_in_the_same_dropdown_and_no_separate_button(self):
        w = self.app.text_fields['nature']
        menu = self._bullet_menu(w)
        idx = [i for i in range(menu.index('end') + 1) if menu.type(i) == 'command'
               and menu.entrycget(i, 'label').startswith('1.')][0]
        w.insert('1.0', 'a\nb')
        w.tag_add(tk.SEL, '1.0', 'end-1c')
        menu.invoke(idx)
        self.assertEqual(w.get('1.0', 'end-1c'), '1. a\n2. b')
        menu.invoke(idx)  # 再選一次取消
        self.assertEqual(w.get('1.0', 'end-1c'), 'a\nb')
        buttons = [c.cget('text') for c in self._all_widgets(self.root) if isinstance(c, ttk.Button)]
        self.assertNotIn('1. 編號', buttons)

    # ---- 預設照片說明
    def test_default_caption_is_editable_and_used_in_report(self):
        self.fill_dates()
        self.assertEqual(self.app.default_caption.get(), g.DEFAULT_CAPTION)
        cat, lb, (a, b) = self._photos(names=('a.jpg', 'b.jpg'))
        self.app.photo_captions[cat][a] = '有說明'
        self.app.default_caption.set('  圖說待補  ')
        data, err = self.app.collect_data()
        self.assertIsNone(err)
        self.assertEqual(data['default_caption'], '圖說待補')  # 前後空白會去掉
        out = os.path.join(self.tmp.name, 'o.docx')
        g.build_report(TEMPLATE, data, {c: list(p) for c, p in self.app.photo_files.items()},
                       {c: dict(v) for c, v in self.app.photo_captions.items()}, out)
        text = '\n'.join(c.text for c in unique_cells(Document(out)))
        self.assertIn('有說明', text)
        self.assertIn('圖說待補', text)
        self.assertNotIn('[照片說明]', text)

    def test_default_caption_can_be_blank(self):
        self.fill_dates()
        cat, lb, (a,) = self._photos(names=('a.jpg',))
        self.app.default_caption.set('')
        data, _ = self.app.collect_data()
        out = os.path.join(self.tmp.name, 'o.docx')
        g.build_report(TEMPLATE, data, {c: list(p) for c, p in self.app.photo_files.items()}, {}, out)
        self.assertNotIn('[照片說明]', '\n'.join(c.text for c in unique_cells(Document(out))))

    def test_default_caption_saved_in_draft_marks_dirty_and_resets_on_clear(self):
        self.app._saved_signature = self.app._signature()
        self.app.default_caption.set('自訂預設')
        self.assertTrue(self.app.is_dirty())
        draft = os.path.join(self.tmp.name, 'd.json')
        self._save_draft_to(draft)
        self.app.default_caption.set('別的')
        self._load_draft_from(draft)
        self.assertEqual(self.app.default_caption.get(), '自訂預設')
        with mock.patch.object(messagebox, 'askyesno', lambda *a, **k: True):
            self.app.clear_all()
        self.assertEqual(self.app.default_caption.get(), g.DEFAULT_CAPTION)

    def test_default_caption_is_blank_by_default(self):
        self.assertEqual(g.DEFAULT_CAPTION, '')
        self.assertEqual(self.app.default_caption.get(), '')

    def test_legacy_default_caption_in_old_draft_becomes_blank(self):
        import json
        draft = os.path.join(self.tmp.name, 'legacy.json')
        self._save_draft_to(draft)
        with open(draft, encoding='utf-8') as f:
            data = json.load(f)
        data['default_caption'] = '[照片說明]'
        with open(draft, 'w', encoding='utf-8') as f:
            json.dump(data, f)
        self._load_draft_from(draft)
        self.assertEqual(self.app.default_caption.get(), '')

    def test_caption_hint_text(self):
        texts = [w.cget('text') for w in self._all_widgets(self.root) if isinstance(w, ttk.Label)]
        self.assertIn('（雙擊檔名輸入照片說明）', texts)
        self.assertFalse(any('可留空' in t for t in texts))

    # ---- 分頁外觀
    # ---- 填表日期自動帶入（民國）
    def test_roc_date_parts(self):
        import datetime as dt
        self.assertEqual(g.roc_date_parts(dt.datetime(2026, 10, 14)), ('115', '10', '14'))
        self.assertEqual(g.roc_date_parts(dt.datetime(2027, 1, 5)), ('116', '1', '5'))
        self.assertEqual(g.roc_date_parts(dt.datetime(2012, 12, 31)), ('101', '12', '31'))

    def test_new_form_gets_todays_roc_date(self):
        import datetime as dt
        self.root.destroy()
        self.root = tk.Tk()
        with mock.patch.object(g, 'current_date', lambda: dt.datetime(2026, 10, 14, 9, 30)):
            self.app = g.ReportGeneratorGUI(self.root, autosave_path=self.autosave, restore_on_start=False,
                                      settings_path=self.settings)
        self.assertEqual((self.app.year.get(), self.app.month.get(), self.app.day.get()), ('115', '10', '14'))
        labels = [w.cget('text') for w in self._all_widgets(self.root) if isinstance(w, ttk.Label)]
        self.assertIn('1. 填表日期:', labels)
        self.assertFalse(any('年月日' in t for t in labels))

    def test_auto_date_alone_is_not_unsaved_content(self):
        self.assertTrue(self.app.year.get())
        self.assertFalse(self.app.is_dirty())
        self.assertFalse(self.app.autosave_now())  # 只有自動日期不算內容，不產生自動存檔
        closed, asked = self._close_with(None)
        self.assertTrue(closed)
        self.assertFalse(asked)

    def test_autosave_starts_once_real_content_is_entered(self):
        self.app.personnel.insert(0, 'x')
        self.assertTrue(self.app.is_dirty())
        self.assertTrue(self.app.autosave_now())

    def test_clear_all_regenerates_date(self):
        import datetime as dt
        for name in g.REPORT_DATE_FIELDS:
            getattr(self.app, name).delete(0, tk.END)
        self.app.year.insert(0, '100')
        with mock.patch.object(messagebox, 'askyesno', lambda *a, **k: True), \
                mock.patch.object(g, 'current_date', lambda: dt.datetime(2026, 10, 20)):
            self.app.clear_all()
        self.assertEqual((self.app.year.get(), self.app.month.get(), self.app.day.get()), ('115', '10', '20'))

    def test_loading_draft_keeps_its_original_fill_date(self):
        import datetime as dt
        draft = os.path.join(self.tmp.name, 'd.json')
        for name, value in zip(g.REPORT_DATE_FIELDS, ('115', '9', '30')):  # 第一次填表是 9/30
            getattr(self.app, name).delete(0, tk.END)
            getattr(self.app, name).insert(0, value)
        self._save_draft_to(draft)
        with mock.patch.object(g, 'current_date', lambda: dt.datetime(2026, 10, 14)):  # 今天已經是 10/14
            self._load_draft_from(draft)
        self.assertEqual((self.app.year.get(), self.app.month.get(), self.app.day.get()), ('115', '9', '30'))

    def test_draft_without_date_gets_todays_date(self):
        import json, datetime as dt
        draft = os.path.join(self.tmp.name, 'nodate.json')
        self._save_draft_to(draft)
        with open(draft, encoding='utf-8') as f:
            data = json.load(f)
        for name in g.REPORT_DATE_FIELDS:
            data['fields'][name] = ''
        with open(draft, 'w', encoding='utf-8') as f:
            json.dump(data, f)
        with mock.patch.object(g, 'current_date', lambda: dt.datetime(2026, 11, 2)):
            self._load_draft_from(draft)
        self.assertEqual((self.app.year.get(), self.app.month.get(), self.app.day.get()), ('115', '11', '2'))

    def test_auto_date_flows_into_report(self):
        import datetime as dt
        self.root.destroy()
        self.root = tk.Tk()
        with mock.patch.object(g, 'current_date', lambda: dt.datetime(2026, 10, 14)):
            self.app = g.ReportGeneratorGUI(self.root, autosave_path=self.autosave, restore_on_start=False,
                                      settings_path=self.settings)
        for e, v in [(self.app.start_year, '115'), (self.app.start_month, '10'), (self.app.start_day, '8'),
                     (self.app.end_year, '115'), (self.app.end_month, '10'), (self.app.end_day, '11')]:
            e.insert(0, v)
        data, err = self.app.collect_data()
        self.assertIsNone(err)
        out = os.path.join(self.tmp.name, 'o.docx')
        g.build_report(TEMPLATE, data, {c: [] for c in self.app.photo_files}, {}, out)
        text = '\n'.join(c.text for c in unique_cells(Document(out)))
        self.assertIn('115年10月14日', text)

    # ---- 清除前的確認
    def _clear_with_answer(self, answer):
        calls = []

        def ask(title, message, **kwargs):
            calls.append((title, message, kwargs))
            return answer

        with mock.patch.object(messagebox, 'askyesno', ask):
            self.app.clear_all()
        return calls

    def test_clear_asks_for_confirmation_with_safe_default(self):
        self.app.personnel.insert(0, '保留')
        calls = self._clear_with_answer(False)
        self.assertEqual(len(calls), 1)
        _, _, kwargs = calls[0]
        self.assertEqual(kwargs.get('default'), 'no')  # 預設「否」，按 Enter 不會誤清除
        self.assertEqual(kwargs.get('icon'), 'warning')
        self.assertEqual(self.app.personnel.get(), '保留')  # 選否：什麼都不動

    def test_clear_confirmed_clears_everything(self):
        self.app.personnel.insert(0, '清掉')
        cat, lb, _ = self._photos()
        self._clear_with_answer(True)
        self.assertEqual(self.app.personnel.get(), '')
        self.assertEqual(lb.size(), 0)

    def test_clear_warning_mentions_scope_and_irreversibility(self):
        _, message, _ = self._clear_with_answer(False)[0]
        for word in ('欄位', '照片', '說明', '無法復原'):
            self.assertIn(word, message)

    def test_clear_warns_extra_when_there_is_unsaved_content(self):
        _, clean_message, _ = self._clear_with_answer(False)[0]
        self.assertNotIn('尚未儲存', clean_message)
        self.app.personnel.insert(0, 'x')
        _, dirty_message, _ = self._clear_with_answer(False)[0]
        self.assertIn('尚未儲存', dirty_message)
        self.assertIn('儲存草稿', dirty_message)
        # 存成草稿之後，又回到一般確認
        self._save_draft_to(os.path.join(self.tmp.name, 'd.json'))
        _, saved_message, _ = self._clear_with_answer(False)[0]
        self.assertNotIn('尚未儲存', saved_message)

    # ---- 外觀主題
    @staticmethod
    def _luminance(color):
        def channel(v):
            v /= 255
            return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
        r, gr, b = (int(color[i:i + 2], 16) for i in (1, 3, 5))
        return 0.2126 * channel(r) + 0.7152 * channel(gr) + 0.0722 * channel(b)

    def _contrast(self, fg, bg):
        hi, lo = sorted((self._luminance(fg), self._luminance(bg)), reverse=True)
        return (hi + 0.05) / (lo + 0.05)

    def test_themes_available(self):
        self.assertEqual(set(g.THEMES), {'柔和淺色', '色弱友善', '高對比'})
        self.assertIn(g.DEFAULT_THEME, g.THEMES)
        self.assertEqual(tuple(self.app.theme_box.cget('values')), tuple(g.THEMES))
        self.assertEqual(self.app.theme_var.get(), g.DEFAULT_THEME)

    def test_every_theme_text_is_readable_wcag_aa(self):
        for name, p in g.THEMES.items():
            pairs = [('primary', p['primary_fg'], p['primary']), ('danger', p['danger_fg'], p['danger']),
                     ('heading', p['heading_fg'], p['bg']), ('field label', p['field_fg'], p['bg']),
                     ('active tab', p['tab_active_fg'], p['tab_active']), ('idle tab', p['tab_idle_fg'], p['tab_idle']),
                     ('primary hover', p['primary_fg'], p['primary_active']), ('danger hover', p['danger_fg'], p['danger_active'])]
            for label, fg, bg in pairs:
                with self.subTest(theme=name, item=label):
                    self.assertGreaterEqual(self._contrast(fg, bg), 4.5)

    def test_default_theme_colors_are_light(self):
        p = g.THEMES[g.DEFAULT_THEME]
        for key in ('bg', 'tab_idle', 'tab_active', 'primary', 'danger'):
            with self.subTest(key=key):
                self.assertGreater(self._luminance(p[key]), 0.5)  # 淺色，不是濃重的底色

    def test_colorblind_theme_avoids_red_green_pair(self):
        p = g.THEMES['色弱友善']
        pr = [int(p['primary_border'][i:i + 2], 16) for i in (1, 3, 5)]
        dg = [int(p['danger_border'][i:i + 2], 16) for i in (1, 3, 5)]
        self.assertGreater(pr[2], pr[0])   # 主要按鈕：藍色（B 大於 R）
        self.assertGreater(pr[2], pr[1])
        self.assertGreater(dg[0], dg[2])   # 危險按鈕：橘色（R 大於 B，且不是綠色）
        self.assertGreater(dg[0], dg[1])
        self.assertGreater(dg[1], dg[2])
        self.assertNotEqual(p['primary'], p['danger'])

    def test_apply_theme_changes_styles_and_widgets_and_is_remembered(self):
        import json
        self.app.apply_theme('色弱友善')
        p = g.THEMES['色弱友善']
        self.assertEqual(style_lookup(self.root, 'Primary.TButton', 'background'), p['primary'])
        self.assertEqual(style_lookup(self.root, 'Danger.TButton', 'background'), p['danger'])
        self.assertEqual(style_lookup(self.root, 'TNotebook.Tab', 'background', ['selected']), p['tab_active'])
        self.assertEqual(str(self.root.cget('background')), p['bg'])
        canvases = [w for w in self._all_widgets(self.root) if isinstance(w, tk.Canvas)]
        self.assertTrue(canvases)
        self.assertTrue(all(str(c.cget('background')) == p['bg'] for c in canvases))
        with open(self.settings, encoding='utf-8') as f:
            self.assertEqual(json.load(f), {'theme': '色弱友善'})
        # 重新開啟時沿用上次選的主題
        self.root.destroy()
        self.root = tk.Tk()
        self.app = g.ReportGeneratorGUI(self.root, autosave_path=self.autosave, restore_on_start=False,
                                        settings_path=self.settings)
        self.assertEqual(self.app.theme_name, '色弱友善')
        self.assertEqual(self.app.theme_var.get(), '色弱友善')
        self.assertEqual(style_lookup(self.root, 'Primary.TButton', 'background'), p['primary'])

    def test_theme_combobox_selection_applies_theme(self):
        self.app.theme_var.set('高對比')
        self.app.theme_box.event_generate('<<ComboboxSelected>>')
        self.root.update()
        self.assertEqual(self.app.theme_name, '高對比')
        self.assertEqual(style_lookup(self.root, 'Danger.TButton', 'background'), g.THEMES['高對比']['danger'])

    def test_unknown_or_broken_settings_fall_back_to_default(self):
        for content in ('{ broken', '{"theme": "不存在"}', '[]'):
            with open(self.settings, 'w', encoding='utf-8') as f:
                f.write(content)
            self.assertEqual(self.app._load_theme_setting(), g.DEFAULT_THEME, content)
        os.remove(self.settings)
        self.assertEqual(self.app._load_theme_setting(), g.DEFAULT_THEME)
        self.app.apply_theme('不存在')  # 無效名稱不處理
        self.assertEqual(self.app.theme_name, g.DEFAULT_THEME)

    def test_switching_theme_does_not_make_form_dirty(self):
        self.app.apply_theme('高對比')
        self.assertFalse(self.app.is_dirty())

    # ---- 日期欄位間距
    def test_date_row_spacing_is_symmetric_and_widths_equal(self):
        self.root.update()
        for prefix in ('', 'start_', 'end_'):
            self.assertEqual(self.app.__dict__[f'{prefix}year'].winfo_width(), self.app.year.winfo_width())
            self.assertEqual(self.app.__dict__[f'{prefix}month'].winfo_width(), self.app.month.winfo_width())
            self.assertEqual(self.app.__dict__[f'{prefix}day'].winfo_width(), self.app.day.winfo_width())
        frame = self.app.start_year.master
        for label in [w for w in frame.winfo_children() if isinstance(w, ttk.Label)]:
            padx = label.pack_info()['padx']
            left, right = (padx if isinstance(padx, tuple) else (padx, padx))
            with self.subTest(label=label.cget('text')):
                self.assertEqual(int(left), int(right))  # 單位字前後間距相同
        texts = [w.cget('text') for w in frame.winfo_children() if isinstance(w, ttk.Label)]
        self.assertEqual(texts, ['年', '月', '日', '至', '年', '月', '日'])  # 「至」是獨立的一個字，不再和「日」黏在一起

    # ---- 照片分頁分組與視窗高度
    def test_natural_resources_moved_to_second_photo_tab(self):
        nature = '自然資源狀況'
        self.assertNotIn(nature, g.PHOTO_CATEGORIES_1)
        self.assertIn(nature, g.PHOTO_CATEGORIES_2)
        first = self.app.previews[g.PHOTO_CATEGORIES_1[0]]
        self.assertIsNot(self.app.previews[nature], first)
        # 實際畫面：「自然資源狀況」的清單在「服勤照片(下)」分頁裡
        tab = self.app._tab_by_title('服勤照片(下)')
        self.assertTrue(str(self.app.listboxes[nature]).startswith(tab))
        tab_up = self.app._tab_by_title('服勤照片(上)')
        self.assertFalse(str(self.app.listboxes[nature]).startswith(tab_up))

    def test_word_order_does_not_depend_on_tab_grouping(self):
        self.assertEqual(list(self.app.photo_files), g.PHOTO_CATEGORIES)
        self.assertEqual(g.PHOTO_CATEGORIES.index('自然資源狀況'), 2)  # 仍在「設施」之後、「環境清理」之前
        self.assertEqual(sorted(g.PHOTO_CATEGORIES), sorted(g.PHOTO_CATEGORIES_1 + g.PHOTO_CATEGORIES_2))
        self.assertEqual(len(set(g.PHOTO_CATEGORIES_1) & set(g.PHOTO_CATEGORIES_2)), 0)

    def test_report_keeps_original_chapter_order(self):
        import glob
        self.fill_dates()
        paths = {}
        for i, cat in enumerate(g.PHOTO_CATEGORIES):
            p = os.path.join(self.tmp.name, f'c{i}.jpg')
            Image.new('RGB', (400, 300), 'red').save(p)
            self.app.photo_files[cat].append(p)
            paths[cat] = p
        data, err = self.app.collect_data()
        self.assertIsNone(err)
        out = os.path.join(self.tmp.name, 'o.docx')
        g.build_report(TEMPLATE, data, {c: list(v) for c, v in self.app.photo_files.items()}, {}, out)
        heads = []
        for table in Document(out).tables[1:]:
            for row in table.rows:
                text = row.cells[0].text.strip()
                if text in g.PHOTO_CATEGORIES and (not heads or heads[-1] != text):
                    heads.append(text)
        self.assertEqual(heads, g.PHOTO_CATEGORIES)

    def test_draft_roundtrip_with_moved_category(self):
        cat = '自然資源狀況'
        lb = self.app.listboxes[cat]
        p = os.path.join(self.tmp.name, 'n.jpg')
        Image.new('RGB', (30, 20)).save(p)
        self.app.photo_files[cat].append(p)
        lb.insert(tk.END, 'n.jpg')
        draft = os.path.join(self.tmp.name, 'd.json')
        self._save_draft_to(draft)
        self.app.clear_photos(cat, lb)
        self._load_draft_from(draft)
        self.assertEqual(self.app.photo_files[cat], [p])
        self.assertEqual(lb.size(), 1)


    def test_scrollbar_appears_only_when_content_does_not_fit(self):
        nb = self.app.notebook
        title = '服勤照片(上)'
        canvas, inner, scrollbar = self.app._scroll_tabs[title]
        nb.select(self.app._tab_by_title(title))
        listboxes = [lb for c, lb in self.app.listboxes.items() if c in g.PHOTO_CATEGORIES_1]
        original = [int(lb.cget('height')) for lb in listboxes]
        try:
            for lb in listboxes:  # 內容變少 -> 放得下 -> 隱藏
                lb.configure(height=1)
            self.root.geometry(f'{self.root.winfo_width()}x{self.root.winfo_height()}')
            self.root.update()
            self.assertLessEqual(inner.winfo_reqheight(), canvas.winfo_height())
            self.assertEqual(scrollbar.winfo_manager(), '', '內容放得下時要隱藏捲動軸')
            for lb in listboxes:  # 內容變多 -> 放不下 -> 顯示
                lb.configure(height=40)
            self.root.update()
            self.assertGreater(inner.winfo_reqheight(), canvas.winfo_height())
            self.assertNotEqual(scrollbar.winfo_manager(), '', '內容放不下時要顯示捲動軸')
        finally:
            for lb, rows in zip(listboxes, original):
                lb.configure(height=rows)

    def test_short_screen_falls_back_to_compact_photo_lists(self):
        """螢幕不夠高、視窗加高也放不下時，改把照片清單縮成較少的列數，並把視窗限制在螢幕高度內。"""
        for listbox in self.app.listboxes.values():
            listbox.configure(height=g.PHOTO_LIST_ROWS)
        self.root.geometry('1050x820')
        self.root.update()
        before = self.app._photo_tabs_extra_height()
        with mock.patch.object(self.root, 'winfo_screenheight', lambda: 930):
            self.app._fit_window_height()
        self.root.update()
        self.assertTrue(all(int(lb.cget('height')) == g.PHOTO_LIST_ROWS_COMPACT for lb in self.app.listboxes.values()))
        self.assertLessEqual(self.root.winfo_height(), 930 - 90)
        self.assertLess(self.app._photo_tabs_extra_height(), before)  # 比原本好（縮小清單 + 加高視窗）

    def test_lists_stay_normal_size_when_the_window_can_hold_everything(self):
        """視窗放得下兩個照片分頁時，不需要縮小照片清單。"""
        for lb in self.app.listboxes.values():
            lb.configure(height=g.PHOTO_LIST_ROWS)
        with mock.patch.object(self.app, '_photo_tabs_extra_height', lambda: 0):
            self.app._fit_window_height()
        self.assertTrue(all(int(lb.cget('height')) == g.PHOTO_LIST_ROWS for lb in self.app.listboxes.values()))

    def test_photo_tabs_need_no_horizontal_room_beyond_window(self):
        for title in g.PHOTO_TAB_TITLES:
            canvas, inner, _ = self.app._scroll_tabs[title]
            self.app.notebook.select(self.app._tab_by_title(title))
            self.root.update()
            with self.subTest(tab=title):
                self.assertLessEqual(inner.winfo_reqwidth(), canvas.winfo_width())

    def _linespace(self, widget):
        import tkinter.font as tkfont
        font = widget.cget('font') or style_lookup(self.root, str(widget.cget('style')), 'font')
        return tkfont.Font(root=self.root, font=font).metrics('linespace')

    def test_long_section_titles_stay_on_one_line(self):
        """三、四、五、八 這些很長的標題，在預設視窗寬度下要一行顯示（不能被固定寬度截斷換行）。"""
        nb = self.app.notebook
        long_titles = ('三、', '四、', '五、', '八、')
        found = []
        for tab_index in (3, 4):  # 自然資源與遊客、意見與心得建議
            nb.select(tab_index)
            self.root.update()
            tab = nb.nametowidget(nb.tabs()[tab_index])
            for w in self._all_widgets(tab):
                if isinstance(w, ttk.Label) and w.cget('text').startswith(long_titles):
                    found.append(w)
                    with self.subTest(title=w.cget('text')[:8]):
                        self.assertLessEqual(w.winfo_reqheight(), self._linespace(w) + 6,
                                             '標題被換成兩行了：' + w.cget('text')[:12])
        self.assertEqual(len(found), 4)

    def test_section_title_wraps_only_when_window_is_too_narrow_and_dropdown_stays_visible(self):
        nb = self.app.notebook
        nb.select(3)
        self.root.update()
        tab = nb.nametowidget(nb.tabs()[3])
        title = next(w for w in self._all_widgets(tab)
                     if isinstance(w, ttk.Label) and w.cget('text').startswith('三、'))
        button = next(w for w in title.master.winfo_children() if isinstance(w, ttk.Menubutton))
        original = self.root.winfo_height()
        try:
            self.root.geometry(f'620x{original}')
            self.root.update()
            self.assertGreater(title.winfo_reqheight(), self._linespace(title) + 6)  # 太窄時才換行
            header = title.master
            self.assertLessEqual(button.winfo_x() + button.winfo_width(), header.winfo_width())  # 下拉按鈕沒被擠出去
        finally:
            self.root.geometry(f'1050x{original}')
            self.root.update()

    # ---- 文字清晰度（對比）與外觀選單置中
    def test_text_colors_have_high_contrast_in_every_theme(self):
        for name, p in g.THEMES.items():
            for label, fg, bg in (('body text', p['text'], '#ffffff'), ('body text on form', p['text'], p['bg']),
                                  ('hint text', p['muted'], p['bg'])):
                with self.subTest(theme=name, item=label):
                    self.assertGreaterEqual(self._contrast(fg, bg), 7.0)  # WCAG AAA 等級

    def test_default_text_and_hint_use_theme_colors_not_pale_gray(self):
        p = g.THEMES[g.DEFAULT_THEME]
        self.assertEqual(style_lookup(self.root, 'TLabel', 'foreground'), p['text'])
        self.assertEqual(style_lookup(self.root, 'TEntry', 'foreground'), p['text'])
        self.assertEqual(style_lookup(self.root, 'Hint.TLabel', 'foreground'), p['muted'])
        for w in self._all_widgets(self.root):
            if isinstance(w, ttk.Label) and w.cget('text').startswith(('（點選預覽', '（雙擊檔名')):
                self.assertEqual(str(w.cget('style')), 'Hint.TLabel')
            if isinstance(w, (tk.Text, tk.Listbox)):
                self.assertEqual(str(w.cget('foreground')), p['text'])

    def test_theme_change_updates_native_text_colors(self):
        self.app.apply_theme('高對比')
        for w in self._all_widgets(self.root):
            if isinstance(w, (tk.Text, tk.Listbox)):
                self.assertEqual(str(w.cget('foreground')), g.THEMES['高對比']['text'])

    def test_theme_dropdown_text_and_options_are_centered(self):
        self.assertEqual(str(self.app.theme_box.cget('justify')), 'center')  # 選取後顯示的文字置中
        popdown = self.root.tk.eval(f'ttk::combobox::PopdownWindow {self.app.theme_box}')
        self.assertEqual(str(self.root.tk.call(f'{popdown}.f.l', 'cget', '-justify')), 'center')  # 展開的選項置中
        self.assertEqual(tuple(self.app.theme_box.cget('values')), tuple(g.THEMES))

    def test_dpi_awareness_helper_is_safe(self):
        with mock.patch.object(g.sys, 'platform', 'linux'):
            self.assertFalse(g.enable_dpi_awareness())
        self.assertIn(g.enable_dpi_awareness.__name__, dir(g))

    def test_font_is_iansui(self):
        self.assertEqual(g.FONT_FAMILY, 'Iansui')
        self.assertEqual(g.FONT_FILES, ('Iansui-Regular.ttf',))
        self.assertEqual(g.FONT_LICENSE, 'OFL-Iansui.txt')
        self.assertEqual(self.app.font_family, 'Iansui')
        old = [n for n in os.listdir(g.FONT_DIR) if any(k in n for k in ('LXGW', 'Cactus', 'Chocolate', 'Chiron'))]
        self.assertEqual(old, [])  # 舊字型已移除

    def test_font_line_height_is_compact_enough_for_photo_tabs(self):
        """字型的行高不能是一般字型的兩倍（之前試過的字型就是這樣），否則服勤照片分頁一定要捲動。"""
        import tkinter.font as tkfont
        font = tkfont.nametofont('TkDefaultFont', root=self.root)
        self.assertLessEqual(font.metrics('linespace'), 30)

    def test_form_uses_no_bold_anywhere(self):
        """表單不使用粗體：所有標籤樣式與介面字型都是一般字重，強調改用字級與顏色。"""
        import tkinter.font as tkfont
        for name in ('TkDefaultFont', 'TkTextFont', 'TkFixedFont', 'TkMenuFont'):
            self.assertEqual(tkfont.nametofont(name, root=self.root).cget('weight'), 'normal')
        for style_name in ('Heading.TLabel', 'Field.TLabel'):
            font = tkfont.Font(root=self.root, font=style_lookup(self.root, style_name, 'font'))
            self.assertEqual(font.cget('weight'), 'normal', style_name)
        for w in self._all_widgets(self.root):
            if isinstance(w, ttk.Label):
                font = str(w.cget('font'))
                self.assertNotIn('bold', font, f'{w.cget("text")[:10]} 使用了粗體')

    def test_headings_are_distinguished_by_size_and_color_instead_of_bold(self):
        import tkinter.font as tkfont
        palette = g.THEMES[g.DEFAULT_THEME]
        heading = tkfont.Font(root=self.root, font=style_lookup(self.root, 'Heading.TLabel', 'font'))
        field = tkfont.Font(root=self.root, font=style_lookup(self.root, 'Field.TLabel', 'font'))
        body = tkfont.nametofont('TkDefaultFont', root=self.root)
        self.assertGreater(abs(heading.cget('size')), abs(field.cget('size')))
        self.assertGreater(abs(field.cget('size')), abs(body.cget('size')))
        self.assertEqual(style_lookup(self.root, 'Heading.TLabel', 'foreground'), palette['heading_fg'])
        self.assertEqual(style_lookup(self.root, 'Field.TLabel', 'foreground'), palette['field_fg'])

    def test_all_labels_use_the_shared_label_styles(self):
        styles = {str(w.cget('style')) for w in self._all_widgets(self.root)
                  if isinstance(w, ttk.Label) and w.cget('text').startswith(('1. 填表', '三、', '九、服勤照片', '二、設施'))}
        self.assertEqual(styles, {'Field.TLabel', 'Heading.TLabel'})

    # ---- 產生前檢查未填欄位
    def _fill_everything(self):
        self.fill_dates()
        self.app.personnel.insert(0, '王小明')
        self.app.location.insert(0, '排雲山莊')
        for w in self.app.text_fields.values():
            w.insert('1.0', '內容')
        for i, cat in enumerate(g.PHOTO_CATEGORIES):
            p = os.path.join(self.tmp.name, f'p{i}.jpg')
            Image.new('RGB', (40, 30)).save(p)
            self.app.photo_files[cat].append(p)

    def test_blank_fields_listed_per_tab(self):
        self.fill_dates()
        data, _ = self.app.collect_data()
        blanks = self.app.find_blank_fields(data)
        self.assertEqual(list(blanks), ['基本資料', '道路步道', '設施', '自然資源與遊客', '意見與心得建議',
                                        '服勤照片(上)', '服勤照片(下)'])
        self.assertEqual(blanks['基本資料'], ['協勤人員姓名', '協勤地點'])
        self.assertEqual(len(blanks['道路步道']), 3)
        self.assertEqual(len(blanks['設施']), 4)
        self.assertEqual(len(blanks['自然資源與遊客']), 3)
        self.assertEqual(len(blanks['意見與心得建議']), 3)
        self.assertEqual(len(blanks['服勤照片(上)']), 3)
        self.assertEqual(len(blanks['服勤照片(下)']), 3)

    def test_partially_filled_form_lists_only_the_blank_ones(self):
        self._fill_everything()
        self.app.text_fields['road_grass'].delete('1.0', tk.END)
        self.app.location.delete(0, tk.END)
        self.app.clear_photos('其他', self.app.listboxes['其他'])
        data, _ = self.app.collect_data()
        blanks = self.app.find_blank_fields(data)
        self.assertEqual(blanks, {
            '基本資料': ['協勤地點'],
            '道路步道': ['2、道路、步道雜草情形'],
            '服勤照片(下)': ['其他（尚未選擇照片）'],
        })

    def test_whitespace_only_counts_as_blank(self):
        self._fill_everything()
        self.app.text_fields['nature'].delete('1.0', tk.END)
        self.app.text_fields['nature'].insert('1.0', '   \n  ')
        data, _ = self.app.collect_data()
        self.assertIn('三、自然資源狀況', self.app.find_blank_fields(data)['自然資源與遊客'])

    def test_generate_warns_about_blanks_with_default_yes_and_proceeds(self):
        self.fill_dates()
        self.app.personnel.insert(0, '王小明')
        out = os.path.join(self.tmp.name, 'o.docx')
        with mock.patch.object(filedialog, 'asksaveasfilename', lambda **k: out), \
                mock.patch.object(g, 'open_file', lambda path: None):
            self.app.generate_report()
            self._wait_until_done()
        self.assertEqual(len(self.confirm_calls), 1)
        title, message, kwargs = self.confirm_calls[0]
        self.assertEqual(title, '有欄位尚未填寫')
        self.assertEqual(kwargs.get('default'), 'yes')   # 預設「是」
        self.assertEqual(kwargs.get('icon'), 'warning')
        self.assertIn('【基本資料】', message)
        self.assertIn('協勤地點', message)
        self.assertNotIn('協勤人員姓名', message)         # 已填的不會列出
        self.assertIn('【服勤照片(下)】', message)
        self.assertIn('仍要產生報告嗎？', message)
        self.assertTrue(os.path.exists(out))             # 回答「是」就照常產生

    def test_answering_no_cancels_and_jumps_to_first_blank_tab(self):
        self.fill_dates()
        saved = []
        with mock.patch.object(messagebox, 'askyesno', lambda *a, **k: False), \
                mock.patch.object(filedialog, 'asksaveasfilename', lambda **k: saved.append(1) or ''):
            self.app.generate_report()
        self.assertEqual(saved, [])  # 沒有進到存檔步驟
        self.assertEqual(self.app.notebook.tab(self.app.notebook.select(), 'text'), '基本資料')
        self.assertEqual(str(self.app.generate_btn['state']), 'normal')
        self.assertEqual(self.pdf_calls, [])

    def test_no_warning_when_everything_is_filled(self):
        self._fill_everything()
        out = os.path.join(self.tmp.name, 'o.docx')
        with mock.patch.object(filedialog, 'asksaveasfilename', lambda **k: out), \
                mock.patch.object(g, 'open_file', lambda path: None):
            self.app.generate_report()
            self._wait_until_done()
        self.assertEqual(self.confirm_calls, [])
        self.assertTrue(os.path.exists(out))

    def test_date_errors_are_reported_before_blank_check(self):
        self.app.generate_report()  # 協勤日期還是空的
        self.assertEqual(self.shown[-1][0], 'err')
        self.assertEqual(self.confirm_calls, [])

    def test_primary_and_danger_buttons_keep_3d_bevel(self):
        """產生報告/清除按鈕要和其他按鈕一樣有凸起的立體邊框（亮邊 + 暗邊，且暗邊比底色深）。"""
        for style_name in ('Primary.TButton', 'Danger.TButton', 'TButton'):
            bg = style_lookup(self.root, style_name, 'background')
            light = style_lookup(self.root, style_name, 'lightcolor')
            dark = style_lookup(self.root, style_name, 'darkcolor')
            with self.subTest(style=style_name):
                self.assertNotEqual(light, bg)
                self.assertNotEqual(dark, bg)
                self.assertGreater(self._luminance(light), self._luminance(dark))
        for name in g.THEMES:
            self.app.apply_theme(name)
            for style_name, key in (('Primary.TButton', 'primary'), ('Danger.TButton', 'danger')):
                with self.subTest(theme=name, style=style_name):
                    fill = g.THEMES[name][key]
                    self.assertEqual(style_lookup(self.root, style_name, 'lightcolor'), g.BEVEL_LIGHT)
                    self.assertLess(self._luminance(style_lookup(self.root, style_name, 'darkcolor')),
                                    self._luminance(fill))

    def test_mix_color(self):
        self.assertEqual(g.mix_color('#ffffff', '#000000', 0.25), '#bfbfbf')
        self.assertEqual(g.mix_color('#102030', '#102030', 0.5), '#102030')
        self.assertEqual(g.mix_color('#000000', '#ffffff', 1), '#ffffff')

    def test_window_title(self):
        self.assertEqual(g.APP_TITLE, '玉山保育巡查回報單生成器')
        self.assertEqual(self.root.title(), '玉山保育巡查回報單生成器')

    def test_date_entries_are_centered(self):
        for name in ('year', 'month', 'day', 'start_year', 'start_month', 'start_day',
                     'end_year', 'end_month', 'end_day'):
            with self.subTest(name=name):
                self.assertEqual(str(getattr(self.app, name).cget('justify')), 'center')

    def test_selected_tab_is_not_smaller_than_other_tabs(self):
        import re
        style = ttk.Style(self.root)
        numbers = lambda value: [int(n) for n in re.findall(r'\d+', str(value))]
        normal = numbers(style.lookup('TNotebook.Tab', 'padding', []))
        active = numbers(style.lookup('TNotebook.Tab', 'padding', ['selected']))
        self.assertEqual(len(normal), 2)
        self.assertEqual(len(active), 2)
        self.assertGreaterEqual(active[0], normal[0])
        self.assertGreaterEqual(active[1], normal[1])

    def test_selected_tab_is_taller_and_not_narrower_on_screen(self):
        nb = next(w for w in self.root.winfo_children()[0].winfo_children() if isinstance(w, ttk.Notebook))
        self.root.update()
        boxes = []
        for i in range(len(nb.tabs())):
            nb.select(i)
            self.root.update()
            boxes.append(nb.bbox(i))
        for i, (x, y, w, h) in enumerate(boxes):
            nb.select(i)
            self.root.update()
            x2, y2, w2, h2 = nb.bbox(i)
            others = [nb.bbox(j) for j in range(len(boxes)) if j != i]
            self.assertGreaterEqual(h2, max(b[3] for b in others), f'分頁 {i} 被選取時比其他分頁矮')
            # 選取時的寬度不應比沒選取時小
            nb.select((i + 1) % len(boxes))
            self.root.update()
            self.assertGreaterEqual(w2, nb.bbox(i)[2], f'分頁 {i} 選取時比沒選取時窄')

    def test_notebook_tabs_are_visually_separated(self):
        style = ttk.Style(self.root)
        self.assertEqual(style.theme_use(), 'clam')
        active = style.lookup('TNotebook.Tab', 'background', ['selected'])
        idle = style.lookup('TNotebook.Tab', 'background', [])
        palette = g.THEMES[g.DEFAULT_THEME]
        self.assertEqual(active, palette['tab_active'])
        self.assertEqual(idle, palette['tab_idle'])
        self.assertNotEqual(active, idle)
        self.assertEqual(style.lookup('TNotebook.Tab', 'foreground', ['selected']), palette['tab_active_fg'])
        self.assertGreaterEqual(int(style.lookup('TNotebook.Tab', 'borderwidth')), 2)

    def test_old_draft_without_default_caption_uses_default(self):
        import json
        draft = os.path.join(self.tmp.name, 'old.json')
        self._save_draft_to(draft)
        with open(draft, encoding='utf-8') as f:
            data = json.load(f)
        data.pop('default_caption')
        with open(draft, 'w', encoding='utf-8') as f:
            json.dump(data, f)
        self.app.default_caption.set('x')
        self._load_draft_from(draft)
        self.assertEqual(self.app.default_caption.get(), g.DEFAULT_CAPTION)

    def test_enter_keeps_the_chosen_symbol(self):
        w = self.app.text_fields['nature']
        w.insert('1.0', '■ 第一項')
        w.mark_set(tk.INSERT, 'end-1c')
        self.app._on_return(w)
        self.assertEqual(w.get('1.0', 'end-1c'), '■ 第一項\n■ ')

    def test_enter_continues_list_and_ends_on_empty_item(self):
        w = self.app.text_fields['nature']
        w.insert('1.0', '1. 第一項')
        w.mark_set(tk.INSERT, 'end-1c')
        self.assertEqual(self.app._on_return(w), 'break')
        self.assertEqual(w.get('1.0', 'end-1c'), '1. 第一項\n2. ')
        self.assertEqual(self.app._on_return(w), 'break')  # 空項目再按 Enter：移除編號、結束清單
        self.assertEqual(w.get('1.0', 'end-1c'), '1. 第一項\n')
        self.assertIsNone(self.app._on_return(w))  # 一般文字維持預設行為

    def test_list_markers_flow_into_report(self):
        self.fill_dates()
        w = self.app.text_fields['nature']
        w.insert('1.0', '山羌\n水鹿')
        w.tag_add(tk.SEL, '1.0', 'end-1c')
        self.app.apply_list(w, g.toggle_numbers)
        data, err = self.app.collect_data()
        self.assertIsNone(err)
        self.assertEqual(data['nature'], '1. 山羌\n2. 水鹿')

    def test_gui_fonts_are_enlarged_by_two(self):
        import tkinter.font as tkfont
        probe = tk.Tk()
        base = tkfont.nametofont('TkDefaultFont', root=probe).cget('size')
        probe.destroy()
        scaled = tkfont.nametofont('TkDefaultFont', root=self.root).cget('size')
        step = 2 if base > 0 else -2
        self.assertEqual(scaled, base + step)
        self.assertEqual(g.GUI_FONT_DELTA, 2)

    def test_gui_fonts_use_configured_weight(self):
        import tkinter.font as tkfont
        self.assertIn(g.GUI_FONT_WEIGHT, ('normal', 'bold'))
        for name in ('TkDefaultFont', 'TkTextFont', 'TkFixedFont', 'TkMenuFont'):
            with self.subTest(font=name):
                self.assertEqual(tkfont.nametofont(name, root=self.root).cget('weight'), g.GUI_FONT_WEIGHT)

    def test_scale_gui_fonts_can_keep_normal_weight(self):
        import tkinter.font as tkfont
        probe = tk.Tk()
        try:
            g.scale_gui_fonts(probe, 0, None, 'normal')
            self.assertEqual(tkfont.nametofont('TkDefaultFont', root=probe).cget('weight'), 'normal')
        finally:
            probe.destroy()

    def test_font_files_bundled_with_license(self):
        for name in g.FONT_FILES + (g.FONT_LICENSE,):
            self.assertTrue(os.path.exists(os.path.join(g.FONT_DIR, name)), name)

    def test_gui_uses_bundled_font_family(self):
        import tkinter.font as tkfont
        self.assertEqual(self.app.font_family, g.FONT_FAMILY)
        for name in ('TkDefaultFont', 'TkTextFont', 'TkFixedFont'):
            self.assertEqual(tkfont.nametofont(name, root=self.root).cget('family'), g.FONT_FAMILY)

    def test_font_family_falls_back_when_not_available(self):
        with mock.patch.object(g, 'load_private_fonts', lambda: False), \
                mock.patch.object(g.tkfont, 'families', lambda root=None: ['Arial']):
            self.assertEqual(g.pick_font_family(self.root), 'Arial')
        with mock.patch.object(g, 'load_private_fonts', lambda: False), \
                mock.patch.object(g.tkfont, 'families', lambda root=None: ['Arial', g.FALLBACK_FONT_FAMILY]):
            self.assertEqual(g.pick_font_family(self.root), g.FALLBACK_FONT_FAMILY)

    # ---- 照片預覽與拖曳排序
    class _Ev:
        def __init__(self, y):
            self.y = y

    def _photos(self, names=('a.jpg', 'b.jpg', 'c.jpg'), cat=None, size=(400, 200), exif=None):
        cat = cat or g.PHOTO_CATEGORIES_1[0]
        lb = self.app.listboxes[cat]
        paths = []
        for n in names:
            p = os.path.join(self.tmp.name, n)
            im = Image.new('RGB', size, 'red')
            if exif:
                ex = im.getexif()
                ex[0x0112] = exif
                im.save(p, exif=ex)
            else:
                im.save(p)
            self.app.photo_files[cat].append(p)
            lb.insert(tk.END, self.app._photo_label(cat, p))
            paths.append(p)
        return cat, lb, paths

    def _drag(self, cat, lb, start, *targets):
        """模擬從 start 位置按下，依序拖到 targets，最後放開。"""
        full = (0, 0, 100, 1000)
        with mock.patch.object(lb, 'bbox', lambda i: full), mock.patch.object(lb, 'nearest') as nearest:
            nearest.return_value = start
            self.app._drag_start(cat, lb, self._Ev(5))
            for t in targets:
                nearest.return_value = t
                self.app._drag_motion(cat, lb, self._Ev(5))
            self.app._drag_end()

    def test_drag_reorders_photos_and_listbox(self):
        cat, lb, (a, b, c) = self._photos()
        lb.selection_set(0)
        self._drag(cat, lb, 0, 1, 2)
        self.assertEqual(self.app.photo_files[cat], [b, c, a])
        self.assertEqual([lb.get(i) for i in range(3)], ['b.jpg', 'c.jpg', 'a.jpg'])
        self.assertEqual(lb.curselection(), (2,))
        self._drag(cat, lb, 2, 0)
        self.assertEqual(self.app.photo_files[cat], [a, b, c])

    def test_drag_order_persists_in_draft_and_marks_dirty(self):
        cat, lb, (a, b, c) = self._photos()
        self.app._saved_signature = self.app._signature()
        self._drag(cat, lb, 2, 0)
        self.assertEqual(self.app.build_draft()['photos'][cat], [c, a, b])
        self.assertTrue(self.app.is_dirty())

    def test_drag_in_blank_area_does_nothing(self):
        cat, lb, (a, b, c) = self._photos()
        with mock.patch.object(lb, 'bbox', lambda i: (0, 0, 100, 20)), mock.patch.object(lb, 'nearest', lambda y: 2):
            self.app._drag_start(cat, lb, self._Ev(500))  # 按在清單最後一項下方的空白處
            self.assertIsNone(self.app._drag)
            self.assertIsNone(self.app._drag_motion(cat, lb, self._Ev(500)))
        self.assertEqual(self.app.photo_files[cat], [a, b, c])

    def test_drag_target_is_clamped(self):
        cat, lb, (a, b, c) = self._photos()
        self._drag(cat, lb, 0, 99)
        self.assertEqual(self.app.photo_files[cat], [b, c, a])

    def test_select_shows_preview_with_info(self):
        cat, lb, (a, b, c) = self._photos(size=(400, 200))
        self.app.photo_captions[cat][a] = '山屋外觀'
        lb.selection_set(0)
        nb = next(w for w in self.root.winfo_children()[0].winfo_children() if isinstance(w, ttk.Notebook))
        nb.select(5)  # 沒顯示出來的元件收不到事件，要先切到照片分頁
        self.root.update()
        lb.event_generate('<<ListboxSelect>>')  # 確認事件有綁到預覽
        self.root.update()
        preview = self.app.previews[cat]
        self.assertEqual(preview.current, (cat, a))
        self.assertIsNotNone(preview.image)
        self.assertLessEqual(preview.image.width(), g.PREVIEW_SIZE[0])
        self.assertLessEqual(preview.image.height(), g.PREVIEW_SIZE[1])
        text = preview.info.cget('text')
        self.assertIn('a.jpg', text)
        self.assertIn('400×200（橫式）', text)
        self.assertIn('山屋外觀', text)

    def test_preview_applies_exif_rotation(self):
        cat, lb, (a,) = self._photos(names=('r.jpg',), size=(400, 200), exif=6)
        lb.selection_set(0)
        self.app._on_photo_select(cat, lb)
        preview = self.app.previews[cat]
        self.assertGreater(preview.image.height(), preview.image.width())  # 轉正後為直式
        self.assertIn('200×400（直式）', preview.info.cget('text'))

    def test_preview_of_unreadable_file_does_not_crash(self):
        cat = g.PHOTO_CATEGORIES_1[0]
        lb = self.app.listboxes[cat]
        bad = os.path.join(self.tmp.name, 'bad.jpg')
        with open(bad, 'wb') as f:
            f.write(b'not an image')
        self.app.photo_files[cat].append(bad)
        lb.insert(tk.END, 'bad.jpg')
        lb.selection_set(0)
        self.app._on_photo_select(cat, lb)
        preview = self.app.previews[cat]
        self.assertIsNone(preview.image)
        self.assertIn('bad.jpg', preview.info.cget('text'))

    def test_preview_cleared_when_photo_removed_or_cleared(self):
        cat, lb, (a, b, c) = self._photos()
        lb.selection_set(1)
        self.app._on_photo_select(cat, lb)
        self.assertEqual(self.app.previews[cat].current, (cat, b))
        lb.selection_clear(0, tk.END)
        lb.selection_set(0)
        self.app.remove_photos(cat, lb)  # 移除 a：預覽的是 b，不受影響
        self.assertEqual(self.app.previews[cat].current, (cat, b))
        lb.selection_clear(0, tk.END)
        lb.selection_set(0)
        self.app.remove_photos(cat, lb)  # 移除 b：預覽要清掉
        self.assertIsNone(self.app.previews[cat].current)
        self.assertIsNone(self.app.previews[cat].image)
        lb.selection_set(0)
        self.app._on_photo_select(cat, lb)
        self.assertEqual(self.app.previews[cat].current, (cat, c))
        self.app.clear_photos(cat, lb)
        self.assertIsNone(self.app.previews[cat].current)

    def test_preview_refreshed_after_caption_edit(self):
        cat, lb, (a, b, c) = self._photos()
        lb.selection_set(0)
        with mock.patch.object(simpledialog, 'askstring', lambda *a, **k: '新說明'):
            self.app.edit_caption(cat, lb)
        self.assertIn('新說明', self.app.previews[cat].info.cget('text'))
        self.assertEqual(lb.curselection(), (0,))

    def test_tab_shares_one_preview_between_its_categories(self):
        self.assertIs(self.app.previews[g.PHOTO_CATEGORIES_1[0]], self.app.previews[g.PHOTO_CATEGORIES_1[2]])
        self.assertIsNot(self.app.previews[g.PHOTO_CATEGORIES_1[0]], self.app.previews[g.PHOTO_CATEGORIES_2[0]])

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
