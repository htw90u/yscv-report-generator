"""資安相關測試：草稿（可能來自別人）、意見反應的輸入過濾與送出內容、端點設定、Apps Script、git、圖片炸彈、相依套件。
執行：python -m unittest discover -s tests"""
import gc
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import tkinter as tk
import types
import unittest
from tkinter import filedialog
from unittest import mock

from PIL import Image

import yscv_report_generator as g
from test_reports import FakeFeedbackServer

messagebox = g.messagebox
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class DraftSafetyTest(unittest.TestCase):
    """草稿檔可能是別人給的：不能因為草稿而連到外部主機、清掉目前的內容或當掉。"""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = tk.Tk()
        self.autosave = os.path.join(self.tmp.name, 'autosave.json')
        self.app = g.ReportGeneratorGUI(self.root, autosave_path=self.autosave, restore_on_start=False,
                                        settings_path=os.path.join(self.tmp.name, 's.json'))
        self.shown = []
        for p in (mock.patch.object(messagebox, 'showerror', lambda t, m: self.shown.append(('err', m))),
                  mock.patch.object(messagebox, 'showinfo', lambda t, m: self.shown.append(('info', m))),
                  mock.patch.object(messagebox, 'askyesno', lambda *a, **k: True)):
            p.start()
            self.addCleanup(p.stop)
        self.addCleanup(gc.collect)
        self.addCleanup(self.root.destroy)
        self.cat = g.PHOTO_CATEGORIES[0]
        self.base = {'version': g.DRAFT_VERSION, 'report': 'patrol'}

    def fill_screen(self):
        g.widget_set(self.app.personnel, '王小明（已填好）')
        self.app.text_fields['road_status'].delete('1.0', tk.END)
        self.app.text_fields['road_status'].insert('1.0', '辛苦寫的內容')

    def screen_intact(self):
        return (self.app.personnel.get() == '王小明（已填好）'
                and self.app.text_fields['road_status'].get('1.0', tk.END).strip() == '辛苦寫的內容')

    def write_draft(self, draft):
        path = os.path.join(self.tmp.name, 'draft.json')
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(draft, f)
        return path

    def test_unsafe_path_detection(self):
        for bad in (r'\\attacker.example\share\x.jpg', '//attacker.example/share/x.jpg', r'\\?\C:\x.jpg',
                    r'\\.\pipe\x', 'x.jpg', r'..\x.jpg', '', None, 123, ['C:\\a.jpg'], 'C:\\a\0b.jpg',
                    r'\\?\UNC\attacker\share\x.jpg'):
            self.assertTrue(g.is_unsafe_draft_path(bad), repr(bad))
        for good in (r'C:\Users\a\x.jpg', 'D:/photos/x.jpg', r'C:\照片\勤務 1.jpg'):
            self.assertFalse(g.is_unsafe_draft_path(good), good)

    def test_network_paths_in_a_draft_are_never_touched(self):
        touched = []
        real_exists = os.path.exists

        def spy(path):
            touched.append(path)
            return real_exists(path)

        paths = [r'\\attacker.example\share\x.jpg', '//attacker.example/share/y.jpg', r'\\?\UNC\attacker\share\z.jpg']
        with mock.patch.object(os.path, 'exists', spy):
            warnings = self.app.apply_draft(dict(self.base, photos={self.cat: paths}))
        self.assertEqual([p for p in touched if 'attacker' in str(p)], [])           # 沒有對那些路徑做任何檔案存取
        self.assertEqual(len(warnings), 3)
        self.assertTrue(all('網路路徑' in w for w in warnings))
        self.assertEqual(self.app.photo_files[self.cat], [])

    def test_local_photos_in_a_draft_still_load(self):
        photo = os.path.join(self.tmp.name, 'ok.jpg')
        Image.new('RGB', (40, 30), 'red').save(photo)
        warnings = self.app.apply_draft(dict(self.base, photos={self.cat: [photo]},
                                             captions={self.cat: {photo: '說明'}}, rotations={self.cat: {photo: 90}}))
        self.assertEqual(warnings, [])
        self.assertEqual(self.app.photo_files[self.cat], [photo])
        self.assertEqual(self.app.photo_captions[self.cat][photo], '說明')
        self.assertEqual(self.app.photo_rotations[self.cat][photo], 90)

    def test_odd_photo_entries_become_warnings_not_crashes(self):
        warnings = self.app.apply_draft(dict(self.base, photos={self.cat: [123, None, 'C:\\a\0b.jpg', ['x'], 'rel.jpg']}))
        self.assertEqual(len(warnings), 5)
        self.assertEqual(self.app.photo_files[self.cat], [])

    def test_malformed_drafts_are_rejected_before_touching_the_screen(self):
        malformed = {
            'fields 是清單': dict(self.base, fields=[1, 2]),
            'texts 是字串': dict(self.base, texts='abc'),
            'photos 是字串': dict(self.base, photos='abc'),
            'photos 的值不是清單': dict(self.base, photos={self.cat: 'abc'}),
            'captions 是清單': dict(self.base, captions=[1]),
            'rotations 是數字': dict(self.base, rotations=5),
            'crops 是字串': dict(self.base, crops='x'),
        }
        for name, draft in malformed.items():
            self.fill_screen()
            with self.assertRaises(ValueError, msg=name):
                self.app.apply_draft(draft)
            self.assertTrue(self.screen_intact(), f'{name}：畫面內容被清掉了')

    def test_loading_a_malformed_draft_shows_an_error_and_keeps_the_screen(self):
        self.fill_screen()
        path = self.write_draft(dict(self.base, fields=[1, 2]))
        with mock.patch.object(filedialog, 'askopenfilename', lambda **k: path):
            self.app.load_draft()
        self.assertEqual(self.shown[-1][0], 'err')
        self.assertIn('無法載入草稿', self.shown[-1][1])
        self.assertTrue(self.screen_intact())

    def test_malformed_autosave_is_discarded_and_keeps_the_screen(self):
        self.fill_screen()
        with open(self.autosave, 'w', encoding='utf-8') as f:
            json.dump(dict(self.base, texts='abc'), f)
        self.app.offer_restore()
        self.assertFalse(os.path.exists(self.autosave))
        self.assertTrue(self.screen_intact())

    def test_tolerant_inner_values_still_load(self):
        self.fill_screen()
        photo = os.path.join(self.tmp.name, 'ok.jpg')
        Image.new('RGB', (40, 30), 'red').save(photo)
        draft = dict(self.base, photos={self.cat: [photo]}, captions={self.cat: 'bad'}, crops={self.cat: {photo: 'bad'}},
                     fields={'personnel': '李小華', 'unknown_field': 'x'}, texts={'not_a_field': 'y'})
        self.assertEqual(self.app.apply_draft(draft), [])
        self.assertEqual(self.app.personnel.get(), '李小華')                          # 不認得的欄位直接忽略
        self.assertEqual(self.app.photo_files[self.cat], [photo])


class FeedbackInputTest(unittest.TestCase):
    """意見反應的輸入整理與送出內容（不開視窗）。"""

    def test_clean_line_and_clean_body(self):
        self.assertEqual(g.clean_line('a\r\nBcc: x\t\u202e  y\x00'), 'a Bcc: x y')
        self.assertEqual(g.clean_line('  ' + 'x' * 500, 10), 'x' * 10)
        self.assertEqual(g.clean_line(None), '')
        self.assertEqual(g.clean_body('a\x00\r\nb\rc\u202e\u2028d\x7f'), 'a\nb\ncd')
        self.assertEqual(g.clean_body('保留\t定位字元\n與換行'), '保留\t定位字元\n與換行')

    def test_valid_email(self):
        for good in ('me@example.com', 'a.b+c@sub.example.co.uk', 'x_y%z@e-x.io'):
            self.assertTrue(g.valid_email(good), good)
        for bad in ('', 'abc', 'a@b', 'a@b.', '@x.io', 'a@@x.io', 'a b@x.io', 'a@x.io\n', 'a@x.io\x00', '<a@x.io>',
                    'a@x.io,b@x.io', 'a@x.io;b@x.io', ('a' * 65) + '@x.io', 'a@' + ('b' * 250) + '.io', '我@x.io'):
            self.assertFalse(g.valid_email(bad), repr(bad))

    def test_send_feedback_cleans_and_caps_even_if_the_caller_does_not(self):
        server = FakeFeedbackServer()
        self.addCleanup(server.close)
        with mock.patch.object(g, 'feedback_endpoint_token', return_value=''):
            ok, _ = g.send_feedback('主旨\r\nBcc: evil@example.com', 'B' * 50000 + '\x00', url=server.url,
                                    reply_to='me@example.com\r\nBcc: evil@example.com')
        self.assertTrue(ok)
        payload = server.received[0]
        self.assertEqual(payload['subject'], '主旨 Bcc: evil@example.com')
        self.assertEqual(len(payload['body']), g.FEEDBACK_BODY_MAX)
        self.assertNotIn('\x00', payload['body'])
        self.assertEqual(payload['reply_to'], '')                                    # 不合格的回信地址直接丟掉，不是修一修再用

    def test_valid_reply_to_is_kept_and_token_is_sent_only_when_configured(self):
        server = FakeFeedbackServer()
        self.addCleanup(server.close)
        with mock.patch.object(g, 'feedback_endpoint_token', return_value=''):
            g.send_feedback('s', 'b', url=server.url, reply_to=' me@example.com ')
        self.assertEqual(server.received[-1]['reply_to'], 'me@example.com')
        self.assertNotIn('token', server.received[-1])
        with mock.patch.object(g, 'feedback_endpoint_token', return_value='s3cret'):
            g.send_feedback('s', 'b', url=server.url)
        self.assertEqual(server.received[-1]['token'], 's3cret')

    def test_endpoint_url_must_be_https(self):
        for url in ('http://example.com/x', 'file:///C:/Windows/win.ini', 'ftp://example.com/x', 'javascript:alert(1)',
                    '//example.com/x', 'example.com/x'):
            fake = types.ModuleType('yscv_feedback_endpoint')
            fake.URL = url
            with mock.patch.dict(sys.modules, {'yscv_feedback_endpoint': fake}):
                self.assertEqual(g.feedback_endpoint_url(), '', url)
        fake = types.ModuleType('yscv_feedback_endpoint')
        fake.URL = 'HTTPS://script.google.com/macros/s/ABC/exec'
        fake.TOKEN = ' abc '
        with mock.patch.dict(sys.modules, {'yscv_feedback_endpoint': fake}):
            self.assertEqual(g.feedback_endpoint_url(), 'HTTPS://script.google.com/macros/s/ABC/exec')
            self.assertEqual(g.feedback_endpoint_token(), 'abc')

    def test_cooldown_remaining(self):
        with mock.patch.object(g, 'FEEDBACK_COOLDOWN', 30):
            g._feedback_last_sent[0] = 0.0
            self.assertEqual(g.feedback_cooldown_remaining(), 0)                      # 還沒寄過
            g._feedback_last_sent[0] = 100.0
            self.addCleanup(lambda: g._feedback_last_sent.__setitem__(0, 0.0))
            with mock.patch.object(g.time, 'monotonic', return_value=110.0):
                self.assertEqual(g.feedback_cooldown_remaining(), 20.0)
            with mock.patch.object(g.time, 'monotonic', return_value=500.0):
                self.assertEqual(g.feedback_cooldown_remaining(), 0)


class GitGuardTest(unittest.TestCase):
    def test_git_is_not_run_outside_a_repository(self):
        with tempfile.TemporaryDirectory() as folder, mock.patch.object(g.subprocess, 'run') as run:
            self.assertEqual(g.git_revision(folder), '')
            self.assertIsNone(g.git_describe(folder))
            self.assertEqual(g.version_label(folder), f'v{g.VERSION}')
            run.assert_not_called()                                                   # 根本不會去執行 git

    def test_git_still_runs_inside_a_repository(self):
        self.assertRegex(g.git_revision(ROOT), r'^[0-9a-f]{7,}$')


class AppsScriptTest(unittest.TestCase):
    """用 Node 載入 Apps Script 原始碼，搭配假的 Google 服務實際執行 doPost。"""

    @unittest.skipUnless(shutil.which('node'), '需要安裝 Node.js 才能測試 Apps Script')
    def test_apps_script_behaviour(self):
        out = subprocess.run(['node', os.path.join(ROOT, 'tests', 'apps_script_harness.js'),
                              os.path.join(ROOT, 'apps_script_feedback.gs')],
                             capture_output=True, encoding='utf-8', timeout=60)
        self.assertEqual(out.returncode, 0, out.stderr[:500])
        results = json.loads(out.stdout)
        self.assertGreaterEqual(len(results), 10)
        failed = {name: value for name, value in results.items() if value is not True}
        self.assertEqual(failed, {})

    def test_template_has_no_real_recipient_or_secret(self):
        with open(os.path.join(ROOT, 'apps_script_feedback.gs'), encoding='utf-8') as f:
            text = f.read()
        self.assertIn('var TO = "請改成收件信箱";', text)
        self.assertIn('var TOKEN = "";', text)
        self.assertNotRegex(text, r'[\w.]+@[\w.]+\.(io|com|tw|org)')


class ImageBombTest(unittest.TestCase):
    """畫素多得驚人的圖片（解壓縮炸彈）不能讓程式當掉或吃光記憶體：載入失敗以警告處理。"""

    def test_huge_image_is_rejected_gracefully(self):
        with tempfile.TemporaryDirectory() as folder:
            bomb = os.path.join(folder, 'bomb.png')
            Image.new('1', (20000, 10000)).save(bomb)                                  # 2 億畫素，檔案卻很小
            self.assertLess(os.path.getsize(bomb), 5_000_000)
            warnings = []
            items = g._load_photos({'分類': [bomb]}, {}, 3.4, 4.7, warnings)['分類']
            self.assertIsNone(items[0].source)                                         # 沒有載入
            self.assertEqual(len(warnings), 1)
            self.assertIn('bomb.png', warnings[0])

    def test_pixel_limit_is_not_disabled(self):
        self.assertIsNotNone(Image.MAX_IMAGE_PIXELS)                                   # 程式沒有把保護關掉


class DependencyTest(unittest.TestCase):
    def test_requirements_file_lists_minimum_versions(self):
        with open(os.path.join(ROOT, 'requirements.txt'), encoding='utf-8') as f:
            text = f.read()
        lines = [ln.strip() for ln in text.splitlines() if ln.strip() and not ln.startswith('#')]
        names = {re.split(r'[<>=!~]', ln)[0].lower(): ln for ln in lines}
        self.assertEqual(set(names), {'python-docx', 'pillow'})
        for line in names.values():
            self.assertRegex(line, r'>=\d')                                            # 每個套件都有最低版本

    def test_installed_versions_satisfy_the_minimums(self):
        from importlib import metadata

        def version(name):
            return tuple(int(x) for x in re.findall(r'\d+', metadata.version(name))[:3])

        self.assertGreaterEqual(version('python-docx'), (1, 1, 0))
        self.assertGreaterEqual(version('Pillow'), (10, 3, 0))


class RobustnessTest(unittest.TestCase):
    """第二輪檢查補上的項目：貼上的控制字元、https 轉址降級、設定檔寫入、PowerShell 啟動方式。"""

    BAD = 'a\x00b\x08c\x0bd\x1fe\ufffe'

    def build(self, which):
        import yscv_reports as r
        from test_reports import TATAKA_DATA, ECO_DATA
        from _common import TEMPLATE, TATAKA_TEMPLATE, ECO_TEMPLATE
        data, template, fn = {'tataka': (TATAKA_DATA, TATAKA_TEMPLATE, r.build_tataka_report),
                              'eco': (ECO_DATA, ECO_TEMPLATE, r.build_ecotoilet_report),
                              'patrol': (None, TEMPLATE, g.build_report)}[which]
        if data is None:
            data = dict(year='115', month='10', day='15', personnel='王小明', location='玉山', start_year='115',
                        start_month='10', start_day='15', end_year='115', end_month='10', end_day='15')
        data = {k: (self.BAD if isinstance(v, str) else v) for k, v in data.items()}
        out = os.path.join(tempfile.mkdtemp(), 'x.docx')
        self.addCleanup(shutil.rmtree, os.path.dirname(out), True)
        fn(template, data, {}, {'cat': {'p': self.BAD}}, out)
        return out

    def test_control_characters_do_not_break_report_generation(self):
        for which in ('tataka', 'eco', 'patrol'):
            with self.subTest(which):
                self.assertTrue(os.path.exists(self.build(which)))

    def test_xml_safe_keeps_normal_text(self):
        self.assertEqual(g.xml_safe('第一行\n第二行\t中文 abc'), '第一行\n第二行\t中文 abc')
        self.assertEqual(g.xml_safe(self.BAD), 'abcde')
        self.assertEqual(g.xml_safe(5), 5)

    def test_control_characters_are_removed_from_file_names(self):
        name = g.default_output_name(dict(year='115', month='10', day='15', start_year='115', start_month='10',
                                          start_day='15', end_year='115', end_month='10', end_day='15',
                                          location='玉山\x00主峰\x0b線', personnel='王\x1f小明'))
        self.assertFalse(any(ord(c) < 32 for c in name), repr(name))

    def test_https_redirect_to_http_is_refused(self):
        import urllib.error
        import urllib.request
        import yscv_updates
        handler = yscv_updates._NoDowngradeRedirect()
        request = urllib.request.Request('https://example.com/a')
        for target in ('http://example.com/b', 'ftp://example.com/b', 'file:///C:/x'):
            with self.assertRaises(urllib.error.URLError, msg=target):
                handler.redirect_request(request, None, 302, 'Found', {}, target)
        self.assertIsNotNone(handler.redirect_request(request, None, 302, 'Found', {}, 'https://other.example/b'))

    def test_settings_are_written_atomically(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, 'settings.json')
            self.assertTrue(g.save_settings(path, theme='x'))
            with mock.patch.object(g.json, 'dump', side_effect=OSError('disk full')):
                self.assertFalse(g.save_settings(path, theme='y'))
            self.assertEqual(g.load_settings(path), {'theme': 'x'})                    # 寫失敗不會弄壞原本的檔案

    def test_powershell_is_started_from_the_system_folder(self):
        if sys.platform != 'win32':
            self.skipTest('只在 Windows')
        path = g._powershell_path()
        self.assertTrue(os.path.isabs(path) and os.path.exists(path), path)
        self.assertIn('AutomationSecurity = 3', g._PDF_SCRIPT)                          # 轉 PDF 時 Word 強制停用巨集


if __name__ == '__main__':
    unittest.main()
