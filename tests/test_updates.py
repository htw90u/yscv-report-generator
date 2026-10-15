"""更新檢查（第一階段：只通知、由使用者決定；程式不下載任何東西）。
所有網路請求都只連本機假的 GitHub API 伺服器（127.0.0.1），不會連到真的 GitHub。"""
import gc
import http.server
import json
import os
import tempfile
import threading
import time
import tkinter as tk
import unittest
from datetime import datetime, timedelta
from unittest import mock

import yscv_report_generator as g
import yscv_updates as u

messagebox = g.messagebox
REPO = 'owner/yscv-report-generator'
NL = chr(10)


def release_json(tag='v0.9', **overrides):
    data = {'tag_name': tag, 'name': f'版本 {tag}', 'body': '- 新增功能 A\n- 修正問題 B', 'draft': False,
            'prerelease': False, 'published_at': '2026-10-20T08:00:00Z',
            'html_url': 'https://evil.example/phishing', 'assets': [{'browser_download_url': 'https://evil.example/x.exe'}]}
    data.update(overrides)
    return data


class FakeGitHub:
    """本機假的 GitHub API：依 mode 回應 /repos/<repo>/releases/latest，並記下收到的請求。"""

    def __init__(self, mode='newer'):
        outer = self
        self.mode = mode
        self.requests = []
        self.latest_tag = 'v0.9'

        class Handler(http.server.BaseHTTPRequestHandler):
            def log_message(self, *args):
                pass

            def reply(self, code, payload, ctype='application/json'):
                data = payload if isinstance(payload, bytes) else json.dumps(payload).encode('utf-8')
                self.send_response(code)
                self.send_header('Content-Type', ctype)
                self.send_header('Content-Length', str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def do_GET(self):
                outer.requests.append({'path': self.path, 'headers': dict(self.headers)})
                mode = outer.mode
                if self.path != f'/repos/{REPO}/releases/latest':
                    return self.reply(404, {'message': 'Not Found'})
                if mode == 'newer':
                    return self.reply(200, release_json(outer.latest_tag))
                if mode == 'same':
                    return self.reply(200, release_json(g.VERSION and f'v{g.VERSION}'))
                if mode == 'older':
                    return self.reply(200, release_json('v0.1'))
                if mode == 'prerelease':
                    return self.reply(200, release_json('v9.9', prerelease=True))
                if mode == 'draft':
                    return self.reply(200, release_json('v9.9', draft=True))
                if mode == 'badtag':
                    return self.reply(200, release_json('latest'))
                if mode == 'rc':
                    return self.reply(200, release_json('v1.0-rc1'))
                if mode == 'garbage':
                    return self.reply(200, '<html>不是 JSON</html>'.encode('utf-8'), 'text/html')
                if mode == 'array':
                    return self.reply(200, [1, 2, 3])
                if mode == 'huge':
                    return self.reply(200, json.dumps(release_json(body='x' * 2_000_000)).encode('utf-8'))
                if mode == 'longnotes':
                    return self.reply(200, release_json(body='行\n' * 5000 + '\x00‮結尾'))
                if mode == '404':
                    return self.reply(404, {'message': 'Not Found'})
                if mode == '403':
                    return self.reply(403, {'message': 'API rate limit exceeded'})
                if mode == '500':
                    return self.reply(500, {})
                return self.reply(200, release_json())

        self.server = http.server.HTTPServer(('127.0.0.1', 0), Handler)
        self.url = f'http://127.0.0.1:{self.server.server_address[1]}'
        threading.Thread(target=self.server.serve_forever, daemon=True).start()

    def close(self):
        self.server.shutdown()
        self.server.server_close()


class VersionLogicTest(unittest.TestCase):
    def test_parse_version(self):
        self.assertEqual(u.parse_version('v0.9'), (0, 9, 0, 0))
        self.assertEqual(u.parse_version('0.9.1'), (0, 9, 1, 0))
        self.assertEqual(u.parse_version(' V1.2.3.4 '), (1, 2, 3, 4))
        for bad in ('', None, 'latest', 'v1.0-rc1', 'v1.0+build', '1.2.3.4.5', 'v', 'vv1.0', '1..2', 'v1.0 beta', 'v٣.1'):
            self.assertIsNone(u.parse_version(bad), repr(bad))

    def test_version_ordering(self):
        order = ['v0.5', 'v0.8', 'v0.9', 'v0.10', 'v1.0', 'v1.0.1', 'v1.1', 'v2.0']
        versions = [u.parse_version(tag) for tag in order]
        self.assertEqual(versions, sorted(versions))
        self.assertEqual(u.parse_version('v0.9'), u.parse_version('0.9.0'))

    def test_repo_name_validation(self):
        for good in ('owner/repo', 'a-b/e.f-g_1', 'A1/x', 'a' * 39 + '/r'):
            self.assertTrue(u.is_valid_repo(good), good)
        for bad in ('', None, 'owner', 'owner/repo/extra', '../x', './x', 'owner/..', 'owner/.', 'owner/repo?x=1',
                    'owner/ repo', 'owner/repo#a', 'https://github.com/owner/repo', '-owner/repo', 'owner-/repo',
                    'own_er/repo', 'own.er/repo', 'a' * 40 + '/r', '/repo', 'owner/'):
            self.assertFalse(u.is_valid_repo(bad), repr(bad))

    def test_release_page_url_is_built_by_us(self):
        self.assertEqual(u.release_page_url(REPO, 'v0.9'), f'https://github.com/{REPO}/releases/tag/v0.9')
        self.assertEqual(u.release_page_url(REPO, 'v0.9/../../x?y#z'),
                         f'https://github.com/{REPO}/releases/tag/v0.9%2F..%2F..%2Fx%3Fy%23z')

    def test_trusted_release_url(self):
        good = f'https://github.com/{REPO}/releases/tag/v0.9'
        self.assertTrue(u.is_trusted_release_url(good, REPO))
        self.assertTrue(u.is_trusted_release_url(f'https://github.com/{REPO}/releases', REPO))
        for bad in (f'http://github.com/{REPO}/releases', 'https://evil.example/phishing', f'https://github.com.evil.example/{REPO}/releases',
                    f'https://evil.example/github.com/{REPO}/releases', f'https://user@github.com/{REPO}/releases',
                    f'https://github.com:8443/{REPO}/releases', f'https://github.com/other/repo/releases/tag/v1',
                    f'https://github.com/{REPO}x/releases', f'https://github.com/{REPO}/issues', '', None,
                    f'file:///github.com/{REPO}/releases', f'javascript:alert(1)//github.com/{REPO}/releases'):
            self.assertFalse(u.is_trusted_release_url(bad, REPO), repr(bad))

    def test_clean_notes(self):
        self.assertEqual(u.clean_notes('a\x00b‮c\r\nd\re'), 'abc\nd\ne')
        long = u.clean_notes('x' * 20000)
        self.assertLessEqual(len(long), u.NOTES_MAX_CHARS + 40)
        self.assertIn('完整說明請看發行頁面', long)
        self.assertEqual(u.clean_notes(None), '')

    def test_render_notes_turns_headings_and_bold_into_styles(self):
        text = NL.join(['## 新功能', '- 加了 **粗體** 與 __另一種粗體__ 的說明', '**完整變更**：見下方', '### 小標題 ###', '一般文字'])
        self.assertEqual(u.render_notes(text), [
            [('新功能', 'heading')],
            [('- 加了 ', ''), ('粗體', 'bold'), (' 與 ', ''), ('另一種粗體', 'bold'), (' 的說明', '')],
            [('完整變更', 'bold'), ('：見下方', '')],
            [('小標題', 'heading')],
            [('一般文字', '')]])

    def test_render_notes_leaves_everything_else_as_plain_text(self):
        text = NL.join(['#沒有空白不是標題', '壞的 **沒有結尾', '**  前後有空白不算粗體  **', '`程式碼` 與 [連結](https://example.com)',
                        '    ## 縮排太多不是標題', '****', '', '# ', '2 ** 3 ** 4'])
        rendered = u.render_notes(text)
        self.assertEqual(rendered[0], [('#沒有空白不是標題', '')])
        self.assertEqual(rendered[1], [('壞的 **沒有結尾', '')])
        self.assertEqual(rendered[2], [('**  前後有空白不算粗體  **', '')])
        self.assertEqual(rendered[3], [('`程式碼` 與 [連結](https://example.com)', '')])
        self.assertEqual(rendered[4], [('    ## 縮排太多不是標題', '')])
        self.assertEqual(rendered[5], [('****', '')])
        self.assertEqual(rendered[6], [])
        self.assertEqual(rendered[8], [('2 ** 3 ** 4', '')])
        self.assertEqual(u.render_notes(''), [[]])
        self.assertEqual(u.render_notes(None), [[]])
        flat = NL.join(''.join(t for t, _ in line) for line in u.render_notes('## A' + NL + '**B** c'))
        self.assertEqual(flat, 'A' + NL + 'B c')                                          # 組回去只少了標記符號，內容沒掉

    def test_render_notes_keeps_all_the_text(self):
        for sample in ('純文字沒有任何記號', '# 只有標題', '**只有粗體**', 'a **b** c __d__ e'):
            marks = ''.join(t for line in u.render_notes(sample) for t, _ in line)
            self.assertEqual(marks, sample.replace('# ', '').replace('**', '').replace('__', ''))

    def test_should_check_throttle(self):
        now = datetime(2026, 10, 20, 12, 0, 0)
        self.assertTrue(u.should_check(None, now))
        self.assertTrue(u.should_check('壞掉的時間', now))
        self.assertEqual(u.CHECK_INTERVAL_SECONDS, 7 * 24 * 3600)                              # 每週一次
        self.assertFalse(u.should_check((now - timedelta(hours=1)).isoformat(), now))
        self.assertFalse(u.should_check((now - timedelta(days=1)).isoformat(), now))
        self.assertFalse(u.should_check((now - timedelta(days=6, hours=23, minutes=59)).isoformat(), now))
        self.assertTrue(u.should_check((now - timedelta(days=7)).isoformat(), now))
        self.assertTrue(u.should_check((now - timedelta(days=30)).isoformat(), now))
        self.assertTrue(u.should_check((now + timedelta(hours=5)).isoformat(), now))           # 時鐘被調回去了

    def test_version_constant_is_not_behind_the_latest_git_tag(self):
        """打了新 tag 卻忘了更新 VERSION，更新檢查就會誤報；這裡擋下來（不在 git 倉庫裡就略過）。"""
        tag = g._git(['describe', '--tags', '--abbrev=0'])
        if not tag:
            self.skipTest('不在 git 倉庫或還沒有 tag')
        self.assertGreaterEqual(u.parse_version(g.VERSION), u.parse_version(tag), f'VERSION={g.VERSION} 落後最新 tag {tag}')


class CheckForUpdateTest(unittest.TestCase):
    def server(self, mode):
        server = FakeGitHub(mode)
        self.addCleanup(server.close)
        return server

    def check(self, mode, current='0.8', skipped=''):
        server = self.server(mode)
        return u.check_for_update(REPO, current, skipped, api_base=server.url, timeout=5), server

    def test_new_release_is_offered(self):
        result, server = self.check('newer')
        self.assertEqual(result.status, 'available')
        release = result.release
        self.assertEqual((release.tag, release.version, release.published), ('v0.9', (0, 9, 0, 0), '2026-10-20'))
        self.assertIn('新增功能 A', release.notes)

    def test_download_link_is_ours_never_the_servers(self):
        result, _ = self.check('newer')
        self.assertEqual(result.release.url, f'https://github.com/{REPO}/releases/tag/v0.9')
        self.assertTrue(u.is_trusted_release_url(result.release.url, REPO))
        self.assertNotIn('evil.example', repr(result.release))                          # 伺服器回傳的網址與檔案連結完全沒被採用

    def test_same_older_prerelease_and_draft_are_not_updates(self):
        self.assertEqual(self.check('same', current='0.8')[0].status, 'up_to_date')
        self.assertEqual(self.check('older')[0].status, 'up_to_date')                   # 不會降版
        self.assertEqual(self.check('prerelease')[0].status, 'up_to_date')
        self.assertEqual(self.check('draft')[0].status, 'up_to_date')
        self.assertEqual(self.check('newer', current='0.9')[0].status, 'up_to_date')
        self.assertEqual(self.check('newer', current='v1.0')[0].status, 'up_to_date')

    def test_skipped_version_is_not_offered_again(self):
        result, _ = self.check('newer', skipped='v0.9')
        self.assertEqual(result.status, 'skipped')
        self.assertEqual(result.release.tag, 'v0.9')
        server = self.server('newer')
        server.latest_tag = 'v0.10'                                                       # 出更新的版本就會再提醒
        self.assertEqual(u.check_for_update(REPO, '0.8', 'v0.9', api_base=server.url).status, 'available')

    def test_unparseable_tags_are_errors_not_updates(self):
        for mode in ('badtag', 'rc'):
            result, _ = self.check(mode)
            self.assertEqual(result.status, 'error', mode)
            self.assertIn('無法辨識', result.message)

    def test_bad_responses_are_reported(self):
        for mode, expected in (('garbage', '格式不正確'), ('array', '格式不正確'), ('huge', '太大'), ('403', '限制查詢'),
                               ('500', '500')):
            result, _ = self.check(mode)
            self.assertEqual(result.status, 'error', mode)
            self.assertIn(expected, result.message, mode)

    def test_no_release_yet_or_private_repo_is_reported_as_such(self):
        result, _ = self.check('404')
        self.assertEqual(result.status, 'no_release')                                    # 404：還沒發行，或 repo 沒公開
        self.assertIn('沒有公開', result.message)
        self.assertIsNone(result.release)

    def test_unreachable_server(self):
        server = FakeGitHub()
        url = server.url
        server.close()
        result = u.check_for_update(REPO, '0.8', api_base=url, timeout=3)
        self.assertEqual(result.status, 'error')
        self.assertIn('無法連線', result.message)

    def test_disabled_without_a_repo_makes_no_request(self):
        server = self.server('newer')
        self.assertEqual(u.check_for_update('', '0.8', api_base=server.url).status, 'disabled')
        self.assertEqual(server.requests, [])
        result = u.check_for_update('不是 repo', '0.8', api_base=server.url)
        self.assertEqual(result.status, 'error')
        self.assertEqual(server.requests, [])

    def test_request_carries_no_personal_data(self):
        result, server = self.check('newer')
        request = server.requests[0]
        self.assertEqual(request['path'], f'/repos/{REPO}/releases/latest')
        headers = {k.lower(): v for k, v in request['headers'].items()}
        self.assertEqual(headers['accept'], 'application/vnd.github+json')
        self.assertEqual(headers['user-agent'], 'yscv-report-generator')
        for forbidden in ('cookie', 'authorization', 'x-forwarded-for', 'referer'):
            self.assertNotIn(forbidden, headers)

    def test_notes_are_plain_text_and_limited(self):
        result, _ = self.check('longnotes')
        notes = result.release.notes
        self.assertLessEqual(len(notes), u.NOTES_MAX_CHARS + 40)
        self.assertFalse(any(c in notes for c in ('\x00', '‮')))

    def test_current_version_must_be_understood(self):
        result, _ = self.check('newer', current='開發版')
        self.assertEqual(result.status, 'error')

    def test_nothing_is_downloaded_or_executed(self):
        with open(u.__file__, encoding='utf-8') as f:
            source = f.read()
        for forbidden in ('urlretrieve', 'subprocess', 'os.system', 'startfile', 'browser_download_url', 'assets'):
            self.assertNotIn(forbidden, source, forbidden)                               # 第一階段：只查詢，不下載、不執行


class SettingsTest(unittest.TestCase):
    def test_save_settings_merges_instead_of_overwriting(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, 's.json')
            self.assertEqual(g.load_settings(path), {})
            g.save_settings(path, theme='高對比')
            g.save_settings(path, update_skipped='v0.9')
            self.assertEqual(g.load_settings(path), {'theme': '高對比', 'update_skipped': 'v0.9'})
            self.assertEqual(g.load_theme_name(path), '高對比')
            with open(path, 'w', encoding='utf-8') as f:
                f.write('壞掉的內容')
            self.assertEqual(g.load_settings(path), {})
            with open(path, 'w', encoding='utf-8') as f:
                f.write('[1, 2]')
            self.assertEqual(g.load_settings(path), {})
            self.assertEqual(g.load_theme_name(path), g.DEFAULT_THEME)

    def test_changing_the_theme_keeps_the_update_settings(self):
        with tempfile.TemporaryDirectory() as folder:
            settings = os.path.join(folder, 's.json')
            g.save_settings(settings, update_skipped='v0.9', update_auto=False)
            root = tk.Tk()
            try:
                app = g.ReportGeneratorGUI(root, autosave_path=os.path.join(folder, 'a.json'), restore_on_start=False,
                                           settings_path=settings)
                app.apply_theme('高對比')
            finally:
                root.destroy()
                gc.collect()
            self.assertEqual(g.load_settings(settings),
                             {'update_skipped': 'v0.9', 'update_auto': False, 'theme': '高對比'})


class LauncherUpdateTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.settings = os.path.join(self.tmp.name, 's.json')
        self.server = FakeGitHub('newer')
        self.addCleanup(self.server.close)
        self.root = tk.Tk()
        self.opened, self.shown = [], []
        for p in (mock.patch.object(g, 'open_url', lambda url: self.opened.append(url) or True),
                  mock.patch.object(messagebox, 'showinfo', lambda t, m: self.shown.append(('info', t, m))),
                  mock.patch.object(messagebox, 'showerror', lambda t, m: self.shown.append(('err', t, m)))):
            p.start()
            self.addCleanup(p.stop)
        self.addCleanup(lambda: messagebox.attach(None))
        self.addCleanup(gc.collect)
        self.addCleanup(self.root.destroy)

    def launcher(self, repo=REPO, auto=False):
        launcher = g.ReportLauncher(self.root, settings_path=self.settings, update_repo=repo, auto_update_check=auto,
                                    update_api_base=self.server.url)
        return launcher

    def wait(self, launcher, timeout=8):
        t0 = time.time()
        while time.time() - t0 < timeout and launcher._update_checking:
            self.root.update()
            time.sleep(0.02)
        self.root.update()

    def banner_visible(self, launcher):
        return bool(launcher.update_banner.winfo_manager())

    def test_no_repo_means_no_update_ui_and_no_request(self):
        launcher = self.launcher(repo='', auto=True)
        self.root.update()
        time.sleep(0.5)
        self.root.update()
        self.assertIsNone(launcher.check_button)
        self.assertIsNone(launcher.auto_check_button)
        self.assertFalse(self.banner_visible(launcher))
        self.assertFalse(launcher.check_updates(manual=True))
        self.assertEqual(self.server.requests, [])

    def test_manual_check_shows_the_banner_and_the_notes(self):
        launcher = self.launcher()
        self.assertFalse(self.banner_visible(launcher))
        self.assertTrue(launcher.check_updates(manual=True))
        self.assertTrue(launcher.check_button.instate(['disabled']))                      # 檢查中：按鈕鎖住
        self.wait(launcher)
        self.assertFalse(launcher.check_button.instate(['disabled']))
        self.assertTrue(self.banner_visible(launcher))
        self.assertIn('v0.9', launcher.update_label.cget('text'))
        self.assertIn(f'v{g.VERSION}', launcher.update_label.cget('text'))
        notes = launcher.last_notes_dialog                                                # 手動檢查發現新版：直接顯示更新內容
        self.assertIn('新增功能 A', notes.text.get('1.0', 'end'))
        self.assertEqual(self.opened, [])                                                 # 還沒有開任何網頁
        notes.destroy()

    def test_download_button_opens_only_our_release_page(self):
        launcher = self.launcher()
        launcher.check_updates(manual=True)
        self.wait(launcher)
        launcher.last_notes_dialog.destroy()
        self.assertTrue(launcher.open_release_page())
        self.assertEqual(self.opened, [f'https://github.com/{REPO}/releases/tag/v0.9'])
        launcher.update_release = u.ReleaseInfo('v0.9', (0, 9, 0, 0), 'x', '', 'https://evil.example/phishing', '')
        self.assertFalse(launcher.open_release_page())                                    # 網址被改過：拒絕開啟
        self.assertEqual(len(self.opened), 1)
        self.assertEqual(self.shown[-1][0], 'err')

    def test_notes_dialog_shows_headings_and_bold_as_styles(self):
        launcher = self.launcher()
        notes = NL.join(['## 新功能', '- 加了 **粗體** 的說明', '', '**完整變更**：見網頁'])
        launcher.update_release = u.ReleaseInfo('v0.9', (0, 9, 0, 0), 'v0.9', u.clean_notes(notes),
                                                u.release_page_url(REPO, 'v0.9'), '2026-10-27')
        dialog = launcher.show_update_notes()
        self.addCleanup(dialog.destroy)
        text = dialog.text
        shown = text.get('1.0', 'end').rstrip(NL)
        self.assertEqual(shown, NL.join(['新功能', '- 加了 粗體 的說明', '', '完整變更：見網頁']))      # 標記符號不見了
        self.assertNotIn('##', shown)
        self.assertNotIn('**', shown)

        def styled(tag):
            ranges = text.tag_ranges(tag)
            return [text.get(ranges[i], ranges[i + 1]) for i in range(0, len(ranges), 2)]

        self.assertEqual(styled('heading'), ['新功能'])
        self.assertEqual(styled('bold'), ['粗體', '完整變更'])
        import tkinter.font as tkfont
        weight = lambda tag: tkfont.Font(root=self.root, font=text.tag_cget(tag, 'font')).actual('weight')
        self.assertEqual((weight('heading'), weight('bold')), ('bold', 'bold'))
        normal = tkfont.Font(root=self.root, font=text.cget('font')).actual('weight')
        self.assertEqual(normal, 'normal')                                                  # 其餘仍是一般字
        self.assertEqual(str(text.cget('state')), 'disabled')                              # 仍然唯讀

    def test_notes_dialog_without_notes(self):
        launcher = self.launcher()
        launcher.update_release = u.ReleaseInfo('v0.9', (0, 9, 0, 0), 'v0.9', '', u.release_page_url(REPO, 'v0.9'), '')
        dialog = launcher.show_update_notes()
        self.addCleanup(dialog.destroy)
        self.assertIn('沒有附更新說明', dialog.text.get('1.0', 'end'))

    def test_download_from_the_notes_dialog(self):
        launcher = self.launcher()
        launcher.check_updates(manual=True)
        self.wait(launcher)
        launcher.last_notes_dialog.download_button.invoke()
        self.assertEqual(self.opened, [f'https://github.com/{REPO}/releases/tag/v0.9'])
        self.assertFalse(launcher.last_notes_dialog.winfo_exists())

    def test_skip_hides_the_banner_and_is_remembered(self):
        launcher = self.launcher()
        launcher.check_updates(manual=True)
        self.wait(launcher)
        launcher.last_notes_dialog.destroy()
        launcher.skip_button.invoke()
        self.assertFalse(self.banner_visible(launcher))
        self.assertEqual(g.load_settings(self.settings)['update_skipped'], 'v0.9')
        # 之後自動檢查：伺服器還是 v0.9，但使用者略過了，所以不再提醒
        self.server.requests.clear()
        again = self.launcher()
        self.root.update()
        again.check_updates(manual=False)
        self.wait(again)
        self.assertEqual(len(self.server.requests), 1)
        self.assertFalse(self.banner_visible(again))
        # 手動檢查會無視略過（使用者主動要看）
        again.check_updates(manual=True)
        self.wait(again)
        self.assertTrue(self.banner_visible(again))
        again.last_notes_dialog.destroy()

    def test_up_to_date_and_error_are_reported_only_for_manual_checks(self):
        self.server.mode = 'older'
        launcher = self.launcher()
        launcher.check_updates(manual=False)
        self.wait(launcher)
        self.assertEqual(self.shown, [])                                                  # 自動檢查：沒新版就安靜
        launcher.check_updates(manual=True)
        self.wait(launcher)
        self.assertEqual(self.shown[-1][0:2], ('info', '檢查更新'))
        self.assertIn('最新版', self.shown[-1][2])
        self.server.mode = '500'
        n = len(self.shown)
        launcher.check_updates(manual=False)
        self.wait(launcher)
        self.assertEqual(len(self.shown), n)                                              # 自動檢查失敗：安靜
        launcher.check_updates(manual=True)
        self.wait(launcher)
        self.assertEqual(self.shown[-1][0:2], ('err', '檢查更新'))
        self.assertIn('500', self.shown[-1][2])

    def test_no_release_message_only_for_manual_checks_and_counts_as_checked(self):
        self.server.mode = '404'
        launcher = self.launcher()
        launcher.check_updates(manual=False)
        self.wait(launcher)
        self.assertEqual(self.shown, [])                                                  # 自動檢查：安靜
        self.assertFalse(self.banner_visible(launcher))
        self.assertIn('update_last_check', g.load_settings(self.settings))                # 有收到回應就算檢查過，不會每次啟動都重查
        launcher.check_updates(manual=True)
        self.wait(launcher)
        self.assertEqual(self.shown[-1][0:2], ('info', '檢查更新'))
        self.assertIn('沒有公開', self.shown[-1][2])
        self.assertNotIn('最新版', self.shown[-1][2])                                      # 不會誤說「已是最新版」

    def test_failed_check_does_not_count_as_checked(self):
        self.server.mode = '500'
        launcher = self.launcher()
        launcher.check_updates(manual=False)
        self.wait(launcher)
        self.assertNotIn('update_last_check', g.load_settings(self.settings))             # 失敗不記時間，下次啟動還會再試
        self.server.mode = 'newer'
        launcher.check_updates(manual=False)
        self.wait(launcher)
        self.assertIn('update_last_check', g.load_settings(self.settings))

    def test_auto_check_runs_once_a_week_and_reminds_from_the_saved_result(self):
        first = self.launcher(auto=True)
        self.root.update()
        t0 = time.time()
        while time.time() - t0 < 8 and (not self.server.requests or first._update_checking):
            self.root.update()
            time.sleep(0.02)
        self.wait(first)
        self.assertEqual(len(self.server.requests), 1)
        self.assertTrue(self.banner_visible(first))
        settings = g.load_settings(self.settings)
        self.assertEqual(settings['update_latest_tag'], 'v0.9')
        self.assertTrue(settings['update_last_check'])
        first.root.update()
        # 一週內再開：不連網，但用上次記下來的結果顯示提示
        second = self.launcher(auto=False)
        second._auto_check_updates()
        self.root.update()
        self.assertEqual(len(self.server.requests), 1)
        self.assertTrue(self.banner_visible(second))
        self.assertEqual(second.update_release.tag, 'v0.9')
        self.assertEqual(second.update_release.url, f'https://github.com/{REPO}/releases/tag/v0.9')
        # 超過一週：重新檢查
        old = (datetime.now() - timedelta(days=8)).isoformat(timespec='seconds')
        g.save_settings(self.settings, update_last_check=old)
        third = self.launcher()
        third._auto_check_updates()
        self.wait(third)
        self.assertEqual(len(self.server.requests), 2)

    def test_cached_reminder_is_dropped_once_you_are_up_to_date_or_skipped(self):
        g.save_settings(self.settings, update_last_check=datetime.now().isoformat(timespec='seconds'),
                        update_latest_tag=f'v{g.VERSION}')                                # 記下來的版本不比目前新
        launcher = self.launcher()
        launcher._auto_check_updates()
        self.assertFalse(self.banner_visible(launcher))
        g.save_settings(self.settings, update_latest_tag='v99.0', update_skipped='v99.0')
        launcher = self.launcher()
        launcher._auto_check_updates()
        self.assertFalse(self.banner_visible(launcher))
        g.save_settings(self.settings, update_latest_tag='../../壞掉', update_skipped='')
        launcher = self.launcher()
        launcher._auto_check_updates()                                                     # 設定檔被亂改也不會出錯
        self.assertFalse(self.banner_visible(launcher))
        self.assertEqual(self.server.requests, [])

    def test_turning_off_auto_check_stops_the_requests(self):
        launcher = self.launcher(auto=False)
        self.assertTrue(launcher.auto_check_var.get())                                    # 預設開啟
        launcher.auto_check_button.invoke()
        self.assertFalse(launcher.auto_check_var.get())
        self.assertFalse(g.load_settings(self.settings)['update_auto'])
        launcher._auto_check_updates()
        self.root.update()
        time.sleep(0.3)
        self.root.update()
        self.assertEqual(self.server.requests, [])
        self.assertTrue(self.launcher().auto_check_var.get() is False)                    # 下次啟動記得這個選擇

    def test_usage_hint_sits_under_the_heading_not_in_the_footer(self):
        launcher = self.launcher()
        self.root.update()
        hint = launcher.hint_label
        first_card = launcher.buttons[g.REPORT_CHOICES[0][0]]
        self.assertIn('點選或按數字鍵 1～3選擇；按 Esc 離開', hint.cget('text'))
        self.assertLess(hint.winfo_rooty(), first_card.winfo_rooty())                       # 在卡片上方（標題下面）
        self.assertGreater(hint.winfo_rooty(), self.root.winfo_rooty())
        footer_labels = [w for w in launcher.check_button.master.winfo_children() if w.winfo_class() == 'TLabel']
        self.assertEqual(footer_labels, [])                                                 # 下方那排只剩按鈕與勾選框
        bottoms = [card.winfo_rooty() + card.winfo_height() for card in launcher.buttons.values()]
        self.assertGreater(launcher.check_button.winfo_rooty(), max(bottoms))

    def test_footer_row_is_not_crowded(self):
        launcher = self.launcher()
        self.root.update()
        widgets = [launcher.auto_check_button, launcher.check_button, launcher.feedback_button]
        boxes = [(w.winfo_rootx(), w.winfo_rootx() + w.winfo_width()) for w in widgets]
        for (_, right), (left, _) in zip(boxes, boxes[1:]):
            self.assertGreaterEqual(left - right, 8)                                        # 彼此至少隔 8 像素
        gap = boxes[1][0] - boxes[0][1]
        self.assertGreaterEqual(gap, 150)                                                   # 左邊的勾選框與右邊的按鈕之間有明顯留白

    def test_footer_layout_keeps_feedback_at_the_right(self):
        launcher = self.launcher()
        self.root.update()
        xs = [w.winfo_rootx() for w in (launcher.auto_check_button, launcher.check_button, launcher.feedback_button)]
        self.assertEqual(xs, sorted(xs))                                                  # 左：自動檢查、右：檢查更新、意見反應
        self.assertGreater(launcher.feedback_button.winfo_rootx(), self.root.winfo_rootx() + self.root.winfo_width() // 2)

    def test_command_line_option(self):
        self.assertTrue(g.parse_args(['--no-update-check']).no_update_check)
        self.assertFalse(g.parse_args([]).no_update_check)


class NoNetworkByDefaultTest(unittest.TestCase):
    def test_configured_update_repo_is_valid(self):
        self.assertTrue(u.UPDATE_REPO == '' or u.is_valid_repo(u.UPDATE_REPO))

    def test_launcher_without_the_flag_never_checks(self):
        server = FakeGitHub()
        self.addCleanup(server.close)
        root = tk.Tk()
        try:
            launcher = g.ReportLauncher(root, settings_path=os.path.join(tempfile.mkdtemp(), 's.json'), update_repo=REPO,
                                        update_api_base=server.url)                       # auto_update_check 預設是 False
            root.update()
            time.sleep(0.6)
            root.update()
            self.assertEqual(server.requests, [])
        finally:
            messagebox.attach(None)
            root.destroy()
            gc.collect()


if __name__ == '__main__':
    unittest.main()
