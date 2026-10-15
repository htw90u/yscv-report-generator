"""檢查 GitHub 上有沒有新版（第一階段：只查詢，不下載、不執行任何東西；要不要更新由使用者自己決定）。

安全原則
- 只向 GitHub 的公開 API 發一個 GET，不附帶任何使用者資料（User-Agent 只有程式名稱與版本）。
- 「前往下載」的網址由程式自己用 repo 名稱與版本標籤組出來，絕不使用伺服器回傳的網址。
- 只接受版本號比目前更新的正式發行（草稿、預告版不算），不會降版。
- 更新說明只當純文字顯示，限制長度並拿掉控制字元。
"""
import json
import re
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import datetime

UPDATE_REPO = 'htw90u/yscv-report-generator'   # 'owner/repo'；空字串＝不啟用更新檢查（不連網、選單也不顯示相關按鈕）
API_BASE = 'https://api.github.com'
CHECK_INTERVAL_SECONDS = 7 * 24 * 3600    # 自動檢查最多每週一次
TIMEOUT = 8                           # 秒
MAX_RESPONSE_BYTES = 1_000_000
NOTES_MAX_CHARS = 4000
NO_RELEASE_MESSAGE = "GitHub 上找不到已發行的版本（可能還沒發行，或這個 repo 目前沒有公開）"
_CONTROL_CHARS = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\x7f  ‪-‮⁦-⁩]')
# GitHub 的帳號只有英數字與連字號（不能以連字號開頭或結尾，最長 39 字）；repo 名稱可以有英數字、-、_、.（但不能是 . 或 ..）
_REPO_PATTERN = re.compile(r'[A-Za-z0-9](?:[A-Za-z0-9\-]{0,37}[A-Za-z0-9])?/[A-Za-z0-9_.\-]{1,100}')


@dataclass(frozen=True)
class ReleaseInfo:
    tag: str            # 例如 v0.9
    version: tuple      # (0, 9, 0, 0)
    name: str
    notes: str          # 純文字的更新說明
    url: str            # 該版的發行頁面（程式自己組出來的 https://github.com/<repo>/releases/tag/<tag>）
    published: str      # 發行日期（只取日期，例如 2026-10-20）


@dataclass(frozen=True)
class UpdateResult:
    status: str         # available（有新版）、up_to_date、skipped（有新版但使用者略過了）、no_release（找不到發行）、error、disabled
    release: ReleaseInfo = None
    message: str = ''


def parse_version(text):
    """「v0.9」「0.9.1」→ (0, 9, 1, 0)；格式不是單純的數字版本（例如 v1.0-rc1）回傳 None。"""
    match = re.fullmatch(r'[vV]?([0-9]{1,6}(?:\.[0-9]{1,6}){0,3})', str(text or '').strip())
    if not match:
        return None
    numbers = [int(part) for part in match.group(1).split('.')]
    return tuple(numbers + [0] * (4 - len(numbers)))


def is_valid_repo(repo):
    repo = str(repo or '')
    return bool(_REPO_PATTERN.fullmatch(repo)) and repo.split('/')[1] not in ('.', '..')


def release_page_url(repo, tag):
    """某一版的發行頁面網址（自己組，不信任伺服器回傳的網址）。"""
    return f"https://github.com/{repo}/releases/tag/{urllib.parse.quote(tag, safe='')}"


def is_trusted_release_url(url, repo):
    """只有 https://github.com/<repo>/releases… 才算可以開啟的發行頁面。"""
    parts = urllib.parse.urlsplit(str(url or ''))
    return (parts.scheme == 'https' and parts.hostname == 'github.com' and parts.port is None and not parts.username
            and parts.path.lower().startswith(f"/{repo.lower()}/releases"))


def clean_notes(text):
    text = _CONTROL_CHARS.sub('', str(text or '').replace('\r\n', '\n').replace('\r', '\n')).strip()
    return text if len(text) <= NOTES_MAX_CHARS else text[:NOTES_MAX_CHARS].rstrip() + '\n…（內容太長，完整說明請看發行頁面）'


def fetch_latest_release(repo, api_base=API_BASE, timeout=TIMEOUT, user_agent='yscv-report-generator'):
    """查詢最新的正式發行。回傳 (ReleaseInfo 或 None, 錯誤訊息)；沒有錯誤時訊息是空字串，
    找不到發行時（404）錯誤訊息是 NO_RELEASE_MESSAGE；草稿與預告版不算發行（回傳 None 與空訊息）。"""
    if not is_valid_repo(repo):
        return None, "更新來源的設定不正確"
    request = urllib.request.Request(
        f"{api_base}/repos/{repo}/releases/latest",
        headers={'Accept': 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28',
                 'User-Agent': user_agent})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read(MAX_RESPONSE_BYTES + 1)
    except urllib.error.HTTPError as e:
        code = e.code
        e.close()
        if code == 404:
            return None, NO_RELEASE_MESSAGE       # 還沒有正式發行，或這個 repo 沒有公開（沒登入的人看不到私人 repo，也是 404）
        if code in (403, 429):
            return None, "GitHub 暫時限制查詢次數，請稍後再試"
        return None, f"GitHub 回應錯誤（{code}）"
    except (urllib.error.URLError, OSError) as e:
        return None, f"無法連線到 GitHub（{getattr(e, 'reason', e)}）"
    if len(raw) > MAX_RESPONSE_BYTES:
        return None, "GitHub 的回應太大，已忽略"
    try:
        data = json.loads(raw.decode('utf-8', errors='replace'))
    except ValueError:
        return None, "GitHub 的回應格式不正確"
    if not isinstance(data, dict):
        return None, "GitHub 的回應格式不正確"
    if data.get('draft') or data.get('prerelease'):
        return None, ''                           # 草稿與預告版不算正式發行
    tag = str(data.get('tag_name') or '').strip()
    version = parse_version(tag)
    if version is None:
        return None, f"最新發行的版本編號無法辨識（{tag[:30]}）"
    return ReleaseInfo(tag=tag, version=version, name=clean_notes(data.get('name') or tag)[:100],
                       notes=clean_notes(data.get('body')), url=release_page_url(repo, tag),
                       published=str(data.get('published_at') or '')[:10]), ''


def check_for_update(repo, current_version, skipped='', api_base=API_BASE, timeout=TIMEOUT,
                     user_agent='yscv-report-generator'):
    """查詢並判斷：有沒有比 current_version 更新的正式發行。skipped 是使用者按過「略過這個版本」的版本標籤。"""
    if not repo:
        return UpdateResult('disabled')
    release, error = fetch_latest_release(repo, api_base, timeout, user_agent)
    if error == NO_RELEASE_MESSAGE:
        return UpdateResult('no_release', message=error)
    if error:
        return UpdateResult('error', message=error)
    if release is None:
        return UpdateResult('up_to_date')
    current = parse_version(current_version)
    if current is None:
        return UpdateResult('error', message="目前的版本編號無法辨識")
    if release.version <= current:
        return UpdateResult('up_to_date', release)
    if skipped and release.tag == skipped:
        return UpdateResult('skipped', release)
    return UpdateResult('available', release)


def should_check(last_check, now=None, interval=CHECK_INTERVAL_SECONDS):
    """自動檢查要不要執行：從沒檢查過、時間格式壞掉、時鐘被調回去，或距離上次超過 interval 秒就檢查。"""
    now = now or datetime.now()
    try:
        last = datetime.fromisoformat(str(last_check))
    except ValueError:
        return True
    elapsed = (now - last).total_seconds()
    return elapsed < 0 or elapsed >= interval
