#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
玉山保育巡查回報單生成器 - GUI版本
提供圖形化介面供使用者輸入資料並選擇照片
"""

import base64
import copy
import io
import json
import logging
import math
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog
from tkinter import font as tkfont

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PIL import Image, ImageOps, ImageTk

APP_TITLE = '玉山保育巡查回報單生成器'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 固定使用的模板（與程式放在同一個資料夾）
TEMPLATE_PATH = os.path.join(BASE_DIR, '玉山國家公園高山保育志工協勤回報單.docx')
GUI_FONT_DELTA = 2  # 產生器介面的字體在系統預設上加大的點數
GUI_FONT_WEIGHT = 'bold'  # 介面字型粗細：'bold'（粗體）或 'normal'（一般）
# 介面字型：霞鶩文楷等寬 TC（SIL OFL 授權，字型檔放在 fonts/ 資料夾，執行時私下載入、不安裝到系統）
FONT_DIR = os.path.join(BASE_DIR, 'fonts')
FONT_FILES = ('LXGWWenKaiMonoTC-Regular.ttf', 'LXGWWenKaiMonoTC-Bold.ttf')
FONT_FAMILY = 'LXGW WenKai Mono TC'
FALLBACK_FONT_FAMILY = 'Microsoft JhengHei UI'
LOG_FILE = os.path.join(BASE_DIR, 'report_generator_gui.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

KAI = '標楷體'
BODY_FONT_PT = 12   # 內容字級（與參考回報單相同，未特別指定時為 12pt）
LINE_TWIPS = 440    # 固定行高 22pt（參考回報單）

DRAFT_VERSION = 1
DEFAULT_AUTOSAVE_PATH = os.path.join(BASE_DIR, 'report_generator_autosave.json')
AUTOSAVE_INTERVAL_MS = 30_000  # 每 30 秒檢查一次，有變動才寫檔
# 基本資料欄位（同時是 ReportGeneratorGUI 上的屬性名稱）
REPORT_DATE_FIELDS = ('year', 'month', 'day')  # 「填表日期」：自動帶入第一次填表的日期
ENTRY_FIELDS = ['year', 'month', 'day', 'personnel',
                'start_year', 'start_month', 'start_day',
                'end_year', 'end_month', 'end_day', 'location']

DEFAULT_CAPTION = ''  # 照片沒有填說明時，文件中顯示的預設字樣（預設空白，可在照片分頁修改）
LEGACY_DEFAULT_CAPTION = '[照片說明]'  # 舊版預設值；載入舊草稿時視為「沒改過」，改成空白
# 外觀主題。每個主題：表單背景、分頁（閒置/目前/滑過）、主要按鈕（產生報告）、危險按鈕（清除）。
# 主要與危險按鈕都是「淺底深字」，文字與底色的對比度都達到 WCAG AA（4.5:1 以上）。
THEMES = {
    '柔和淺色': dict(
        bg='#f6f8fb', border='#b5c0d0',
        tab_idle='#e6ecf4', tab_idle_fg='#475569', tab_hover='#d9e3f1',
        tab_active='#cfe0fb', tab_active_fg='#1e3a8a',
        primary='#cdeedd', primary_fg='#14532d', primary_active='#b4e3c9', primary_border='#6fbf96',
        danger='#fde2e2', danger_fg='#991b1b', danger_active='#fbc9c9', danger_border='#e89a9a'),
    # 色弱友善：避開紅綠搭配，改用藍色（主要）與橘色（危險），取自 Okabe-Ito 色盲友善色盤
    '色弱友善': dict(
        bg='#f7f7f5', border='#a8a29e',
        tab_idle='#ececea', tab_idle_fg='#44403c', tab_hover='#e0e0dd',
        tab_active='#cfe6f5', tab_active_fg='#003a5c',
        primary='#b8dcf2', primary_fg='#00314d', primary_active='#9fd0ee', primary_border='#0072B2',
        danger='#fbdcc0', danger_fg='#6b2d00', danger_active='#f6c79d', danger_border='#D55E00'),
    '高對比': dict(
        bg='#ffffff', border='#000000',
        tab_idle='#ffffff', tab_idle_fg='#000000', tab_hover='#e5e5e5',
        tab_active='#000000', tab_active_fg='#ffffff',
        primary='#0b3d91', primary_fg='#ffffff', primary_active='#082a66', primary_border='#000000',
        danger='#b34700', danger_fg='#ffffff', danger_active='#8a3700', danger_border='#000000'),
}
DEFAULT_THEME = '柔和淺色'
DEFAULT_SETTINGS_PATH = os.path.join(BASE_DIR, 'report_generator_settings.json')
DATE_UNIT_PADX = 4     # 年/月/日 單位字左右的間距（對稱）
DATE_RANGE_PADX = 12   # 「至」左右的間距（對稱）
TAB_PADDING = (18, 8)         # 一般分頁內距（左右, 上下）
TAB_PADDING_ACTIVE = (18, 12)  # 目前分頁：左右一樣、上下多一點，看起來比其他分頁更突出
PREVIEW_SIZE = (360, 300)  # 照片預覽區大小（像素）
MAX_IMAGE_WIDTH_PX = 1200  # 照片在文件中寬 2.8 吋，1200px 約 430dpi，已足夠清晰
JPEG_QUALITY = 85

# 照片版面（單位：吋）。橫式/直式照片都以填滿儲存格寬度為原則，並模擬分頁以盡量排滿整頁
CELL_SIDE_PAD = 0.2     # 儲存格左右內距與餘裕
HEADER_H = 0.32         # 分類標題列
SPACER_H = 0.15         # 分類表格之間的間隔
HEADING_H = 0.45        # 「九、服勤照片」標題
PHOTO_ROW_PAD = 0.04
CAPTION_LINE_H = 0.2    # 說明文字每行高度（12pt）
CAPTION_PAD = 0.08
MIN_PHOTO_SCALE = 0.8   # 為了排滿整頁，照片最多等比例縮小到 80%
PAGE_SAFETY = 0.25      # 每頁預留的誤差空間，避免估算偏差造成多出一頁

PHOTO_CATEGORIES_1 = ['道路、步道狀況及維護情形', '設施狀況及維護情形', '自然資源狀況', '環境清理及遊客宣導情形']
PHOTO_CATEGORIES_2 = ['遊客違法情形', '其他']

# 模板中「標籤後面接內容」的欄位：(標籤文字, 資料 key)
FILL_LABELS = [
    ('1、道路、步道狀況及維護情形：', 'road_status'),
    ('2、道路、步道雜草情形：', 'road_grass'),
    ('3、其他：', 'road_other'),
    ('1、山屋：', 'facility_house'),
    ('2、牌示：', 'facility_sign'),
    ('3、木棧道：', 'facility_bridge'),
    ('4、其他：', 'facility_other'),
    ('數量及周圍環境紀錄等）：', 'nature'),
    ('護宣導等情形）：', 'environment'),
    ('團體與人員）：', 'violation'),
    ('六、登山遊客反應意見：', 'feedback'),
    ('七、協勤心得與建議：', 'reflection'),
]
QUESTION_MARKER = '八、相關疑問與解答：'
QUESTION_PLACEHOLDER = '1.相關疑問？答：因為......。'

FIELD_NAMES = {
    'road_status': '1、道路、步道狀況及維護情形', 'road_grass': '2、道路、步道雜草情形',
    'road_other': '3、其他（道路步道）', 'facility_house': '1、山屋', 'facility_sign': '2、牌示',
    'facility_bridge': '3、木棧道', 'facility_other': '4、其他（設施）', 'nature': '三、自然資源狀況',
    'environment': '四、環境清理及遊客宣導情形', 'violation': '五、遊客違法情形',
    'feedback': '六、登山遊客反應意見', 'reflection': '七、協勤心得與建議', 'questions': '八、相關疑問與解答',
}


# ---------------------------------------------------------------- 文件處理（不依賴 tkinter）

# 下拉選單提供的項目符號：(符號, 名稱)
BULLET_CHOICES = [('•', '圓點'), ('◆', '實心菱形'), ('■', '實心方塊')]
BULLET_SYMBOLS = [sym for sym, _ in BULLET_CHOICES]
# 其他常見符號（使用者自己貼上的）也當成項目符號處理；連字號類必須後面接空白，避免把「-3度」「—」誤判
_SYMBOL_CHARS = '•·‧◆◇■□●○★☆▲△▼▽※＊*'
_MARKER_PATTERN = rf'(?:[{_SYMBOL_CHARS}]|[-–—](?=\s))'
LIST_MARKER_RE = re.compile(rf'^(?:{_MARKER_PATTERN}|\d+\s*[.、)）]|[一二三四五六七八九十]+、)')
NUMBER_MARKER_RE = re.compile(r'^\d+\s*[.、)）]')
BULLET = '• '


def content_lines(text):
    """把使用者輸入拆成要寫入文件的行。項目符號與編號由使用者在編輯區自行加入，這裡不再自動補。"""
    return [line.strip() for line in (text or '').split('\n')]


# ---- 編輯區的項目符號 / 編號（純文字處理，方便測試）
_STRIP_MARKER_RE = re.compile(rf'^(?:{_MARKER_PATTERN}\s*|\d+\s*[.、)）]\s*)')


def strip_list_marker(line):
    return _STRIP_MARKER_RE.sub('', line.lstrip(), count=1)


def toggle_bullets(lines, symbol=BULLET_SYMBOLS[0]):
    """全部有內容的行都已是該符號開頭就取消；否則全部換成「符號 + 空白」。空白行不動。"""
    filled = [ln for ln in lines if ln.strip()]
    if filled and all(ln.lstrip().startswith(symbol) for ln in filled):
        return [strip_list_marker(ln) if ln.strip() else ln for ln in lines]
    return [f'{symbol} {strip_list_marker(ln)}' if ln.strip() else ln for ln in lines]


def remove_list_markers(lines):
    return [strip_list_marker(ln) if ln.strip() else ln for ln in lines]


def toggle_numbers(lines):
    """全部有內容的行都已有編號就取消；否則依序加上 1. 2. 3.（空白行不編號）。"""
    filled = [ln for ln in lines if ln.strip()]
    if filled and all(NUMBER_MARKER_RE.match(ln.lstrip()) for ln in filled):
        return [strip_list_marker(ln) if ln.strip() else ln for ln in lines]
    result, n = [], 1
    for ln in lines:
        if ln.strip():
            result.append(f'{n}. {strip_list_marker(ln)}')
            n += 1
        else:
            result.append(ln)
    return result


def continue_list(text_before_cursor):
    """在項目符號/編號的行按 Enter：回傳 ('continue', 下一行前綴)、('end', '')（空項目，結束清單）或 (None, '')。"""
    line = text_before_cursor.lstrip()
    symbol = next((sym for sym in BULLET_SYMBOLS if line.startswith(sym)), None)
    if symbol and (symbol != '-' or line[1:2].isspace()):
        rest, prefix = line[len(symbol):].strip(), f'{symbol} '
    else:
        m = re.match(r'^(\d+)\s*[.、)）]\s*', line)
        if not m:
            return None, ''
        rest, prefix = line[m.end():].strip(), f'{int(m.group(1)) + 1}. '
    return ('end', '') if not rest else ('continue', prefix)


def set_cell_border(cell):
    """設定儲存格邊框"""
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tcBorders.append(border)
    tcPr.append(tcBorders)


def set_font_kai(run, size=BODY_FONT_PT):
    """標楷體（中英文、數字都用標楷體，與參考回報單一致）。"""
    run.font.name = KAI
    run.font.size = Pt(size)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), KAI)


def _kai_rpr():
    rpr = OxmlElement('w:rPr')
    rfonts = OxmlElement('w:rFonts')
    for attr in ('ascii', 'hAnsi', 'eastAsia'):
        rfonts.set(qn(f'w:{attr}'), KAI)
    rpr.append(rfonts)
    for tag in ('w:sz', 'w:szCs'):
        size = OxmlElement(tag)
        size.set(qn('w:val'), str(BODY_FONT_PT * 2))
        rpr.append(size)
    return rpr


def _build_content_paragraph(text):
    """內容段落：標楷體 12pt、固定行高 22pt（格式依參考回報單）。

    開頭是項目符號或編號的行用懸掛縮排（換行後文字對齊在符號之後），其餘行直接縮排。
    """
    p = OxmlElement('w:p')
    ppr = OxmlElement('w:pPr')
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), str(LINE_TWIPS))
    spacing.set(qn('w:lineRule'), 'exact')
    ind = OxmlElement('w:ind')
    if text and NUMBER_MARKER_RE.match(text):
        ind.set(qn('w:left'), '840')
        ind.set(qn('w:hanging'), '360')
    elif text and LIST_MARKER_RE.match(text):
        ind.set(qn('w:left'), '720')
        ind.set(qn('w:hanging'), '240')
    else:
        ind.set(qn('w:left'), '480')
    ppr.append(spacing)
    ppr.append(ind)
    ppr.append(_kai_rpr())
    p.append(ppr)
    if text:
        run = OxmlElement('w:r')
        run.append(_kai_rpr())
        t = OxmlElement('w:t')
        t.text = text
        t.set(qn('xml:space'), 'preserve')
        run.append(t)
        p.append(run)
    return p


def _paragraph_text(element):
    return ''.join(t.text or '' for t in element.iter(qn('w:t')))


def _insert_content_after(paragraph, lines, drop_text=None):
    """在 paragraph 後面插入內容段落，並移除其後的空白佔位段落（含 drop_text 的佔位文字）。"""
    anchor = paragraph._p
    nxt = anchor.getnext()
    while nxt is not None and nxt.tag == qn('w:p') and (
            _is_empty_paragraph(nxt) or (drop_text and drop_text in _paragraph_text(nxt))):
        anchor.getparent().remove(nxt)
        nxt = anchor.getnext()
    for text in lines:
        new = _build_content_paragraph(text)
        anchor.addnext(new)
        anchor = new


def _set_cell_value(cell, text, center=False):
    """填入基本資料：沿用儲存格原有的段落設定與字型大小（對齊、14pt 等），字型改為標楷體。"""
    paragraphs = cell.paragraphs
    for extra in paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    p = paragraphs[0]
    runs = p.runs
    if runs:
        run = runs[0]
        run.text = text
        for extra in runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        run = p.add_run(text)
        ppr = p._p.pPr
        mark = ppr.find(qn('w:rPr')) if ppr is not None else None
        if mark is not None:  # 空儲存格的字型大小記錄在段落標記上
            rpr = run._r.get_or_add_rPr()
            for child in mark:
                rpr.append(copy.deepcopy(child))
    rfonts = run._r.get_or_add_rPr().get_or_add_rFonts()
    for attr in ('ascii', 'hAnsi', 'eastAsia'):
        rfonts.set(qn(f'w:{attr}'), KAI)
    rfonts.attrib.pop(qn('w:hint'), None)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def prepare_image(path):
    """回傳 (來源, (寬px, 高px))；來源可直接給 add_picture 使用。

    有 EXIF 旋轉時先轉正；寬度超過 MAX_IMAGE_WIDTH_PX 時縮小，避免 docx 過大。
    不需要處理時來源直接是原檔路徑（不重新壓縮）。回傳的尺寸是處理後（轉正後）的尺寸。
    """
    with Image.open(path) as img:
        needs_rotate = img.getexif().get(0x0112, 1) not in (0, 1)
        if needs_rotate:
            img = ImageOps.exif_transpose(img)
        needs_resize = img.width > MAX_IMAGE_WIDTH_PX
        if not needs_rotate and not needs_resize:
            return path, (img.width, img.height)
        if needs_resize:
            new_height = max(1, round(img.height * MAX_IMAGE_WIDTH_PX / img.width))
            img = img.resize((MAX_IMAGE_WIDTH_PX, new_height), Image.LANCZOS)
        buf = io.BytesIO()
        if img.mode in ('RGBA', 'LA', 'P'):
            img.convert('RGBA').save(buf, 'PNG', optimize=True)
        else:
            img.convert('RGB').save(buf, 'JPEG', quality=JPEG_QUALITY, optimize=True)
        buf.seek(0)
        return buf, (img.width, img.height)


def _all_text(table):
    return '\n'.join(c.text for row in table.rows for c in row.cells)


LABEL_COLUMN_TWIPS = 2300  # 「協勤人員姓名：」等標籤欄的寬度（14pt × 7 字 + 儲存格邊距，原本 2093 會被換成兩行）


def widen_label_column(table, target=LABEL_COLUMN_TWIPS):
    """把基本資料表格第一欄加寬到 target（twips），從第二欄扣掉，整個表格寬度不變。"""
    grid = table._tbl.tblGrid.findall(qn('w:gridCol'))
    if len(grid) < 3:
        return
    widths = [int(c.get(qn('w:w'))) for c in grid]
    delta = target - widths[0]
    if delta <= 0 or widths[1] - delta < 1000:
        return
    change = [delta, -delta] + [0] * (len(grid) - 2)
    for col, c in zip(change, grid):
        c.set(qn('w:w'), str(int(c.get(qn('w:w'))) + col))
    for tr in table._tbl.findall(qn('w:tr')):
        position = 0
        for tc in tr.findall(qn('w:tc')):
            tcPr = tc.find(qn('w:tcPr'))
            span_el = tcPr.find(qn('w:gridSpan')) if tcPr is not None else None
            span = int(span_el.get(qn('w:val'))) if span_el is not None else 1
            tcW = tcPr.find(qn('w:tcW')) if tcPr is not None else None
            shift = sum(change[position:position + span])
            if tcW is not None and shift and tcW.get(qn('w:type')) == 'dxa':
                tcW.set(qn('w:w'), str(int(tcW.get(qn('w:w'))) + shift))
            position += span


def _fill_main_table(main_table, data):
    """把文字資料填進主表格，回傳警告清單。

    以段落為單位插入內容，保留模板原有的標題（粗體/14pt）等格式；
    插入的內容採標楷體 12pt、固定行高 22pt（參考他人填寫的回報單）。
    """
    warnings = []
    widen_label_column(main_table)

    try:
        _set_cell_value(main_table.rows[0].cells[2], f"{data['year']}年{data['month']}月{data['day']}日", center=True)
        _set_cell_value(main_table.rows[1].cells[1], data['personnel'])
        _set_cell_value(main_table.rows[2].cells[1],
                        f"{data['start_year']}年{data['start_month']}月{data['start_day']}日至"
                        f"{data['end_year']}年{data['end_month']}月{data['end_day']}日", center=True)
        _set_cell_value(main_table.rows[3].cells[1], data['location'])
    except IndexError:
        raise ValueError("模板格式錯誤：基本資料表格的列/欄數量與預期不符（需至少 4 列、第一列至少 3 欄）")

    filled = set()
    seen = set()
    for row in main_table.rows:
        for cell in row.cells:
            # 合併儲存格會在 row.cells 中重複出現，同一格只處理一次
            if cell._tc in seen:
                continue
            seen.add(cell._tc)

            for marker, key in FILL_LABELS:
                if key in filled:
                    continue
                paragraph = next((p for p in cell.paragraphs if marker in p.text), None)
                if paragraph is None:
                    continue
                lines = content_lines(data.get(key, ''))
                if data.get(key):
                    _insert_content_after(paragraph, lines)
                filled.add(key)

            if 'questions' not in filled:
                head = next((p for p in cell.paragraphs if QUESTION_MARKER in p.text), None)
                if head is not None:
                    if data.get('questions'):
                        _insert_content_after(head, content_lines(data['questions']),
                                              drop_text=QUESTION_PLACEHOLDER)
                    filled.add('questions')

    for key, name in FIELD_NAMES.items():
        if data.get(key) and key not in filled:
            warnings.append(f"模板中找不到「{name}」欄位，這部分內容未填入")
    return warnings


@dataclass
class PhotoItem:
    path: str
    caption: str
    source: object = None  # add_picture 可用的來源；None 表示載入失敗
    px_w: int = 1
    px_h: int = 1
    img_w: float = 0.0     # 嵌入文件的尺寸（吋）
    img_h: float = 0.0

    @property
    def landscape(self):
        return self.px_w >= self.px_h


def size_photo(item, fit_width, max_portrait_h):
    """橫式照片寬度填滿儲存格；直式照片同樣以填滿寬度為原則，過高時改以高度上限縮小。"""
    aspect = item.px_h / item.px_w
    if item.landscape or fit_width * aspect <= max_portrait_h:
        item.img_w, item.img_h = fit_width, fit_width * aspect
    else:
        item.img_h = max_portrait_h
        item.img_w = max_portrait_h / aspect


def caption_height(text, line_chars):
    """估算說明文字占的高度（吋）。全形字算 1，半形字算 0.5。"""
    lines = 0
    for part in (text or ' ').split('\n'):
        units = sum(1 if ord(ch) > 255 else 0.5 for ch in part)
        lines += max(1, math.ceil(units / line_chars))
    return lines * CAPTION_LINE_H + CAPTION_PAD


def row_height(row, line_chars):
    return (max(p.img_h for p in row) + PHOTO_ROW_PAD
            + max(caption_height(p.caption, line_chars) for p in row))


class _Page:
    """模擬中的一頁：固定高度（標題列、說明文字等）與可縮放的照片高度分開計算。"""

    def __init__(self, fixed=0.0):
        self.fixed = fixed
        self.img = 0.0
        self.rows = []       # [(分類, 列)]
        self.last_category = None

    def _extra(self, category):
        """換到另一個分類時要多出的分類標題列與間隔。"""
        if category == self.last_category:
            return 0.0
        return HEADER_H + (SPACER_H if self.last_category is not None else 0.0)

    def fits(self, category, row, line_chars, page_h):
        fixed = (self.fixed + self._extra(category) + PHOTO_ROW_PAD
                 + max(caption_height(p.caption, line_chars) for p in row))
        img = self.img + max(p.img_h for p in row)
        return fixed + MIN_PHOTO_SCALE * img <= page_h

    def add(self, category, row, line_chars):
        self.fixed += (self._extra(category) + PHOTO_ROW_PAD
                       + max(caption_height(p.caption, line_chars) for p in row))
        self.img += max(p.img_h for p in row)
        self.rows.append((category, row))
        self.last_category = category

    def finish(self, page_h):
        """這頁照片放不下時，等比例縮小（不低於 MIN_PHOTO_SCALE），讓版面剛好排滿。"""
        if self.img <= 0:
            return
        scale = min(1.0, (page_h - self.fixed) / self.img)
        if scale < 1.0:
            for _, row in self.rows:
                for p in row:
                    p.img_w *= scale
                    p.img_h *= scale


def build_rows(items):
    """把一個分類的照片排成列：橫式在前、直式在後，每列 2 張。

    橫式與直式各剩一張落單時，兩張拼成一列（放在橫式與直式之間）。
    """
    land = [p for p in items if p.landscape]
    port = [p for p in items if not p.landscape]
    land_rows = [land[i:i + 2] for i in range(0, len(land), 2)]
    port_rows = [port[i:i + 2] for i in range(0, len(port), 2)]
    if len(land) % 2 and len(port) % 2:
        mixed = land_rows.pop() + port_rows.pop()
        return land_rows + [mixed] + port_rows
    return land_rows + port_rows


def plan_layout(categories, page_h, start_y, line_chars):
    """安排所有照片列，回傳依序排列的 [(分類, 列)]。

    - 每個分類（章節）的照片一定連在一起，分類之間依原本順序，不會交錯混放。
    - 分類內橫式排在一起、直式排在一起（見 build_rows）。
    - 模擬分頁：目前這一頁放不下下一列時，在「同一個分類」裡往後找放得下的列補位；
      照片等比例縮小一點（最多縮到 MIN_PHOTO_SCALE）就能多放一列時也會這樣做，
      讓整頁盡量排滿。都放不下才換頁。
    - page_h 為內文可用高度，start_y 為照片區開始時已用掉的高度（吋）。
    """
    page = _Page(fixed=start_y)
    sequence = []
    for name, items in categories.items():
        rows = build_rows(items)
        while rows:
            index = next((i for i, r in enumerate(rows) if page.fits(name, r, line_chars, page_h)), None)
            if index is None:
                if page.rows:
                    page.finish(page_h)
                    page = _Page()
                    continue
                index = 0  # 連空白頁都放不下（說明文字特別長）：硬放，讓 Word 自行處理
            row = rows.pop(index)
            page.add(name, row, line_chars)
            sequence.append((name, row))
    page.finish(page_h)
    return sequence


def _is_empty_paragraph(element):
    if element.tag != qn('w:p'):
        return False
    has_text = any((t.text or '').strip() for t in element.iter(qn('w:t')))
    has_object = any(True for _ in element.iter(qn('w:drawing'))) or any(True for _ in element.iter(qn('w:br')))
    return not has_text and not has_object


def _tighten(paragraph, keep_with_next=False):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    pf.keep_with_next = keep_with_next


def _add_spacer(doc):
    """表格之間的小間隔段落（相鄰表格會被 Word 合併，所以一定要有）。"""
    p = doc.add_paragraph()
    _tighten(p)
    rpr = OxmlElement('w:rPr')
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '12')  # 6pt
    rpr.append(sz)
    p._p.get_or_add_pPr().append(rpr)
    return p


def _mark_row(row, *flags):
    trPr = row._tr.get_or_add_trPr()
    for flag in flags:
        trPr.append(OxmlElement(flag))


def _load_photos(photo_files, captions, fit_width, max_portrait_h, warnings, default_caption=DEFAULT_CAPTION):
    """載入照片並計算尺寸。載入失敗的照片以文字佔位，不中斷整份報告。"""
    categories = {}
    for category, paths in photo_files.items():
        items = []
        for path in paths:
            caption = captions.get(category, {}).get(path) or default_caption
            item = PhotoItem(path=path, caption=caption)
            name = os.path.basename(path)
            try:
                item.source, (item.px_w, item.px_h) = prepare_image(path)
                size_photo(item, fit_width, max_portrait_h)
            except Exception as e:
                logging.error(f"無法載入照片 {name}: {e}")
                warnings.append(f"照片「{name}」無法載入：{e}")
                item.img_w, item.img_h = fit_width, 0.4
            items.append(item)
        categories[category] = items
    return categories


def _add_photo_section(doc, main_table, photo_files, captions, warnings, default_caption=DEFAULT_CAPTION):
    body = doc.element.body

    # 只刪除內容含「照片」的舊照片表格，避免誤刪其他表格
    for table in list(doc.tables)[1:]:
        if '照片' in _all_text(table):
            table._element.getparent().remove(table._element)

    # 頁面尺寸（吋）
    section = doc.sections[-1]
    usable_w = (section.page_width - section.left_margin - section.right_margin) / 914400
    page_h = (section.page_height - section.top_margin - section.bottom_margin) / 914400 - PAGE_SAFETY
    col_w = usable_w / 2
    fit_width = col_w - CELL_SIDE_PAD
    max_portrait_h = (page_h - HEADER_H - 2 * (CAPTION_LINE_H + CAPTION_PAD)) / 2 - PHOTO_ROW_PAD
    line_chars = fit_width * 72 / BODY_FONT_PT  # 全形字寬 = 字級

    # 模板內可能已有「九、服勤照片」標題（內文段落或主表格內），有的話不重複新增
    heading_p = next((p for p in doc.paragraphs if '九、服勤照片' in p.text), None)
    in_table = '九、服勤照片' in _all_text(main_table)
    anchor = None
    start_y = 0.0
    break_on_first_table = False

    if heading_p is not None:
        # 模板用大量空白段落把標題推到下一頁，改成「標題前換頁」，並清掉標題前後的空白段落
        prev = heading_p._p.getprevious()
        while prev is not None and _is_empty_paragraph(prev):
            body.remove(prev)
            prev = heading_p._p.getprevious()
        heading_p.paragraph_format.page_break_before = True
        while True:
            last = [e for e in body if e.tag != qn('w:sectPr')][-1]
            if last is not heading_p._p and _is_empty_paragraph(last):
                body.remove(last)
            else:
                break
        anchor = heading_p._p
        start_y = HEADING_H
    elif in_table:
        break_on_first_table = True
    else:
        heading = doc.add_paragraph('九、服勤照片')
        heading.paragraph_format.page_break_before = True
        heading_run = heading.runs[0]
        heading_run.font.bold = True
        heading_run.font.size = Pt(14)
        set_font_kai(heading_run)
        start_y = HEADING_H

    categories = _load_photos(photo_files, captions, fit_width, max_portrait_h, warnings, default_caption)
    sequence = plan_layout(categories, page_h, start_y, line_chars)

    # 連續同分類的列合成一張表格（跨頁時由 tblHeader 重複標題列）
    segments = []
    for category, row in sequence:
        if segments and segments[-1][0] == category:
            segments[-1][1].append(row)
        else:
            segments.append((category, [row]))

    before = set(body)
    first_table = True
    for category, rows in segments:
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        header_row = table.rows[0]
        for cell in header_row.cells:
            cell.width = Inches(col_w)
        _mark_row(header_row, 'w:cantSplit', 'w:tblHeader')
        header_cell = header_row.cells[0].merge(header_row.cells[1])
        header_cell.text = category
        set_cell_border(header_cell)
        for paragraph in header_cell.paragraphs:
            _tighten(paragraph, keep_with_next=True)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if first_table and break_on_first_table:
                paragraph.paragraph_format.page_break_before = True
            for run in paragraph.runs:
                set_font_kai(run)
                run.font.bold = True
        first_table = False

        for row in rows:
            photo_row = table.add_row()
            caption_row = table.add_row()
            _mark_row(photo_row, 'w:cantSplit')
            _mark_row(caption_row, 'w:cantSplit')
            for col_idx in range(2):
                for r in (photo_row, caption_row):
                    r.cells[col_idx].width = Inches(col_w)
                    set_cell_border(r.cells[col_idx])

            for col_idx, item in enumerate(row):
                paragraph = photo_row.cells[col_idx].paragraphs[0]
                _tighten(paragraph, keep_with_next=True)  # 照片與說明不分頁
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if item.source is not None:
                    paragraph.add_run().add_picture(item.source, width=Inches(item.img_w),
                                                    height=Inches(item.img_h))
                    logging.info(f"✓ SUCCESS: 插入照片 {os.path.basename(item.path)}")
                else:
                    set_font_kai(paragraph.add_run(f"[照片無法載入：{os.path.basename(item.path)}]"))

                caption_cell = caption_row.cells[col_idx]
                caption_cell.text = item.caption
                for p in caption_cell.paragraphs:
                    _tighten(p)
                    for r in p.runs:
                        set_font_kai(r)

        _add_spacer(doc)
        logging.info(f"✓ SUCCESS: 處理完成 {category}")

    if anchor is not None:
        # 新增的表格/間隔段落原本接在文件最後，搬到標題後面（舊照片表格原本的位置）
        created = [e for e in body if e not in before and e.tag != qn('w:sectPr')]
        for element in reversed(created):
            anchor.addnext(element)


def open_file(path):
    """用系統預設程式開啟檔案；失敗只記錄，不影響主程式。"""
    try:
        if hasattr(os, 'startfile'):
            os.startfile(path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])
    except OSError as e:
        logging.error(f"無法自動開啟 {path}: {e}")


def current_date():
    """今天的日期（獨立成函式，方便測試時替換）。"""
    return datetime.now()


def roc_date_parts(day=None):
    """回傳 (民國年, 月, 日) 的字串，例如 2026-10-14 -> ('115', '10', '14')。"""
    day = day or current_date()
    return str(day.year - 1911), str(day.month), str(day.day)


_INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')


def roc_date_range(data):
    """協勤日期轉成檔名用的字串，例如 1150925-28（同月）、1150930-1002（跨月）。"""
    y1, m1, d1 = int(data['start_year']), int(data['start_month']), int(data['start_day'])
    y2, m2, d2 = int(data['end_year']), int(data['end_month']), int(data['end_day'])
    start = f'{y1}{m1:02d}{d1:02d}'
    if (y1, m1, d1) == (y2, m2, d2):
        return start
    if (y1, m1) == (y2, m2):
        return f'{start}-{d2:02d}'
    if y1 == y2:
        return f'{start}-{m2:02d}{d2:02d}'
    return f'{start}-{y2}{m2:02d}{d2:02d}'


def default_output_name(data):
    """<協勤日期>_<協勤地點>協勤回報單（不含副檔名），日期放最前面方便依日期排序。"""
    location = _INVALID_FILENAME_CHARS.sub('', data.get('location', '')).strip()
    return f"{roc_date_range(data)}_{location}協勤回報單"


_PDF_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($env:REPORT_DOCX, $false, $true)
    $doc.ExportAsFixedFormat($env:REPORT_PDF, 17)
    $doc.Close($false)
} finally {
    # 只在沒有其他文件開著時才關閉 Word，避免影響使用者正在用的 Word
    if ($word -ne $null -and $word.Documents.Count -eq 0) { $word.Quit() }
}
"""


def export_pdf(docx_path, pdf_path, timeout=180):
    """用 Word 把 docx 轉成 PDF（需要 Windows 與已安裝的 Word）。失敗時丟出 RuntimeError。"""
    if sys.platform != 'win32':
        raise RuntimeError("目前只支援在 Windows 上用 Word 轉換 PDF")
    encoded = base64.b64encode(_PDF_SCRIPT.encode('utf-16-le')).decode('ascii')
    env = dict(os.environ, REPORT_DOCX=os.path.abspath(docx_path), REPORT_PDF=os.path.abspath(pdf_path))
    try:
        result = subprocess.run(
            ['powershell', '-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
            env=env, capture_output=True, timeout=timeout,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    except (OSError, subprocess.TimeoutExpired) as e:
        raise RuntimeError(f"無法執行 Word 轉換：{e}")
    if result.returncode != 0 or not os.path.exists(pdf_path):
        detail = (result.stderr or result.stdout or b'').decode('utf-8', errors='replace').strip()
        raise RuntimeError(f"Word 轉換 PDF 失敗（是否已安裝 Word？）{detail[:200]}")


def build_report(template_path, data, photo_files, captions, output_path):
    """產生報告並存檔，回傳警告訊息清單。"""
    doc = Document(template_path)
    if len(doc.tables) < 1:
        raise ValueError("模板格式錯誤：找不到表格")
    main_table = doc.tables[0]

    warnings = _fill_main_table(main_table, data)
    _add_photo_section(doc, main_table, photo_files, captions, warnings,
                       data.get('default_caption', DEFAULT_CAPTION))
    doc.save(output_path)
    return warnings


def move_item(seq, src, dst):
    """把 seq[src] 移到 dst 的位置（原地修改），其餘順序不變。"""
    seq.insert(dst, seq.pop(src))


# ---------------------------------------------------------------- GUI

class PhotoPreview(ttk.Frame):
    """照片預覽區：顯示縮圖（已依 EXIF 轉正）、檔名、尺寸與說明。"""

    PLACEHOLDER = "點選左邊的照片\n即可預覽"

    def __init__(self, master, font_family):
        super().__init__(master, padding=6, relief=tk.GROOVE)
        box = tk.Frame(self, width=PREVIEW_SIZE[0], height=PREVIEW_SIZE[1])
        box.pack_propagate(False)
        box.pack()
        self.picture = tk.Label(box, text=self.PLACEHOLDER, fg='gray', justify=tk.CENTER)
        self.picture.pack(fill=tk.BOTH, expand=True)
        self.info = ttk.Label(self, text='', wraplength=PREVIEW_SIZE[0], justify=tk.LEFT,
                              font=(font_family, 9 + GUI_FONT_DELTA))
        self.info.pack(anchor=tk.W, pady=(6, 0))
        self.image = None      # 要保留參照，否則 PhotoImage 會被回收而不顯示
        self.current = None    # (分類, 路徑)

    def clear(self, message=None):
        self.image = None
        self.current = None
        self.picture.config(image='', text=message or self.PLACEHOLDER)
        self.info.config(text='')

    def show(self, category, path, caption=None):
        name = os.path.basename(path)
        try:
            with Image.open(path) as img:
                width, height = img.size
                if img.getexif().get(0x0112, 1) in (5, 6, 7, 8):  # EXIF 旋轉後寬高互換
                    width, height = height, width
                img.draft('RGB', (PREVIEW_SIZE[0] * 2, PREVIEW_SIZE[1] * 2))  # JPEG 先縮小再解碼，大圖也很快
                thumb = ImageOps.exif_transpose(img)
                thumb.thumbnail(PREVIEW_SIZE, Image.LANCZOS)
                photo = ImageTk.PhotoImage(thumb.copy())
        except Exception as e:
            self.clear(f"無法預覽\n{e}")
            self.info.config(text=name)
            return
        self.image = photo
        self.current = (category, path)
        self.picture.config(image=photo, text='')
        orient = '橫式' if width >= height else '直式'
        self.info.config(text=f"{name}\n{width}×{height}（{orient}）\n說明：{caption or '（無）'}")


def load_private_fonts():
    """Windows：只在本程式執行期間載入 fonts/ 裡的字型（不安裝、不需要系統管理員權限）。回傳是否有載入成功。"""
    if sys.platform != 'win32':
        return False
    import ctypes
    loaded = False
    for name in FONT_FILES:
        path = os.path.join(FONT_DIR, name)
        if os.path.exists(path):
            # FR_PRIVATE = 0x10：字型只對本程序有效
            loaded |= ctypes.windll.gdi32.AddFontResourceExW(path, 0x10, 0) > 0
    return loaded


def pick_font_family(root):
    """有載入到指定字型就用它，否則退回系統字型。"""
    load_private_fonts()
    families = set(tkfont.families(root))
    if FONT_FAMILY in families:
        return FONT_FAMILY
    return FALLBACK_FONT_FAMILY if FALLBACK_FONT_FAMILY in families else 'Arial'


def scale_gui_fonts(root, delta, family=None, weight=None):
    """把 Tk 的預設字型（按鈕、輸入框、文字區、分頁等）加大 delta 點，並可一併換字型。"""
    for name in ('TkDefaultFont', 'TkTextFont', 'TkFixedFont', 'TkMenuFont', 'TkHeadingFont',
                 'TkCaptionFont', 'TkSmallCaptionFont', 'TkIconFont', 'TkTooltipFont'):
        try:
            font = tkfont.nametofont(name, root=root)
        except tk.TclError:
            continue
        size = font.cget('size')
        font.configure(size=size + delta if size > 0 else size - delta)  # 負值代表像素
        if family:
            font.configure(family=family)
        if weight:
            font.configure(weight=weight)

class ReportGeneratorGUI:
    def __init__(self, root, autosave_path=DEFAULT_AUTOSAVE_PATH, restore_on_start=True,
                 autosave_interval_ms=AUTOSAVE_INTERVAL_MS, template_path=TEMPLATE_PATH,
                 settings_path=DEFAULT_SETTINGS_PATH):
        self.root = root
        self.settings_path = settings_path
        self.font_family = pick_font_family(root)
        scale_gui_fonts(root, GUI_FONT_DELTA, self.font_family, GUI_FONT_WEIGHT)
        self.theme_name = self._load_theme_setting()
        self.palette = THEMES[self.theme_name]
        self._setup_style()
        self.autosave_path = autosave_path
        self.autosave_interval_ms = autosave_interval_ms
        self.default_caption = tk.StringVar(master=root, value=DEFAULT_CAPTION)
        self.root.title(APP_TITLE)
        self.root.geometry("1050x820")

        self.template_path = template_path
        self.photo_files = {c: [] for c in PHOTO_CATEGORIES_1 + PHOTO_CATEGORIES_2}
        self.photo_captions = {c: {} for c in self.photo_files}  # category -> {path: caption}
        self._last_dir = BASE_DIR  # 存檔對話框預設資料夾
        self.listboxes = {}
        self.previews = {}   # category -> PhotoPreview（同一個分頁的分類共用）
        self._drag = None    # 拖曳排序中：[分類, 目前索引]
        self.text_fields = {}

        self._ui_queue = queue.Queue()
        self.create_widgets()
        self._fill_report_date()
        self._drain_ui_queue()

        # 草稿狀態：與「上次儲存」或「空白」比較，判斷是否有未存內容
        self._saved_signature = self._signature()
        self._blank_signature = self._signature(with_report_date=False)
        self._last_autosave_signature = None
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(self.autosave_interval_ms, self._autosave_tick)
        if restore_on_start:
            self.root.after(300, self.offer_restore)
        if not os.path.exists(self.template_path):
            self.root.after(200, lambda: messagebox.showerror("錯誤", self._missing_template_message()))

    def _missing_template_message(self):
        return f"找不到模板檔案：\n{self.template_path}\n\n請確認它和程式放在同一個資料夾。"

    def _load_theme_setting(self):
        try:
            with open(self.settings_path, encoding='utf-8') as f:
                name = json.load(f).get('theme')
        except (OSError, ValueError, AttributeError):
            return DEFAULT_THEME
        return name if name in THEMES else DEFAULT_THEME

    def _save_theme_setting(self):
        try:
            with open(self.settings_path, 'w', encoding='utf-8') as f:
                json.dump({'theme': self.theme_name}, f, ensure_ascii=False)
        except OSError as e:
            logging.error(f"無法儲存外觀設定: {e}")

    def _setup_style(self):
        """依目前主題設定外觀：分頁是一個個分開的標籤（目前分頁較突出），按鈕採淺底深字。"""
        p = self.palette
        style = ttk.Style(self.root)
        style.theme_use('clam')
        self.root.configure(background=p['bg'])
        style.configure('.', background=p['bg'])
        style.configure('TNotebook', background=p['bg'], bordercolor=p['border'], lightcolor=p['border'],
                        darkcolor=p['border'], borderwidth=2, tabmargins=(4, 6, 4, 0), tabposition='nw')
        style.configure('TNotebook.Tab', padding=TAB_PADDING, background=p['tab_idle'], foreground=p['tab_idle_fg'],
                        bordercolor=p['border'], lightcolor=p['border'], darkcolor=p['border'], borderwidth=2)
        style.map('TNotebook.Tab',
                  background=[('selected', p['tab_active']), ('active', p['tab_hover'])],
                  foreground=[('selected', p['tab_active_fg'])],
                  padding=[('selected', TAB_PADDING_ACTIVE)],
                  expand=[('selected', (0, 0, 0, 0))])
        for name, key in (('Primary.TButton', 'primary'), ('Danger.TButton', 'danger')):
            fill, fg = p[key], p[f'{key}_fg']
            style.configure(name, background=fill, foreground=fg, bordercolor=p[f'{key}_border'],
                            lightcolor=fill, darkcolor=fill)
            style.map(name, background=[('disabled', '#e5e7eb'), ('active', p[f'{key}_active']),
                                        ('pressed', p[f'{key}_active'])],
                      foreground=[('disabled', '#9ca3af'), ('active', fg), ('pressed', fg)])

    def apply_theme(self, name):
        """切換外觀主題並記住選擇。"""
        if name not in THEMES:
            return
        self.theme_name = name
        self.palette = THEMES[name]
        self._setup_style()
        self._recolor(self.root)
        self._save_theme_setting()
        if hasattr(self, 'theme_var'):
            self.theme_var.set(name)
        self.status_label.config(text=f"已套用外觀：{name}")

    def _recolor(self, widget):
        """ttk 元件會自動跟著樣式變色；原生 tk 的容器（Canvas/Frame/Label）要手動換底色。"""
        for child in widget.winfo_children():
            if child.winfo_class() in ('Canvas', 'Frame', 'Label'):
                child.configure(background=self.palette['bg'])
            self._recolor(child)

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        row = 0
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        main_frame.rowconfigure(row, weight=1)

        self.create_basic_info_tab(notebook)
        self.create_text_tab(notebook, "道路步道", [
            ("一、道路、步道狀況及維護情形（如有坍方、雜草等地點及維護處理等）", None),
            ("1、道路、步道狀況及維護情形：", "road_status"),
            ("2、道路、步道雜草情形：", "road_grass"),
            ("3、其他：", "road_other"),
        ], height=5)
        self.create_text_tab(notebook, "設施", [
            ("二、設施狀況及維護情形（如山屋、牌誌、木棧道等設施損壞紀錄地點、種類數量及維護處理情形等）", None),
            ("1、山屋：", "facility_house"),
            ("2、牌示：", "facility_sign"),
            ("3、木棧道：", "facility_bridge"),
            ("4、其他：", "facility_other"),
        ], height=5)
        self.create_text_tab(notebook, "自然資源與遊客", [
            ("三、自然資源狀況（如特殊植物、動物、人文、地質地形、水文等資源之地點、數量及周圍環境紀錄等）：", "nature"),
            ("四、環境清理及遊客宣導情形（如垃圾清理及帶下山、遊客登山安全、環境保護宣導等情形）：", "environment"),
            ("五、遊客違法情形（如非法入園、亂丟垃圾、走捷徑等登山違法情形紀錄登山團體與人員）：", "violation"),
        ], height=6)
        self.create_text_tab(notebook, "意見與心得建議", [
            ("六、登山遊客反應意見：", "feedback"),
            ("七、協勤心得與建議：", "reflection"),
            ("八、相關疑問與解答：（將服勤期間相關疑問，以及透過職員或其他志工尋得之答案記錄下來）", "questions"),
        ], height=6)
        self.create_photo_tab(notebook, "服勤照片(上)", "九、服勤照片", PHOTO_CATEGORIES_1)
        self.create_photo_tab(notebook, "服勤照片(下)", "九、服勤照片（續）", PHOTO_CATEGORIES_2)

        row += 1
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, columnspan=3, pady=10)

        ttk.Button(button_frame, text="載入草稿", command=self.load_draft).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="儲存草稿", command=self.save_draft).pack(side=tk.LEFT, padx=5)
        self.generate_btn = ttk.Button(button_frame, text="產生報告", command=self.generate_report,
                                       style='Primary.TButton')
        self.generate_btn.pack(side=tk.LEFT, padx=5)
        self.clear_btn = ttk.Button(button_frame, text="清除", command=self.clear_all, style='Danger.TButton')
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="離開", command=self.on_close).pack(side=tk.LEFT, padx=5)

        row += 1
        self.status_label = ttk.Label(main_frame, text="準備就緒", relief=tk.SUNKEN)
        self.status_label.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E))
        theme_frame = ttk.Frame(main_frame)
        theme_frame.grid(row=row, column=2, sticky=tk.E, padx=(10, 0))
        ttk.Label(theme_frame, text="外觀：").pack(side=tk.LEFT)
        self.theme_var = tk.StringVar(master=self.root, value=self.theme_name)
        theme_box = ttk.Combobox(theme_frame, textvariable=self.theme_var, values=list(THEMES),
                                 state='readonly', width=9)
        theme_box.pack(side=tk.LEFT)
        theme_box.bind('<<ComboboxSelected>>', lambda e: self.apply_theme(self.theme_var.get()))
        self.theme_box = theme_box

    # ---- 共用：可捲動分頁
    def make_scrollable_tab(self, notebook, title):
        frame = ttk.Frame(notebook, padding="10")
        notebook.add(frame, text=title)

        canvas = tk.Canvas(frame, highlightthickness=0, background=self.palette['bg'])
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)

        window = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        def on_wheel(event):
            # 文字框/清單自己有捲動，不要連帶捲動分頁
            if isinstance(event.widget, (tk.Text, tk.Listbox)):
                return
            # 只處理滑鼠位於本分頁內的事件（用 bind_all，避免 Enter/Leave 在子元件間誤觸發）
            w = event.widget
            while w is not None and w is not frame:
                w = getattr(w, 'master', None)
            if w is frame and frame.winfo_ismapped():
                canvas.yview_scroll(int(-event.delta / 120), "units")

        canvas.bind_all("<MouseWheel>", on_wheel, add="+")

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        inner.columnconfigure(0, weight=1)
        return inner

    def create_basic_info_tab(self, notebook):
        frame = ttk.Frame(notebook, padding="10")
        notebook.add(frame, text="基本資料")
        frame.columnconfigure(1, weight=1)

        def date_entries(parent, prefix):
            """年/月/日輸入框；每個單位字前後的間距相同（左右各 4 像素），起迄兩組看起來一樣寬。"""
            for key, unit, width in (('year', '年', 8), ('month', '月', 5), ('day', '日', 5)):
                entry = ttk.Entry(parent, width=width, justify=tk.CENTER)
                entry.pack(side=tk.LEFT)
                ttk.Label(parent, text=unit).pack(side=tk.LEFT, padx=DATE_UNIT_PADX)
                setattr(self, f'{prefix}{key}', entry)

        row = 0
        ttk.Label(frame, text="1. 填表日期:", font=(self.font_family, 9 + GUI_FONT_DELTA, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        date_frame = ttk.Frame(frame)
        date_frame.grid(row=row, column=1, sticky=tk.W, padx=5)
        date_entries(date_frame, '')

        row += 1
        ttk.Label(frame, text="2. 協勤人員姓名:", font=(self.font_family, 9 + GUI_FONT_DELTA, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        self.personnel = ttk.Entry(frame, width=50)
        self.personnel.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5)

        row += 1
        ttk.Label(frame, text="3. 協勤日期:", font=(self.font_family, 9 + GUI_FONT_DELTA, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        date_range_frame = ttk.Frame(frame)
        date_range_frame.grid(row=row, column=1, sticky=tk.W, padx=5)
        date_entries(date_range_frame, 'start_')
        ttk.Label(date_range_frame, text="至").pack(side=tk.LEFT, padx=DATE_RANGE_PADX)
        date_entries(date_range_frame, 'end_')

        row += 1
        ttk.Label(frame, text="4. 協勤地點:", font=(self.font_family, 9 + GUI_FONT_DELTA, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        self.location = ttk.Entry(frame, width=50)
        self.location.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5)

    def create_text_tab(self, notebook, title, fields, height):
        inner = self.make_scrollable_tab(notebook, title)
        row_idx = 0
        for label, key in fields:
            if key is None:
                ttk.Label(inner, text=label, font=(self.font_family, 10 + GUI_FONT_DELTA, 'bold'), wraplength=880).grid(
                    row=row_idx, column=0, sticky=tk.W, pady=(15, 5))
                row_idx += 1
                continue
            header = ttk.Frame(inner)
            header.grid(row=row_idx, column=0, sticky=(tk.W, tk.E), pady=(10, 2))
            header.columnconfigure(0, weight=1)
            ttk.Label(header, text=label, font=(self.font_family, 9 + GUI_FONT_DELTA, 'bold'), wraplength=640).grid(
                row=0, column=0, sticky=tk.W)
            text_widget = scrolledtext.ScrolledText(inner, height=height, width=70, undo=True)
            self._make_list_menu(header, text_widget).grid(row=0, column=1, padx=2)
            text_widget.grid(row=row_idx + 1, column=0, sticky=(tk.W, tk.E), padx=5)
            text_widget.bind('<Return>', lambda e, w=text_widget: self._on_return(w))
            self.text_fields[key] = text_widget
            row_idx += 2

    def _make_list_menu(self, parent, widget):
        """「項目符號/編號 ▾」下拉選單：選一種就套用到選取的行（再選同一種則取消）。"""
        button = ttk.Menubutton(parent, text="項目符號/編號", width=14)
        menu = tk.Menu(button, tearoff=0)
        for symbol, name in BULLET_CHOICES:
            menu.add_command(label=f"{symbol}  {name}",
                             command=lambda w=widget, sym=symbol: self.apply_list(
                                 w, lambda lines: toggle_bullets(lines, sym)))
        menu.add_separator()
        menu.add_command(label="1. 2. 3.  編號", command=lambda w=widget: self.apply_list(w, toggle_numbers))
        menu.add_separator()
        menu.add_command(label="取消項目符號/編號", command=lambda w=widget: self.apply_list(w, remove_list_markers))
        button['menu'] = menu
        button.menu = menu
        return button

    def apply_list(self, widget, toggler):
        """對選取的行（沒有選取就是游標所在行）加上/取消項目符號或編號。"""
        try:
            first = int(widget.index(tk.SEL_FIRST).split('.')[0])
            last_index = widget.index(tk.SEL_LAST)
            last = int(last_index.split('.')[0])
            if last_index.endswith('.0') and last > first:
                last -= 1  # 選取範圍結尾剛好在下一行開頭，不算那一行
        except tk.TclError:
            first = last = int(widget.index(tk.INSERT).split('.')[0])
        old = widget.get(f'{first}.0', f'{last}.end')
        new = '\n'.join(toggler(old.split('\n')))
        widget.delete(f'{first}.0', f'{last}.end')
        widget.insert(f'{first}.0', new)
        end_line = first + new.count('\n')
        widget.tag_remove(tk.SEL, '1.0', tk.END)
        widget.tag_add(tk.SEL, f'{first}.0', f'{end_line}.end')
        widget.mark_set(tk.INSERT, f'{end_line}.end')
        widget.focus_set()

    def _on_return(self, widget):
        """在項目符號/編號的行按 Enter：自動接上下一個符號；空項目再按 Enter 則結束清單。"""
        line_start = widget.index('insert linestart')
        action, prefix = continue_list(widget.get(line_start, 'insert'))
        if action is None:
            return None
        if action == 'end':
            widget.delete(line_start, 'insert')
        else:
            widget.insert('insert', '\n' + prefix)
        return 'break'

    def create_photo_tab(self, notebook, title, heading, categories):
        inner = self.make_scrollable_tab(notebook, title)
        ttk.Label(inner, text=heading, font=(self.font_family, 11 + GUI_FONT_DELTA, 'bold')).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 10))
        ttk.Label(inner, text="（點選預覽、拖曳調整順序）", foreground="gray").grid(
            row=0, column=0, sticky=tk.E)

        caption_row = ttk.Frame(inner)
        caption_row.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))
        ttk.Label(caption_row, text="未填說明時的預設字樣：").pack(side=tk.LEFT)
        ttk.Entry(caption_row, textvariable=self.default_caption, width=24).pack(side=tk.LEFT, padx=5)
        ttk.Label(caption_row, text="（雙擊檔名輸入照片說明）", foreground="gray").pack(side=tk.LEFT)

        preview = PhotoPreview(inner, self.font_family)
        preview.grid(row=2, column=1, rowspan=len(categories) * 2, sticky=tk.N, padx=(10, 0), pady=(10, 0))

        for idx, category in enumerate(categories):
            row = idx * 2 + 2
            ttk.Label(inner, text=category, font=(self.font_family, 9 + GUI_FONT_DELTA, 'bold')).grid(
                row=row, column=0, sticky=tk.W, pady=(10, 5))

            list_frame = ttk.Frame(inner)
            list_frame.grid(row=row + 1, column=0, sticky=(tk.W, tk.E), padx=20)
            list_frame.columnconfigure(0, weight=1)

            listbox = tk.Listbox(list_frame, height=5, selectmode=tk.EXTENDED, exportselection=False)
            listbox.grid(row=0, column=0, sticky=(tk.W, tk.E))
            scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
            scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
            listbox.configure(yscrollcommand=scrollbar.set)
            listbox.bind("<Double-Button-1>", lambda e, c=category, lb=listbox: self.edit_caption(c, lb))
            listbox.bind("<<ListboxSelect>>", lambda e, c=category, lb=listbox: self._on_photo_select(c, lb))
            listbox.bind("<Button-1>", lambda e, c=category, lb=listbox: self._drag_start(c, lb, e))
            listbox.bind("<B1-Motion>", lambda e, c=category, lb=listbox: self._drag_motion(c, lb, e))
            listbox.bind("<ButtonRelease-1>", lambda e: self._drag_end())

            btn_frame = ttk.Frame(list_frame)
            btn_frame.grid(row=1, column=0, columnspan=2, pady=5)
            ttk.Button(btn_frame, text="選擇照片",
                       command=lambda c=category, lb=listbox: self.select_photos(c, lb)).pack(side=tk.LEFT, padx=2)
            ttk.Button(btn_frame, text="移除選取",
                       command=lambda c=category, lb=listbox: self.remove_photos(c, lb)).pack(side=tk.LEFT, padx=2)
            ttk.Button(btn_frame, text="清除全部",
                       command=lambda c=category, lb=listbox: self.clear_photos(c, lb)).pack(side=tk.LEFT, padx=2)

            self.listboxes[category] = listbox
            self.previews[category] = preview

    # ---- 事件處理
    def _photo_label(self, category, filepath):
        caption = self.photo_captions[category].get(filepath)
        name = os.path.basename(filepath)
        return f"{name} — {caption}" if caption else name

    # ---- 預覽與拖曳排序
    def _on_photo_select(self, category, listbox):
        selected = listbox.curselection()
        preview = self.previews[category]
        if not selected:
            return
        path = self.photo_files[category][selected[0]]
        preview.show(category, path, self.photo_captions[category].get(path))

    def _refresh_previews(self):
        """照片被移除/清除後，若預覽的照片已不在清單中就清掉預覽。"""
        for preview in set(self.previews.values()):
            if preview.current and preview.current[1] not in self.photo_files[preview.current[0]]:
                preview.clear()

    def _move_photo(self, category, listbox, src, dst):
        move_item(self.photo_files[category], src, dst)
        label = listbox.get(src)
        listbox.delete(src)
        listbox.insert(dst, label)
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(dst)
        listbox.activate(dst)

    def _drag_start(self, category, listbox, event):
        """按下滑鼠：記錄要拖曳的項目（按在清單空白處不算）。不攔截事件，讓一般的選取照常運作。"""
        self._drag = None
        if listbox.size() == 0:
            return None
        index = listbox.nearest(event.y)
        box = listbox.bbox(index)
        if box and box[1] <= event.y < box[1] + box[3]:
            self._drag = [category, index]
        return None

    def _drag_motion(self, category, listbox, event):
        if not self._drag or self._drag[0] != category:
            return None
        target = max(0, min(listbox.nearest(event.y), listbox.size() - 1))
        if target != self._drag[1]:
            self._move_photo(category, listbox, self._drag[1], target)
            self._drag[1] = target
            listbox.see(target)
        return 'break'  # 拖曳時不要讓清單把它當成「擴大選取」

    def _drag_end(self):
        self._drag = None

    def select_photos(self, category, listbox):
        filepaths = filedialog.askopenfilenames(
            title=f"選擇照片 - {category}",
            filetypes=[
                ("圖片檔案", "*.jpg *.jpeg *.png *.gif *.bmp"),
                ("所有檔案", "*.*")
            ]
        )
        if not filepaths:
            return
        added = 0
        for filepath in filepaths:
            if filepath not in self.photo_files[category]:
                self.photo_files[category].append(filepath)
                listbox.insert(tk.END, self._photo_label(category, filepath))
                added += 1
        self.status_label.config(text=f"已新增 {added} 張照片到 {category}")
        logging.info(f"新增 {added} 張照片到 {category}")

    def edit_caption(self, category, listbox):
        selected = listbox.curselection()
        if not selected:
            return
        idx = selected[0]
        filepath = self.photo_files[category][idx]
        caption = simpledialog.askstring(
            "照片說明", f"請輸入「{os.path.basename(filepath)}」的說明：",
            initialvalue=self.photo_captions[category].get(filepath, ""), parent=self.root)
        if caption is None:
            return
        caption = caption.strip()
        if caption:
            self.photo_captions[category][filepath] = caption
        else:
            self.photo_captions[category].pop(filepath, None)
        listbox.delete(idx)
        listbox.insert(idx, self._photo_label(category, filepath))
        listbox.selection_set(idx)
        self._on_photo_select(category, listbox)

    def remove_photos(self, category, listbox):
        selected = listbox.curselection()
        if not selected:
            return
        for idx in reversed(selected):
            filepath = self.photo_files[category].pop(idx)
            self.photo_captions[category].pop(filepath, None)
            listbox.delete(idx)
        self._refresh_previews()
        self.status_label.config(text=f"已移除 {len(selected)} 張照片")

    def clear_photos(self, category, listbox):
        self.photo_files[category].clear()
        self.photo_captions[category].clear()
        listbox.delete(0, tk.END)
        self._refresh_previews()
        self.status_label.config(text=f"已清除 {category} 的所有照片")

    def _clear_widgets(self):
        for name in ENTRY_FIELDS:
            getattr(self, name).delete(0, tk.END)
        for text_widget in self.text_fields.values():
            text_widget.delete('1.0', tk.END)
        for category, listbox in self.listboxes.items():
            self.photo_files[category].clear()
            self.photo_captions[category].clear()
            listbox.delete(0, tk.END)
        for preview in set(self.previews.values()):
            preview.clear()
        self.default_caption.set(DEFAULT_CAPTION)

    def clear_all(self):
        if self.is_dirty():
            message = ("⚠ 目前有尚未儲存的內容！\n\n"
                       "清除會移除所有欄位、照片與說明，無法復原。\n建議先按「儲存草稿」。\n\n確定要清除嗎？")
        else:
            message = ("清除會移除所有欄位、照片與說明（填表日期會重設為今天），無法復原。\n\n確定要清除嗎？")
        # 預設選「否」，不小心按到 Enter 也不會清掉資料
        if not messagebox.askyesno("清除全部內容", message, icon='warning', default='no'):
            return
        self._clear_widgets()
        self._fill_report_date()
        self.status_label.config(text="已清除所有資料")

    # ---- 草稿存檔
    def build_draft(self):
        """把目前畫面上的內容（原始輸入，不含 bullet 處理）整理成可存成 JSON 的 dict。"""
        return {
            'version': DRAFT_VERSION,
            'saved_at': datetime.now().isoformat(timespec='seconds'),
            'fields': {name: getattr(self, name).get() for name in ENTRY_FIELDS},
            'texts': {key: w.get('1.0', tk.END).rstrip('\n') for key, w in self.text_fields.items()},
            'photos': {c: list(p) for c, p in self.photo_files.items()},
            'captions': {c: dict(v) for c, v in self.photo_captions.items()},
            'default_caption': self.default_caption.get(),
        }

    def apply_draft(self, draft):
        """把草稿內容還原到畫面，回傳警告訊息清單（例如找不到的照片）。"""
        if not isinstance(draft, dict) or draft.get('version') != DRAFT_VERSION:
            raise ValueError("不是有效的草稿檔，或草稿版本不相容")
        warnings = []
        self._clear_widgets()

        for name in ENTRY_FIELDS:
            getattr(self, name).insert(0, str(draft.get('fields', {}).get(name, '')))
        default_caption = str(draft.get('default_caption', DEFAULT_CAPTION))
        self.default_caption.set(DEFAULT_CAPTION if default_caption == LEGACY_DEFAULT_CAPTION else default_caption)
        self._fill_report_date()  # 草稿裡有日期就沿用（第一次填表的日期）；沒有才補上今天
        for key, widget in self.text_fields.items():
            widget.insert('1.0', str(draft.get('texts', {}).get(key, '')))

        for category, paths in draft.get('photos', {}).items():
            if category not in self.photo_files:
                continue
            captions = draft.get('captions', {}).get(category, {})
            for path in paths:
                if not os.path.exists(path):
                    warnings.append(f"找不到照片，已略過：{path}")
                    continue
                self.photo_files[category].append(path)
                if captions.get(path):
                    self.photo_captions[category][path] = captions[path]
                self.listboxes[category].insert(tk.END, self._photo_label(category, path))
        return warnings

    def save_draft(self):
        path = filedialog.asksaveasfilename(
            title="儲存草稿",
            defaultextension=".json",
            initialdir=self._last_dir,
            initialfile=f"巡查回報單草稿_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            filetypes=[("草稿檔", "*.json")]
        )
        if not path:
            return False
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self.build_draft(), f, ensure_ascii=False, indent=2)
        except OSError as e:
            messagebox.showerror("錯誤", f"草稿儲存失敗：{e}")
            return False
        self._saved_signature = self._signature()
        self._remove_autosave()
        logging.info(f"儲存草稿: {path}")
        self.status_label.config(text=f"已儲存草稿: {os.path.basename(path)}")
        return True

    def load_draft(self):
        path = filedialog.askopenfilename(title="載入草稿", filetypes=[("草稿檔", "*.json")])
        if not path:
            return
        try:
            with open(path, encoding='utf-8') as f:
                draft = json.load(f)
            if not isinstance(draft, dict) or draft.get('version') != DRAFT_VERSION:
                raise ValueError("不是有效的草稿檔，或草稿版本不相容")
        except (OSError, ValueError) as e:  # JSONDecodeError 是 ValueError 的子類別
            messagebox.showerror("錯誤", f"無法載入草稿：{e}")
            return
        if not messagebox.askyesno("確認", "載入草稿會覆蓋目前畫面上的內容，確定嗎？"):
            return
        warnings = self.apply_draft(draft)
        self._saved_signature = self._signature()
        logging.info(f"載入草稿: {path}")
        self.status_label.config(text=f"已載入草稿: {os.path.basename(path)}")
        if warnings:
            messagebox.showinfo("草稿已載入", "⚠ 注意：\n" + "\n".join(f"- {w}" for w in warnings))

    # ---- 未存內容偵測、自動存檔、關閉提示
    def _fill_report_date(self):
        """「填表日期」自動帶入第一次填表的日期（今天，民國格式）。已經有日期（例如載入草稿）就不動。"""
        if any(getattr(self, name).get().strip() for name in REPORT_DATE_FIELDS):
            return
        for name, value in zip(REPORT_DATE_FIELDS, roc_date_parts()):
            getattr(self, name).insert(0, value)

    def _signature(self, with_report_date=True):
        """目前內容的指紋（不含存檔時間），用來判斷內容是否改變。

        with_report_date=False 時不含自動帶入的「填表日期」，用來判斷畫面是不是空白。
        """
        draft = self.build_draft()
        draft.pop('saved_at')
        if not with_report_date:
            for name in REPORT_DATE_FIELDS:
                draft['fields'].pop(name, None)
        return json.dumps(draft, ensure_ascii=False, sort_keys=True)

    def is_dirty(self):
        return self._signature() != self._saved_signature

    def autosave_now(self):
        """有未存且非空白的內容時，寫入自動存檔。回傳是否有寫入。"""
        signature = self._signature()
        if (signature == self._saved_signature or self._signature(with_report_date=False) == self._blank_signature
                or signature == self._last_autosave_signature):
            return False
        tmp_path = self.autosave_path + '.tmp'
        try:
            with open(tmp_path, 'w', encoding='utf-8') as f:
                json.dump(self.build_draft(), f, ensure_ascii=False, indent=2)
            os.replace(tmp_path, self.autosave_path)  # 原子替換，避免寫到一半當機留下壞檔
        except OSError as e:
            logging.error(f"自動存檔失敗: {e}")
            return False
        self._last_autosave_signature = signature
        self.status_label.config(text=f"已自動存檔 {datetime.now().strftime('%H:%M:%S')}")
        return True

    def _autosave_tick(self):
        try:
            self.autosave_now()
        finally:
            self.root.after(self.autosave_interval_ms, self._autosave_tick)

    def _remove_autosave(self):
        self._last_autosave_signature = None
        try:
            os.remove(self.autosave_path)
        except FileNotFoundError:
            pass
        except OSError as e:
            logging.error(f"無法刪除自動存檔: {e}")

    def offer_restore(self):
        """啟動時若有上次異常結束留下的自動存檔，詢問是否還原。"""
        if not os.path.exists(self.autosave_path):
            return
        try:
            with open(self.autosave_path, encoding='utf-8') as f:
                draft = json.load(f)
            if not isinstance(draft, dict) or draft.get('version') != DRAFT_VERSION:
                raise ValueError("版本不相容")
            saved_at = str(draft.get('saved_at', '未知時間')).replace('T', ' ')
        except (OSError, ValueError) as e:
            logging.error(f"自動存檔無法讀取，已刪除: {e}")
            self._remove_autosave()
            return
        if messagebox.askyesno("還原上次的內容", f"發現上次未儲存的自動存檔（{saved_at}），要還原嗎？\n\n選「否」會刪除這份自動存檔。"):
            warnings = self.apply_draft(draft)
            self.status_label.config(text=f"已還原自動存檔（{saved_at}）")
            logging.info(f"還原自動存檔: {saved_at}")
            if warnings:
                messagebox.showinfo("自動存檔已還原", "⚠ 注意：\n" + "\n".join(f"- {w}" for w in warnings))
        else:
            self._remove_autosave()

    def on_close(self):
        if self.is_dirty():
            answer = messagebox.askyesnocancel(
                "尚未儲存", "有尚未儲存的內容，要先儲存草稿再離開嗎？\n\n"
                           "是：儲存草稿後離開\n否：不儲存直接離開（下次開啟時可從自動存檔還原）\n取消：回到程式")
            if answer is None:
                return
            if answer is False:
                # 不存草稿離開：保留自動存檔（先補存最新內容；畫面已是空白就不留舊內容）
                if self._signature(with_report_date=False) == self._blank_signature:
                    self._remove_autosave()
                else:
                    self.autosave_now()
                self.root.destroy()
                return
            if not self.save_draft():
                return  # 取消存檔或存檔失敗，留在程式裡
        self._remove_autosave()
        self.root.destroy()

    def collect_data(self):
        """在主執行緒讀取所有欄位。驗證失敗時回傳 (None, 錯誤訊息)。"""
        data = {
            'year': self.year.get().strip(),
            'month': self.month.get().strip(),
            'day': self.day.get().strip(),
            'personnel': self.personnel.get().strip(),
            'start_year': self.start_year.get().strip(),
            'start_month': self.start_month.get().strip(),
            'start_day': self.start_day.get().strip(),
            'end_year': self.end_year.get().strip(),
            'end_month': self.end_month.get().strip(),
            'end_day': self.end_day.get().strip(),
            'location': self.location.get().strip(),
        }

        dates = {}
        for label, prefix in [("填表日期", ''), ("協勤日期（起）", 'start_'), ("協勤日期（迄）", 'end_')]:
            try:
                dates[prefix] = datetime(int(data[prefix + 'year']), int(data[prefix + 'month']),
                                         int(data[prefix + 'day']))
            except ValueError:
                return None, f"「{label}」不是有效的日期，請檢查"
        if dates['start_'] > dates['end_']:
            return None, "協勤日期的起日不可晚於迄日"

        for key, widget in self.text_fields.items():
            data[key] = widget.get('1.0', tk.END).strip()
        data['default_caption'] = self.default_caption.get().strip()
        return data, None

    def generate_report(self):
        if not os.path.exists(self.template_path):
            messagebox.showerror("錯誤", self._missing_template_message())
            return

        data, error = self.collect_data()
        if error:
            messagebox.showerror("錯誤", error)
            return

        photo_files = {c: list(p) for c, p in self.photo_files.items()}
        captions = {c: dict(v) for c, v in self.photo_captions.items()}
        missing = [p for photos in photo_files.values() for p in photos if not os.path.exists(p)]
        if missing:
            messagebox.showerror("錯誤", "以下照片檔案不存在：\n" + "\n".join(missing[:10]))
            return

        output_path = filedialog.asksaveasfilename(
            title="儲存報告（同時會產生同名的 PDF）",
            defaultextension=".docx",
            initialdir=self._last_dir,
            initialfile=default_output_name(data) + ".docx",
            filetypes=[("Word文件", "*.docx")]
        )
        if not output_path:
            return
        if os.path.exists(output_path) and os.path.samefile(output_path, self.template_path):
            messagebox.showerror("錯誤", "輸出檔案不能和模板是同一個檔案，否則會覆蓋模板。\n請換一個檔名。")
            return
        self._last_dir = os.path.dirname(output_path)

        self.generate_btn.config(state=tk.DISABLED)
        self.status_label.config(text="正在產生報告與 PDF...")
        threading.Thread(target=self._generate_report_worker,
                         args=(data, photo_files, captions, output_path),
                         daemon=True).start()

    def _post(self, func):
        """背景執行緒把 GUI 更新交給主執行緒執行（tkinter 不是 thread-safe）。"""
        self._ui_queue.put(func)

    def _drain_ui_queue(self):
        try:
            while True:
                self._ui_queue.get_nowait()()
        except queue.Empty:
            pass
        self.root.after(100, self._drain_ui_queue)

    def _generate_report_worker(self, data, photo_files, captions, output_path):
        try:
            warnings = build_report(self.template_path, data, photo_files, captions, output_path)
            logging.info(f"✓ SUCCESS: 文檔已產生 {output_path}")

            pdf_path = os.path.splitext(output_path)[0] + '.pdf'
            try:
                export_pdf(output_path, pdf_path)
                logging.info(f"✓ SUCCESS: PDF 已產生 {pdf_path}")
            except RuntimeError as e:
                logging.error(f"PDF 轉換失敗: {e}")
                warnings.append(f"PDF 未產生：{e}")
                pdf_path = None

            # 產生後自動開啟檔案；本工具維持開啟，方便修改後再次產生
            self._post(lambda: open_file(output_path))
            files = f"Word: {output_path}" + (f"\nPDF: {pdf_path}" if pdf_path else "")
            if warnings:
                message = f"報告已產生！\n\n{files}\n\n⚠ 注意：\n" + "\n".join(f"- {w}" for w in warnings)
                logging.warning("警告: " + "; ".join(warnings))
                self._post(lambda: messagebox.showinfo("成功（有警告）", message))
            self._post(lambda: self.status_label.config(text=f"✓ SUCCESS: 已產生 {output_path}"))
        except Exception as e:
            error_msg = f"✗ FAIL: {e}"
            logging.error(error_msg, exc_info=True)
            self._post(lambda: messagebox.showerror("錯誤", error_msg))
            self._post(lambda: self.status_label.config(text=error_msg))
        finally:
            self._post(lambda: self.generate_btn.config(state=tk.NORMAL))


def main():
    root = tk.Tk()
    ReportGeneratorGUI(root)

    logging.info("=" * 60)
    logging.info(f"{APP_TITLE} 啟動")
    logging.info(f"執行命令: {sys.executable} {' '.join(sys.argv)}")
    logging.info("=" * 60)

    root.mainloop()


if __name__ == '__main__':
    main()
