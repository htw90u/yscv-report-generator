#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
保育巡查回報單生成器 - GUI版本
提供圖形化介面供使用者輸入資料並選擇照片
"""

import io
import json
import logging
import math
import os
import queue
import sys
import threading
import tkinter as tk
from dataclasses import dataclass
from datetime import datetime
from tkinter import ttk, filedialog, messagebox, scrolledtext, simpledialog

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PIL import Image, ImageOps

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'report_generator_gui.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

DRAFT_VERSION = 1
DEFAULT_AUTOSAVE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'report_generator_autosave.json')
AUTOSAVE_INTERVAL_MS = 60_000
# 基本資料欄位（同時是 ReportGeneratorGUI 上的屬性名稱）
ENTRY_FIELDS = ['year', 'month', 'day', 'personnel',
                'start_year', 'start_month', 'start_day',
                'end_year', 'end_month', 'end_day', 'location']

MAX_IMAGE_WIDTH_PX = 1200  # 照片在文件中寬 2.8 吋，1200px 約 430dpi，已足夠清晰
JPEG_QUALITY = 85

# 照片版面（單位：吋）。橫式/直式照片都以填滿儲存格寬度為原則，並模擬分頁以盡量排滿整頁
CELL_SIDE_PAD = 0.2     # 儲存格左右內距與餘裕
HEADER_H = 0.32         # 分類標題列
SPACER_H = 0.15         # 分類表格之間的間隔
HEADING_H = 0.45        # 「九、服勤照片」標題
PHOTO_ROW_PAD = 0.04
CAPTION_LINE_H = 0.2    # 說明文字每行高度（11pt）
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

def add_bullet_points(text):
    """為每一行加上bullet前綴並縮排（使用者已輸入 bullet 的行不重複加）"""
    if not text:
        return ""
    result = []
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            result.append('')
            continue
        if line[0] in '•·':
            line = line[1:].strip()
        result.append(f'    • {line}')
    return '\n'.join(result)


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


def set_font_kai(run):
    """設定字型為標楷體11pt"""
    run.font.name = '標楷體'
    run.font.size = Pt(11)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')


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


def _fill_main_table(main_table, data):
    """把文字資料填進主表格，回傳警告清單。"""
    warnings = []

    try:
        main_table.rows[0].cells[2].text = f"{data['year']}年{data['month']}月{data['day']}日"
        main_table.rows[1].cells[1].text = data['personnel']
        main_table.rows[2].cells[1].text = (
            f"{data['start_year']}年{data['start_month']}月{data['start_day']}日至"
            f"{data['end_year']}年{data['end_month']}月{data['end_day']}日")
        main_table.rows[3].cells[1].text = data['location']
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

            text = new_text = cell.text
            for marker, key in FILL_LABELS:
                if marker in new_text and key not in filled:
                    content = data.get(key, '')
                    if content:
                        new_text = new_text.replace(marker, f'{marker}\n{content}', 1)
                    filled.add(key)

            if QUESTION_MARKER in new_text and 'questions' not in filled:
                content = data.get('questions', '')
                if QUESTION_PLACEHOLDER in new_text:
                    new_text = new_text.replace(QUESTION_PLACEHOLDER, content, 1)
                elif content:
                    new_text = new_text + '\n' + content
                filled.add('questions')

            if new_text != text:
                cell.text = new_text

    for key, name in FIELD_NAMES.items():
        if data.get(key) and key not in filled:
            warnings.append(f"模板中找不到「{name}」欄位，這部分內容未填入")

    for row in main_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    set_font_kai(run)
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
        self.rows = []

    def fits(self, row, line_chars, page_h):
        fixed = self.fixed + PHOTO_ROW_PAD + max(caption_height(p.caption, line_chars) for p in row)
        img = self.img + max(p.img_h for p in row)
        return fixed + MIN_PHOTO_SCALE * img <= page_h

    def add(self, row, line_chars):
        self.fixed += PHOTO_ROW_PAD + max(caption_height(p.caption, line_chars) for p in row)
        self.img += max(p.img_h for p in row)
        self.rows.append(row)

    def finish(self, page_h):
        """這頁照片放不下時，等比例縮小（不低於 MIN_PHOTO_SCALE），讓版面剛好排滿。"""
        if self.img <= 0:
            return
        scale = min(1.0, (page_h - self.fixed) / self.img)
        if scale < 1.0:
            for row in self.rows:
                for p in row:
                    p.img_w *= scale
                    p.img_h *= scale


def plan_layout(categories, page_h, start_y, line_chars):
    """安排每個分類的照片列，回傳 {分類: [[照片, 照片], ...]}。

    - 橫式與直式分開，同方向的照片排在一起（橫式在前），每列 2 張。
    - 模擬分頁：目前這一頁放不下下一列時，往後找放得下的列補進來；
      照片等比例縮小一點（最多縮到 MIN_PHOTO_SCALE）就能多放一列時，也會這樣做，讓整頁盡量排滿。
      都放不下才換頁。
    - page_h 為內文可用高度，start_y 為照片區開始時已用掉的高度（吋）。
    """
    page = _Page(fixed=start_y)
    plan = {}
    for name, items in categories.items():
        if not items:
            continue
        land = [p for p in items if p.landscape]
        port = [p for p in items if not p.landscape]
        rows = ([land[i:i + 2] for i in range(0, len(land), 2)]
                + [port[i:i + 2] for i in range(0, len(port), 2)])

        probe = _Page(fixed=page.fixed + HEADER_H)
        probe.img = page.img
        if not probe.fits(rows[0], line_chars, page_h):
            page.finish(page_h)  # 標題列加第一列放不下，整個分類從新的一頁開始
            page = _Page()
        page.fixed += HEADER_H

        ordered = []
        while rows:
            pick = next((i for i, r in enumerate(rows) if page.fits(r, line_chars, page_h)), None)
            if pick is None:
                page.finish(page_h)
                page = _Page(fixed=HEADER_H)  # 換頁；標題列會在新頁重複出現
                pick = 0
            row = rows.pop(pick)
            ordered.append(row)
            page.add(row, line_chars)
        plan[name] = ordered
        page.fixed += SPACER_H
    page.finish(page_h)
    return plan


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


def _load_photos(photo_files, captions, fit_width, max_portrait_h, warnings):
    """載入照片並計算尺寸。載入失敗的照片以文字佔位，不中斷整份報告。"""
    categories = {}
    for category, paths in photo_files.items():
        items = []
        for path in paths:
            caption = captions.get(category, {}).get(path) or "[照片說明]"
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


def _add_photo_section(doc, main_table, photo_files, captions, warnings):
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
    line_chars = fit_width * 72 / 11  # 11pt 全形字寬 = 11pt

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

    categories = _load_photos(photo_files, captions, fit_width, max_portrait_h, warnings)
    plan = plan_layout(categories, page_h, start_y, line_chars)

    before = set(body)
    first_table = True
    for category, rows in plan.items():
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


def build_report(template_path, data, photo_files, captions, output_path):
    """產生報告並存檔，回傳警告訊息清單。"""
    doc = Document(template_path)
    if len(doc.tables) < 1:
        raise ValueError("模板格式錯誤：找不到表格")
    main_table = doc.tables[0]

    warnings = _fill_main_table(main_table, data)
    _add_photo_section(doc, main_table, photo_files, captions, warnings)
    doc.save(output_path)
    return warnings


# ---------------------------------------------------------------- GUI

class ReportGeneratorGUI:
    def __init__(self, root, autosave_path=DEFAULT_AUTOSAVE_PATH, restore_on_start=True,
                 autosave_interval_ms=AUTOSAVE_INTERVAL_MS):
        self.root = root
        self.autosave_path = autosave_path
        self.autosave_interval_ms = autosave_interval_ms
        self.root.title("保育巡查回報單生成器")
        self.root.geometry("900x700")

        self.template_path = None
        self.photo_files = {c: [] for c in PHOTO_CATEGORIES_1 + PHOTO_CATEGORIES_2}
        self.photo_captions = {c: {} for c in self.photo_files}  # category -> {path: caption}
        self.listboxes = {}
        self.text_fields = {}

        self._ui_queue = queue.Queue()
        self.create_widgets()
        self._drain_ui_queue()

        # 草稿狀態：與「上次儲存」或「空白」比較，判斷是否有未存內容
        self._saved_signature = self._signature()
        self._blank_signature = self._saved_signature
        self._last_autosave_signature = None
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(self.autosave_interval_ms, self._autosave_tick)
        if restore_on_start:
            self.root.after(300, self.offer_restore)

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(main_frame, text="模板檔案:", font=('Arial', 10, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        self.template_label = ttk.Label(main_frame, text="未選擇", foreground="red")
        self.template_label.grid(row=row, column=1, sticky=tk.W, padx=5)
        ttk.Button(main_frame, text="選擇模板", command=self.select_template).grid(row=row, column=2, padx=5)

        row += 1
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

        self.generate_btn = ttk.Button(button_frame, text="產生報告", command=self.generate_report)
        self.generate_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="儲存草稿", command=self.save_draft).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="載入草稿", command=self.load_draft).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清除", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="離開", command=self.on_close).pack(side=tk.LEFT, padx=5)

        row += 1
        self.status_label = ttk.Label(main_frame, text="準備就緒", relief=tk.SUNKEN)
        self.status_label.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E))

    # ---- 共用：可捲動分頁
    def make_scrollable_tab(self, notebook, title):
        frame = ttk.Frame(notebook, padding="10")
        notebook.add(frame, text=title)

        canvas = tk.Canvas(frame, highlightthickness=0)
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

        def date_entries(parent, prefix, last_label):
            y = ttk.Entry(parent, width=8)
            y.pack(side=tk.LEFT)
            ttk.Label(parent, text="年").pack(side=tk.LEFT, padx=2)
            m = ttk.Entry(parent, width=5)
            m.pack(side=tk.LEFT)
            ttk.Label(parent, text="月").pack(side=tk.LEFT, padx=2)
            d = ttk.Entry(parent, width=5)
            d.pack(side=tk.LEFT)
            ttk.Label(parent, text=last_label).pack(side=tk.LEFT, padx=2 if last_label == "日" else 5)
            setattr(self, f'{prefix}year', y)
            setattr(self, f'{prefix}month', m)
            setattr(self, f'{prefix}day', d)

        row = 0
        ttk.Label(frame, text="1. 年月日:", font=('Arial', 9, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        date_frame = ttk.Frame(frame)
        date_frame.grid(row=row, column=1, sticky=tk.W, padx=5)
        date_entries(date_frame, '', "日")

        row += 1
        ttk.Label(frame, text="2. 協勤人員姓名:", font=('Arial', 9, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        self.personnel = ttk.Entry(frame, width=50)
        self.personnel.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5)

        row += 1
        ttk.Label(frame, text="3. 協勤日期:", font=('Arial', 9, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        date_range_frame = ttk.Frame(frame)
        date_range_frame.grid(row=row, column=1, sticky=tk.W, padx=5)
        date_entries(date_range_frame, 'start_', "日 至")
        date_entries(date_range_frame, 'end_', "日")

        row += 1
        ttk.Label(frame, text="4. 協勤地點:", font=('Arial', 9, 'bold')).grid(row=row, column=0, sticky=tk.W, pady=5)
        self.location = ttk.Entry(frame, width=50)
        self.location.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5)

    def create_text_tab(self, notebook, title, fields, height):
        inner = self.make_scrollable_tab(notebook, title)
        row_idx = 0
        for label, key in fields:
            if key is None:
                ttk.Label(inner, text=label, font=('Arial', 10, 'bold'), wraplength=760).grid(
                    row=row_idx, column=0, sticky=tk.W, pady=(15, 5))
                row_idx += 1
                continue
            ttk.Label(inner, text=label, font=('Arial', 9, 'bold'), wraplength=760).grid(
                row=row_idx, column=0, sticky=tk.W, pady=(10, 2))
            text_widget = scrolledtext.ScrolledText(inner, height=height, width=70)
            text_widget.grid(row=row_idx + 1, column=0, sticky=(tk.W, tk.E), padx=5)
            self.text_fields[key] = text_widget
            row_idx += 2

    def create_photo_tab(self, notebook, title, heading, categories):
        inner = self.make_scrollable_tab(notebook, title)
        ttk.Label(inner, text=heading, font=('Arial', 11, 'bold')).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 10))
        ttk.Label(inner, text="（雙擊照片可編輯說明）", foreground="gray").grid(
            row=0, column=0, sticky=tk.E)

        for idx, category in enumerate(categories):
            row = idx * 2 + 1
            ttk.Label(inner, text=category, font=('Arial', 9, 'bold')).grid(
                row=row, column=0, sticky=tk.W, pady=(10, 5))

            list_frame = ttk.Frame(inner)
            list_frame.grid(row=row + 1, column=0, sticky=(tk.W, tk.E), padx=20)
            list_frame.columnconfigure(0, weight=1)

            listbox = tk.Listbox(list_frame, height=4, selectmode=tk.EXTENDED)
            listbox.grid(row=0, column=0, sticky=(tk.W, tk.E))
            scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
            scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
            listbox.configure(yscrollcommand=scrollbar.set)
            listbox.bind("<Double-Button-1>", lambda e, c=category, lb=listbox: self.edit_caption(c, lb))

            btn_frame = ttk.Frame(list_frame)
            btn_frame.grid(row=1, column=0, columnspan=2, pady=5)
            ttk.Button(btn_frame, text="選擇照片",
                       command=lambda c=category, lb=listbox: self.select_photos(c, lb)).pack(side=tk.LEFT, padx=2)
            ttk.Button(btn_frame, text="移除選取",
                       command=lambda c=category, lb=listbox: self.remove_photos(c, lb)).pack(side=tk.LEFT, padx=2)
            ttk.Button(btn_frame, text="清除全部",
                       command=lambda c=category, lb=listbox: self.clear_photos(c, lb)).pack(side=tk.LEFT, padx=2)

            self.listboxes[category] = listbox

    # ---- 事件處理
    def select_template(self):
        filepath = filedialog.askopenfilename(
            title="選擇模板檔案",
            filetypes=[("Word文件", "*.docx"), ("所有檔案", "*.*")]
        )
        if filepath:
            self.template_path = filepath
            self.template_label.config(text=os.path.basename(filepath), foreground="green")
            self.status_label.config(text=f"已選擇模板: {os.path.basename(filepath)}")
            logging.info(f"選擇模板: {filepath}")

    def _photo_label(self, category, filepath):
        caption = self.photo_captions[category].get(filepath)
        name = os.path.basename(filepath)
        return f"{name} — {caption}" if caption else name

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

    def remove_photos(self, category, listbox):
        selected = listbox.curselection()
        if not selected:
            return
        for idx in reversed(selected):
            filepath = self.photo_files[category].pop(idx)
            self.photo_captions[category].pop(filepath, None)
            listbox.delete(idx)
        self.status_label.config(text=f"已移除 {len(selected)} 張照片")

    def clear_photos(self, category, listbox):
        self.photo_files[category].clear()
        self.photo_captions[category].clear()
        listbox.delete(0, tk.END)
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

    def clear_all(self):
        if not messagebox.askyesno("確認", "確定要清除所有資料嗎？"):
            return
        self._clear_widgets()
        self.status_label.config(text="已清除所有資料")

    # ---- 草稿存檔
    def build_draft(self):
        """把目前畫面上的內容（原始輸入，不含 bullet 處理）整理成可存成 JSON 的 dict。"""
        return {
            'version': DRAFT_VERSION,
            'saved_at': datetime.now().isoformat(timespec='seconds'),
            'template_path': self.template_path,
            'fields': {name: getattr(self, name).get() for name in ENTRY_FIELDS},
            'texts': {key: w.get('1.0', tk.END).rstrip('\n') for key, w in self.text_fields.items()},
            'photos': {c: list(p) for c, p in self.photo_files.items()},
            'captions': {c: dict(v) for c, v in self.photo_captions.items()},
        }

    def apply_draft(self, draft):
        """把草稿內容還原到畫面，回傳警告訊息清單（例如找不到的照片）。"""
        if not isinstance(draft, dict) or draft.get('version') != DRAFT_VERSION:
            raise ValueError("不是有效的草稿檔，或草稿版本不相容")
        warnings = []
        self._clear_widgets()

        for name in ENTRY_FIELDS:
            getattr(self, name).insert(0, str(draft.get('fields', {}).get(name, '')))
        for key, widget in self.text_fields.items():
            widget.insert('1.0', str(draft.get('texts', {}).get(key, '')))

        # 已在本次手動選好模板時，以目前選擇為準；草稿記錄的模板只在尚未選擇時才帶入
        template = draft.get('template_path')
        if not self.template_path and template:
            if os.path.exists(template):
                self.template_path = template
                self.template_label.config(text=os.path.basename(template), foreground="green")
            else:
                warnings.append(f"找不到草稿記錄的模板：{template}，請重新選擇模板")

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
            initialdir=os.path.dirname(self.template_path) if self.template_path else None,
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
    def _signature(self):
        """目前內容的指紋（不含存檔時間與模板路徑），用來判斷內容是否改變。"""
        draft = self.build_draft()
        draft.pop('saved_at')
        draft.pop('template_path')
        return json.dumps(draft, ensure_ascii=False, sort_keys=True)

    def is_dirty(self):
        return self._signature() != self._saved_signature

    def autosave_now(self):
        """有未存且非空白的內容時，寫入自動存檔。回傳是否有寫入。"""
        signature = self._signature()
        if (signature == self._saved_signature or signature == self._blank_signature
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
                           "是：儲存草稿後離開\n否：不儲存直接離開\n取消：回到程式")
            if answer is None:
                return
            if answer and not self.save_draft():
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
        for label, prefix in [("年月日", ''), ("協勤日期（起）", 'start_'), ("協勤日期（迄）", 'end_')]:
            try:
                dates[prefix] = datetime(int(data[prefix + 'year']), int(data[prefix + 'month']),
                                         int(data[prefix + 'day']))
            except ValueError:
                return None, f"「{label}」不是有效的日期，請檢查"
        if dates['start_'] > dates['end_']:
            return None, "協勤日期的起日不可晚於迄日"

        for key, widget in self.text_fields.items():
            data[key] = add_bullet_points(widget.get('1.0', tk.END).strip())
        return data, None

    def generate_report(self):
        if not self.template_path:
            messagebox.showerror("錯誤", "請先選擇模板檔案")
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
            title="儲存報告",
            defaultextension=".docx",
            initialdir=os.path.dirname(self.template_path),
            initialfile=f"保育巡查回報單_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
            filetypes=[("Word文件", "*.docx")]
        )
        if not output_path:
            return
        if os.path.exists(output_path) and os.path.samefile(output_path, self.template_path):
            messagebox.showerror("錯誤", "輸出檔案不能和模板是同一個檔案，否則會覆蓋模板，"
                                         "下次再用它當模板時內容會重複疊加。\n請換一個檔名。")
            return

        self.generate_btn.config(state=tk.DISABLED)
        self.status_label.config(text="正在產生報告...")
        threading.Thread(target=self._generate_report_worker,
                         args=(data, photo_files, captions, output_path), daemon=True).start()

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

            message = f"報告已產生！\n\n檔案: {output_path}\n日誌: {LOG_FILE}"
            if warnings:
                message += "\n\n⚠ 注意：\n" + "\n".join(f"- {w}" for w in warnings)
                logging.warning("警告: " + "; ".join(warnings))
            self._post(lambda: messagebox.showinfo("成功", message))
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
    logging.info("保育巡查回報單生成器 GUI 版本啟動")
    logging.info(f"執行命令: {sys.executable} {' '.join(sys.argv)}")
    logging.info("=" * 60)

    root.mainloop()


if __name__ == '__main__':
    main()
