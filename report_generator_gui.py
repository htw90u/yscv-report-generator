#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
保育巡查回報單生成器 - GUI版本
提供圖形化介面供使用者輸入資料並選擇照片
"""

import io
import logging
import os
import queue
import sys
import threading
import tkinter as tk
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
    """回傳可供 add_picture 使用的來源；有 EXIF 旋轉時先轉正。"""
    with Image.open(path) as img:
        orientation = img.getexif().get(0x0112, 1)
        if orientation in (1, 0):
            return path
        img = ImageOps.exif_transpose(img)
        buf = io.BytesIO()
        if img.mode in ('RGBA', 'LA', 'P'):
            img.convert('RGBA').save(buf, 'PNG')
        else:
            img.convert('RGB').save(buf, 'JPEG', quality=92)
        buf.seek(0)
        return buf


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


def _add_photo_section(doc, main_table, photo_files, captions, warnings):
    # 只刪除內容含「照片」的舊照片表格，避免誤刪其他表格
    for table in list(doc.tables)[1:]:
        if '照片' in _all_text(table):
            table._element.getparent().remove(table._element)

    # 模板內可能已有「九、服勤照片」標題（表格內或內文段落），有的話不重複新增
    has_heading = ('九、服勤照片' in _all_text(main_table)
                   or any('九、服勤照片' in p.text for p in doc.paragraphs))
    if not has_heading:
        doc.add_paragraph()
        heading = doc.add_paragraph('九、服勤照片')
        heading_run = heading.runs[0]
        heading_run.font.bold = True
        heading_run.font.size = Pt(14)
        set_font_kai(heading_run)

    for category, photos in photo_files.items():
        if not photos:
            continue

        table = doc.add_table(rows=1, cols=2)
        header_cell = table.rows[0].cells[0].merge(table.rows[0].cells[1])
        header_cell.text = category
        set_cell_border(header_cell)
        for paragraph in header_cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                set_font_kai(run)
                run.font.bold = True

        for i in range(0, len(photos), 2):
            pair = photos[i:i + 2]
            photo_row = table.add_row()
            caption_row = table.add_row()

            for col_idx in range(2):
                set_cell_border(photo_row.cells[col_idx])
                set_cell_border(caption_row.cells[col_idx])

            for col_idx, photo_path in enumerate(pair):
                cell = photo_row.cells[col_idx]
                paragraph = cell.paragraphs[0]
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                name = os.path.basename(photo_path)
                try:
                    source = prepare_image(photo_path)
                    paragraph.add_run().add_picture(source, width=Inches(2.8))
                    logging.info(f"✓ SUCCESS: 插入照片 {name}")
                except Exception as e:
                    logging.error(f"無法插入照片 {name}: {e}")
                    warnings.append(f"照片「{name}」無法載入：{e}")
                    run = paragraph.add_run(f"[照片無法載入：{name}]")
                    set_font_kai(run)

                caption = captions.get(category, {}).get(photo_path) or f"[照片{col_idx + 1}說明]"
                caption_cell = caption_row.cells[col_idx]
                caption_cell.text = caption
                for p in caption_cell.paragraphs:
                    for r in p.runs:
                        set_font_kai(r)

        doc.add_paragraph()
        logging.info(f"✓ SUCCESS: 處理完成 {category}")


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
    def __init__(self, root):
        self.root = root
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
        ttk.Button(button_frame, text="清除", command=self.clear_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="離開", command=self.root.destroy).pack(side=tk.LEFT, padx=5)

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

    def clear_all(self):
        if not messagebox.askyesno("確認", "確定要清除所有資料嗎？"):
            return

        for entry in (self.year, self.month, self.day, self.personnel,
                      self.start_year, self.start_month, self.start_day,
                      self.end_year, self.end_month, self.end_day, self.location):
            entry.delete(0, tk.END)
        for text_widget in self.text_fields.values():
            text_widget.delete('1.0', tk.END)
        for category, listbox in self.listboxes.items():
            self.photo_files[category].clear()
            self.photo_captions[category].clear()
            listbox.delete(0, tk.END)

        self.status_label.config(text="已清除所有資料")

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
