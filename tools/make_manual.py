"""產生給使用者看的 PDF 使用說明（docs/使用說明.pdf）：用 docs/images/ 的畫面截圖加上標註，排進 Word 文件，再用 Word 轉成 PDF。
用法：python tools/make_manual.py（需要 Windows 與 Microsoft Word）。打包時 build_exe.py 會把這份 PDF 放進 zip。"""
import os
import sys
import tempfile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import yscv_report_generator as g   # noqa: E402

IMAGES = os.path.join(ROOT, 'docs', 'images')
OUTPUT = os.path.join(ROOT, 'docs', '使用說明.pdf')
FONT = 'Microsoft JhengHei'
RED = '#d92d20'


LEFT_TRIM = 8      # 左邊的視窗陰影也裁掉；標註座標都已扣掉這 8 像素


def load(name, trim=7):
    """載入截圖；視窗截圖右邊與下面各有幾個像素的黑邊（視窗陰影），裁掉。"""
    image = Image.open(os.path.join(IMAGES, name)).convert('RGB')
    return image.crop((LEFT_TRIM, 0, image.width - trim, image.height - trim))


def annotate(image, boxes, notes):
    """boxes：要圈起來的 (x0, y0, x1, y1)；notes：(x, y, 文字)，紅字加白底。"""
    image = image.copy()
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(os.path.join(ROOT, 'fonts', 'Iansui-Regular.ttf'), 25)
    for box in boxes:
        draw.rounded_rectangle(box, radius=8, outline=RED, width=3)
    for x, y, text in notes:
        left, top, right, bottom = draw.textbbox((x, y), text, font=font)
        draw.rounded_rectangle((left - 8, top - 5, right + 8, bottom + 5), radius=6, fill='white', outline=RED, width=2)
        draw.text((x, y), text, fill=RED, font=font)
    return image


def set_font(run, size=None, bold=None, color=None):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def para(doc, text, size=11, bold=False, color=None, after=4, style=None, align=None):
    p = doc.add_paragraph(style=style)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(0)
    if align:
        p.alignment = align
    set_font(p.add_run(text), size, bold, color)
    return p


def heading(doc, text):
    p = para(doc, text, size=15, bold=True, color='1F3A93', after=4)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.keep_with_next = True


def bullets(doc, items):
    for item in items:
        p = para(doc, item, size=11, after=2, style='List Bullet')
        p.paragraph_format.left_indent = Cm(0.9)


FULL_WIDTH_CM = 18.4        # A4 寬 21 cm − 左右邊界各 1.2 cm


def picture(doc, image, caption, width_cm=FULL_WIDTH_CM):
    path = os.path.join(tempfile.mkdtemp(), 'figure.png')
    image.save(path)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(path, width=Cm(width_cm))
    para(doc, caption, size=9.5, color='666666', after=8, align=WD_ALIGN_PARAGRAPH.CENTER)


def build(docx_path):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.left_margin = section.right_margin = Cm(1.2)      # 邊界縮小，圖才放得夠大
    section.top_margin = section.bottom_margin = Cm(1.5)

    para(doc, '玉山保育志工回報單產生器', size=22, bold=True, color='1F3A93', after=0)
    para(doc, f'使用說明（適用 v{g.VERSION}）', size=13, color='555555', after=8)

    heading(doc, '一、開始使用')
    bullets(doc, [
        '把 zip 解壓縮到任一資料夾，雙擊「玉山保育志工回報單產生器.exe」就能使用，不需要安裝 Python。',
        '要產生 PDF 需要安裝 Microsoft Word；沒有 Word 時仍會產生 Word 檔（.docx），只是沒有 PDF。',
        '第一次執行時，Windows 若出現「Windows 已保護您的電腦」，請按「其他資訊」，再按「仍要執行」。',
    ])

    heading(doc, '二、選擇回報單')
    para(doc, '開啟後先選要填寫的回報單：用滑鼠點選，或按數字鍵 1～3；按 Esc 離開。右下角的「檢查更新」可以查看有沒有新版，'
              '「意見反應」可以直接把問題或建議送給開發者，最右下角的灰色小字是目前的版本。')
    picture(doc, load('01_入口選單.png'), '圖 1　入口選單', 14.5)

    heading(doc, '三、填寫內容')
    bullets(doc, [
        '依分頁由左到右填寫。分頁名稱後面的「●」表示這一頁還有欄位沒填，填完就會消失。',
        '日期會自動帶入今天，可以直接修改；志工姓名有多位時，用空白、頓號（、）或逗號隔開即可。',
    ])
    top = load('02_基本資料.png')
    picture(doc, top.crop((0, 0, 1105, 255)), '圖 2　基本資料分頁（只截上半部）')
    para(doc, '需要寫比較多字時，可以調整欄位大小：每個填寫欄下方有一條小橫條，用滑鼠往上或往下拖曳，就能改變那個欄位的高度（圖 3 紅框處）。',
         after=2)
    third = annotate(load('03_填寫欄位.png').crop((0, 0, 1105, 520)), [(626, 306, 714, 326), (626, 485, 714, 505)],
                     [(222, 303, '拖曳這條橫條，上下調整欄位高度')])
    picture(doc, third, '圖 3　填寫欄位（可上下調整高度）')

    heading(doc, '四、加入與整理照片')
    bullets(doc, [
        '到「服勤照片」分頁，先選好上方的小分頁（照片類別），再按「加入照片」，或直接從檔案總管把照片拖進視窗（放在哪個小分頁就加到哪一類）。',
        '點選清單裡的照片，右邊會預覽；在預覽下方的欄位輸入說明，按 Enter 會儲存並跳到下一張。',
        '用滑鼠在清單上按住照片上下拖曳，可以調整順序；也可以把照片拖到上方其他小分頁，移到別的類別。',
        '預覽區的按鈕可以「向左轉／向右轉 90°」與「裁切」，都不會改到原本的照片檔。',
        '要移除照片：選取後按 Delete，或按「移除所選照片」。報告裡的照片會依你排的順序，每 2 張一列。',
    ])
    photo = annotate(load('04_照片.png'), [(32, 224, 692, 290), (960, 779, 1319, 815)],
                     [(62, 560, '拖曳照片可調整順序，也可拖到上方其他小分頁'),
                      (62, 610, '也可以從檔案總管把照片拖進視窗'),
                      (742, 690, '裁切、旋轉：不會改到原本的照片檔')])
    picture(doc, photo, '圖 4　服勤照片分頁')

    heading(doc, '五、儲存草稿與產生報告')
    bullets(doc, [
        '「儲存草稿（Ctrl+S）」可以把填到一半的內容存成檔案，之後用「載入草稿（Ctrl+L）」接著填。草稿只能載入同一種回報單。',
        '程式每 30 秒會自動存檔；如果當機或不小心關掉，下次開啟會詢問是否還原。',
        '填好後按「產生報告（Ctrl+G）」：若有欄位沒填會先提醒；選好存檔位置後，會同時產生 Word 檔與同名的 PDF，並自動開啟。',
        '檔名是自動取的：報告是「日期_地點回報單」（不放人員姓名）；草稿是「日期_地點回報單草稿_全部人員姓名（以 _ 分隔）_儲存時間」，可以在存檔視窗修改。',
        '「清除所有資料」會清空目前的內容且無法復原，請先儲存草稿。畫面右下角可以切換外觀（柔和淺色、色弱友善、高對比）。',
    ])

    heading(doc, '六、其他')
    bullets(doc, [
        '設定與自動存檔放在「%LOCALAPPDATA%\\yscv-report-generator」資料夾，裡面可能有姓名等資料，分享前請先刪除。',
        '遇到問題或有建議，請用入口選單右下角的「意見反應」；回報時附上聯絡 Email（選填）比較方便回覆。',
    ])
    doc.save(docx_path)


def main():
    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, '使用說明.docx')
        build(docx_path)
        g.export_pdf(docx_path, OUTPUT)
    print('完成：', OUTPUT)


if __name__ == '__main__':
    main()
