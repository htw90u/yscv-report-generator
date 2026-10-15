# yscv-report-generator

玉山保育巡查回報單生成器：用圖形介面填寫玉山國家公園高山保育志工協勤回報單，自動產生 Word（.docx），並用 Word 轉成同名的 PDF。

## 功能

- 依固定範本 `玉山國家公園高山保育志工協勤回報單.docx` 填入基本資料與各項內容
- 服勤照片：從檔案總管拖進視窗、拖曳調整順序、拖到其他分類、旋轉、裁切（不改原檔）、填寫說明
- 報告裡的照片依你排的順序放置（橫式與直式不並排），並放大到填滿儲存格
- 草稿存檔／載入、每 30 秒自動存檔
- 產生前檢查未填欄位；產生後自動開啟檔案
- 外觀主題（柔和淺色、色弱友善、高對比）

## 需求

- Windows、Python 3（開發與測試使用 3.14）
- 套件：`python-docx`、`Pillow`
- 轉成 PDF 需要安裝 Microsoft Word；沒有 Word 時仍會產生 .docx，只是不會有 PDF

```bash
pip install python-docx Pillow
python yscv_report_generator.py
```

## 測試

```bash
python -m unittest discover -s tests
```

## 字型

介面使用「芫荽 Iansui」（`fonts/Iansui-Regular.ttf`，SIL Open Font License，授權文字見 `fonts/OFL-Iansui.txt`），只在程式執行期間載入，不會安裝到系統。
