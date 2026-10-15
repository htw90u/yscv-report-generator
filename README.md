# yscv-report-generator

玉山保育志工回報單生成器：用圖形介面填寫玉山國家公園志工的回報單，自動產生 Word（.docx），並用 Word 轉成同名的 PDF。

程式啟動後先選擇要填寫的回報單：

| 回報單 | 範本檔 |
| --- | --- |
| 保育巡查回報單 | `保育巡查回報單.docx` |
| 塔塔加保育宣導回報單 | `塔塔加保育宣導回報單.docx` |
| 生態廁所回報單（玉山主峰線） | `生態廁所回報單.docx` |

## 功能

- 依固定範本填入基本資料與各項內容（範本放在程式旁邊，要更新表單格式直接換掉範本檔）
- 服勤照片：從檔案總管拖進視窗、拖曳調整順序、拖到其他分類、旋轉、裁切（不改原檔）、填寫說明
- 報告裡的照片依你排的順序放置（橫式與直式不並排），並放大到填滿儲存格
- 草稿存檔／載入、每 30 秒自動存檔（三種回報單各自存檔，草稿也不會載入到別種回報單）
- 視窗標題顯示版本與 git 編號：`<tag>.<tag 之後第幾個 commit>-<sha>`，剛好在 tag 上是 `v0.5-76b5acc`，之後第 12 個 commit 是 `v0.5.12-3c1f9ab`
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

也可以在命令列直接指定回報單，跳過選單：`python yscv_report_generator.py tataka`（`patrol`、`tataka`、`ecotoilet`）。

## 測試

```bash
python -m unittest discover -s tests
```

## 字型

介面使用「芫荽 Iansui」（`fonts/Iansui-Regular.ttf`，SIL Open Font License，授權文字見 `fonts/OFL-Iansui.txt`），只在程式執行期間載入，不會安裝到系統。

## 授權

程式碼以 [MIT License](LICENSE) 授權。

- 字型「芫荽 Iansui」維持它自己的 SIL Open Font License，不受 MIT 涵蓋。
- 三份範本（`保育巡查回報單.docx`、`塔塔加保育宣導回報單.docx`、`生態廁所回報單.docx`）是玉山國家公園志工回報單的表單，其著作權屬於原機關，不在 MIT 授權的範圍內。
