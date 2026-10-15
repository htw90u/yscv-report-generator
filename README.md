# yscv-report-generator

玉山保育志工回報單產生器：用圖形介面填寫玉山國家公園志工的回報單，自動產生 Word（.docx），並用 Word 轉成同名的 PDF。

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
- 入口選單右下角的「意見反應」：填標題與內容，按寄出就直接送出（主旨是「[回報單產生器 v1.0-96aa62e] 你填的標題」，內文最後也附上版本；可選填聯絡 Email，會寫進信裡並設成回信地址），不用使用者的郵件程式；做法是 POST 到 Google Apps Script，由 Google 端寄信，設定方式見下方「意見反應的設定」
- 草稿存檔／載入、每 30 秒自動存檔（三種回報單各自存檔，草稿也不會載入到別種回報單）
- 版本與 git 編號用灰色小字顯示在視窗右下角（入口選單與三種回報單）：`<tag>.<tag 之後第幾個 commit>-<sha>`，剛好在 tag 上是 `v1.0-96aa62e`，之後第 12 個 commit 是 `v1.0.12-3c1f9ab`；打包成 .exe 時會把打包當下的版本烙進去
- 產生前檢查未填欄位；產生後自動開啟檔案
- 外觀主題（柔和淺色、色弱友善、高對比）

## 下載使用（不需要安裝 Python）

到 GitHub 的 Releases 下載最新的 `yscv-report-generator-v<版本>-win64.zip`，解壓縮後執行資料夾裡的「玉山保育志工回報單產生器.exe」。下載後可以用 `Get-FileHash <zip 檔> -Algorithm SHA256` 核對 Release 說明裡的雜湊值。

- 需要 Windows 10 或 11；要產生 PDF 需要安裝 Microsoft Word（沒有 Word 仍會產生 .docx）。
- 這個程式沒有程式碼簽章，第一次執行 Windows SmartScreen 可能警告，按「其他資訊」→「仍要執行」；部分防毒軟體可能誤判，請先核對 SHA-256。
- 設定、自動存檔與除錯記錄放在 `%LOCALAPPDATA%\yscv-report-generator\`（直接跑 .py 時則放在程式資料夾）。

## 使用說明（基本流程）

完整的圖文說明在 [`docs/使用說明.pdf`](docs/使用說明.pdf)（下載的 zip 裡也有一份）。基本流程：

1. **選回報單**：啟動後點選卡片或按數字鍵 1～3。
2. **填寫內容**：依分頁由左到右填寫，分頁名稱後面的「●」表示還有欄位沒填；文字填寫欄下方的小橫條可以上下拖曳，調整欄位高度。
3. **加入與整理照片**：在「服勤照片」分頁按「加入照片」，或從檔案總管把照片拖進視窗；點選照片預覽並輸入說明；用滑鼠拖曳調整順序（也能拖到其他小分頁）；預覽區可以旋轉與裁切（不改原檔）。
4. **儲存草稿**：Ctrl+S 儲存、Ctrl+L 載入；程式每 30 秒自動存檔，當機後重開會詢問是否還原。
5. **產生報告**：Ctrl+G；選好存檔位置後同時產生 Word 檔與同名的 PDF（PDF 需要安裝 Word），並自動開啟。

PDF 說明由 `python tools/make_manual.py` 產生（用 `docs/images/` 的截圖，需要 Word）。

## 從原始碼執行

### 需求

- Windows、Python 3（開發與測試使用 3.14）
- 套件：`python-docx`、`Pillow`（見 `requirements.txt`，建議保持最新版）
- 轉成 PDF 需要安裝 Microsoft Word；沒有 Word 時仍會產生 .docx，只是不會有 PDF

```bash
pip install -r requirements.txt
python yscv_report_generator.py
```

也可以在命令列直接指定回報單，跳過選單：`python yscv_report_generator.py tataka`（`patrol`、`tataka`、`ecotoilet`）。

平常不會產生記錄檔；要除錯時加上 `-d`（或 `-D`、`--debug`），完整的執行記錄會寫到 `yscv_report_generator.log`：

```bash
python yscv_report_generator.py --debug
```

## 打包成 .exe

```bash
pip install pyinstaller
python build_exe.py
```

`build_exe.py` 會用 PyInstaller 打包成資料夾版（`dist/玉山保育志工回報單產生器/`，裡面的執行檔就叫「玉山保育志工回報單產生器.exe」），並壓成 `dist/yscv-report-generator-v<版本>-win64.zip`、印出 SHA-256。圖示用 `icons/yscv.ico`；範本、字型與圖示都會一起包進去；打包當下的 `<tag>.<次數>-<sha>` 版本會烙進 .exe（打包後沒有 git）。
如果本機有 `yscv_feedback_endpoint.py`，意見反應的網址與密鑰會被一起打包進 .exe（沒有就不內建寄送功能）。

## 測試

```bash
python -m unittest discover -s tests
```

## 字型

介面使用「芫荽 Iansui」（`fonts/Iansui-Regular.ttf`，SIL Open Font License，授權文字見 `fonts/OFL-Iansui.txt`），只在程式執行期間載入，不會安裝到系統。

## 意見反應的設定

意見反應是 POST 到一個 Google Apps Script 網頁應用程式，由它用部署者的帳號寄信，所以程式裡沒有任何郵件帳號密碼，收件信箱也只設定在 Apps Script 裡。

1. 部署 `apps_script_feedback.gs`（步驟寫在檔案開頭的註解，約 5 分鐘），拿到網頁應用程式網址。
2. 複製 `yscv_feedback_endpoint.example.py` 為 `yscv_feedback_endpoint.py`，把 `URL` 換成那個網址（這個檔案已被 `.gitignore` 排除）。

沒有設定時，按「寄出」會提示這個版本沒有設定寄送功能，其他功能不受影響。

## 隱私與安全

- 自動存檔、草稿與除錯記錄（`-d`）都是明文，裡面有姓名等資料，只放在你自己的電腦上；它們已被 `.gitignore` 排除，**但直接壓縮整個資料夾分享時會被一起帶出去**，分享前請先刪掉。
- 草稿檔可能來自別人：載入時先檢查格式，格式不對不會動到畫面；草稿裡的照片只接受本機路徑，網路路徑（`\\主機\分享`）會被略過（Windows 開啟網路路徑時會向對方主機送出登入雜湊）。
- 意見反應只會傳你填的標題、內容、聯絡 Email 和程式版本，用 https 送出；標題與 Email 會過濾換行與控制字元，標題 100 字、內容 5000 字為上限，成功寄出後 30 秒內不能再寄。端點的密鑰與頻率限制設定在 Apps Script 裡（見 `apps_script_feedback.gs`）。

## 更新檢查

入口選單可以檢查 GitHub 上有沒有新版：啟動時自動檢查（最多每週一次）或按「檢查更新」。有新版只會在選單上方顯示提示，你可以「查看更新內容」「前往下載」（用瀏覽器開啟發行頁面）或「略過這個版本」；**程式不會自己下載或安裝任何東西**。

- 只向 GitHub 的公開 API 發出一個查詢，不傳任何個人資料。
- 不想自動檢查：取消選單左下角的「啟動時自動檢查更新」，或用 `--no-update-check`、環境變數 `YSCV_NO_UPDATE_CHECK=1`。
- 更新來源是 `yscv_updates.py` 的 `UPDATE_REPO`（目前是 `htw90u/yscv-report-generator`）；設成空字串，整個更新功能就不啟用，也不會連網。
- **repo 必須公開，使用者才查得到新版**：私人 repo 對沒登入的人是 404，選單會顯示「找不到已發行的版本」，不會誤報成「已是最新版」。
- 發行新版：先把 `yscv_report_generator.py` 的 `VERSION` 改成新版號再打同名 tag（例如 `VERSION = '1.1'`、tag `v1.1`），執行 `python build_exe.py` 打包，再在 GitHub 建立 Release 並上傳 zip（說明裡附上 SHA-256）；測試會檢查 `VERSION` 沒有落後最新的 tag。

## 授權

程式碼以 [MIT License](LICENSE) 授權。

- 字型「芫荽 Iansui」維持它自己的 SIL Open Font License，不受 MIT 涵蓋。
- 視窗與 .exe 圖示用 Google 的 Noto Emoji 圖片（Apache License 2.0，說明見 `icons/NOTICE.txt`），不受 MIT 涵蓋。
- 三份範本（`保育巡查回報單.docx`、`塔塔加保育宣導回報單.docx`、`生態廁所回報單.docx`）是玉山國家公園志工回報單的表單，其著作權屬於原機關，不在 MIT 授權的範圍內。
