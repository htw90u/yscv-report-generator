# AGENTS.md

給 AI 協作工具與新接手的開發者看的專案說明。使用者文件見 `README.md` 與 `docs/使用說明.pdf`；功能與規則的完整整理見 `skills/yscv-report-generator/SKILL.md`。

## 這是什麼

玉山保育志工回報單產生器：Tkinter 圖形介面，填固定的 Word 範本（保育巡查、塔塔加保育宣導、生態廁所三種），輸出 `.docx`，再用 Word COM 轉成同名 PDF。只支援 Windows。GitHub：`htw90u/yscv-report-generator`。

## 檔案地圖

| 檔案 | 用途 |
| --- | --- |
| `yscv_report_generator.py` | 主程式：共用元件（主題、字型、分頁標記、照片面板、草稿與自動存檔、PDF 轉檔）、入口選單 `ReportLauncher`、`VERSION` |
| `yscv_reports.py` | 塔塔加、生態廁所兩種回報單的畫面與 Word 填寫（保育巡查在主程式） |
| `yscv_updates.py` | 更新檢查（只查 GitHub `releases/latest`，不下載、不安裝） |
| `build_exe.py` | PyInstaller 打包（資料夾版）並壓成 `dist/yscv-report-generator-v<版本>-win64.zip`、印出 SHA-256 |
| `tools/make_manual.py` | 產生 `docs/使用說明.pdf`（用 `docs/images/` 的截圖，需要 Word） |
| `apps_script_feedback.gs` | 意見反應的 Google Apps Script 端 |
| `yscv_feedback_endpoint.py` | 意見反應的網址與密鑰，**git 忽略**，範本是 `yscv_feedback_endpoint.example.py` |
| `tests/` | `unittest`：`test_gui`、`test_reports`、`test_security`、`test_report`、`test_updates` |
| `*.docx`（根目錄三份） | Word 範本，著作權屬原機關 |

## 常用指令

```bash
pip install -r requirements.txt
python yscv_report_generator.py              # 執行；加 patrol / tataka / ecotoilet 直接開某種回報單
python yscv_report_generator.py --debug      # 寫 yscv_report_generator.log
python -m unittest discover -s tests         # 全部測試（test_gui 很慢，約 13 分鐘）
python -m unittest discover -s tests -p test_security.py   # 只跑相關的
python tools/make_manual.py                  # 重做 PDF 說明
python build_exe.py                          # 打包（先 pip install pyinstaller）
```

改動後只跑相關的測試檔；動到共用的版面、主題、字型、分頁時才跑全部。測試會開真的 Tk 視窗，**跑測試時不要同時編輯原始碼**（背景程序可能讀到改一半的檔案）。

## 慣例

- 介面、訊息、註解、commit 訊息一律用**繁體中文**；程式識別字維持英文。
- 使用者看得到的名稱統一叫「產生器」（不是「生成器」）。
- 每個改動都 commit；不要自己 push，也不要改寫已推送的歷史，除非使用者明確要求。
- 版本號：`VERSION` 不能落後最新的 git tag（測試會檢查）。畫面上的版本標籤是 `<tag>.<之後第幾個 commit>-<sha>`，打包時由 `build_exe.py` 寫進 `yscv_build_info.py`（git 忽略）。
- 檔名規則：報告 `<協勤日期>_<地點>回報單`，不放人員姓名；草稿放全部人員姓名並以 `_` 分隔。
- 介面字型預設「微軟正黑體」，可選「芫荽」；選項在入口選單，設定檔鍵 `font`。Word 報告字型固定標楷體。
- 分頁標記：未填＝橘色小圓點、填完＝勾；每個主題有自己的顏色（色弱友善用藍勾，高對比改用 `●` `✓` 文字）。
- 開啟回報單時只量「最高的那一頁」決定視窗高度（`MEASURE_TABS`），量的期間蓋載入中遮罩；文字較少的分頁用 `FILL_TEXT_TABS` 往下長滿。改版面後要確認量出來的那一頁真的最高、其他頁不用捲動（`StartupLayoutTest`）。

## 注意事項（踩過的坑）

- 使用者貼上的文字可能含 XML 不允許的控制字元，寫進 Word 前一律經 `xml_safe*`。
- 草稿檔可能來自別人：載入前先整份檢查格式，通過才動畫面；照片路徑只接受本機路徑，網路路徑（UNC）略過。
- 更新檢查只接受 https 與較新的正式版；「前往下載」的網址由程式用 repo 名稱和 tag 組出，不信任伺服器回傳的網址。
- 重建分頁後 `<Configure>` 不一定會觸發，需要版面計算時用 `after(80, …)` 補一次。
- 在 shell 的 heredoc 裡寫含 `\n` 或反斜線的 Python，內容容易被改壞；改程式請用編輯工具。
- 打包後資源讀自 `BASE_DIR`（`sys._MEIPASS`），使用者資料在 `%LOCALAPPDATA%\yscv-report-generator`；從原始碼執行則都在程式資料夾。

## 隱私與安全（不要違反）

- **不要提交**：`yscv_feedback_endpoint.py`（含網址與密鑰）、設定檔、自動存檔、草稿、記錄檔、任何含真實姓名或照片的檔案、`dist/`、`build/`。`.gitignore` 已排除，提交前仍要看一眼 `git status`。
- 不要在 repo、commit 訊息、文件或畫面上寫出意見反應的收件信箱；信箱只存在 Apps Script 端。
- 不要加入「用使用者的郵件程式寄信」的備援；意見反應只走 HTTPS POST 到 Apps Script。
- git 作者信箱使用 GitHub 的 noreply 位址，不要換成真實信箱。
- 示範用的截圖或檔案，檔名前加「示範_」，不要改動應用程式本身。

## 發行流程

1. 改 `VERSION`，更新 README／說明書／skill 的範例與內容，測試通過後 commit。
2. 打同名 annotated tag（例如 `v1.2`）。
3. `python build_exe.py`：產生 zip 與 SHA-256，並確認 .exe 能開啟。
4. 使用者手動推送、在 GitHub 建立 Release（上傳 zip、附 SHA-256）、上傳 VirusTotal，再把掃描結果補進 Release 說明。
5. 發行包只包含：執行檔資料夾、`使用說明.pdf`、`README.md`、`LICENSE`；zip 檔名維持英文。
