"""意見反應的 Google Apps Script 端點設定範本。

使用方式：複製本檔為 yscv_feedback_endpoint.py（已在 .gitignore，不會被提交），
把 URL 換成部署後的網頁應用程式網址（https://script.google.com/macros/s/.../exec，一定要是 https）。
部署步驟見 apps_script_feedback.gs 檔頭。

程式裡只有這個網址：寄信由 Google 端的 Apps Script 執行，所以沒有任何信箱帳號密碼，
收件信箱也只設定在 Apps Script 裡。沒有這個檔案時，按「寄出」會提示尚未設定，其他功能不受影響。
"""
URL = "https://script.google.com/macros/s/XXXXXXXX/exec"

# 選用：共用密鑰。Apps Script 裡的 TOKEN 設成一樣的字，不一樣的請求就不會寄信。
# 留空表示不使用。注意：密鑰會跟著程式一起發出去，只能擋「只知道網址的隨手濫用」，被濫用時請換掉。
TOKEN = ""
