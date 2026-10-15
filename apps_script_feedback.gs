/**
 * 玉山保育志工回報單產生器 — 意見反應代寄端點（Google Apps Script）
 *
 * 作用：接收程式 POST 來的 JSON（subject / body / app），用「部署者的 Google 帳號」寄信到 TO。
 * 郵件帳號密碼完全不進程式，也不管使用者有沒有安裝郵件程式。
 *
 * 部署步驟（約 5 分鐘）：
 * 1. 用要寄信的 Google 帳號開 https://script.google.com → 新增專案，
 *    把本檔內容全部貼進 Code.gs（覆蓋預設內容），把下面的 TO 改成要收意見的信箱，存檔。
 * 2. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」：
 *    - 執行身分：我（你的帳號）
 *    - 具有存取權的使用者：任何人
 * 3. 按「部署」→ 首次會要求授權（允許代表你傳送電子郵件）→
 *    複製產生的網頁應用程式網址（https://script.google.com/macros/s/.../exec）。
 * 4. 回到專案：複製 yscv_feedback_endpoint.example.py 為 yscv_feedback_endpoint.py，
 *    把 URL 填成上面的網址。
 *
 * 之後要更新這段程式：「部署」→「管理部署作業」→ 鉛筆 → 版本選「新版本」→ 部署（網址不變）。
 * 若端點被濫用，封存該部署再重新部署，換一個新網址即可。
 *
 * 配額：免費 Gmail 帳號經 MailApp 每日約可寄 100 封，意見反應綽綽有餘。
 */
var TO = "請改成收件信箱";
var MAX_SUBJECT = 200;
var MAX_BODY = 10000;

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var subject = String(data.subject || "意見反應").slice(0, MAX_SUBJECT);
    var body = String(data.body || "").slice(0, MAX_BODY);
    var name = String(data.app || "玉山保育志工回報單產生器").slice(0, 60);
    MailApp.sendEmail(TO, subject, body, { name: name });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
}
