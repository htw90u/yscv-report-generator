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
var DEFAULT_NAME = "玉山保育志工回報單產生器";
var DEFAULT_SUBJECT = "意見反應";

var MAX_SUBJECT = 200;       // 主旨最多幾個字
var MAX_BODY = 10000;        // 內文最多幾個字
var MAX_REPLY_TO = 254;
var MAX_PER_MINUTE = 5;      // 整個端點每分鐘最多寄幾封（不分來源，保護信箱與每日寄信配額）
var MAX_PER_HOUR = 30;
// 選用的共用密鑰：設成一串自訂的字，程式端（yscv_feedback_endpoint.py 的 TOKEN）要填一樣的，不一樣就不寄。
// 留空就不檢查。注意密鑰會跟著程式發出去，能擋「只知道網址的隨手濫用」，擋不了有心人，被濫用時換掉即可。
var TOKEN = "";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (TOKEN && String(data.token || "") !== TOKEN) {
      return json_({ ok: false, error: "未授權" });
    }
    if (!allow_()) {
      return json_({ ok: false, error: "寄送太頻繁，請稍後再試" });
    }
    var opts = {
      // 寄件人名稱：程式有傳 app 欄位就用它，沒傳就是預設名稱，所以舊版小工具不受影響
      name: oneLine_(data.app || DEFAULT_NAME).slice(0, 60) || DEFAULT_NAME
    };
    var reply = oneLine_(data.reply_to || "");
    if (reply && reply.length <= MAX_REPLY_TO &&
        /^[A-Za-z0-9._%+\-]{1,64}@[A-Za-z0-9\-]+(\.[A-Za-z0-9\-]+)+$/.test(reply)) {
      opts.replyTo = reply;
    }
    var subject = oneLine_(data.subject || DEFAULT_SUBJECT).slice(0, MAX_SUBJECT) || DEFAULT_SUBJECT;
    var body = String(data.body || "").replace(/\r\n?/g, "\n")
        .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f\u2028\u2029\u202a-\u202e\u2066-\u2069]/g, "")
        .slice(0, MAX_BODY);
    MailApp.sendEmail(TO, subject, body, opts);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// 單行文字：拿掉控制字元與雙向文字覆寫字元，換行／Tab／連續空白都變成一個空白（避免在主旨裡夾帶標頭）
function oneLine_(text) {
  return String(text)
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f\u202a-\u202e\u2066-\u2069]/g, "")
      .replace(/\s+/g, " ").trim();
}

// 整個端點的頻率限制：每分鐘、每小時各一個計數器（存在 CacheService），超過就拒絕
function allow_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var cache = CacheService.getScriptCache();
    var now = new Date().getTime();
    var minuteKey = "fb_m_" + Math.floor(now / 60000);
    var hourKey = "fb_h_" + Math.floor(now / 3600000);
    var perMinute = Number(cache.get(minuteKey) || 0);
    var perHour = Number(cache.get(hourKey) || 0);
    if (perMinute >= MAX_PER_MINUTE || perHour >= MAX_PER_HOUR) return false;
    cache.put(minuteKey, String(perMinute + 1), 120);
    cache.put(hourKey, String(perHour + 1), 7200);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON);
}
