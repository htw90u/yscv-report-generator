// 在 Node 裡載入 Apps Script（apps_script_feedback.gs），用假的 MailApp／ContentService／LockService／CacheService
// 實際跑 doPost，檢查輸入過濾、長度上限、密鑰、頻率限制。用法：node apps_script_harness.js <gs 檔案路徑>
// 輸出一行 JSON：{ 檢查名稱: true | 錯誤訊息 }
const fs = require('fs');
const vm = require('vm');

const code = fs.readFileSync(process.argv[2], 'utf8');

function fresh(overrides) {
  const sent = [];
  const store = new Map();
  const ctx = {
    MailApp: { sendEmail: (to, subject, body, opts) => sent.push({ to, subject, body, opts }) },
    ContentService: {
      MimeType: { JSON: 'json' },
      createTextOutput: (text) => ({ text, setMimeType() { return this; } }),
    },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    CacheService: {
      getScriptCache: () => ({ get: (k) => (store.has(k) ? store.get(k) : null), put: (k, v) => store.set(k, v) }),
    },
  };
  vm.createContext(ctx);
  vm.runInContext(code, ctx);
  for (const [name, value] of Object.entries(overrides || {})) {
    vm.runInContext(`${name} = ${JSON.stringify(value)}`, ctx);
  }
  const post = (payload) => JSON.parse(ctx.doPost({
    postData: { contents: typeof payload === 'string' ? payload : JSON.stringify(payload) },
  }).text);
  return { ctx, sent, post };
}

const results = {};
function check(name, fn) {
  try {
    const value = fn();
    results[name] = value === true ? true : `不符：${JSON.stringify(value)}`;
  } catch (e) {
    results[name] = `例外：${e}`;
  }
}

check('正常寄出（收件人、寄件人名稱、回信地址）', () => {
  const { ctx, sent, post } = fresh();
  const r = post({ subject: '[v1] 標題', body: '內容', app: '測試程式', reply_to: 'me@example.com' });
  return r.ok === true && sent.length === 1 && sent[0].to === ctx.TO && sent[0].subject === '[v1] 標題'
    && sent[0].body === '內容' && sent[0].opts.name === '測試程式' && sent[0].opts.replyTo === 'me@example.com';
});

check('主旨夾帶換行與標頭：換行被拿掉', () => {
  const { sent, post } = fresh();
  post({ subject: 'a\r\nBcc: x@y.com\r\nSubject: z', body: 'b' });
  return sent.length === 1 && !/[\r\n]/.test(sent[0].subject) && sent[0].subject === 'a Bcc: x@y.com Subject: z';
});

check('內文的控制字元與雙向文字覆寫字元被拿掉，換行統一', () => {
  const { sent, post } = fresh();
  post({ subject: 's', body: '內\u0000容‮反轉\r\n第二行\r第三行 ' });
  return sent[0].body === '內容反轉\n第二行\n第三行';
});

check('長度上限：主旨 200、內文 10000、寄件人名稱 60', () => {
  const { sent, post } = fresh();
  post({ subject: 'S'.repeat(1000), body: 'B'.repeat(50000), app: 'A'.repeat(500) });
  return sent[0].subject.length === 200 && sent[0].body.length === 10000 && sent[0].opts.name.length === 60;
});

check('不正確的回信地址不會被設成 Reply-To（但信照寄）', () => {
  const bad = ['a@b', 'x@y.com,evil@z.com', 'me@example.com\r\nBcc: e@x.com', '<me@example.com>', 'a'.repeat(300) + '@x.io',
               'me@example.com;e@x.com', 'me@@example.com', '@example.com'];
  const { sent, post } = fresh({ MAX_PER_MINUTE: 1000, MAX_PER_HOUR: 1000 });
  const failures = [];
  for (const reply_to of bad) {
    const n = sent.length;
    post({ subject: 's', body: 'b', reply_to });
    if (sent.length !== n + 1 || sent[n].opts.replyTo !== undefined) failures.push(reply_to);
  }
  return failures.length === 0 ? true : failures;
});

check('沒有 app、subject 時用預設值', () => {
  const { ctx, sent, post } = fresh();
  post({ body: 'x' });
  return sent[0].opts.name === ctx.DEFAULT_NAME && sent[0].subject === ctx.DEFAULT_SUBJECT;
});

check('設了密鑰：沒帶、帶錯都不寄，帶對才寄', () => {
  const { sent, post } = fresh({ TOKEN: 'secret' });
  const none = post({ subject: 's', body: 'b' });
  const wrong = post({ subject: 's', body: 'b', token: 'nope' });
  const right = post({ subject: 's', body: 'b', token: 'secret' });
  return none.ok === false && wrong.ok === false && right.ok === true && sent.length === 1;
});

check('沒設密鑰：不檢查', () => {
  const { sent, post } = fresh();
  return post({ subject: 's', body: 'b', token: 'whatever' }).ok === true && sent.length === 1;
});

check('每分鐘頻率限制：第 6 封被擋', () => {
  const { sent, post } = fresh();
  const results = Array.from({ length: 7 }, () => post({ subject: 's', body: 'b' }));
  return results.slice(0, 5).every((r) => r.ok) && results[5].ok === false && /頻繁/.test(results[5].error)
    && results[6].ok === false && sent.length === 5;
});

check('每小時頻率限制', () => {
  const { sent, post } = fresh({ MAX_PER_MINUTE: 1000, MAX_PER_HOUR: 3 });
  const results = Array.from({ length: 5 }, () => post({ subject: 's', body: 'b' }));
  return results.slice(0, 3).every((r) => r.ok) && !results[3].ok && !results[4].ok && sent.length === 3;
});

check('被擋的請求不會消耗寄信（含被密鑰擋下的）', () => {
  const { sent, post } = fresh({ TOKEN: 'secret', MAX_PER_MINUTE: 2 });
  for (let i = 0; i < 10; i++) post({ subject: 's', body: 'b', token: 'bad' });      // 10 次錯誤密鑰
  const a = post({ subject: 's', body: 'b', token: 'secret' });
  const b = post({ subject: 's', body: 'b', token: 'secret' });
  return a.ok && b.ok && sent.length === 2;                                           // 正常使用者不受影響
});

check('壞掉的 JSON：回報失敗而不是丟例外', () => {
  const { sent, post } = fresh();
  const r = post('這不是 JSON');
  return r.ok === false && typeof r.error === 'string' && sent.length === 0;
});

check('只有控制字元的主旨／名稱：改用預設值', () => {
  const { ctx, sent, post } = fresh();
  post({ subject: '\r\n\t\u0000', body: 'b', app: '‮\r\n' });
  return sent[0].subject === ctx.DEFAULT_SUBJECT && sent[0].opts.name === ctx.DEFAULT_NAME;
});

console.log(JSON.stringify(results));
