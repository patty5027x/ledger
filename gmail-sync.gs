/**
 * AI 記帳 — Gmail 自動記帳（Google Apps Script）
 *
 * 做的事：
 *   每 15 分鐘看一次 Gmail，找「刷卡消費通知」和「訂閱收據」，
 *   用 Gemini 讀出 商店／金額／日期／哪張卡，存在一張 Google 試算表裡。
 *   記帳 app 打開時會來這裡拿新的紀錄。
 *
 * 設定步驟（README 有圖文版）：
 *   1. 到 script.google.com 建立新專案，把這整個檔案貼進去。
 *   2. 左邊「專案設定」→「指令碼屬性」新增兩個：
 *        GEMINI_KEY = 你的 Gemini API key
 *        TOKEN      = 自己想一組密碼（例如 patty-ledger-8273），等等也要貼到 app
 *   3. 上面選 setup 函式 → 執行，第一次會要求授權 Gmail / 試算表，按允許。
 *   4. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *        執行身分：我
 *        誰可以存取：任何人
 *      部署後複製「網頁應用程式網址」，貼到 app 的 設定 → Gmail。
 */

// 要找的信。可以依你的銀行通知信主旨再調整。
var QUERY = 'newer_than:7d -in:spam -in:trash (' +
  'subject:(消費 OR 刷卡 OR 交易 OR 授權 OR 簽帳 OR 扣款 OR 帳單) OR ' +
  'subject:(receipt OR invoice OR 收據 OR 訂閱 OR subscription OR "your payment" OR 付款) OR ' +
  'from:(no_reply@email.apple.com OR info@account.netflix.com OR no-reply@spotify.com OR googleplay-noreply@google.com)' +
  ')';
var MODEL = 'gemini-2.5-flash';
var CATEGORIES = ['餐飲', '生活用品', '交通', '娛樂', '訂閱', '學習', '醫療', '其他'];

function setup() {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('GEMINI_KEY') || !p.getProperty('TOKEN')) throw new Error('請先在「專案設定 → 指令碼屬性」加入 GEMINI_KEY 和 TOKEN');
  sheet_();
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'syncMail') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncMail').timeBased().everyMinutes(15).create();
  syncMail();
  Logger.log('設定完成，試算表：' + sheet_().getParent().getUrl());
}

/** 每 15 分鐘自動執行 */
function syncMail() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    var sh = sheet_();
    var done = {};
    var ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
    ids.forEach(function (r) { if (r[0]) done[r[0]] = true; });
    var threads = GmailApp.search(QUERY, 0, 40);
    var count = 0;
    threads.forEach(function (th) {
      th.getMessages().forEach(function (m) {
        var id = m.getId();
        if (done[id] || count >= 25) return;
        var body = (m.getPlainBody() || '').replace(/\s+/g, ' ').slice(0, 3500);
        var info;
        try { info = parse_(m.getSubject(), m.getFrom(), m.getDate(), body); }
        catch (e) { Logger.log('解析失敗 ' + id + '：' + e); return; }
        count++;
        sh.appendRow([id, info.kind || 'ignore', info.merchant || '', +info.amount || 0, info.currency || 'TWD',
          info.date || Utilities.formatDate(m.getDate(), 'Asia/Taipei', 'yyyy-MM-dd'), info.cardHint || '', info.last4 ? "'" + info.last4 : '',
          info.category || '', info.subscriptionName || '', m.getSubject(), Date.now()]);
        done[id] = true;
      });
    });
  } finally { lock.releaseLock(); }
}

/** 給 app 拿資料：?token=...&since=毫秒 */
function doGet(e) {
  var p = PropertiesService.getScriptProperties();
  var token = (e && e.parameter && e.parameter.token) || '';
  if (!token || token !== p.getProperty('TOKEN')) return json_({ ok: false, error: '密碼不對' });
  var since = +((e.parameter && e.parameter.since) || 0);
  var sh = sheet_(); var n = sh.getLastRow() - 1;
  var rows = n > 0 ? sh.getRange(2, 1, n, 12).getValues() : [];
  var items = rows.filter(function (r) { return r[1] !== 'ignore' && (+r[11] || 0) >= since; }).map(function (r) {
    return { id: String(r[0]), kind: r[1], merchant: r[2], amount: +r[3] || 0, currency: r[4], date: fmtDate_(r[5]), cardHint: r[6], last4: String(r[7] || ''), category: r[8], subscriptionName: r[9] };
  }).filter(function (it) { return it.currency === 'TWD' || !it.currency; });
  return json_({ ok: true, items: items, now: Date.now() });
}

/* ---------- 內部 ---------- */
function parse_(subject, from, date, body) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  var prompt = '下面是一封 Email。判斷它是不是：\n' +
    '(a) 信用卡「刷卡／消費／交易成功」通知 → kind = "card"\n' +
    '(b) 訂閱服務的「付款收據」（Netflix、Spotify、Apple、Google Play、YouTube 等每月扣款）→ kind = "subscription"\n' +
    '(c) 其他（廣告、帳單總額通知、繳款成功、退款、OTP 驗證碼…）→ kind = "ignore"\n' +
    '讀出：merchant（商店或服務名稱，簡短）、amount（新台幣金額數字；若是外幣，填外幣數字並把 currency 設成該幣別代碼）、currency（TWD 等）、' +
    'date（交易日期 YYYY-MM-DD）、cardHint（銀行或卡片名稱，例如「國泰 CUBE」「富邦 J 卡」）、last4（卡號末四碼）、' +
    'category（只能從：' + CATEGORIES.join('、') + '）、subscriptionName（訂閱服務名稱，只有 kind=subscription 才填）。\n' +
    '寄件者：' + from + '\n主旨：' + subject + '\n日期：' + date + '\n內容：' + body;
  var schema = { type: 'OBJECT', properties: {
    kind: { type: 'STRING', enum: ['card', 'subscription', 'ignore'] }, merchant: { type: 'STRING' }, amount: { type: 'NUMBER' },
    currency: { type: 'STRING' }, date: { type: 'STRING' }, cardHint: { type: 'STRING' }, last4: { type: 'STRING' },
    category: { type: 'STRING' }, subscriptionName: { type: 'STRING' } }, required: ['kind'] };
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + MODEL + ':generateContent?key=' + encodeURIComponent(key), {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: schema } })
  });
  var j = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) throw new Error((j.error && j.error.message) || res.getResponseCode());
  return JSON.parse(j.candidates[0].content.parts[0].text);
}
function sheet_() {
  var p = PropertiesService.getScriptProperties(); var id = p.getProperty('SHEET_ID'); var ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('AI 記帳 - Gmail 同步紀錄');
    p.setProperty('SHEET_ID', ss.getId());
    ss.getSheets()[0].appendRow(['gmailId', 'kind', 'merchant', 'amount', 'currency', 'date', 'cardHint', 'last4', 'category', 'subscriptionName', 'subject', 'savedAt']);
  }
  return ss.getSheets()[0];
}
function fmtDate_(v) { return v instanceof Date ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd') : String(v || ''); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
