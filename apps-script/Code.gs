/**
 * 教學訓練計畫主持人工作坊報名接收端。
 * 貼到 Apps Script 的 Code.gs，執行 setupSheet 後部署為網頁應用程式。
 * 詳細操作請見 README.md。試算表不需要公開分享。
 */
const CONFIG = Object.freeze({
  spreadsheetId: '1LEllNReIM9pXQ6M3Jac7CgOgfJ29gNWz5t9qlcP0D1Y',
  sheetName: '工作坊報名資料',
  timezone: 'Asia/Taipei',
  websiteUrl: 'https://linmenggann.github.io/TrainingWorkshop/'
});

const HEADERS = Object.freeze([
  '送出時間', '姓名', '機構名', '職稱', '負責的職類',
  '目前是否擔任教學訓練計畫主持人', '預計參與方式', 'Email', '聯繫電話', '報名編號'
]);
const PROFESSIONS = Object.freeze([
  '藥事', '醫事放射', '醫事檢驗', '護理', '營養', '呼吸治療',
  '聽力', '物理治療', '職能治療', '臨床心理', '語言治療', '其他'
]);

/** 第一次部署前，從編輯器手動執行一次，授權並建立／檢查表頭。 */
function setupSheet() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getRegistrationSheet_();
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#673cd0').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, HEADERS.length);
    SpreadsheetApp.flush();
    console.log('工作坊報名資料分頁與表頭已就緒。');
  } finally {
    lock.releaseLock();
  }
}

/** 開啟部署網址僅顯示狀態，不讀取或公開報名資料。 */
function doGet() {
  return receipt_('報名接收服務', '請從工作坊網頁填寫資料並送出報名。', false);
}

/** 使用一般 HTML 表單 POST，不依賴 fetch 的跨來源讀取權限。 */
function doPost(e) {
  let lock;
  try {
    const data = validateRegistration_(e);
    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      return receipt_('目前送出人數較多', '尚未確認資料寫入，請稍後回原頁以相同資料重試。', false);
    }
    const sheet = getRegistrationSheet_();
    // 在鎖內檢查並寫入，避免同一筆重試或同時送出造成重複列。
    if (sheet.getLastRow() > 1) {
      const existing = sheet.getRange(2, HEADERS.length, sheet.getLastRow() - 1, 1)
        .createTextFinder(data.requestId).matchEntireCell(true).findNext();
      if (existing) {
        return receipt_('這筆報名已收到 ✅', '相同報名編號已存在，未重複新增。', true, data.requestId);
      }
    }
    const submittedAt = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy/MM/dd HH:mm:ss');
    const row = [submittedAt, data.name, data.organization, data.job, data.profession,
      data.host, data.attendance, data.email, data.phone, data.requestId];
    const rowNumber = sheet.getLastRow() + 1;
    if (rowNumber > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    // 以文字寫入，保留電話開頭的 0；另外轉義可能被當作公式的內容。
    sheet.getRange(rowNumber, 1, 1, HEADERS.length)
      .setNumberFormat('@').setValues([row.map(sheetText_)]);
    SpreadsheetApp.flush();
    return receipt_('報名已送出 🎉', '資料已寫入「工作坊報名資料」分頁，請保留以下報名編號。', true, data.requestId, submittedAt);
  } catch (error) {
    // 不記錄姓名、Email、電話或整份請求。
    const message = error && error.userMessage
      ? error.userMessage
      : '無法確認資料寫入。請聯繫主辦單位，或回原頁以相同資料重試；相同編號不會重複新增。';
    return receipt_('報名尚未確認', message, false);
  } finally {
    if (lock && lock.hasLock()) lock.releaseLock();
  }
}

function getRegistrationSheet_() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sheet = spreadsheet.getSheetByName(CONFIG.sheetName) || spreadsheet.insertSheet(CONFIG.sheetName);
  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS.slice()]);
  } else {
    const actual = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    if (!HEADERS.every((header, index) => actual[index] === header)) {
      fail_('試算表表頭與程式不一致，請主辦單位依 README 檢查 A1:J1；既有資料未被覆蓋。');
    }
  }
  return sheet;
}

function validateRegistration_(e) {
  if (!e || !e.parameter || !e.postData || e.postData.length > 10000) {
    fail_('請從工作坊網頁提交有效的報名資料。');
  }
  const fields = {name: 80, organization: 150, job: 100, profession: 20,
    host: 1, attendance: 2, email: 254, phone: 40, requestId: 36};
  const data = {};
  Object.keys(fields).forEach(function (key) {
    if (e.parameters && e.parameters[key] && e.parameters[key].length !== 1) {
      fail_('資料格式不正確，請回原頁重新確認。');
    }
    const value = e.parameter[key];
    if (typeof value !== 'string' || !value.trim() || value.trim().length > fields[key]) {
      fail_('有欄位空白或超過長度限制，請回原頁重新確認。');
    }
    data[key] = value.trim();
  });
  if (PROFESSIONS.indexOf(data.profession) < 0 || ['是', '否'].indexOf(data.host) < 0 ||
      ['實體', '線上'].indexOf(data.attendance) < 0) {
    fail_('職類、主持人身分或參與方式不正確，請重新選擇。');
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) fail_('Email 格式不正確，請重新確認。');
  if (!/^[+\d][\d\s()+#xX.,\-分機轉]{5,39}$/.test(data.phone)) fail_('聯繫電話格式不正確，請填入手機或市話（可含分機）。');
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(data.requestId)) {
    fail_('報名編號格式不正確，請回工作坊網頁重新填寫。');
  }
  return data;
}

function fail_(message) {
  const error = new Error(message);
  error.userMessage = message;
  throw error;
}

function sheetText_(value) {
  const text = String(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function escapeHtml_(value) {
  return String(value).replace(/[&<>"']/g, function (char) {
    return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[char];
  });
}

function receipt_(title, message, success, requestId, submittedAt) {
  const details = requestId ? '<p class="code">報名編號<br><strong>' + escapeHtml_(requestId) + '</strong></p>' : '';
  const time = submittedAt ? '<p>送出時間：' + escapeHtml_(submittedAt) + '（台北時間）</p>' : '';
  const html = '<!doctype html><html lang="zh-Hant"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"><title>' + escapeHtml_(title) + '</title>' +
    '<style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:linear-gradient(135deg,#34106c,#9655dd);' +
    'font-family:system-ui,"Microsoft JhengHei",sans-serif;color:#24163f;line-height:1.8}' +
    'main{width:min(480px,calc(100% - 72px));margin:24px;padding:28px;border-radius:22px;background:white;text-align:center}' +
    'h1{font-size:26px}p{font-size:14px}.code{background:#f4effc;border-radius:10px;padding:14px;overflow-wrap:anywhere}' +
    'a{display:inline-block;padding:12px 22px;background:#673cd0;color:white;border-radius:10px;text-decoration:none}' +
    'small{display:block;color:#776b8c;margin-top:18px}</style></head><body><main>' +
    '<div style="font-size:40px" aria-hidden="true">' + (success ? '💜' : '📝') + '</div>' +
    '<h1>' + escapeHtml_(title) + '</h1><p>' + escapeHtml_(message) + '</p>' + details + time +
    '<a href="' + escapeHtml_(CONFIG.websiteUrl) + '" target="_top" rel="noopener">返回工作坊網頁</a>' +
    '<small>也可關閉此分頁，返回原本的填寫頁面。<br>奇美醫院教學部｜06-281-2811 分機 57112233</small>' +
    '</main></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
