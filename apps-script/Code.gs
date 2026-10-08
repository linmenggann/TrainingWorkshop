/**
 * 教學訓練計畫主持人工作坊報名接收端。
 * 貼到 Apps Script 的 Code.gs，執行 setupSheet 後部署為網頁應用程式。
 * 詳細操作請見 README.md。試算表不需要公開分享。
 */
const CONFIG = Object.freeze({
  spreadsheetId: '1LEllNReIM9pXQ6M3Jac7CgOgfJ29gNWz5t9qlcP0D1Y',
  sheetName: '工作坊報名資料',
  registrationLimit: 4,
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

/** 儀表板只取得彙總統計；不傳回報名名單或個人聯絡資料。 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'capacity') {
    return capacityResponse_(e);
  }
  if (e && e.parameter && e.parameter.action === 'dashboard') {
    return dashboardResponse_(e);
  }
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
    // 同一把鎖涵蓋名額檢查、寫入與 flush；最後一席只會由一筆新報名取得。
    const registered = sheet.getLastRow() > 1
      ? sheet.getRange(2, 2, sheet.getLastRow() - 1, 8).getValues().filter(hasRegistrationFields_).length
      : 0;
    if (registered >= CONFIG.registrationLimit) {
      return receipt_('額滿', '本活動限額 4 名，目前已額滿，停止受理新報名。此筆資料未新增。', false);
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

/** 唯讀統計介面。JSONP 供 GitHub Pages 跨來源讀取；不需要公開整份試算表。 */
function dashboardResponse_(e) {
  const callback = e.parameter.callback || '';
  const validCallback = /^workshopDashboard_[A-Za-z0-9_]{1,80}$/.test(callback);
  let result;
  try {
    if (callback && !validCallback) fail_('統計請求格式不正確。');
    result = registrationStats_(e.parameter);
  } catch (error) {
    result = {type: 'workshop-dashboard', version: 1, ok: false,
      message: error && error.userMessage ? error.userMessage : '無法讀取報名統計，請主辦單位確認部署版本與試算表存取權限。'};
  }
  const json = JSON.stringify(result).replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  return ContentService.createTextOutput(validCallback ? callback + '(' + json + ');' : json)
    .setMimeType(validCallback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function registrationStats_(params) {
  const filters = {};
  ['profession', 'attendance', 'host', 'from', 'to'].forEach(function (key) {
    filters[key] = String(params[key] || '').trim();
  });
  if ((filters.profession && PROFESSIONS.indexOf(filters.profession) < 0) ||
      (filters.attendance && ['實體', '線上'].indexOf(filters.attendance) < 0) ||
      (filters.host && ['是', '否'].indexOf(filters.host) < 0)) fail_('篩選條件不正確。');
  ['from', 'to'].forEach(function (key) {
    if (filters[key] && !validDateKey_(filters[key])) fail_('請使用有效的篩選日期。');
  });
  if (filters.from && filters.to && filters.from > filters.to) fail_('起始日期不能晚於結束日期。');

  // 統計請求不可建立分頁、重設表頭或寫入任何資料。
  const values = readRegistrationValues_();
  const stats = {type: 'workshop-dashboard', version: 1, ok: true,
    timezone: CONFIG.timezone, generatedAt: new Date().toISOString(), filters: filters,
    total: 0, overallTotal: 0, today: 0, organizations: 0, lastRegistrationAt: null,
    professions: {}, attendance: {'實體': 0, '線上': 0, '未分類': 0},
    hosts: {'是': 0, '否': 0, '未分類': 0}, daily: [],
    quality: {incomplete: 0, unknownDate: 0}};
  PROFESSIONS.forEach(function (profession) { stats.professions[profession] = 0; });
  stats.professions['未分類'] = 0;
  const organizations = new Set();
  const days = {};
  const today = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');
  values.slice(1).forEach(function (row) {
    // 八個報名欄位全空的列不列入；部分缺值仍計入並提示資料品質。
    const fields = row.slice(1, 9).map(function (value) { return String(value == null ? '' : value).trim(); });
    if (!hasRegistrationFields_(fields)) return;
    stats.overallTotal++;
    const time = registrationTime_(row[0]);
    if ((filters.profession && fields[3] !== filters.profession) ||
        (filters.attendance && fields[5] !== filters.attendance) ||
        (filters.host && fields[4] !== filters.host) ||
        ((filters.from || filters.to) && !time) ||
        (filters.from && time.date < filters.from) || (filters.to && time.date > filters.to)) return;
    stats.total++;
    if (fields.some(function (value) { return value === ''; })) stats.quality.incomplete++;
    if (fields[1]) organizations.add(fields[1]);
    const profession = PROFESSIONS.indexOf(fields[3]) >= 0 ? fields[3] : '未分類';
    const attendance = ['實體', '線上'].indexOf(fields[5]) >= 0 ? fields[5] : '未分類';
    const host = ['是', '否'].indexOf(fields[4]) >= 0 ? fields[4] : '未分類';
    stats.professions[profession]++;
    stats.attendance[attendance]++;
    stats.hosts[host]++;
    if (time) {
      days[time.date] = (days[time.date] || 0) + 1;
      if (time.date === today) stats.today++;
      if (!stats.lastRegistrationAt || time.timestamp > stats.lastRegistrationAt) stats.lastRegistrationAt = time.timestamp;
    } else stats.quality.unknownDate++;
  });
  stats.organizations = organizations.size;
  stats.daily = Object.keys(days).sort().map(function (date) { return {date: date, count: days[date]}; });
  // 活動名額不受日期、職類或參與方式篩選影響。
  stats.capacity = capacityForCount_(stats.overallTotal);
  return stats;
}

function hasRegistrationFields_(fields) {
  return fields.some(function (value) { return String(value == null ? '' : value).trim() !== ''; });
}

function readRegistrationValues_() {
  const sheet = SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.sheetName);
  if (!sheet) fail_('找不到「工作坊報名資料」分頁，請主辦單位確認試算表設定。');
  if (sheet.getLastRow() === 0 || sheet.getMaxColumns() < HEADERS.length) fail_('統計表頭尚未設定，請主辦單位先執行 setupSheet。');
  const values = sheet.getRange(1, 1, sheet.getLastRow(), HEADERS.length).getValues();
  if (!HEADERS.every(function (header, index) { return values[0][index] === header; })) {
    fail_('試算表表頭不符，請主辦單位確認 A1:J1。');
  }
  return values;
}

function capacityForCount_(registered) {
  return {limit: CONFIG.registrationLimit, registered: registered,
    remaining: Math.max(0, CONFIG.registrationLimit - registered), full: registered >= CONFIG.registrationLimit};
}

/** 報名頁唯讀名額查詢：不回傳名單、個資或分類統計。 */
function capacityResponse_(e) {
  const callback = e.parameter.callback || '';
  const validCallback = /^workshopCapacity_[A-Za-z0-9_]{1,80}$/.test(callback);
  let result;
  try {
    if (callback && !validCallback) fail_('名額請求格式不正確。');
    const values = readRegistrationValues_();
    const registered = values.slice(1).filter(function (row) { return hasRegistrationFields_(row.slice(1, 9)); }).length;
    result = {type: 'workshop-capacity', version: 1, ok: true,
      generatedAt: new Date().toISOString(), capacity: capacityForCount_(registered)};
  } catch (error) {
    result = {type: 'workshop-capacity', version: 1, ok: false,
      message: error && error.userMessage ? error.userMessage : '暫時無法確認名額，請稍後再試。'};
  }
  const json = JSON.stringify(result).replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  return ContentService.createTextOutput(validCallback ? callback + '(' + json + ');' : json)
    .setMimeType(validCallback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function validDateKey_(date) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return false;
  const parsed = new Date(date + 'T00:00:00Z');
  return !isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === date;
}

function registrationTime_(value) {
  let text;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return null;
    text = Utilities.formatDate(value, CONFIG.timezone, 'yyyy/MM/dd HH:mm:ss');
  } else text = String(value == null ? '' : value).trim();
  const match = /^(\d{4})[\/-](\d{2})[\/-](\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(text);
  if (!match) return null;
  const date = match[1] + '-' + match[2] + '-' + match[3];
  const hour = match[4] || '00', minute = match[5] || '00', second = match[6] || '00';
  if (!validDateKey_(date) || Number(hour) > 23 || Number(minute) > 59 || Number(second) > 59) return null;
  return {date: date, timestamp: date + ' ' + hour + ':' + minute + ':' + second};
}
