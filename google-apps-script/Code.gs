const SPREADSHEET_ID = '1v0_9-XG5DikOT4-Kk1mUToUXATkleCutJi4dwB3HMBA';
const SHEET_NAME = '工作坊報名資料';
const HEADERS = [
  '送出時間',
  '姓名',
  '機構名',
  '職稱',
  '負責的職類',
  '是否擔任教學訓練計畫主持人',
  '參與方式',
  'Email',
  '聯繫電話'
];

function doGet() {
  return jsonResponse_({
    ok: true,
    message: '工作坊報名 API 運作中'
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('未收到報名資料');
    }

    const data = JSON.parse(e.postData.contents);
    const requiredFields = [
      'name', 'organization', 'title', 'profession',
      'host', 'attendance', 'email', 'phone'
    ];
    const missingFields = requiredFields.filter(function (field) {
      return !String(data[field] || '').trim();
    });

    if (missingFields.length) {
      throw new Error('缺少必填欄位：' + missingFields.join(', '));
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('找不到分頁：' + SHEET_NAME);
    }

    ensureHeaders_(sheet);
    sheet.appendRow([
      new Date(),
      safeCell_(data.name),
      safeCell_(data.organization),
      safeCell_(data.title),
      safeCell_(data.profession),
      safeCell_(data.host),
      safeCell_(data.attendance),
      safeCell_(data.email),
      safeCell_(data.phone)
    ]);

    return jsonResponse_({ ok: true, message: '報名資料已儲存' });
  } catch (error) {
    console.error(error);
    return jsonResponse_({ ok: false, message: error.message });
  } finally {
    lock.releaseLock();
  }
}

function ensureHeaders_(sheet) {
  const range = sheet.getRange(1, 1, 1, HEADERS.length);
  const current = range.getDisplayValues()[0];
  const matches = HEADERS.every(function (header, index) {
    return current[index] === header;
  });

  if (!matches) {
    range.setValues([HEADERS]);
    range
      .setBackground('#7c3aed')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }
}

function safeCell_(value) {
  const text = String(value || '').trim();
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
