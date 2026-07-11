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

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'dashboard') {
    return dashboardResponse_(e.parameter.callback);
  }

  return jsonResponse_({
    ok: true,
    message: '工作坊報名 API 運作中'
  });
}

function dashboardResponse_(callback) {
  try {
    const data = getDashboardData_();
    const payload = { ok: true, data: data };

    if (callback && /^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(callback)) {
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return jsonResponse_(payload);
  } catch (error) {
    console.error(error);
    return jsonResponse_({ ok: false, message: error.message });
  }
}

function getDashboardData_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('找不到分頁：' + SHEET_NAME);
  }

  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getDisplayValues()
    : [];

  const professionCounts = countBy_(rows, 4);
  const institutionCounts = countBy_(rows, 2);
  const attendanceCounts = countBy_(rows, 6);
  const hostCounts = countBy_(rows, 5);
  const timeZone = spreadsheet.getSpreadsheetTimeZone() || 'Asia/Taipei';

  return {
    total: rows.length,
    physical: attendanceCounts['實體'] || 0,
    online: attendanceCounts['線上'] || 0,
    hostYes: hostCounts['是'] || 0,
    hostNo: hostCounts['否'] || 0,
    uniqueInstitutions: Object.keys(institutionCounts).length,
    professions: toSortedItems_(professionCounts),
    institutions: toSortedItems_(institutionCounts),
    latestSubmission: rows.length ? rows[rows.length - 1][0] : '',
    updatedAt: Utilities.formatDate(new Date(), timeZone, 'yyyy/MM/dd HH:mm:ss')
  };
}

function countBy_(rows, columnIndex) {
  return rows.reduce(function (counts, row) {
    const label = String(row[columnIndex] || '未填寫').trim() || '未填寫';
    counts[label] = (counts[label] || 0) + 1;
    return counts;
  }, {});
}

function toSortedItems_(counts) {
  return Object.keys(counts)
    .map(function (label) {
      return { label: label, value: counts[label] };
    })
    .sort(function (a, b) {
      return b.value - a.value || a.label.localeCompare(b.label, 'zh-Hant');
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
