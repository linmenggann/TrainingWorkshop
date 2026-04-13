/**
 * 教學訓練計畫主持人工作坊 — Google Apps Script
 *
 * 【設定步驟】
 * 1. 開啟 Google Sheets：
 *    https://docs.google.com/spreadsheets/d/1DSm07gh94ber84KwkDE-638mRjU17dk_O5vExkT-zxc/edit
 *
 * 2. 確認分頁名稱為「工作坊報名資料」
 *
 * 3. 在第一列（A1~I1）貼上以下表頭：
 *    時間戳記 | 姓名 | 機構名 | 職稱 | 負責職類 | 是否擔任主持人 | 參與方式 | Email | 聯繫電話
 *
 * 4. 點選上方選單「擴充功能」>「Apps Script」
 *
 * 5. 將此檔案的全部程式碼貼入編輯器，取代原有內容，然後儲存
 *
 * 6. 點選「部署」>「新增部署作業」
 *    - 類型選擇「網頁應用程式」
 *    - 執行身分：選「我」
 *    - 誰可以存取：選「所有人」
 *    - 點選「部署」
 *
 * 7. 複製產生的 Web App URL
 *
 * 8. 將該 URL 貼回 index.html 中的 APPS_SCRIPT_URL 變數
 *    （取代 'YOUR_APPS_SCRIPT_WEB_APP_URL'）
 */

const SHEET_NAME = '工作坊報名資料';

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: '找不到分頁：' + SHEET_NAME }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = JSON.parse(e.postData.contents);

    const timestamp = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

    const row = [
      timestamp,
      data.name || '',
      data.institution || '',
      data.title || '',
      data.category || '',
      data.isHost || '',
      data.mode || '',
      data.email || '',
      data.phone || ''
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: '報名成功' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // action=dashboard：回傳所有報名資料供儀表板使用
  if (e && e.parameter && e.parameter.action === 'dashboard') {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_NAME);

      if (!sheet) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', message: '找不到分頁：' + SHEET_NAME }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'success', data: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const range = sheet.getRange(2, 1, lastRow - 1, 9);
      const values = range.getValues();

      const records = values.map(function(row) {
        return {
          timestamp: row[0],
          name: row[1],
          institution: row[2],
          title: row[3],
          category: row[4],
          isHost: row[5],
          mode: row[6],
          email: row[7],
          phone: row[8]
        };
      });

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success', data: records }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 如果有帶參數，視為報名提交（解決跨域 POST 重導向問題）
  if (e && e.parameter && e.parameter.name) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_NAME);

      if (!sheet) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'error', message: '找不到分頁：' + SHEET_NAME }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      const p = e.parameter;
      const timestamp = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

      const row = [
        timestamp,
        p.name || '',
        p.institution || '',
        p.title || '',
        p.category || '',
        p.isHost || '',
        p.mode || '',
        p.email || '',
        p.phone || ''
      ];

      sheet.appendRow(row);

      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success', message: '報名成功' }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '工作坊報名 API 運作中 ✅' }))
    .setMimeType(ContentService.MimeType.JSON);
}
