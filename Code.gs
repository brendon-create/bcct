// =============================================
//  BCCT 森入心流 點餐表單 — Google Apps Script
//  貼到 Extensions > Apps Script > Code.gs
// =============================================

const SHEET_ID = '1XkoxvJfl4fhp4lTBjVHqkY29ixPJZGrEeiF9UoZQqwg'; // ← 貼上你的 Google Sheet ID（網址列 /d/ 後面那段）

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const activity = data.activity || '未命名活動';
    const name     = data.name;
    const meal     = data.meal;
    const price    = data.price;
    const drink    = data.drink;
    const time     = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName(activity);

    // 如果這個活動的 tab 不存在，自動建立並加上標題列
    if (!sheet) {
      sheet = ss.insertSheet(activity);
      sheet.appendRow(['送出時間', '活動', '姓名', '餐點', '價格', '飲料']);
      sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }

    sheet.appendRow([time, activity, name, meal, price, drink]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 允許跨來源請求（CORS）
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'BCCT 點餐表單 API 正常運作中' }))
    .setMimeType(ContentService.MimeType.JSON);
}
