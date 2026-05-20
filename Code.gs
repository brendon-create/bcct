// =============================================
//  BCCT 森入心流 — Google Apps Script
//  支援：點餐表單 / 團體意外險資料
//  貼到 Extensions > Apps Script > Code.gs
// =============================================

const SHEET_ID = '1L69fkJL4x4IxXtP5Gx1S68glfc8Jo6y-K85fwvVCdsg';

function doPost(e) {
  try {
    const data     = JSON.parse(e.postData.contents);
    const type     = data.type || 'order';
    const activity = data.activity || '未命名活動';
    const time     = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

    const ss = SpreadsheetApp.openById(SHEET_ID);

    if (type === 'insurance') {
      handleInsurance(ss, data, activity, time);
    } else {
      handleOrder(ss, data, activity, time);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleOrder(ss, data, activity, time) {
  const sheetName = '點餐｜' + activity;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['送出時間', '活動', '姓名', '餐點', '價格', '飲料']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  sheet.appendRow([time, activity, data.name, data.meal, data.price, data.drink]);
}

function handleInsurance(ss, data, activity, time) {
  const sheetName = '意外險｜' + activity;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['送出時間', '活動', '姓名', '身分證字號', '出生日期', '性別', '聯絡電話', '緊急聯絡人', '緊急聯絡人電話', '關係']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
  }
  sheet.appendRow([
    time, activity,
    data.name, data.idNumber, data.dob, data.gender, data.phone,
    data.emergencyName, data.emergencyPhone, data.relationship,
  ]);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'BCCT API 正常運作中' }))
    .setMimeType(ContentService.MimeType.JSON);
}
