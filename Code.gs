/**
 * Oil Price Dashboard - Google Apps Script
 * ดึงราคาน้ำมันจาก Thai Oil API → เก็บใน Google Sheets
 */

var SHEET_NAME = 'ราคาน้ำมัน';
var API_URL = 'https://api.chnwt.dev/thai-oil-api/latest';

// ==================== Web App ====================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('dashboard')
    .setTitle('Dashboard ราคาน้ำมัน ( Apps Script )')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==================== ดึงและบันทึกข้อมูล ====================

function fetchAndSaveOilPrices() {
  var sheet = getOrCreateSheet();
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy');

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var lastDate = sheet.getRange(lastRow, 1).getValue();
    if (lastDate == today) {
      return { status: 'skip', message: 'วันนี้บันทึกแล้ว: ' + today };
    }
  }

  var response = UrlFetchApp.fetch(API_URL, { muteHttpExceptions: true });
  var json = JSON.parse(response.getContentText());
  if (json.status !== 'success') {
    return { status: 'error', message: 'API ตอบกลับผิดพลาด' };
  }

  if (lastRow <= 1) {
    createHeaders(sheet);
  }

  var fuelOrder = [
    'gasoline_95', 'gasohol_95', 'gasohol_91', 'gasohol_e20', 'gasohol_e85',
    'diesel', 'diesel_b7', 'diesel_b20', 'premium_diesel',
    'premium_gasohol_95', 'superpower_gasohol_95', 'vpower_gasohol_95',
    'vpower_diesel', 'vpower_diesel_b7', 'premium_gasohol_97', 'ngv'
  ];
  var stationOrder = ['ptt', 'bcp', 'shell', 'esso', 'caltex', 'irpc', 'pt', 'susco', 'pure', 'susco_dealers'];

  var row = [today, json.response.date];
  stationOrder.forEach(function(station) {
    fuelOrder.forEach(function(fuel) {
      if (json.response.stations[station] && json.response.stations[station][fuel]) {
        row.push(parseFloat(json.response.stations[station][fuel].price) || 0);
      } else {
        row.push('');
      }
    });
  });

  sheet.appendRow(row);
  return { status: 'success', message: 'บันทึกสำเร็จ: ' + today };
}

function createHeaders(sheet) {
  var fuelOrder = [
    'gasoline_95', 'gasohol_95', 'gasohol_91', 'gasohol_e20', 'gasohol_e85',
    'diesel', 'diesel_b7', 'diesel_b20', 'premium_diesel',
    'premium_gasohol_95', 'superpower_gasohol_95', 'vpower_gasohol_95',
    'vpower_diesel', 'vpower_diesel_b7', 'premium_gasohol_97', 'ngv'
  ];
  var stationOrder = ['ptt', 'bcp', 'shell', 'esso', 'caltex', 'irpc', 'pt', 'susco', 'pure', 'susco_dealers'];

  var headers = ['วันที่', 'วันที่ (API)'];
  stationOrder.forEach(function(station) {
    fuelOrder.forEach(function(fuel) {
      headers.push(station.toUpperCase() + ' - ' + fuel);
    });
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

// ==================== ข้อมูลสำหรับ Dashboard ====================

function getLatestPrices() {
  var response = UrlFetchApp.fetch(API_URL, { muteHttpExceptions: true });
  return JSON.parse(response.getContentText());
}

function getHistoryFromSheet() {
  var sheet = getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var fuelOrder = [
    'gasoline_95', 'gasohol_95', 'gasohol_91', 'gasohol_e20', 'gasohol_e85',
    'diesel', 'diesel_b7', 'diesel_b20', 'premium_diesel',
    'premium_gasohol_95', 'superpower_gasohol_95', 'vpower_gasohol_95',
    'vpower_diesel', 'vpower_diesel_b7', 'premium_gasohol_97', 'ngv'
  ];
  var stationOrder = ['ptt', 'bcp', 'shell', 'esso', 'caltex', 'irpc', 'pt', 'susco', 'pure', 'susco_dealers'];

  return data.map(function(row) {
    var entry = { date: row[0] };
    var colIndex = 2;
    stationOrder.forEach(function(station) {
      fuelOrder.forEach(function(fuel) {
        var key = station + '_' + fuel;
        entry[key] = row[colIndex] !== '' ? parseFloat(row[colIndex]) : null;
        colIndex++;
      });
    });
    return entry;
  });
}

// ==================== Trigger ====================

function setupDailyTrigger() {
  var existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'fetchAndSaveOilPrices') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('fetchAndSaveOilPrices')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .nearMinute(0)
    .inTimezone('Asia/Bangkok')
    .create();
}

// ==================== Test ====================

function testFetch() {
  var result = fetchAndSaveOilPrices();
  Logger.log(result);
}

// ==================== Backfill ข้อมูลย้อนหลัง ====================

function backfillHistory() {
  var sheet = getOrCreateSheet();
  if (sheet.getLastRow() > 1) {
    return { status: 'skip', message: 'มีข้อมูลอยู่แล้ว — ต้องล้างข้อมูลเก่าก่อน หรือใช้ backfillAppend()' };
  }
  return backfillAppend();
}

function backfillAppend() {
  var sheet = getOrCreateSheet();
  if (sheet.getLastRow() <= 1) {
    createHeaders(sheet);
  }

  var fuelOrder = [
    'gasoline_95', 'gasohol_95', 'gasohol_91', 'gasohol_e20', 'gasohol_e85',
    'diesel', 'diesel_b7', 'diesel_b20', 'premium_diesel',
    'premium_gasohol_95', 'superpower_gasohol_95', 'vpower_gasohol_95',
    'vpower_diesel', 'vpower_diesel_b7', 'premium_gasohol_97', 'ngv'
  ];
  var stationOrder = ['ptt', 'bcp', 'shell', 'esso', 'caltex', 'irpc', 'pt', 'susco', 'pure', 'susco_dealers'];

  // ข้อมูลราคาน้ำมันย้อนหลังจาก Bangchak (ราคาเฉลี่ยทุกปั๊ม)
  // ที่มา: bangchak.co.th/th/oilprice/historical
  var history = [
    { date: '09/01/2569', bcp: { gasoline_95: 45.64, gasohol_91: 29.94, gasohol_e20: 26.59, premium_diesel: 30.48, diesel: 28.64 }, ptt: { gasoline_95: 45.64, gasohol_95: 38.55, gasohol_91: 38.18, gasohol_e20: 33.55, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '18/02/2569', bcp: { gasoline_95: 45.64, gasohol_91: 29.94, gasohol_e20: 26.29, premium_diesel: 30.18, diesel: 28.34 }, ptt: { gasoline_95: 45.64, gasohol_95: 38.55, gasohol_91: 38.18, gasohol_e20: 33.55, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '10/03/2569', bcp: { gasoline_95: 45.64, gasohol_91: 29.94, gasohol_e20: 25.79, premium_diesel: 30.68, diesel: 27.84 }, ptt: { gasoline_95: 45.64, gasohol_95: 38.55, gasohol_91: 38.18, gasohol_e20: 33.55, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '18/03/2569', bcp: { gasoline_95: 46.14, gasohol_91: 30.44, gasohol_e20: 23.79, premium_diesel: 31.68, diesel: 27.05 }, ptt: { gasoline_95: 46.14, gasohol_95: 38.80, gasohol_91: 38.43, gasohol_e20: 33.80, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '21/03/2569', bcp: { gasoline_95: 46.84, gasohol_91: 31.14, gasohol_e20: 24.79, premium_diesel: 32.68, diesel: 28.05 }, ptt: { gasoline_95: 46.84, gasohol_95: 39.15, gasohol_91: 38.78, gasohol_e20: 34.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '24/03/2569', bcp: { gasoline_95: 48.84, gasohol_91: 32.94, gasohol_e20: 26.79, premium_diesel: 34.68, diesel: 30.05 }, ptt: { gasoline_95: 48.84, gasohol_95: 40.15, gasohol_91: 39.78, gasohol_e20: 35.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '26/03/2569', bcp: { gasoline_95: 56.84, gasohol_91: 38.94, gasohol_e20: 32.79, premium_diesel: 40.68, diesel: 36.05 }, ptt: { gasoline_95: 56.84, gasohol_95: 44.15, gasohol_91: 43.78, gasohol_e20: 39.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '31/03/2569', bcp: { gasoline_95: 58.64, gasohol_91: 40.74, gasohol_e20: 33.79, premium_diesel: 41.68, diesel: 37.05 }, ptt: { gasoline_95: 58.64, gasohol_95: 45.15, gasohol_91: 44.78, gasohol_e20: 40.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '02/04/2569', bcp: { gasoline_95: 62.14, gasohol_91: 44.24, gasohol_e20: 34.99, premium_diesel: 42.88, diesel: 38.25 }, ptt: { gasoline_95: 62.14, gasohol_95: 47.15, gasohol_91: 46.78, gasohol_e20: 42.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '03/04/2569', bcp: { gasoline_95: 66.14, gasohol_91: 47.74, gasohol_e20: 35.69, premium_diesel: 43.58, diesel: 38.95 }, ptt: { gasoline_95: 66.14, gasohol_95: 49.15, gasohol_91: 48.78, gasohol_e20: 44.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '04/04/2569', bcp: { gasoline_95: 66.14, gasohol_91: 47.74, gasohol_e20: 34.89, premium_diesel: 43.58, diesel: 38.95 }, ptt: { gasoline_95: 66.14, gasohol_95: 49.15, gasohol_91: 48.78, gasohol_e20: 44.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '05/04/2569', bcp: { gasoline_95: 70.94, gasohol_91: 50.54, gasohol_e20: 34.89, premium_diesel: 43.58, diesel: 38.95 }, ptt: { gasoline_95: 70.94, gasohol_95: 51.15, gasohol_91: 50.78, gasohol_e20: 46.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '09/04/2569', bcp: { gasoline_95: 68.80, gasohol_91: 48.40, gasohol_e20: 34.89, premium_diesel: 43.58, diesel: 38.95 }, ptt: { gasoline_95: 68.80, gasohol_95: 50.15, gasohol_91: 49.78, gasohol_e20: 45.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '11/04/2569', bcp: { gasoline_95: 66.80, gasohol_91: 44.40, gasohol_e20: 31.89, premium_diesel: 42.58, diesel: 35.95 }, ptt: { gasoline_95: 66.80, gasohol_95: 48.15, gasohol_91: 47.78, gasohol_e20: 43.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '17/04/2569', bcp: { gasoline_95: 65.30, gasohol_91: 42.90, gasohol_e20: 31.39, premium_diesel: 42.08, diesel: 35.45 }, ptt: { gasoline_95: 65.30, gasohol_95: 47.15, gasohol_91: 46.78, gasohol_e20: 42.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '21/04/2569', bcp: { gasoline_95: 64.10, gasohol_91: 41.70, gasohol_e20: 31.39, premium_diesel: 42.08, diesel: 35.45 }, ptt: { gasoline_95: 64.10, gasohol_95: 46.15, gasohol_91: 45.78, gasohol_e20: 41.15, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '24/04/2569', bcp: { gasoline_95: 62.10, gasohol_91: 40.20, gasohol_e20: 31.39, premium_diesel: 42.08, diesel: 35.45 }, ptt: { gasoline_95: 62.10, gasohol_95: 44.65, gasohol_91: 44.28, gasohol_e20: 39.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '01/05/2569', bcp: { gasoline_95: 62.10, gasohol_91: 40.80, gasohol_e20: 32.24, premium_diesel: 42.93, diesel: 36.30 }, ptt: { gasoline_95: 62.10, gasohol_95: 44.65, gasohol_91: 44.28, gasohol_e20: 39.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '08/05/2569', bcp: { gasoline_95: 61.25, gasohol_91: 39.95, gasohol_e20: 31.39, premium_diesel: 42.08, diesel: 35.45 }, ptt: { gasoline_95: 61.25, gasohol_95: 43.65, gasohol_91: 43.28, gasohol_e20: 38.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '13/05/2569', bcp: { gasoline_95: 61.25, gasohol_91: 40.75, gasohol_e20: 32.29, premium_diesel: 42.98, diesel: 36.35 }, ptt: { gasoline_95: 61.25, gasohol_95: 43.65, gasohol_91: 43.28, gasohol_e20: 38.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '14/05/2569', bcp: { gasoline_95: 61.25, gasohol_91: 41.45, gasohol_e20: 32.99, premium_diesel: 43.68, diesel: 37.05 }, ptt: { gasoline_95: 61.25, gasohol_95: 43.65, gasohol_91: 43.28, gasohol_e20: 38.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '19/05/2569', bcp: { gasoline_95: 61.25, gasohol_91: 42.20, gasohol_e20: 33.84, premium_diesel: 44.53, diesel: 37.90 }, ptt: { gasoline_95: 61.25, gasohol_95: 43.65, gasohol_91: 43.28, gasohol_e20: 38.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '20/05/2569', bcp: { gasoline_95: 61.25, gasohol_91: 42.20, gasohol_e20: 33.84, premium_diesel: 44.53, diesel: 37.90 }, ptt: { gasoline_95: 61.25, gasohol_95: 43.65, gasohol_91: 43.28, gasohol_e20: 38.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '26/05/2569', bcp: { gasoline_95: 60.25, gasohol_91: 41.20, gasohol_e20: 33.84, premium_diesel: 43.93, diesel: 37.90 }, ptt: { gasoline_95: 60.25, gasohol_95: 42.65, gasohol_91: 42.28, gasohol_e20: 37.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '29/05/2569', bcp: { gasoline_95: 59.25, gasohol_91: 41.20, gasohol_e20: 33.84, premium_diesel: 43.23, diesel: 37.90 }, ptt: { gasoline_95: 59.25, gasohol_95: 41.65, gasohol_91: 41.28, gasohol_e20: 36.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '30/05/2569', bcp: { gasoline_95: 58.25, gasohol_91: 40.70, gasohol_e20: 33.84, premium_diesel: 42.53, diesel: 37.90 }, ptt: { gasoline_95: 58.25, gasohol_95: 40.65, gasohol_91: 40.28, gasohol_e20: 35.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '04/06/2569', bcp: { gasoline_95: 58.25, gasohol_91: 41.50, gasohol_e20: 34.24, premium_diesel: 42.93, diesel: 38.30 }, ptt: { gasoline_95: 58.25, gasohol_95: 40.65, gasohol_91: 40.28, gasohol_e20: 35.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '06/06/2569', bcp: { gasoline_95: 57.25, gasohol_91: 40.80, gasohol_e20: 33.54, premium_diesel: 42.23, diesel: 37.60 }, ptt: { gasoline_95: 57.25, gasohol_95: 39.65, gasohol_91: 39.28, gasohol_e20: 34.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '09/06/2569', bcp: { gasoline_95: 57.25, gasohol_91: 41.30, gasohol_e20: 34.04, premium_diesel: 42.73, diesel: 38.10 }, ptt: { gasoline_95: 57.25, gasohol_95: 39.65, gasohol_91: 39.28, gasohol_e20: 34.65, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '11/06/2569', bcp: { gasoline_95: 56.25, gasohol_91: 40.80, gasohol_e20: 34.04, premium_diesel: 42.73, diesel: 38.10 }, ptt: { gasoline_95: 56.25, gasohol_95: 38.85, gasohol_91: 38.48, gasohol_e20: 33.85, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '13/06/2569', bcp: { gasoline_95: 55.25, gasohol_91: 39.80, gasohol_e20: 33.24, premium_diesel: 41.93, diesel: 37.30 }, ptt: { gasoline_95: 55.25, gasohol_95: 38.85, gasohol_91: 38.48, gasohol_e20: 33.85, diesel: 37.50, premium_diesel: 54.25 } },
    { date: '19/06/2569', bcp: { gasoline_95: 55.25, gasohol_91: 39.80, gasohol_e20: 33.24, premium_diesel: 41.93, diesel: 37.30 }, ptt: { gasoline_95: 55.25, gasohol_95: 38.85, gasohol_91: 38.48, gasohol_e20: 33.85, diesel: 37.50, premium_diesel: 54.25 } }
  ];

  var count = 0;
  history.forEach(function(h) {
    var dateStr = h.date;
    // ตรวจสอบว่าวันที่นี้มีอยู่แล้วหรือไม่
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    var exists = data.some(function(r) { return r[0] === dateStr; });
    if (exists) return;

    var row = [dateStr, dateStr];
    stationOrder.forEach(function(station) {
      fuelOrder.forEach(function(fuel) {
        if (h[station] && h[station][fuel]) {
          row.push(h[station][fuel]);
        } else {
          row.push('');
        }
      });
    });
    sheet.appendRow(row);
    count++;
  });

  return { status: 'success', message: 'เพิ่มข้อมูลย้อนหลัง ' + count + ' วัน (ข้ามวันที่มีอยู่แล้ว)' };
}
