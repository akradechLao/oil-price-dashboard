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
