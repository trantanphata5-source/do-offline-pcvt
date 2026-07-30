/**
 * D-Office PCVT - Google Apps Script Backend
 * ============================================
 * Deploy this as a Web App from Google Sheets.
 * 
 * Setup:
 * 1. Mở Google Sheet: https://docs.google.com/spreadsheets/d/14GcrYlzTZ2CzwRFXcHmAGcqDWKVj4c551QmOORTkIN8
 * 2. Extensions → Apps Script
 * 3. Paste toàn bộ code này vào
 * 4. Deploy → Manage deployments → Edit → Version: New version → Deploy
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy URL → paste vào app.js (biến GAS_URL)
 */

var SS_ID = '14GcrYlzTZ2CzwRFXcHmAGcqDWKVj4c551QmOORTkIN8';
var DRIVE_FOLDER_ID = '1x9JJl4rIW8sJxYiXasvvGLLradzpQ0Bl';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SS_ID);
}

function getOrCreateSheet(name, headers) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ============================================================
// SHEET DEFINITIONS
// ============================================================

var VANBAN_HEADERS = [
  'id', 'soVanBan', 'coQuanBanHanh', 'ngayVanBan', 
  'trichYeu', 'doKhan', 'doMat', 'files', 'zipName', 'folderName'
];

var CHIDAO_HEADERS = [
  'id', 'vanBanId', 'noiDung', 'chuTri', 'phoiHop', 'xemDeBiet',
  'hanGiaiQuyet', 'doKhan', 'yeuCauTraLoi', 'chuyenThuKy', 
  'theoDoiVanBan', 'timestamp', 'nguoiChiDao'
];

var CHUYENCHIDAO_HEADERS = [
  'id', 'vanBanId', 'noiDungChuyen', 'chuTri', 'xemDeBiet',
  'timestamp', 'nguoiChuyen'
];

// ============================================================
// INITIALIZE SHEETS (Run manually from menu)
// ============================================================

function initSheets() {
  getOrCreateSheet('VanBan', VANBAN_HEADERS);
  getOrCreateSheet('ChiDao', CHIDAO_HEADERS);
  getOrCreateSheet('ChuyenChiDao', CHUYENCHIDAO_HEADERS);
}

// ============================================================
// API HANDLERS
// ============================================================

function doGet(e) {
  var action = (e && e.parameter) ? e.parameter.action : '';
  var result;
  
  try {
    switch (action) {
      case 'getDocuments':
        result = getDocuments();
        break;
      case 'getChiDao':
        result = getChiDao(e.parameter.vanBanId);
        break;
      case 'getChuyenChiDao':
        result = getChuyenChiDao(e.parameter.vanBanId);
        break;
      case 'getAllChiDao':
        result = getAllChiDao();
        break;
      case 'getAllChuyenChiDao':
        result = getAllChuyenChiDao();
        break;
      case 'getDriveMapping':
        result = getDriveMapping();
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Invalid JSON: ' + err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var action = data.action;
  var result;
  
  try {
    switch (action) {
      case 'saveChiDao':
        result = saveChiDao(data);
        break;
      case 'updateChiDao':
        result = updateChiDao(data);
        break;
      case 'saveChuyenChiDao':
        result = saveChuyenChiDao(data);
        break;
      case 'importDocuments':
        result = importDocuments(data.documents, data.clearFirst);
        break;
      case 'scanDriveFolder':
        result = scanDriveFolder();
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// DATA FUNCTIONS
// ============================================================

function sheetToArray(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    rows.push(obj);
  }
  return rows;
}

function getDocuments() {
  var sheet = getOrCreateSheet('VanBan', VANBAN_HEADERS);
  var docs = sheetToArray(sheet);
  
  // Parse files JSON string back to array
  for (var i = 0; i < docs.length; i++) {
    try {
      if (typeof docs[i].files === 'string' && docs[i].files) {
        docs[i].files = JSON.parse(docs[i].files);
      }
    } catch (ex) {
      docs[i].files = [];
    }
  }
  
  return { success: true, data: docs };
}

function getChiDao(vanBanId) {
  var sheet = getOrCreateSheet('ChiDao', CHIDAO_HEADERS);
  var all = sheetToArray(sheet);
  var filtered = vanBanId ? all.filter(function(r) { return r.vanBanId === vanBanId; }) : all;
  return { success: true, data: filtered };
}

function getAllChiDao() {
  var sheet = getOrCreateSheet('ChiDao', CHIDAO_HEADERS);
  return { success: true, data: sheetToArray(sheet) };
}

function getChuyenChiDao(vanBanId) {
  var sheet = getOrCreateSheet('ChuyenChiDao', CHUYENCHIDAO_HEADERS);
  var all = sheetToArray(sheet);
  var filtered = vanBanId ? all.filter(function(r) { return r.vanBanId === vanBanId; }) : all;
  return { success: true, data: filtered };
}

function getAllChuyenChiDao() {
  var sheet = getOrCreateSheet('ChuyenChiDao', CHUYENCHIDAO_HEADERS);
  return { success: true, data: sheetToArray(sheet) };
}

function importDocuments(docArray, clearFirst) {
  var sheet = getOrCreateSheet('VanBan', VANBAN_HEADERS);
  
  // Only clear if explicitly requested (first batch)
  if (clearFirst) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, VANBAN_HEADERS.length).clear();
    }
  }
  
  var rows = [];
  for (var i = 0; i < docArray.length; i++) {
    var doc = docArray[i];
    rows.push([
      doc.id || '',
      doc.soVanBan || '',
      doc.coQuanBanHanh || '',
      doc.ngayVanBan || '',
      doc.trichYeu || '',
      doc.doKhan || 'Bình thường',
      doc.doMat || 'Bình thường',
      JSON.stringify(doc.files || []),
      doc.zipName || '',
      doc.folderName || ''
    ]);
  }
  
  if (rows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, VANBAN_HEADERS.length).setValues(rows);
  }
  
  return { success: true, count: rows.length };
}

/**
 * Re-import documents from Vercel deployment.
 * Run this manually from Apps Script editor to update Google Sheets
 * with the latest documents.json data.
 */
function reimportFromVercel() {
  var url = 'https://do-offline-pcvt.vercel.app/documents.json';
  var response = UrlFetchApp.fetch(url);
  var docs = JSON.parse(response.getContentText());
  
  Logger.log('Fetched ' + docs.length + ' documents from Vercel');
  
  // Clear existing VanBan data
  var sheet = getOrCreateSheet('VanBan', VANBAN_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, VANBAN_HEADERS.length).clear();
  }
  
  // Import in batches of 100
  var batchSize = 100;
  var totalImported = 0;
  for (var i = 0; i < docs.length; i += batchSize) {
    var batch = docs.slice(i, i + batchSize);
    var rows = [];
    for (var j = 0; j < batch.length; j++) {
      var doc = batch[j];
      rows.push([
        doc.id || '',
        doc.soVanBan || '',
        doc.coQuanBanHanh || '',
        doc.ngayVanBan || '',
        doc.trichYeu || '',
        doc.doKhan || 'Bình thường',
        doc.doMat || 'Bình thường',
        JSON.stringify(doc.files || []),
        doc.zipName || '',
        doc.folderName || ''
      ]);
    }
    if (rows.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, VANBAN_HEADERS.length).setValues(rows);
      totalImported += rows.length;
    }
  }
  
  Logger.log('Imported ' + totalImported + ' documents to VanBan sheet');
  return { success: true, count: totalImported };
}

function saveChiDao(data) {
  var sheet = getOrCreateSheet('ChiDao', CHIDAO_HEADERS);
  var id = 'CD_' + new Date().getTime();
  var timestamp = new Date().toISOString();
  
  var row = [
    id,
    data.vanBanId || '',
    data.noiDung || '',
    data.chuTri || '',
    data.phoiHop || '',
    data.xemDeBiet || '',
    data.hanGiaiQuyet || '',
    data.doKhan || 'Bình thường',
    data.yeuCauTraLoi ? 'Có' : 'Không',
    data.chuyenThuKy ? 'Có' : 'Không',
    data.theoDoiVanBan ? 'Có' : 'Không',
    timestamp,
    data.nguoiChiDao || ''
  ];
  
  sheet.appendRow(row);
  return { success: true, id: id, message: 'Đã lưu chỉ đạo' };
}

function updateChiDao(data) {
  var sheet = getOrCreateSheet('ChiDao', CHIDAO_HEADERS);
  var allData = sheet.getDataRange().getValues();
  var timestamp = new Date().toISOString();
  
  // Find existing row by vanBanId
  var foundRow = -1;
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][1] === data.vanBanId) {
      foundRow = i + 1; // 1-indexed for sheet
      break;
    }
  }
  
  if (foundRow > 0) {
    // Update existing row (keep original ID)
    var existingId = allData[foundRow - 1][0];
    var updatedRow = [
      existingId,
      data.vanBanId || '',
      data.noiDung || '',
      data.chuTri || '',
      data.phoiHop || '',
      data.xemDeBiet || '',
      data.hanGiaiQuyet || '',
      data.doKhan || 'Bình thường',
      data.yeuCauTraLoi ? 'Có' : 'Không',
      data.chuyenThuKy ? 'Có' : 'Không',
      data.theoDoiVanBan ? 'Có' : 'Không',
      timestamp,
      data.nguoiChiDao || ''
    ];
    sheet.getRange(foundRow, 1, 1, CHIDAO_HEADERS.length).setValues([updatedRow]);
    return { success: true, id: existingId, message: 'Đã cập nhật chỉ đạo' };
  } else {
    // Fallback: create new
    return saveChiDao(data);
  }
}

function saveChuyenChiDao(data) {
  var sheet = getOrCreateSheet('ChuyenChiDao', CHUYENCHIDAO_HEADERS);
  var id = 'CCD_' + new Date().getTime();
  var timestamp = new Date().toISOString();
  
  var row = [
    id,
    data.vanBanId || '',
    data.noiDungChuyen || '',
    data.chuTri || '',
    data.xemDeBiet || '',
    timestamp,
    data.nguoiChuyen || ''
  ];
  
  sheet.appendRow(row);
  return { success: true, id: id, message: 'Đã lưu chuyển chỉ đạo' };
}

// ============================================================
// GOOGLE DRIVE FILE MAPPING
// ============================================================

var DRIVE_HEADERS = ['folderName', 'fileName', 'fileId', 'mimeType', 'previewUrl'];

function scanDriveFolder() {
  var parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var sheet = getOrCreateSheet('DriveFiles', DRIVE_HEADERS);
  
  // Clear existing data
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, DRIVE_HEADERS.length).clear();
  }
  
  var rows = [];
  var subFolders = parentFolder.getFolders();
  var folderCount = 0;
  var fileCount = 0;
  
  while (subFolders.hasNext()) {
    var folder = subFolders.next();
    var folderName = folder.getName();
    var files = folder.getFiles();
    folderCount++;
    
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      var fileId = file.getId();
      var mimeType = file.getMimeType();
      var previewUrl = 'https://drive.google.com/file/d/' + fileId + '/preview';
      
      rows.push([folderName, fileName, fileId, mimeType, previewUrl]);
      fileCount++;
    }
  }
  
  // Write all rows at once for performance
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, DRIVE_HEADERS.length).setValues(rows);
  }
  
  return { success: true, folders: folderCount, files: fileCount };
}

function getDriveMapping() {
  var sheet = getOrCreateSheet('DriveFiles', DRIVE_HEADERS);
  var data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return { success: true, data: {} };
  }
  
  // Build mapping: folderName/fileName -> previewUrl
  var mapping = {};
  for (var i = 1; i < data.length; i++) {
    var folderName = data[i][0];
    var fileName = data[i][1];
    var previewUrl = data[i][4];
    var key = folderName + '/' + fileName;
    mapping[key] = previewUrl;
  }
  
  return { success: true, data: mapping };
}
