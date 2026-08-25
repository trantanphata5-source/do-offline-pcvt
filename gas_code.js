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
var DRIVE_FOLDER_ID = '1CNV2KFYscyVyD5Pxbums3Yof8RaJE9Nd';

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
  'trichYeu', 'doKhan', 'doMat', 'files', 'zipName', 'folderName', 'ngayTaiLen'
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
  
  // Check duplicate: if vanBanId already exists, update instead
  var allData = sheet.getDataRange().getValues();
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][1] === data.vanBanId) {
      return updateChiDao(data); // Redirect to update
    }
  }
  
  var id = 'CD_' + new Date().getTime();
  var timestamp = formatTimestamp(new Date());
  
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

function formatTimestamp(d) {
  var dd = ('0' + d.getDate()).slice(-2);
  var mm = ('0' + (d.getMonth() + 1)).slice(-2);
  var yyyy = d.getFullYear();
  var hh = ('0' + d.getHours()).slice(-2);
  var mi = ('0' + d.getMinutes()).slice(-2);
  return dd + '/' + mm + '/' + yyyy + ' ' + hh + ':' + mi;
}

function updateChiDao(data) {
  var sheet = getOrCreateSheet('ChiDao', CHIDAO_HEADERS);
  var allData = sheet.getDataRange().getValues();
  var timestamp = formatTimestamp(new Date());
  
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
  var timestamp = formatTimestamp(new Date());
  
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
  var folderCount = 0;
  var fileCount = 0;
  
  // === SCAN FILES TRUC TIEP TRONG FOLDER GOC ===
  var rootFiles = parentFolder.getFiles();
  while (rootFiles.hasNext()) {
    var file = rootFiles.next();
    var fileName = file.getName();
    var fileId = file.getId();
    var mimeType = file.getMimeType();
    var previewUrl = 'https://drive.google.com/file/d/' + fileId + '/preview';
    
    // folderName = ten file khong co duoi mo rong (de map voi VanBan.folderName)
    var baseName = fileName.replace(/\.[^.]+$/, '');
    rows.push([baseName, fileName, fileId, mimeType, previewUrl]);
    fileCount++;
  }
  
  // === SCAN CAC SUBFOLDER (tuong thich cu) ===
  var subFolders = parentFolder.getFolders();
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
  
  Logger.log('Scanned ' + folderCount + ' folders, ' + fileCount + ' files');
  return { success: true, folders: folderCount, files: fileCount };
}

function getDriveMapping() {
  var sheet = getOrCreateSheet('DriveFiles', DRIVE_HEADERS);
  var data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return { success: true, data: {} };
  }
  
  // Build mapping: multiple keys -> previewUrl
  var mapping = {};
  for (var i = 1; i < data.length; i++) {
    var folderName = data[i][0];
    var fileName = data[i][1];
    var previewUrl = data[i][4];
    
    // Key 1: folderName/fileName (tuong thich cu - subfolder)
    var key1 = folderName + '/' + fileName;
    mapping[key1] = previewUrl;
    
    // Key 2: fileName truc tiep (cho file upload flat)
    mapping[fileName] = previewUrl;
    
    // Key 3: folderName (basename) truc tiep
    mapping[folderName] = previewUrl;
  }
  
  return { success: true, data: mapping };
}

/**
 * Import 29 van ban tu SOURCE-DOGD.
 * XOA du lieu cu truoc, chi giu 29 VB moi.
 * Chay thu cong tu Apps Script editor.
 */
function importDOGD() {
  var sheet = getOrCreateSheet('VanBan', VANBAN_HEADERS);
  
  // === XOA DU LIEU CU ===
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, VANBAN_HEADERS.length).clear();
    Logger.log('Da xoa ' + (lastRow - 1) + ' dong cu');
  }
  
  var docs = [
    ['1067gm','GM-PCVT','Công ty Điện lực Vũng Tàu','25/08/2026','Giấy mời họp','1067gm.pdf'],
    ['13831_PC07_D2','PC07-Đ2','Công an TP Hồ Chí Minh - Phòng Cảnh sát PCCC và CNCH','25/08/2026','V/v phối hợp tuyên truyền điện gắn với công tác PCCC và CNCH','13831-PC07-Đ2.pdf'],
    ['CV_EVNHCMC_thiet_tri_dau_noi','EVNHCMC-KT','Tổng công ty Điện lực TP Hồ Chí Minh','24/08/2026','V/v tuân thủ thiết trí trong công tác khảo sát, đấu nối nhánh dây mắc điện khách hàng','2026.0824_CV-EVNHCMC_Chuẩn hoá công tác thiết trí-đấu nối nhánh dây MĐ.pdf'],
    ['2129_BQLDA','2129/BQLDA-QLDA','UBND Phường Tam Thắng - Ban QLDA Đầu tư Xây dựng','24/08/2026','V/v đề nghị khẩn trương khôi phục, cấp điện Trạm biến áp phục vụ Trạm xử lý nước thải thuộc dự án Hạ tầng kỹ thuật khu tiểu thủ công nghiệp Phước Thắng phục vụ di dời các cơ sở gây ô nhiễm trên địa bàn TP Vũng Tàu (giai đoạn 1)','24-8-2026 SỐ 2129.pdf'],
    ['QC_BCH_PTDS_EVN','QĐ-EVN','Tập đoàn Điện lực Việt Nam','06/08/2026','Quyết định ban hành Quy chế hoạt động của Ban Chỉ huy Phòng thủ dân sự Tập đoàn Điện lực Việt Nam','26.08.06 Quy che hoat dong BCH PTDS EVN .pdf'],
    ['TB_EVNHCMC_an_toan','TB-EVNHCMC','Tổng công ty Điện lực TP Hồ Chí Minh','24/08/2026','Thông báo nội dung kết luận của Phó Tổng giám đốc Luân Quốc Hưng tại cuộc họp về công tác an toàn lao động','260824_TB_EVNHCMC (an toàn).pdf'],
    ['45_CV_BH','CV-BH','Công ty TNHH DVKT Bách Hạnh Công','24/08/2026','V/v thông báo thay thế thiết bị hệ thống điện mặt trời mái nhà','45.CV BH.pdf'],
    ['BB_hop_thang_09','TB-ĐĐHTĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','25/08/2026','Thông báo nội dung họp kế hoạch bảo dưỡng, sửa chữa các trang thiết bị điện và lưới điện 110/220kV tháng 9/2026','BB hop thang 09-2026.pdf'],
    ['BC_relay_F81','ĐĐHTĐ-ĐĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','25/08/2026','V/v rà soát, đề xuất khóa các mạch sa thải F81 để đảm bảo vận hành cấp điện dịp Lễ Quốc khánh 02/9 năm 2026','Báo cáo rà soát hệ thống relay F81 dịp Lễ 02-9-2026.pdf'],
    ['CV_3025_PHU_MY','CV-UBND','UBND Thị xã Phú Mỹ','25/08/2026','V/v đề nghị phối hợp cung cấp điện','CV 3025 PHU MY de nghi phoi hop cung cap dien.pdf'],
    ['CV_BQLDA_xac_nhan_VT','BQLDA-QLDA','UBND Phường Tam Thắng - Ban QLDA Đầu tư Xây dựng','25/08/2026','V/v xác nhận khối lượng vật tư thu hồi gói thầu số 23: Di dời và xây dựng lưới điện hạ thế thuộc dự án đường Lê Quang Định (từ đường 30/4 đến đường Bình Giã), phường Thắng Nhất, TP Vũng Tàu (giai đoạn 2)','CV gui Công ty ĐL VT.signed.signed.signed.pdf'],
    ['EVNHCMC_moi_SCT_ND243','EVNHCMC','Tổng công ty Điện lực TP Hồ Chí Minh','25/08/2026','V/v mời Sở Công Thương tham dự hội nghị triển khai Nghị định 243/2026/NĐ-CP về cơ chế mua bán điện trực tiếp giữa đơn vị phát điện năng lượng tái tạo và khách hàng sử dụng điện lớn','EVNHCMC- moi SCT tham du hoi nghi ND 243.pdf'],
    ['EVNHCMC_TCNS_NGB_dot2','EVNHCMC-TCNS','Tổng công ty Điện lực TP Hồ Chí Minh - Phòng Tổ chức Nhân sự','25/08/2026','V/v thông qua danh sách thi nâng, giữ bậc lương khối gián tiếp đợt 2 năm 2026','EVNHCMC-TCNS thong qua ds thi NGB dot 2-2026.pdf'],
    ['BC_nhanh_PCAPD_SCT_SNV','EVNHCMC-AT','Tổng công ty Điện lực TP Hồ Chí Minh','25/08/2026','V/v báo cáo nhanh về tai nạn lao động tại Công ty Điện lực An Phú Đông (gửi Sở Công Thương, Sở Nội vụ)','EVNHCMC_ BC nhanh - PC An Phú Đông (Sở Công thương - Sở Nội vụ).pdf'],
    ['BC_nhanh_PCAPD_EVN','EVNHCMC-AT','Tổng công ty Điện lực TP Hồ Chí Minh','25/08/2026','V/v báo cáo nhanh về tai nạn lao động tại Công ty Điện lực An Phú Đông (gửi Tập đoàn Điện lực Việt Nam)','EVNHCMC_ BC nhanh - PC An Phú Đông.pdf'],
    ['HL11_XLSC_474_2408','PA-DVĐL','Công ty Dịch vụ Điện lực TP Hồ Chí Minh','24/08/2026','Phương án thi công đường dây đang mang điện đến 22kV (Hotline) - Vũng Tàu ngày 24/08/2026 (XLSC 474 Long Hương)','HL11 .PATC VT 24-08-26 (XLSC 474 Long hương).pdf'],
    ['HL11_XLSC_476_2408','PA-DVĐL','Công ty Dịch vụ Điện lực TP Hồ Chí Minh','24/08/2026','Phương án thi công đường dây đang mang điện đến 22kV (Hotline) - Vũng Tàu ngày 24/08/2026 (XLSC 476 Điện Biên)','HL11 .PATC VT 24-08-26 (XLSC 476 Điện Biên).pdf'],
    ['HL11_475_Do_luong_2508','PA-DVĐL','Công ty Dịch vụ Điện lực TP Hồ Chí Minh','25/08/2026','Phương án thi công đường dây đang mang điện đến 22kV (Hotline) - Vũng Tàu ngày 25/08/2026 (475 Đô Lương)','HL11 .PATC VT 25-08-26 (475 Đô lương).pdf'],
    ['KH_cat_dien_T09','KH-ĐĐHTĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','25/08/2026','Kế hoạch cắt điện tháng 09/2026','KH cắt điện tháng 09-2026 - ss.pdf'],
    ['PA_CCD_Le_02_9','PA-ĐĐHTĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','25/08/2026','Phương án đảm bảo cung cấp điện phục vụ Lễ Quốc khánh 02/9 năm 2026','PA cung cap dien Le 02-9-2026.pdf'],
    ['PA_CCD_T09','PA-ĐĐHTĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','25/08/2026','Phương án đảm bảo cung cấp điện tháng 09 năm 2026','PA cung cap dien thang 09-2026.pdf'],
    ['PT_Tuan36_A2','ĐĐHTĐ-ĐĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','25/08/2026','V/v Phương thức vận hành hệ thống điện TP.HCM tuần 36 (từ ngày 31/08/2026 đến 06/09/2026)','Phuong thuc Tuan 36-2026 - A2.pdf'],
    ['PA_DMTMN_Tuan35','ĐĐHTĐ-PT','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','25/08/2026','V/v phương án phân bổ công suất tối đa ĐMTMN tuần 35 năm 2026','Phương án huy động công suất tối đa ĐMTMN tuần 35 năm 2026_HC.pdf'],
    ['QD_MyXuanA','QĐ-ĐĐHTĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','24/08/2026','Quyết định về việc đánh số thiết bị trạm biến thế 110kV Mỹ Xuân A','QD-MyXuanA260824.pdf'],
    ['QD_MyXuanB1','QĐ-ĐĐHTĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','24/08/2026','Quyết định về việc đánh số thiết bị trạm biến thế 110kV Mỹ Xuân B1','QD-MyXuanB1_260824.pdf'],
    ['QD_TanHanh','QĐ-ĐĐHTĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','22/08/2026','Quyết định về việc đánh số thiết bị trạm biến thế 110kV Tân Hạnh','QD-TanHanh260822.pdf'],
    ['QD_TanPhuoc','QĐ-ĐĐHTĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','24/08/2026','Quyết định về việc đánh số thiết bị trạm biến thế 110kV Tân Phước','QD-TanPhuoc260824.pdf'],
    ['QD_ThanhBinh','QĐ-ĐĐHTĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','24/08/2026','Quyết định về việc đánh số thiết bị trạm biến thế 110kV Thanh Bình','QD-ThanhBinh260824.pdf'],
    ['TB_PA_CCD_T09','ĐĐHTĐ-ĐĐ','Trung tâm Điều độ Hệ thống Điện TP Hồ Chí Minh','25/08/2026','V/v phổ biến Phương án đảm bảo cung cấp điện tháng 09/2026 và cập nhật Phương án đảm bảo cung cấp điện tháng 10/2026','Thong bao PA cung cap dien Thang 09-2026.pdf'],
    ['TB_hoan_thanh','TB-PCVT','Công ty Điện lực Vũng Tàu','25/08/2026','Thông báo hoàn thành','thong bao hoan thanh.pdf']
  ];
  
  // Build rows: [id, soVanBan, coQuan, ngay, trichYeu, doKhan, doMat, files, zipName, folderName, ngayTaiLen]
  var now = formatTimestamp(new Date());
  var rows = [];
  for (var i = 0; i < docs.length; i++) {
    var d = docs[i];
    rows.push([d[0], d[1], d[2], d[3], d[4], 'Bình thường', 'Bình thường', '["' + d[5] + '"]', '', d[0], now]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, VANBAN_HEADERS.length).setValues(rows);
  }
  
  Logger.log('Da xoa du lieu cu va import ' + rows.length + ' van ban moi');
  return { success: true, count: rows.length };
}
