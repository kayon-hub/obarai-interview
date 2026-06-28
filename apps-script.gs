/**
 * OBARAI INTELLIGENCE — HR 資料 Google Drive 後端
 *
 * 部署方式：
 * 1. 到 https://script.google.com 建立新專案
 * 2. 貼上此程式碼
 * 3. 點選「部署」→「新增部署作業」
 * 4. 選擇類型「網頁應用程式」
 * 5. 執行身分：選「我」
 * 6. 誰可以存取：選「所有人」
 * 7. 部署後複製網址，貼回面試表/入職報到頁面
 *
 * 資料夾結構：
 *   我的雲端硬碟
 *   └── OBARAI-HR
 *       ├── 面試表
 *       │   └── 姓名_2026-06-28
 *       │       ├── 大頭照.jpg
 *       │       └── ...
 *       ├── 入職資料
 *       │   └── 姓名_2026-06-28
 *       │       ├── 身分證正面.jpg
 *       │       ├── 身分證反面.jpg
 *       │       ├── 畢業證書/
 *       │       ├── 證照/
 *       │       └── 存摺/
 *       └── HR紀錄 (Google Sheet — 自動建立)
 */

var ROOT_FOLDER_NAME = 'OBARAI-HR';
var SHEET_NAME = 'HR紀錄';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var root = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);

    if (data.type === 'interview') {
      return handleInterview(data, root);
    } else if (data.type === 'onboarding') {
      return handleOnboarding(data, root);
    }

    return jsonResponse({ success: false, error: '未知的表單類型' });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function handleInterview(data, root) {
  var parentFolder = getOrCreateFolder(root, '面試表');
  var folderName = (data.name || '未具名') + '_' + (data.interviewDate || today());
  var folder = parentFolder.createFolder(folderName);

  // Save photo
  if (data.photo) {
    saveFile(folder, data.photo.data, data.photo.mime, '大頭照_' + getExt(data.photo.name));
  }

  // Signatures
  if (data.sigCandidate && data.sigCandidate.data) {
    saveFile(folder, data.sigCandidate.data, data.sigCandidate.mime, '簽名_面試者.png');
  }
  if (data.sigInterviewer && data.sigInterviewer.data) {
    saveFile(folder, data.sigInterviewer.data, data.sigInterviewer.mime, '簽名_面試官.png');
  }

  // Log to sheet
  var sheet = getOrCreateSheet(root, SHEET_NAME);
  ensureInterviewHeaders(sheet);
  sheet.appendRow([
    new Date(),
    '面試表',
    data.name,
    data.ename,
    data.gender,
    data.dob,
    data.idno,
    data.phone,
    data.email,
    data.position,
    data.salary,
    data.available,
    data.interviewDate,
    data.interviewer,
    data.result,
    JSON.stringify(data.evalScores),
    data.evalNotes,
    folder.getUrl()
  ]);

  // Save form data as JSON
  folder.createFile('form-data.json', JSON.stringify(data, null, 2), 'application/json');

  return jsonResponse({ success: true, folderId: folder.getId(), folderUrl: folder.getUrl() });
}

function handleOnboarding(data, root) {
  var parentFolder = getOrCreateFolder(root, '入職資料');
  var folderName = (data.name || '未具名') + '_' + (data.startDate || today());
  var folder = parentFolder.createFolder(folderName);

  // Signature
  if (data.signature && data.signature.data) {
    saveFile(folder, data.signature.data, data.signature.mime, '簽名.png');
  }

  // ID cards
  if (data.idFront) {
    saveFile(folder, data.idFront.data, data.idFront.mime, '身分證正面' + getExtFromMime(data.idFront.mime));
  }
  if (data.idBack) {
    saveFile(folder, data.idBack.data, data.idBack.mime, '身分證反面' + getExtFromMime(data.idBack.mime));
  }

  // Attachments
  if (data.files) {
    if (data.files.diploma && data.files.diploma.length) {
      var diplomaFolder = folder.createFolder('畢業證書');
      data.files.diploma.forEach(function(f) { saveFile(diplomaFolder, f.data, f.mime, f.name); });
    }
    if (data.files.cert && data.files.cert.length) {
      var certFolder = folder.createFolder('證照');
      data.files.cert.forEach(function(f) { saveFile(certFolder, f.data, f.mime, f.name); });
    }
    if (data.files.bankbook && data.files.bankbook.length) {
      var bankFolder = folder.createFolder('存摺');
      data.files.bankbook.forEach(function(f) { saveFile(bankFolder, f.data, f.mime, f.name); });
    }
  }

  // Log to sheet
  var sheet = getOrCreateSheet(root, SHEET_NAME);
  ensureOnboardingHeaders(sheet);
  sheet.appendRow([
    new Date(),
    '入職資料',
    data.name,
    data.ename,
    data.gender,
    data.dob,
    data.idno,
    data.phone,
    data.email,
    data.startDate,
    data.department,
    data.jobTitle,
    data.empId,
    data.bankName + ' ' + data.bankBranch,
    data.bankAccount,
    data.emerName + '(' + data.emerRelation + ') ' + data.emerPhone,
    data.signature,
    folder.getUrl()
  ]);

  // Save form data as JSON
  folder.createFile('form-data.json', JSON.stringify(data, null, 2), 'application/json');

  return jsonResponse({ success: true, folderId: folder.getId(), folderUrl: folder.getUrl() });
}

// === Utilities ===

function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function getOrCreateSheet(folder, name) {
  var files = folder.getFilesByName(name);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next()).getActiveSheet();
  }
  var ss = SpreadsheetApp.create(name);
  var file = DriveApp.getFileById(ss.getId());
  file.moveTo(folder);
  return ss.getActiveSheet();
}

function ensureInterviewHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '時間戳記', '類型', '姓名', '英文名', '性別', '出生日期',
      '身分證字號', '電話', 'Email', '應徵職位', '期望薪資',
      '可到職日', '面試日期', '面試官', '面試結果',
      '評分(JSON)', '備註', '資料夾連結'
    ]);
    sheet.getRange(1, 1, 1, 18).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  }
}

function ensureOnboardingHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '時間戳記', '類型', '姓名', '英文名', '性別', '出生日期',
      '身分證字號', '電話', 'Email', '到職日', '部門', '職稱',
      '員工編號', '銀行/分行', '帳號', '緊急聯絡人', '簽名', '資料夾連結'
    ]);
    sheet.getRange(1, 1, 1, 18).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.setFrozenRows(1);
  }
}

function saveFile(folder, base64Data, mimeType, fileName) {
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  folder.createFile(blob);
}

function getExt(filename) {
  var parts = filename.split('.');
  return parts.length > 1 ? parts[parts.length - 1] : 'jpg';
}

function getExtFromMime(mime) {
  var map = { 'image/jpeg': '.jpg', 'image/png': '.png', 'image/webp': '.webp', 'image/heic': '.heic', 'application/pdf': '.pdf' };
  return map[mime] || '.jpg';
}

function today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

