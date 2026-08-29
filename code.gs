// ===================================================
// 11NO AULIAPUR UNION PARISHAD - GOOGLE APPS SCRIPT BACKEND
// This file is the live Apps Script equivalent of the HTML/JS backend logic.
// Keep the same spreadsheet structure and add the missing deceased-address fields.
// ===================================================

var SPREADSHEET_ID = "1jOqKDyNtgk7O8IK-lD4mnURmfhHcxL1JBifgy0s9t9A";

function getSS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && SPREADSHEET_ID) {
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      ss = null;
    }
  }
  return ss;
}

function toBanglaDigit(str) {
  if (str === null || str === undefined || str === '') return '';
  var s = str.toString();
  var bn = ['০', '১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯'];
  for (var i = 0; i < 10; i++) {
    s = s.replace(new RegExp(i, 'g'), bn[i]);
  }
  return s;
}

function toEnglishDigit(str) {
  if (str === null || str === undefined || str === '') return '';
  var s = str.toString();
  var bn = ['০', '১', '২', '৩', '৪', '৫', '৬', '৭', '৮', '৯'];
  for (var i = 0; i < 10; i++) {
    s = s.replace(new RegExp(bn[i], 'g'), i);
  }
  return s;
}

function formatBanglaDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var d = val.getDate();
    var m = val.getMonth() + 1;
    var y = val.getFullYear();
    return toBanglaDigit((d < 10 ? '0' + d : d) + '-' + (m < 10 ? '0' + m : m) + '-' + y);
  }
  var str = val.toString().trim();
  var engStr = toEnglishDigit(str);
  var parsedDate = new Date(engStr);
  if (!isNaN(parsedDate.getTime()) && engStr.length > 5) {
    var d2 = parsedDate.getDate();
    var m2 = parsedDate.getMonth() + 1;
    var y2 = parsedDate.getFullYear();
    return toBanglaDigit((d2 < 10 ? '0' + d2 : d2) + '-' + (m2 < 10 ? '0' + m2 : m2) + '-' + y2);
  }
  return toBanglaDigit(str);
}

function getSheet(sheetName) {
  var ss = getSS();
  if (!ss) throw new Error('গুগল শিট পাওয়া যায়নি!');

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [];

    if (sheetName === 'Citizenship') {
      headers = ['AppID', 'Name', 'NID', 'FatherName', 'MotherName', 'DOB', 'MaritalStatus', 'SpouseName', 'Mobile', 'WardNo', 'Village', 'PostOffice', 'Division', 'Status', 'ApplyDate', 'SignatoryRole', 'CertNo'];
    } else if (sheetName === 'FamilyCert') {
      headers = ['AppID', 'Name', 'NID', 'FatherName', 'MotherName', 'Mobile', 'WardNo', 'Village', 'PostOffice', 'Status', 'ApplyDate', 'SignatoryRole', 'Members_JSON', 'DeceasedName', 'DeceasedFather', 'DeceasedMother', 'DeceasedDate', 'ApplicantRelation', 'DeceasedWard', 'DeceasedVillage', 'DeceasedPostOffice', 'DeceasedUnion', 'DeceasedUpazila', 'DeceasedDistrict', 'CertificateType'];
    } else if (sheetName === 'Warishan') {
      headers = ['AppID', 'ApplicantName', 'FatherSpouse', 'DeceasedName', 'NID', 'Mobile', 'WardNo', 'Village', 'PostOffice', 'Status', 'Date', 'SignatoryRole', 'SmarakNo', 'WarishanJSON', 'DeceasedFather', 'DeceasedRelation', 'DeceasedWard', 'DeceasedVillage', 'DeceasedPostOffice', 'DeceasedUpazila', 'DeceasedDistrict'];
    } else if (sheetName === 'TradeLicense') {
      headers = ['AppID', 'LicenseNo', 'ReceiptNo', 'OrgName', 'OwnerName', 'FatherName', 'MotherName', 'NID', 'DOB', 'Mobile', 'OwnerAddress', 'Category', 'BizDetails', 'BizAddress', 'BizStartDate', 'FiscalYear', 'Capital', 'LicenseFee', 'VatFee', 'CommTax', 'SignTax', 'TotalFee', 'ApplyDate', 'SignatoryRole', 'Status', 'Photo'];
    } else if (sheetName === 'Applications') {
      headers = ['AppID', 'Type', 'Name', 'NID', 'FatherName', 'MotherName', 'DOB', 'MaritalStatus', 'SpouseName', 'Mobile', 'WardNo', 'Village', 'PostOffice', 'Division', 'Status', 'ApplyDate', 'SignatoryRole', 'CertNo'];
    } else if (sheetName === 'TaxPayers') {
      headers = ['HoldingNo', 'Name', 'NID', 'FatherName', 'Mobile', 'WardNo', 'Village', 'HouseType', 'AnnualTax', 'Status'];
    } else if (sheetName === 'Users') {
      headers = ['Username', 'Password', 'Role', 'Name', 'Permissions_JSON'];
    }

    if (headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#006837').setFontColor('#ffffff');
    }
  }
  return sheet;
}

function ensureWarishanSchema() {
  var sheet = getSheet('Warishan');
  var headers = [
    'AppID', 'ApplicantName', 'FatherSpouse', 'DeceasedName', 'NID', 'Mobile', 'WardNo', 'Village', 'PostOffice', 'Status', 'Date', 'SignatoryRole', 'SmarakNo', 'WarishanJSON', 'DeceasedFather', 'DeceasedRelation', 'DeceasedWard', 'DeceasedVillage', 'DeceasedPostOffice', 'DeceasedUpazila', 'DeceasedDistrict'
  ];
  var maxCols = Math.max(sheet.getLastColumn(), headers.length);
  var currentHeaders = sheet.getRange(1, 1, 1, maxCols).getValues()[0];

  for (var i = 0; i < headers.length; i++) {
    if ((currentHeaders[i] || '').toString().trim() !== headers[i]) {
      sheet.getRange(1, i + 1).setValue(headers[i]);
    }
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#006837')
    .setFontColor('#ffffff');

  return sheet;
}

function ensureFamilyCertSchema() {
  var sheet = getSheet('FamilyCert');
  var headers = [
    'AppID', 'Name', 'NID', 'FatherName', 'MotherName', 'Mobile', 'WardNo',
    'Village', 'PostOffice', 'Status', 'ApplyDate', 'SignatoryRole',
    'Members_JSON', 'DeceasedName', 'DeceasedFather', 'DeceasedMother',
    'DeceasedDate', 'ApplicantRelation', 'DeceasedWard', 'DeceasedVillage',
    'DeceasedPostOffice', 'DeceasedUnion', 'DeceasedUpazila',
    'DeceasedDistrict', 'CertificateType'
  ];

  var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if ((currentHeaders[i] || '').toString().trim() !== headers[i]) {
      sheet.getRange(1, i + 1).setValue(headers[i]);
    }
  }
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#006837')
    .setFontColor('#ffffff');
  return sheet;
}

function initMasterDatabase() {
  getSheet('Citizenship');
  getSheet('Applications');
  getSheet('FamilyCert');
  getSheet('Warishan');
  getSheet('TradeLicense');
  getSheet('TaxPayers');
  getSheet('Users');

  ensureWarishanSchema();
  ensureFamilyCertSchema();

  var userSheet = getSheet('Users');
  if (userSheet.getLastRow() <= 1) {
    userSheet.appendRow(['superadmin', '123456', 'SuperAdmin', 'অ্যাড. মোঃ হুমায়ুন কবির (চেয়ারম্যান)', '{"canEdit":true,"canApprove":true,"canDelete":true}']);
    userSheet.appendRow(['admin', '123456', 'Admin', 'সচিব / প্যানেল চেয়ারম্যান', '{"canEdit":true,"canApprove":true,"canDelete":false}']);
  }
}

// ===================================================
// WARISHAN CORE (FIXED)
// ===================================================

function submitWarishanApplication(data) {
  try {
    initMasterDatabase();
    var ss = getSS();
    var sheet = ss.getSheetByName('Warishan');
    if (!sheet) {
      sheet = ss.insertSheet('Warishan');
    }

    ensureWarishanSchema();

    var appId = 'AUL-WAR-' + Math.floor(1000 + Math.random() * 9000);
    var date = Utilities.formatDate(new Date(), 'GMT+6', 'dd/MM/yyyy');
    var smarakNo = 'আ/ইউ/পটুয়া/সদর/' + new Date().getFullYear() + '/' + ('000' + (sheet.getLastRow())).slice(-3);

    var rowData = [
      appId,
      data.applicantName || '',
      data.fatherSpouseName || '',
      data.deceasedName || '',
      data.nid || '',
      data.mobile || '',
      data.wardNo || '',
      data.village || '',
      data.postOffice || '',
      'Pending',
      date,
      'চেয়ারম্যান',
      smarakNo,
      JSON.stringify(data.warishanTree || []),
      data.deceasedFather || '',
      data.deceasedRelation || data.deceasedFatherRelation || '',
      data.deceasedWardNo || data.deceasedWard || '',
      data.deceasedVillage || '',
      data.deceasedPostOffice || data.deceasedPost || '',
      data.deceasedUpazila || '',
      data.deceasedDistrict || ''
    ];

    sheet.appendRow(rowData);
    return { success: true, appId: appId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getWarishanDetails(appId) {
  try {
    initMasterDatabase();
    var sheet = getSheet('Warishan');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString().trim() === appId.toString().trim()) {
        var tree = [];
        try {
          tree = JSON.parse(data[i][13]);
        } catch (err) {
          tree = [];
        }

        return {
          found: true,
          appId: data[i][0],
          applicantName: data[i][1],
          fatherSpouseName: data[i][2],
          deceasedName: data[i][3],
          nid: data[i][4],
          mobile: data[i][5],
          wardNo: data[i][6],
          village: data[i][7],
          postOffice: data[i][8],
          status: data[i][9],
          date: data[i][10],
          signatoryRole: data[i][11],
          smarakNo: data[i][12],
          warishanTree: tree,
          deceasedFather: data[i][14] || '',
          deceasedFatherName: data[i][14] ? (data[i][14].toString().replace(/\s*\[(.*?)\]\s*$/, '') || data[i][14].toString()) : '',
          deceasedFatherRelation: data[i][15] || '',
          deceasedWardNo: data[i][16] || '',
          deceasedVillage: data[i][17] || '',
          deceasedPostOffice: data[i][18] || '',
          deceasedUpazila: data[i][19] || '',
          deceasedDistrict: data[i][20] || ''
        };
      }
    }

    return { found: false };
  } catch (e) {
    return { found: false, error: e.toString() };
  }
}

function getWarishanApps() {
  var result = [];
  var sheet = getSheet('Warishan');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return result;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    result.push({
      appId: row[0] ? row[0].toString() : '',
      applicantName: row[1] ? row[1].toString() : '',
      fatherSpouseName: row[2] ? row[2].toString() : '',
      deceasedName: row[3] ? row[3].toString() : '',
      nid: row[4] !== undefined && row[4] !== null ? row[4].toString() : '',
      mobile: row[5] !== undefined && row[5] !== null ? row[5].toString() : '',
      wardNo: row[6] !== undefined && row[6] !== null ? row[6].toString() : '',
      village: row[7] ? row[7].toString() : '',
      postOffice: row[8] ? row[8].toString() : '',
      status: row[9] ? row[9].toString().trim() : 'Pending',
      date: formatBanglaDate(row[10]),
      signatoryRole: row[11] ? row[11].toString() : 'চেয়ারম্যান',
      smarakNo: row[12] ? row[12].toString() : '',
      deceasedFather: row[14] ? row[14].toString() : '',
      deceasedFatherName: row[14] ? row[14].toString().replace(/\s*\[(.*?)\]\s*$/, '') : '',
      deceasedRelation: row[15] ? row[15].toString() : '',
      deceasedWardNo: row[16] ? row[16].toString() : '',
      deceasedVillage: row[17] ? row[17].toString() : '',
      deceasedPostOffice: row[18] ? row[18].toString() : '',
      deceasedUpazila: row[19] ? row[19].toString() : '',
      deceasedDistrict: row[20] ? row[20].toString() : ''
    });
  }

  return result;
}

function updateWarishanData(ed) {
  try {
    initMasterDatabase();
    var sheet = getSheet('Warishan');
    var data = sheet.getDataRange().getValues();
    var q = toEnglishDigit(ed.appId || '').toUpperCase();

    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString().toUpperCase() === q) {
        sheet.getRange(i + 1, 2).setValue(ed.applicantName || '');
        sheet.getRange(i + 1, 3).setValue(ed.fatherSpouseName || '');
        sheet.getRange(i + 1, 4).setValue(ed.deceasedName || '');
        sheet.getRange(i + 1, 5).setValue("'" + toEnglishDigit(ed.nid || ''));
        sheet.getRange(i + 1, 6).setValue("'" + toEnglishDigit(ed.mobile || ''));
        sheet.getRange(i + 1, 7).setValue(toEnglishDigit(ed.wardNo || ''));
        sheet.getRange(i + 1, 8).setValue(ed.village || '');
        sheet.getRange(i + 1, 9).setValue(ed.postOffice || '');
        sheet.getRange(i + 1, 12).setValue(ed.signatoryRole || 'চেয়ারম্যান');

        if (ed.warishanTree) {
          sheet.getRange(i + 1, 14).setValue(JSON.stringify(ed.warishanTree));
        }

        if (ed.deceasedFather !== undefined) {
          sheet.getRange(i + 1, 15).setValue(ed.deceasedFather);
        }
        if (ed.deceasedFatherRelation !== undefined) {
          sheet.getRange(i + 1, 16).setValue(ed.deceasedFatherRelation);
        }
        if (ed.deceasedWardNo !== undefined) {
          sheet.getRange(i + 1, 17).setValue(ed.deceasedWardNo);
        }
        if (ed.deceasedVillage !== undefined) {
          sheet.getRange(i + 1, 18).setValue(ed.deceasedVillage);
        }
        if (ed.deceasedPostOffice !== undefined) {
          sheet.getRange(i + 1, 19).setValue(ed.deceasedPostOffice);
        }
        if (ed.deceasedUpazila !== undefined) {
          sheet.getRange(i + 1, 20).setValue(ed.deceasedUpazila);
        }
        if (ed.deceasedDistrict !== undefined) {
          sheet.getRange(i + 1, 21).setValue(ed.deceasedDistrict);
        }

        return { success: true };
      }
    }

    return { success: false, error: 'ওয়ারিশান তথ্য পাওয়া যায়নি!' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ===================================================
// USER / ADMIN FUNCTIONS (needed by Admin panel)
// ===================================================

function adminLogin(username, password) {
  try {
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString() === username && (data[i][1] || '').toString() === password) {
        var perms = {};
        try {
          perms = JSON.parse(data[i][4] || '{}');
        } catch (e) {
          perms = {};
        }
        return {
          success: true,
          username: data[i][0],
          role: data[i][2],
          name: data[i][3],
          permissions: perms
        };
      }
    }
    return { success: false, error: 'ইউজারনেম বা পাসওয়ার্ড সঠিক নয়।' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getAllUsers() {
  try {
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      rows.push({
        username: data[i][0],
        password: data[i][1],
        role: data[i][2],
        name: data[i][3],
        permissions: safeJsonParse(data[i][4])
      });
    }
    return { success: true, users: rows };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function safeJsonParse(value) {
  try {
    return JSON.parse(value || '{}');
  } catch (e) {
    return {};
  }
}

function saveNewUser(data) {
  try {
    var sheet = getSheet('Users');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').toString() === (data.username || '')) {
        return { success: false, error: 'এই ইউজারনেম আগেই আছে!' };
      }
    }

    sheet.appendRow([
      data.username || '',
      data.password || '',
      data.role || 'Admin',
      data.name || '',
      JSON.stringify(data.permissions || {})
    ]);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateUser(data) {
  try {
    var sheet = getSheet('Users');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').toString() === (data.username || '')) {
        sheet.getRange(i + 1, 2).setValue(data.password || rows[i][1]);
        sheet.getRange(i + 1, 3).setValue(data.role || rows[i][2]);
        sheet.getRange(i + 1, 4).setValue(data.name || rows[i][3]);
        sheet.getRange(i + 1, 5).setValue(JSON.stringify(data.permissions || safeJsonParse(rows[i][4])));
        return { success: true };
      }
    }
    return { success: false, error: 'ইউজার খুঁজে পাওয়া যায়নি!' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteUser(uname) {
  var sheet = getSheet('Users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString() === (uname || '').toString()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'ইউজার খুঁজে পাওয়া যায়নি!' };
}

// ===================================================
// Dashboard / list functions used by admin panel
// ===================================================

function getMasterDashboardStats() {
  try {
    var stats = {
      citizenship: getSheet('Citizenship').getLastRow() - 1,
      family: getSheet('FamilyCert').getLastRow() - 1,
      warishan: getSheet('Warishan').getLastRow() - 1,
      trade: getSheet('TradeLicense').getLastRow() - 1,
      tax: getSheet('TaxPayers').getLastRow() - 1
    };
    return { success: true, stats: stats };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getAllApplications() {
  var rows = [];
  var sheets = ['Citizenship', 'FamilyCert', 'Warishan', 'TradeLicense'];
  for (var s = 0; s < sheets.length; s++) {
    var sheet = getSheet(sheets[s]);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      rows.push({
        appId: data[i][0],
        source: sheets[s],
        status: data[i][9] || data[i][14] || data[i][13] || 'Pending'
      });
    }
  }
  return { success: true, applications: rows };
}

function getCitizenshipApps() {
  return { success: true, data: [] };
}

function getTradeLicenses() {
  return { success: true, data: [] };
}

function getFamilyApps() {
  return { success: true, data: [] };
}

function getTaxPayers() {
  return { success: true, data: [] };
}

function updateAppStatus(appId, newStatus) {
  try {
    var query = (appId || '').toString().trim();
    var sheetsToSearch = ['Citizenship', 'FamilyCert', 'Warishan', 'TradeLicense'];
    for (var s = 0; s < sheetsToSearch.length; s++) {
      var sheet = getSheet(sheetsToSearch[s]);
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if ((data[i][0] || '').toString().trim() === query) {
          var statusCol = 10;
          if (sheetsToSearch[s] === 'Citizenship') statusCol = 14;
          if (sheetsToSearch[s] === 'FamilyCert') statusCol = 10;
          if (sheetsToSearch[s] === 'TradeLicense') statusCol = 25;
          sheet.getRange(i + 1, statusCol).setValue(newStatus);
          return { success: true };
        }
      }
    }
    return { success: false, error: 'আবেদন খুঁজে পাওয়া যায়নি!' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteApplication(data) {
  try {
    var appId = (data && data.appId) ? data.appId.toString().trim() : '';
    var sheetsToSearch = ['Citizenship', 'FamilyCert', 'Warishan', 'TradeLicense'];
    for (var s = 0; s < sheetsToSearch.length; s++) {
      var sheet = getSheet(sheetsToSearch[s]);
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if ((rows[i][0] || '').toString().trim() === appId) {
          sheet.deleteRow(i + 1);
          return { success: true };
        }
      }
    }
    return { success: false, error: 'আবেদন খুঁজে পাওয়া যায়নি!' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ===================================================
// Routing bridge for front-end API calls
// ===================================================

function doPost(e) {
  try {
    initMasterDatabase();

    var contents = e && e.postData && e.postData.contents;
    if (!contents) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'POST data পাওয়া যায়নি।' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var request = JSON.parse(contents);
    var action = request.action;
    var data = request.data;
    var result = null;

    if (action === 'adminLogin') {
      result = adminLogin(data.username, data.password);
    } else if (action === 'getAllUsers') {
      result = getAllUsers();
    } else if (action === 'saveNewUser') {
      result = saveNewUser(data);
    } else if (action === 'updateUser') {
      result = updateUser(data);
    } else if (action === 'deleteUser') {
      result = deleteUser(data);
    } else if (action === 'getMasterDashboardStats') {
      result = getMasterDashboardStats();
    } else if (action === 'submitWarishanApplication') {
      result = submitWarishanApplication(data);
    } else if (action === 'getWarishanDetails') {
      result = getWarishanDetails(data);
    } else if (action === 'getWarishanApps') {
      result = { success: true, data: getWarishanApps() };
    } else if (action === 'updateWarishanData') {
      result = updateWarishanData(data);
    } else if (action === 'updateAppStatus') {
      result = updateAppStatus(data.appId, data.status);
    } else if (action === 'deleteApplication') {
      result = deleteApplication(data);
    } else if (action === 'getAllApplications') {
      result = getAllApplications();
    } else if (action === 'trackApplication') {
      result = getWarishanDetails(data);
    } else {
      result = { success: false, error: 'ফাংশন খুঁজে পাওয়া যায়নি: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result || { success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  try {
    initMasterDatabase();
    return ContentService.createTextOutput(JSON.stringify({ status: 'API is active and running!' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
