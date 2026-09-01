// ===================================================
// 11NO AULIAPUR UNION PARISHAD - MASTER BACKEND
// Col A = AppID | Col N = WarishanJSON | Col O = DeceasedFather
// Exact 1-to-1 Database Synchronized Engine
// ===================================================

var SPREADSHEET_ID = "1jOqKDyNtgk7O8IK-lD4mnURmfhHcxL1JBifgy0s9t9A"; 

function getSS() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && SPREADSHEET_ID) {
    try { ss = SpreadsheetApp.openById(SPREADSHEET_ID); } catch(e) {}
  }
  return ss;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 🟢 DIGIT CONVERTERS
function toBanglaDigit(str) {
  if (str === null || str === undefined || str === '') return '';
  var s = str.toString();
  var bn = ['০','১','২','৩','৪','৫','৬','৭','৮','৯'];
  for(var i=0; i<10; i++) s = s.replace(new RegExp(i, 'g'), bn[i]);
  return s;
}

function toEnglishDigit(str) {
  if (str === null || str === undefined || str === '') return '';
  var s = str.toString();
  var bn = ['০','১','২','৩','৪','৫','৬','৭','৮','৯'];
  for(var i=0; i<10; i++) s = s.replace(new RegExp(bn[i], 'g'), i);
  return s;
}

// 🟢 BANGLA DATE FORMATTER
function formatBanglaDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var d = val.getDate(); var m = val.getMonth() + 1; var y = val.getFullYear();
    return (d<10?'0'+d:d) + '/' + (m<10?'0'+m:m) + '/' + y;
  }
  var str = val.toString().trim();
  var engStr = toEnglishDigit(str);
  var parsedDate = new Date(engStr);
  if (!isNaN(parsedDate.getTime()) && engStr.length > 5) {
    var d = parsedDate.getDate(); var m = parsedDate.getMonth() + 1; var y = parsedDate.getFullYear();
    return (d<10?'0'+d:d) + '/' + (m<10?'0'+m:m) + '/' + y;
  }
  return toEnglishDigit(str);
}

// 🟢 ফ্লেক্সিবল শিট ফাইন্ডার (যাতে কোনো শিট মিস না হয়)
function getSheet(sheetName) {
  var ss = getSS();
  if (!ss) throw new Error("গুগল শিট পাওয়া যায়নি!");
  
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    var allSheets = ss.getSheets();
    var cleanTarget = sheetName.toLowerCase().replace(/[^a-z0-9]/g, '');
    for (var i = 0; i < allSheets.length; i++) {
      var sName = allSheets[i].getName().toLowerCase().replace(/[^a-z0-9]/g, '');
      if (sName === cleanTarget) return allSheets[i];
    }
  }

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [];
    if (sheetName === 'Citizenship') {
      headers = ['AppID', 'Name', 'NID', 'FatherName', 'MotherName', 'DOB', 'MaritalStatus', 'SpouseName', 'Mobile', 'WardNo', 'Village', 'PostOffice', 'Division', 'Status', 'ApplyDate', 'SignatoryRole', 'CertNo'];
    } else if (sheetName === 'FamilyCert') {
      headers = ['AppID','Name','NID','FatherName','MotherName','Mobile','WardNo','Village',
      'PostOffice','Status','ApplyDate','SignatoryRole','Members_JSON','DeceasedName','DeceasedFather','DeceasedMother','DeceasedDate','ApplicantRelation','DeceasedWard','DeceasedVillage','DeceasedPostOffice','DeceasedUnion','DeceasedUpazila','DeceasedDistrict','CertificateType'];
    } else if (sheetName === 'Warishan') {
      headers = ['AppID', 'ApplicantName', 'FatherSpouse', 'DeceasedName', 'NID', 'Mobile', 'WardNo', 'Village', 'PostOffice', 'Status', 'Date', 'SignatoryRole', 'SmarakNo', 'WarishanJSON', 'DeceasedFather', 'DeceasedRelation', 'DeceasedWard', 'DeceasedVillage', 'DeceasedPostOffice', 'DeceasedUpazila', 'DeceasedDistrict'];
    } else if (sheetName === 'TradeLicense') {
      headers = ['AppID', 'LicenseNo', 'ReceiptNo', 'OrgName', 'OwnerName', 'FatherName', 'MotherName', 'NID', 'DOB', 'Mobile', 'OwnerAddress', 'Category', 'BizDetails', 'BizAddress', 'BizStartDate', 'FiscalYear', 'Capital', 'LicenseFee', 'VatFee', 'CommTax', 'SignTax', 'TotalFee', 'ApplyDate', 'SignatoryRole', 'Status', 'Photo'];
    } else if (sheetName === 'Applications') {
      headers = ['AppID', 'Type', 'Name', 'NID', 'FatherName', 'MotherName', 'DOB', 'MaritalStatus', 'SpouseName', 'Mobile', 'WardNo', 'Village', 'PostOffice', 'Division', 'Status', 'ApplyDate', 'SignatoryRole', 'CertNo'];
    } else if (sheetName === 'TaxPayers') {
      headers = ['HoldingNo', 'Name', 'NID', 'FatherName', 'Mobile', 'WardNo', 'Village', 'HouseType', 'AnnualTax', 'Status'];
    } else if (sheetName === 'Users') {
      headers = ['Username', 'Password', 'Role', 'Name', 'Permition', 'Photos'];
    }
    if (headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#006837').setFontColor('#ffffff');
    }
  }
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
  ensureSearchIndexSheet();
  syncSearchIndexFromSheets();
}

function ensureSearchIndexSheet() {
  var sheet = getSheet('SearchIndex');
  var headers = ['AppID', 'SourceSheet', 'ServiceType', 'ApplicantName', 'Mobile', 'NID', 'Status', 'ApplyDate', 'SearchText', 'PayloadJSON', 'CreatedAt'];
  var rows = sheet.getDataRange().getValues();

  if (rows.length === 0 || rows[0].length === 0 || (rows[0][0] || '').toString().trim() !== 'AppID') {
    sheet.clear();
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0d6efd').setFontColor('#ffffff');
  }

  return sheet;
}

function normalizeSearchText(value) {
  var v = (value === null || value === undefined) ? '' : value.toString();
  return toEnglishDigit(v).replace(/[^a-z0-9]/gi, '').toLowerCase();
}

function upsertSearchIndexRecord(record) {
  if (!record || !record.appId) return;
  var indexSheet = ensureSearchIndexSheet();
  var rows = indexSheet.getDataRange().getValues();
  var appId = (record.appId || '').toString().trim();
  var rowValues = [
    appId,
    record.sourceSheet || '',
    record.serviceType || '',
    record.applicantName || '',
    record.mobile || '',
    record.nid || '',
    record.status || 'Pending',
    record.applyDate || '',
    normalizeSearchText([appId, record.applicantName, record.mobile, record.nid, record.serviceType].join(' ')),
    JSON.stringify(record.payload || {}),
    new Date().toISOString()
  ];

  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString().trim() === appId) {
      indexSheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);
      return;
    }
  }

  indexSheet.appendRow(rowValues);
}

function deleteSearchIndexRecord(appId) {
  var raw = (appId || '').toString().trim();
  if (!raw) return;
  var indexSheet = ensureSearchIndexSheet();
  var rows = indexSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString().trim() === raw) {
      indexSheet.deleteRow(i + 1);
      return;
    }
  }
}

function addSearchIndexFromServiceSheet(sheetName, rowData) {
  if (!rowData || !rowData.appId) return;
  var appId = (rowData.appId || '').toString().trim();
  var serviceType = rowData.serviceType || rowData.type || sheetName;
  var applicantName = rowData.applicantName || rowData.name || rowData.ownerName || rowData.applicantName || '';
  var mobile = toEnglishDigit(rowData.mobile || rowData.Mobile || '');
  var nid = toEnglishDigit(rowData.nid || rowData.NID || '');
  var status = rowData.status || 'Pending';
  var applyDate = rowData.applyDate || rowData.date || '';

  upsertSearchIndexRecord({
    appId: appId,
    sourceSheet: sheetName,
    serviceType: serviceType,
    applicantName: applicantName,
    mobile: mobile,
    nid: nid,
    status: status,
    applyDate: applyDate,
    payload: rowData
  });
}

function registerSearchIndexRecordFromSubmission(sheetName, record) {
  if (!record || !record.appId) return;

  var payload = record.payload || record;
  upsertSearchIndexRecord({
    appId: record.appId,
    sourceSheet: sheetName,
    serviceType: record.serviceType || record.type || sheetName,
    applicantName: record.applicantName || record.name || record.ownerName || '',
    mobile: toEnglishDigit(record.mobile || record.Mobile || ''),
    nid: toEnglishDigit(record.nid || record.NID || ''),
    status: record.status || 'Pending',
    applyDate: record.applyDate || record.date || '',
    payload: payload
  });
}

function syncSearchIndexFromSheets() {
  var sheets = ['Citizenship', 'FamilyCert', 'Warishan', 'TradeLicense', 'Applications'];
  for (var s = 0; s < sheets.length; s++) {
    var sheet = getSheet(sheets[s]);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || !row[0]) continue;
      var appId = (row[0] || '').toString().trim();
      if (!appId) continue;

      var item = {};
      if (sheets[s] === 'Citizenship') {
        item = {
          appId: appId,
          serviceType: 'নাগরিকত্ব সনদ',
          applicantName: row[1] || '',
          mobile: toEnglishDigit(row[8] || ''),
          nid: toEnglishDigit(row[2] || ''),
          status: row[13] || 'Pending',
          applyDate: formatBanglaDate(row[14]),
          type: 'নাগরিকত্ব সনদ'
        };
      } else if (sheets[s] === 'Warishan') {
        item = {
          appId: appId,
          serviceType: 'ওয়ারিশান সনদ',
          applicantName: row[1] || '',
          mobile: toEnglishDigit(row[5] || ''),
          nid: toEnglishDigit(row[4] || ''),
          status: row[9] || 'Pending',
          applyDate: formatBanglaDate(row[10]),
          type: 'ওয়ারিশান সনদ'
        };
      } else if (sheets[s] === 'FamilyCert') {
        item = {
          appId: appId,
          serviceType: (appId.indexOf('AUL-UW-') === 0) ? 'উত্তরাধিকারী সনদ' : (row[24] || 'পারিবারিক সনদ'),
          applicantName: row[1] || '',
          mobile: toEnglishDigit(row[5] || ''),
          nid: toEnglishDigit(row[2] || ''),
          status: row[9] || 'Pending',
          applyDate: formatBanglaDate(row[10]),
          type: (appId.indexOf('AUL-UW-') === 0) ? 'উত্তরাধিকারী সনদ' : (row[24] || 'পারিবারিক সনদ')
        };
      } else if (sheets[s] === 'TradeLicense') {
        item = {
          appId: appId,
          serviceType: 'ট্রেড লাইসেন্স',
          applicantName: row[4] || '',
          mobile: toEnglishDigit(row[9] || ''),
          nid: toEnglishDigit(row[7] || ''),
          status: row[24] || 'Pending',
          applyDate: formatBanglaDate(row[22]),
          type: 'ট্রেড লাইসেন্স'
        };
      } else {
        item = {
          appId: appId,
          serviceType: row[1] || 'সাধারণ প্রত্যয়ন',
          applicantName: row[2] || '',
          mobile: toEnglishDigit(row[9] || row[8] || ''),
          nid: toEnglishDigit(row[3] || ''),
          status: row[14] || row[13] || 'Pending',
          applyDate: formatBanglaDate(row[15] || row[14]),
          type: row[1] || 'সাধারণ প্রত্যয়ন'
        };
      }

      addSearchIndexFromServiceSheet(sheets[s], item);
    }
  }
}

function normalizeFamilyCertificateType(data) {
  var value = String(data && (data.type || data.serviceType || data.certificateType) || '').trim();
  var prefix = String(data && data.appPrefix || '').trim().toUpperCase();
  if (value === 'উত্তরাধিকারী সনদ' || value === 'উত্তরাধিকারী সনদপত্র' || prefix === 'AUL-UW-') {
    return 'উত্তরাধিকারী সনদ';
  }
  return 'পারিবারিক সনদ';
}

function generateNextSmarakNo() {
  var sheet = getSheet('Warishan');
  var count = sheet.getLastRow();
  var currentYear = new Date().getFullYear();
  var paddedCount = ("000" + count).slice(-3);
  return 'আ/ইউ/পটুয়া/সদর/' + toBanglaDigit(currentYear) + '/' + toBanglaDigit(paddedCount);
}

// 🟢 WARISHAN SUBMISSION
function submitWarishanApplication(data) {
  try {
    var sheet = getSheet('Warishan');
    var appId = "AUL-WAR-" + Math.floor(1000 + Math.random() * 9000);
    var date = Utilities.formatDate(new Date(), "GMT+6", "dd/MM/yyyy");
    var smarakNo = "আ/ইউ/পটুয়া/সদর/" + new Date().getFullYear() + "/" + ("000" + (sheet.getLastRow())).slice(-3);
    
    var rowData = [
      appId,
      data.applicantName || "",
      data.fatherSpouseName || "",
      data.deceasedName || "",
      data.nid || "",
      data.mobile || "",
      data.wardNo || "",
      data.village || "",
      data.postOffice || "আউলিয়াপুর ময়দান",
      "Pending",
      date,
      "চেয়ারম্যান",
      smarakNo,
      JSON.stringify(data.warishanTree || []),
      data.deceasedFather || "",
      data.deceasedRelation || data.deceasedFatherRelation || "",
      data.deceasedWardNo || data.deceasedWard || "",
      data.deceasedVillage || "",
      data.deceasedPostOffice || data.deceasedPost || "",
      data.deceasedUnion || "১১নং আউলিয়াপুর ইউনিয়ন",
      data.deceasedUpazila || "",
      data.deceasedDistrict || ""
    ];
    
    sheet.appendRow(rowData);
    registerSearchIndexRecordFromSubmission('Warishan', {
      appId: appId,
      serviceType: 'ওয়ারিশান সনদ',
      applicantName: data.applicantName || '',
      mobile: data.mobile || '',
      nid: data.nid || '',
      status: 'Pending',
      applyDate: date,
      date: date,
      payload: rowData
    });
    return { success: true, appId: appId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getWarishanDetails(appId) {
  try {
    var sheet = getSheet('Warishan');
    var data = sheet.getDataRange().getValues();
    var q = toEnglishDigit(appId || '').toUpperCase();
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString().toUpperCase() === q) {
        var tree = [];
        try { tree = JSON.parse(data[i][13]); } catch(err) { tree = []; }
        return {
          found: true,
          appId: data[i][0],
          applicantName: data[i][1],
          fatherSpouseName: data[i][2],
          deceasedName: data[i][3],
          nid: toEnglishDigit(data[i][4] || ''),
          mobile: toEnglishDigit(data[i][5] || ''),
          wardNo: toEnglishDigit(data[i][6] || ''),
          village: data[i][7],
          postOffice: data[i][8],
          status: data[i][9],
          date: formatBanglaDate(data[i][10]),
          signatoryRole: data[i][11],
          smarakNo: data[i][12],
          warishanTree: tree,
          deceasedFather: data[i][14] || "",
          deceasedWardNo: data[i][16] || "",
          deceasedVillage: data[i][17] || "",
          deceasedPostOffice: data[i][18] || "",
          deceasedUnion: data[i][19] || '১১নং আউলিয়াপুর ইউনিয়ন',
          deceasedUpazila: data[i][20] || "",
          deceasedDistrict: data[i][21] || ""
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
      appId: row[0].toString(),
      applicantName: row[1] || '',
      fatherSpouseName: row[2] || '',
      deceasedName: row[3] || '',
      nid: toEnglishDigit(row[4] || ''),
      mobile: toEnglishDigit(row[5] || ''),
      wardNo: toEnglishDigit(row[6] || ''),
      village: row[7] || '',
      postOffice: row[8] || 'আউলিয়াপুর ময়দান',
      status: row[9] ? row[9].toString().trim() : 'Pending',
      date: formatBanglaDate(row[10]),
      signatoryRole: row[11] || 'চেয়ারম্যান',
      smarakNo: row[12] || '',
      deceasedFather: row[14] || '',
      deceasedWardNo: row[16] || '',
      deceasedVillage: row[17] || '',
      deceasedPostOffice: row[18] || '',
      deceasedUnion: row[19] || '১১নং আউলিয়াপুর ইউনিয়ন',
      deceasedUpazila: row[20] || '',
      deceasedDistrict: row[21] || ''
    });
  }
  return result;
}

function updateAppStatus(appId, newStatus) {
  var rawQuery = (appId || "").toString().trim();
  var queryEng = toEnglishDigit(rawQuery).toUpperCase();
  var sheetsToSearch = ['Citizenship', 'Applications', 'FamilyCert', 'Warishan', 'TradeLicense'];
  
  for (var s = 0; s < sheetsToSearch.length; s++) {
    var sheet = getSheet(sheetsToSearch[s]);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowA = (data[i][0] || "").toString().trim();
      var rowAEng = toEnglishDigit(rowA).toUpperCase();
      if (rowAEng === queryEng || rowAEng.indexOf(queryEng) > -1) {
        var statusCol = (sheetsToSearch[s] === 'Citizenship') ? 14 : 
                        (sheetsToSearch[s] === 'FamilyCert') ? 10 : 
                        (sheetsToSearch[s] === 'Warishan') ? 10 : 
                        (sheetsToSearch[s] === 'TradeLicense' ? 25 : 15);
        sheet.getRange(i + 1, statusCol).setValue(newStatus);
        syncSearchIndexFromSheets();
        return { success: true };
      }
    }
  }
  return { success: false };
}

function updateWarishanData(ed) {
  try {
    var sheet = getSheet('Warishan');
    var data = sheet.getDataRange().getValues();
    var q = toEnglishDigit(ed.appId || '').toUpperCase();

    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString().toUpperCase() === q) {
        sheet.getRange(i + 1, 2).setValue(ed.applicantName);
        sheet.getRange(i + 1, 3).setValue(ed.fatherSpouseName || '');
        sheet.getRange(i + 1, 4).setValue(ed.deceasedName);
        sheet.getRange(i + 1, 5).setValue("'" + toEnglishDigit(ed.nid || ''));
        sheet.getRange(i + 1, 6).setValue("'" + toEnglishDigit(ed.mobile));
        sheet.getRange(i + 1, 7).setValue(toEnglishDigit(ed.wardNo));
        sheet.getRange(i + 1, 8).setValue(ed.village);
        sheet.getRange(i + 1, 9).setValue(ed.postOffice);
        sheet.getRange(i + 1, 12).setValue(ed.signatoryRole || 'চেয়ারম্যান');
        if (ed.warishanTree) sheet.getRange(i + 1, 14).setValue(JSON.stringify(ed.warishanTree));
        if (ed.deceasedFather !== undefined) sheet.getRange(i + 1, 15).setValue(ed.deceasedFather);
        if (ed.deceasedFatherRelation !== undefined) sheet.getRange(i + 1, 16).setValue(ed.deceasedFatherRelation);
        if (ed.deceasedWardNo !== undefined) sheet.getRange(i + 1, 17).setValue(ed.deceasedWardNo);
        if (ed.deceasedVillage !== undefined) sheet.getRange(i + 1, 18).setValue(ed.deceasedVillage);
        if (ed.deceasedPostOffice !== undefined) sheet.getRange(i + 1, 19).setValue(ed.deceasedPostOffice);
        if (ed.deceasedUnion !== undefined) sheet.getRange(i + 1, 20).setValue(ed.deceasedUnion);
        if (ed.deceasedUpazila !== undefined) sheet.getRange(i + 1, 21).setValue(ed.deceasedUpazila);
        if (ed.deceasedDistrict !== undefined) sheet.getRange(i + 1, 22).setValue(ed.deceasedDistrict);
        return { success: true };
      }
    }
    return { success: false, error: 'ওয়ারিশান তথ্য পাওয়া যায়নি!' };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function updateApplicationData(ed) {
  try {
    var rawQuery = (ed.appId || "").toString().trim();
    var queryEng = toEnglishDigit(rawQuery).toUpperCase();
    var sheetsToSearch = ['Citizenship', 'FamilyCert', 'Applications'];
    
    for (var s = 0; s < sheetsToSearch.length; s++) {
      var sheet = getSheet(sheetsToSearch[s]);
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var rowA = (data[i][0] || "").toString().trim();
        var rowAEng = toEnglishDigit(rowA).toUpperCase();
        if (rowAEng === queryEng || rowAEng.indexOf(queryEng) > -1) {
          if (sheetsToSearch[s] === 'Citizenship') {
            sheet.getRange(i + 1, 2).setValue(ed.name);
            sheet.getRange(i + 1, 3).setValue("'" + toEnglishDigit(ed.nid || ''));
            sheet.getRange(i + 1, 4).setValue(ed.fatherName);
            sheet.getRange(i + 1, 9).setValue("'" + toEnglishDigit(ed.mobile));
            sheet.getRange(i + 1, 10).setValue(toEnglishDigit(ed.wardNo));
            sheet.getRange(i + 1, 11).setValue(ed.village);
            sheet.getRange(i + 1, 12).setValue(ed.postOffice);
            sheet.getRange(i + 1, 16).setValue(ed.signatoryRole || 'চেয়ারম্যান');
          } else if (sheetsToSearch[s] === 'FamilyCert') {
            var editType = ed.type || ed.serviceType || ed.certificateType || 'পারিবারিক সনদ';
            sheet.getRange(i + 1, 2).setValue(ed.name || '');
            sheet.getRange(i + 1, 3).setValue("'" + toEnglishDigit(ed.nid || ''));
            sheet.getRange(i + 1, 4).setValue(ed.fatherName || '');
            sheet.getRange(i + 1, 5).setValue(ed.motherName || '');
            sheet.getRange(i + 1, 6).setValue("'" + toEnglishDigit(ed.mobile || ''));
            sheet.getRange(i + 1, 7).setValue(toEnglishDigit(ed.wardNo || ''));
            sheet.getRange(i + 1, 8).setValue(ed.village || '');
            sheet.getRange(i + 1, 9).setValue(ed.postOffice || '');
            sheet.getRange(i + 1, 12).setValue(ed.signatoryRole || 'চেয়ারম্যান');
            sheet.getRange(i + 1, 13).setValue(JSON.stringify(ed.members || []));
            sheet.getRange(i + 1, 25).setValue(editType);
          } else if (sheetsToSearch[s] === 'Applications') {
            sheet.getRange(i + 1, 3).setValue(ed.name);
            sheet.getRange(i + 1, 4).setValue("'" + toEnglishDigit(ed.nid || ''));
            sheet.getRange(i + 1, 5).setValue(ed.fatherName);
            sheet.getRange(i + 1, 10).setValue("'" + toEnglishDigit(ed.mobile));
            sheet.getRange(i + 1, 11).setValue(toEnglishDigit(ed.wardNo));
            sheet.getRange(i + 1, 12).setValue(ed.village);
            sheet.getRange(i + 1, 13).setValue(ed.postOffice);
            sheet.getRange(i + 1, 17).setValue(ed.signatoryRole || 'চেয়ারম্যান');
          }
          return { success: true };
        }
      }
    }
    return { success: false, error: 'তথ্য পাওয়া যায়নি!' };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// 🔍 🟢 সম্পূর্ণ শক্তিশালী মাল্টি-ফিল্ড ইউনিভার্সাল ট্র্যাকিং ইঞ্জিন (Mobile, NID, Name, AppID)
// 🔍 🟢 সম্পূর্ণ শক্তিশালী ইউনিভার্সাল মাল্টি-ফিল্ড ট্র্যাকিং ইঞ্জিন (Mobile, NID, Name, AppID)
function trackApplication(searchKey) {
  try {
    function normalizeStatusName(value) {
      var status = (value === undefined || value === null) ? 'Pending' : value.toString().trim();
      var lowered = status.toLowerCase();

      if (lowered.indexOf('approved') !== -1 || lowered.indexOf('approve') !== -1 || lowered.indexOf('accept') !== -1 || lowered.indexOf('verified') !== -1 || lowered.indexOf('অনুমোদিত') !== -1 || lowered.indexOf('অনুমতি') !== -1) {
        return 'Approved';
      }
      if (lowered.indexOf('rejected') !== -1 || lowered.indexOf('reject') !== -1 || lowered.indexOf('বাতিল') !== -1 || lowered.indexOf('রিজেক্ট') !== -1 || lowered.indexOf('cancel') !== -1) {
        return 'Rejected';
      }
      return 'Pending';
    }

    function buildSearchVariants(value) {
      var variants = [];
      if (value === undefined || value === null) return variants;
      var raw = value.toString().trim();
      if (!raw) return variants;

      var rawEnglish = toEnglishDigit(raw);
      var values = [
        raw,
        raw.toLowerCase(),
        raw.toUpperCase(),
        rawEnglish,
        rawEnglish.toLowerCase(),
        rawEnglish.toUpperCase()
      ];

      for (var i = 0; i < values.length; i++) {
        var v = values[i];
        if (v && variants.indexOf(v) === -1) variants.push(v);
      }

      var digitsOnly = rawEnglish.replace(/\D/g, '');
      if (digitsOnly) {
        variants.push(digitsOnly);
        if (digitsOnly.length >= 10) variants.push(digitsOnly.slice(-10));
        if (digitsOnly.length >= 9) variants.push(digitsOnly.slice(-9));
        if (digitsOnly.length >= 8) variants.push(digitsOnly.slice(-8));
      }

      return variants;
    }

    function getSheetMatchColumns(sheetName) {
      var name = (sheetName || '').toLowerCase();
      if (name.indexOf('citizen') > -1) {
        return [0, 1, 2, 3, 8, 16, 9, 10, 11, 12, 13, 14, 15];
      }
      if (name.indexOf('family') > -1) {
        return [0, 1, 2, 3, 5, 9, 10, 13, 24];
      }
      if (name.indexOf('warishan') > -1) {
        return [0, 1, 2, 4, 5, 9, 10];
      }
      if (name.indexOf('trade') > -1) {
        return [0, 1, 4, 5, 7, 9, 24];
      }
      return [0, 1, 2, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14, 15, 16];
    }

    function rowMatchesQuery(row, queryVariants, sheetName) {
      if (!row || !row.length || !row[0]) return false;

      var matchColumns = getSheetMatchColumns(sheetName);
      var cellTexts = [];
      var rowText = '';

      for (var c = 0; c < row.length; c++) {
        var cell = row[c];
        if (cell !== null && cell !== undefined && cell !== '') {
          var cellText = cell.toString().trim();
          if (!cellText) continue;
          rowText += ' ' + cellText.toLowerCase();
          if (matchColumns.indexOf(c) > -1) {
            cellTexts.push(cellText);
            cellTexts.push(toEnglishDigit(cellText));
          }
        }
      }

      if (cellTexts.length === 0) return false;

      for (var i = 0; i < queryVariants.length; i++) {
        var candidate = (queryVariants[i] || '').toString().trim();
        if (!candidate) continue;

        var candUpper = toEnglishDigit(candidate).toUpperCase();
        var candLower = candidate.toLowerCase();
        var candDigits = candUpper.replace(/\D/g, '');

        for (var j = 0; j < cellTexts.length; j++) {
          var cellVal = (cellTexts[j] || '').toString();
          var cellUpper = toEnglishDigit(cellVal).toUpperCase();
          var cellLower = cellVal.toLowerCase();
          var cellDigits = cellUpper.replace(/\D/g, '');

          if (cellUpper.indexOf(candUpper) > -1 || cellLower.indexOf(candLower) > -1) {
            return true;
          }

          if (candDigits && cellDigits) {
            if (cellDigits.indexOf(candDigits) > -1 || candDigits.indexOf(cellDigits) > -1 || cellDigits.indexOf(candDigits.slice(-9)) > -1 || candDigits.indexOf(cellDigits.slice(-9)) > -1) {
              return true;
            }
          }
        }
      }

      var queryString = (queryVariants[0] || '').toString().trim().toLowerCase();
      if (queryString.length >= 3 && rowText.indexOf(queryString) > -1) {
        return true;
      }

      return false;
    }

    var rawQuery = "";
    if (typeof searchKey === 'object' && searchKey !== null) {
      rawQuery = searchKey.query || searchKey.appId || searchKey.searchKey || searchKey.mobile || searchKey.nid || searchKey.name || "";
    } else {
      rawQuery = (searchKey || "").toString();
    }
    rawQuery = rawQuery.trim();
    if (!rawQuery) return { found: false };

    var searchVariants = buildSearchVariants(rawQuery);
    if (!searchVariants.length) return { found: false };

    var results = [];
    var seenAppIds = {};

    if (getSheet('SearchIndex').getLastRow() <= 1) {
      syncSearchIndexFromSheets();
    }

    var indexSheet = ensureSearchIndexSheet();
    var indexRows = indexSheet.getDataRange().getValues();
    if (indexRows.length > 1) {
      for (var idx = 1; idx < indexRows.length; idx++) {
        var indexRow = indexRows[idx];
        if (!indexRow || !indexRow[0]) continue;
        var rowText = '';
        for (var c = 0; c < indexRow.length; c++) {
          if (indexRow[c] !== null && indexRow[c] !== undefined && indexRow[c] !== '') {
            rowText += ' ' + indexRow[c].toString().toLowerCase();
          }
        }
        var matched = false;
        for (var q = 0; q < searchVariants.length; q++) {
          var candidate = (searchVariants[q] || '').toString().trim();
          if (!candidate) continue;
          var candLower = candidate.toLowerCase();
          var candDigits = toEnglishDigit(candidate).replace(/\D/g, '');
          if (rowText.indexOf(candLower) > -1) {
            matched = true;
            break;
          }
          if (candDigits && rowText.replace(/\D/g, '').indexOf(candDigits) > -1) {
            matched = true;
            break;
          }
        }

        if (!matched) continue;

        var appId = (indexRow[0] || '').toString().trim();
        if (!appId || seenAppIds[appId]) continue;
        seenAppIds[appId] = true;

        results.push({
          appId: appId,
          sourceSheet: indexRow[1] || '',
          type: indexRow[2] || 'সাধারণ প্রত্যয়ন',
          serviceType: indexRow[2] || 'সাধারণ প্রত্যয়ন',
          applicantName: indexRow[3] || '',
          name: indexRow[3] || '',
          mobile: toEnglishDigit(indexRow[4] || ''),
          nid: toEnglishDigit(indexRow[5] || ''),
          status: indexRow[6] || 'Pending',
          applyDate: indexRow[7] || '',
          date: indexRow[7] || ''
        });
      }
    }

    if (results.length > 0) {
      var resObj = { found: true, total: results.length, list: results };
      for (var k in results[0]) {
        if (!resObj.hasOwnProperty(k)) {
          resObj[k] = results[0][k];
        }
      }
      return resObj;
    }

    var ss = getSS();
    var allSheets = ss.getSheets();
    for (var s = 0; s < allSheets.length; s++) {
      var curSheet = allSheets[s];
      var sName = curSheet.getName().toLowerCase().replace(/[^a-z0-9]/g, '');

      if (sName === 'users' || sName === 'accounts' || sName === 'projects' || sName === 'beneficiaries' || sName === 'taxpayers' || sName === 'taxcollection' || sName === 'searchindex') {
        continue;
      }

      var sheetData = curSheet.getDataRange().getValues();
      for (var r = 1; r < sheetData.length; r++) {
        var row = sheetData[r];
        if (!row || !row[0]) continue;
        if (!rowMatchesQuery(row, searchVariants, curSheet.getName())) continue;

        var appId = row[0].toString().trim();
        if (!appId || seenAppIds[appId]) continue;
        seenAppIds[appId] = true;

        var serviceType = 'প্রত্যয়নপত্র';
        var licNo = '';
        var certNo = '';
        var appName = row[1] || '';
        var father = '';
        var nidVal = '';
        var mobVal = '';
        var applyDateVal = '';
        var statusVal = 'Pending';

        if (sName.indexOf('trade') > -1) {
          var isRenew = appId.indexOf('AUL-RN-') === 0;
          serviceType = isRenew ? 'নবায়নকৃত ট্রেড লাইসেন্স' : 'ট্রেড লাইসেন্স';
          licNo = toEnglishDigit(row[1] || '');
          appName = (row[4] || '') + (row[3] ? ' (' + row[3] + ')' : '');
          father = row[5] || '';
          nidVal = toEnglishDigit(row[7] || '');
          mobVal = toEnglishDigit(row[9] || '');
          applyDateVal = formatBanglaDate(row[22]);
          statusVal = normalizeStatusName(row[24] || 'Pending');
        } else if (sName.indexOf('citizen') > -1) {
          serviceType = 'নাগরিকত্ব সনদ';
          certNo = toEnglishDigit(row[16] || '');
          appName = row[1] || '';
          father = row[3] || '';
          nidVal = toEnglishDigit(row[2] || '');
          mobVal = toEnglishDigit(row[8] || '');
          applyDateVal = formatBanglaDate(row[14]);
          statusVal = normalizeStatusName(row[13] || 'Pending');
        } else if (sName.indexOf('family') > -1) {
          serviceType = (appId.indexOf('AUL-UW-') === 0) ? 'উত্তরাধিকারী সনদ' : (row[24] || 'পারিবারিক সনদ');
          appName = row[1] || '';
          father = row[3] || '';
          nidVal = toEnglishDigit(row[2] || '');
          mobVal = toEnglishDigit(row[5] || '');
          applyDateVal = formatBanglaDate(row[10]);
          statusVal = normalizeStatusName(row[9] || 'Pending');
        } else if (sName.indexOf('warishan') > -1) {
          serviceType = 'ওয়ারিশান সনদ';
          appName = row[1] || '';
          father = row[2] || '';
          nidVal = toEnglishDigit(row[4] || '');
          mobVal = toEnglishDigit(row[5] || '');
          applyDateVal = formatBanglaDate(row[10]);
          statusVal = normalizeStatusName(row[9] || 'Pending');
        } else {
          serviceType = row[1] || 'সাধারণ প্রত্যয়ন';
          appName = row[2] || row[1] || '';
          father = row[4] || row[3] || '';
          nidVal = toEnglishDigit(row[3] || row[2] || '');
          mobVal = toEnglishDigit(row[9] || row[8] || row[5] || '');
          applyDateVal = formatBanglaDate(row[15] || row[14] || row[10]);
          statusVal = normalizeStatusName(row[14] || row[13] || row[9] || 'Pending');
        }

        results.push({
          appId: appId,
          licNo: licNo,
          certNo: certNo,
          type: serviceType,
          serviceType: serviceType,
          applicantName: appName,
          name: appName,
          fatherName: father,
          fatherSpouseName: father,
          nid: nidVal,
          mobile: mobVal,
          applyDate: applyDateVal,
          date: applyDateVal,
          status: statusVal
        });
      }
    }

    if (results.length === 0) {
      return { found: false };
    }

    var resObj = {
      found: true,
      total: results.length,
      list: results
    };

    for (var k in results[0]) {
      if (!resObj.hasOwnProperty(k)) {
        resObj[k] = results[0][k];
      }
    }

    return resObj;
  } catch (err) {
    return { found: false, error: err.toString() };
  }
}


function submitCitizenshipDirect(formData) {
  try {
    var sheet = getSheet('Citizenship');
    var date = Utilities.formatDate(new Date(), "GMT+6", "dd/MM/yyyy");
    var marital = formData.maritalStatus || 'অবিবাহিত';
    var spouse = (marital === 'বিবাহিত') ? (formData.spouseName || '') : '';
    var rand6 = Math.floor(100000 + Math.random() * 900000).toString();
    var trackingId = 'AUL-' + rand6;

    var dobYear = "2000";
    if (formData.dob) {
      var match = toEnglishDigit(formData.dob).match(/\d{4}/);
      if (match) dobYear = match[0];
    }
    var certNo = dobYear + "7819510" + rand6;

    var newRow = [
      trackingId, formData.name, "'" + toEnglishDigit(formData.nid), formData.fatherName,
      formData.motherName, formatBanglaDate(formData.dob), marital, spouse,
      "'" + toEnglishDigit(formData.mobile), toEnglishDigit(formData.wardNo), formData.village,
      formData.postOffice || 'আউলিয়াপুর ময়দান', 'বরিশাল', 'Pending', date, 'চেয়ারম্যান', "'" + certNo
    ];

    sheet.appendRow(newRow);
    registerSearchIndexRecordFromSubmission('Citizenship', {
      appId: trackingId,
      serviceType: 'নাগরিকত্ব সনদ',
      applicantName: formData.name || '',
      mobile: formData.mobile || '',
      nid: formData.nid || '',
      status: 'Pending',
      applyDate: date,
      date: date,
      payload: { ...formData, appId: trackingId, certNo: certNo }
    });
    return { success: true, appId: trackingId, certNo: certNo };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function submitFamilyDirect(formData) {
  try {
    var sheet = getSheet('FamilyCert');
    var date = Utilities.formatDate(new Date(), "GMT+6", "dd/MM/yyyy");
    var certificateType = normalizeFamilyCertificateType(formData);
    var isSuccession = certificateType === 'উত্তরাধিকারী সনদ';
    var prefix = isSuccession ? 'AUL-UW-' : 'AUL-FW-';
    var rand6 = Math.floor(100000 + Math.random() * 900000).toString();
    var trackingId = prefix + rand6;
    var membersJson = JSON.stringify(formData.members || []);

    var row = [
      trackingId,
      formData.name || '',
      "'" + toEnglishDigit(formData.nid || ''),
      formData.fatherName || '',
      formData.motherName || '',
      "'" + toEnglishDigit(formData.mobile || ''),
      toEnglishDigit(formData.wardNo || ''),
      formData.village || '',
      formData.postOffice || 'আউলিয়াপুর ময়দান',
      'Pending',
      date,
      'চেয়ারম্যান',
      membersJson,
      formData.deceasedName || '',
      formData.deceasedFather || '',
      formData.deceasedMother || '',
      formData.deceasedDate || '',
      formData.applicantRelation || '',
      toEnglishDigit(formData.deceasedWard || ''),
      formData.deceasedVillage || '',
      formData.deceasedPostOffice || '',
      formData.deceasedUnion || '১১নং আউলিয়াপুর ইউনিয়ন',
      formData.deceasedUpazila || 'পটুয়াখালী সদর',
      formData.deceasedDistrict || 'পটুয়াখালী',
      certificateType
    ];
    sheet.appendRow(row);
    registerSearchIndexRecordFromSubmission('FamilyCert', {
      appId: trackingId,
      serviceType: certificateType,
      applicantName: formData.name || '',
      mobile: formData.mobile || '',
      nid: formData.nid || '',
      status: 'Pending',
      applyDate: date,
      date: date,
      payload: { ...formData, appId: trackingId, type: certificateType }
    });

    return {
      success: true,
      appId: trackingId,
      type: certificateType,
      serviceType: certificateType
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function submitApplication(formData) {
  if(formData.type === 'নাগরিকত্ব সনদ' || formData.type === 'জাতীয়তা সনদ') {
    return submitCitizenshipDirect(formData);
  }
  if(formData.type === 'পারিবারিক সনদ' || formData.type === 'উত্তরাধিকারী সনদ') {
    return submitFamilyDirect(formData);
  }
  if(formData.type === 'ওয়ারিশান সনদ') {
    return submitWarishanApplication(formData);
  }
  try {
    var date = Utilities.formatDate(new Date(), "GMT+6", "dd/MM/yyyy");
    var type = formData.type || 'সাধারণ প্রত্যয়ন';
    var rand6 = Math.floor(100000 + Math.random() * 900000).toString();
    var trackingId = 'AUL-' + rand6;

    var sheet = getSheet('Applications');
    sheet.appendRow([trackingId, type, formData.name, "'" + toEnglishDigit(formData.nid), formData.fatherName, formData.motherName, formatBanglaDate(formData.dob || ''), formData.maritalStatus || 'অবিবাহিত', formData.spouseName || '', "'" + toEnglishDigit(formData.mobile), toEnglishDigit(formData.wardNo), formData.village, formData.postOffice || 'আউলিয়াপুর ময়দান', 'বরিশাল', 'Pending', date, 'চেয়ারম্যান', '']);
    registerSearchIndexRecordFromSubmission('Applications', {
      appId: trackingId,
      serviceType: type,
      applicantName: formData.name || '',
      mobile: formData.mobile || '',
      nid: formData.nid || '',
      status: 'Pending',
      applyDate: date,
      date: date,
      payload: { ...formData, appId: trackingId, type: type }
    });
    return { success: true, appId: trackingId };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function submitTradeLicenseApplication(data) {
  try {
    var sheet = getSheet('TradeLicense');
    var date = Utilities.formatDate(new Date(), "GMT+6", "dd/MM/yyyy");
    var rand6 = Math.floor(100000 + Math.random() * 900000).toString();
    var trackingId = 'AUL-TR-' + rand6;
    var licNo = '199278195100' + rand6.substring(0, 4);
    var receiptNo = (sheet.getLastRow() < 10) ? '0' + sheet.getLastRow() : sheet.getLastRow().toString();

    sheet.appendRow([
      trackingId, "'" + licNo, "'" + receiptNo, data.orgName, data.ownerName,
      data.fatherName, data.motherName, "'" + toEnglishDigit(data.nid), formatBanglaDate(data.dob),
      "'" + toEnglishDigit(data.mobile), data.ownerAddress, data.category, data.bizDetails || '',
      data.bizAddress, formatBanglaDate(data.bizStartDate), data.fiscalYear || '২০২৫-২০২৬',
      data.capital || '', '200', '30', data.commTax || '0', data.signTax || '0',
      toEnglishDigit(data.totalFee || '0'), date, 'চেয়ারম্যান', 'Pending', data.photo || ''
    ]);
    registerSearchIndexRecordFromSubmission('TradeLicense', {
      appId: trackingId,
      serviceType: 'ট্রেড লাইসেন্স',
      applicantName: data.ownerName || data.orgName || '',
      mobile: data.mobile || '',
      nid: data.nid || '',
      status: 'Pending',
      applyDate: date,
      date: date,
      payload: { ...data, appId: trackingId, type: 'ট্রেড লাইসেন্স' }
    });

    return { success: true, appId: trackingId, licNo: licNo };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function submitTradeRenewalApplication(data) {
  try {
    var sheet = getSheet('TradeLicense');
    var date = Utilities.formatDate(new Date(), "GMT+6", "dd/MM/yyyy");
    var rand6 = Math.floor(100000 + Math.random() * 900000).toString();
    var renewalAppId = 'AUL-RN-' + rand6;
    var receiptNo = (sheet.getLastRow() < 10) ? '0' + sheet.getLastRow() : sheet.getLastRow().toString();

    sheet.appendRow([
      renewalAppId, "'" + toEnglishDigit(data.originalLicNo), "'" + receiptNo, data.orgName, data.ownerName,
      data.fatherName, data.motherName, "'" + toEnglishDigit(data.nid), formatBanglaDate(data.dob),
      "'" + toEnglishDigit(data.mobile), data.ownerAddress, data.category, data.bizDetails || '',
      data.bizAddress, formatBanglaDate(data.bizStartDate), data.fiscalYear, data.capital || '',
      '200', '30', data.commTax || '0', data.signTax || '0', toEnglishDigit(data.totalFee || '0'),
      date, 'চেয়ারম্যান', 'Pending', data.photo || ''
    ]);
    registerSearchIndexRecordFromSubmission('TradeLicense', {
      appId: renewalAppId,
      serviceType: 'ট্রেড লাইসেন্স',
      applicantName: data.ownerName || data.orgName || '',
      mobile: data.mobile || '',
      nid: data.nid || '',
      status: 'Pending',
      applyDate: date,
      date: date,
      payload: { ...data, appId: renewalAppId, type: 'ট্রেড লাইসেন্স' }
    });

    return { success: true, appId: renewalAppId, licNo: data.originalLicNo };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function quickAdminRenewTradeLicense(appId, newFiscal) {
  try {
    var details = getTradeLicenseDetails(appId);
    if(!details.found) return { success: false, error: 'লাইসেন্স পাওয়া যায়নি!' };
    details.fiscalYear = newFiscal;
    details.originalLicNo = details.licNo;
    return submitTradeRenewalApplication(details);
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getCitizenshipApps() {
  var result = [];
  var sheet = getSheet('Citizenship');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return result;

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      appId: data[i][0], certNo: toEnglishDigit(data[i][16]), name: data[i][1], nid: toEnglishDigit(data[i][2]),
      fatherName: data[i][3], motherName: data[i][4], dob: formatBanglaDate(data[i][5]), maritalStatus: data[i][6] || 'অবিবাহিত',
      spouseName: data[i][7] || '', mobile: toEnglishDigit(data[i][8]), wardNo: toEnglishDigit(data[i][9]), village: data[i][10],
      postOffice: data[i][11] || 'আউলিয়াপুর ময়দান', status: data[i][13] || 'Pending', applyDate: formatBanglaDate(data[i][14]), signatoryRole: data[i][15] || 'চেয়ারম্যান'
    });
  }
  return result;
}

function getTradeLicenses() {
  var sheet = getSheet('TradeLicense');
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      appId: data[i][0],
      licNo: toEnglishDigit(data[i][1]),
      receiptNo: toEnglishDigit(data[i][2]),
      orgName: data[i][3],
      ownerName: data[i][4],
      fatherName: data[i][5] || '',
      mobile: toEnglishDigit(data[i][9]),
      fiscalYear: data[i][15] || '২০২৫-২০২৬',
      totalFee: toEnglishDigit(data[i][21]),
      status: data[i][24] || 'Pending'
    });
  }
  return list;
}

function getTradeRenewals() {
  var sheet = getSheet('TradeLicense');
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var appId = (data[i][0] || '').toString();
    if (appId.indexOf('AUL-RN-') !== 0) continue;
    list.push({
      appId: appId,
      licNo: toEnglishDigit(data[i][1] || ''),
      receiptNo: toEnglishDigit(data[i][2] || ''),
      orgName: data[i][3] || '',
      ownerName: data[i][4] || '',
      fatherName: data[i][5] || '',
      mobile: toEnglishDigit(data[i][9] || ''),
      fiscalYear: data[i][15] || '২০২৫-২০২৬',
      totalFee: toEnglishDigit(data[i][21] || ''),
      status: data[i][24] || 'Pending',
      isRenewal: true,
      requestType: 'renewal',
      serviceType: 'ট্রেড লাইসেন্স নবায়ন'
    });
  }
  return list;
}

function getTradeLicenseRenewals() { return getTradeRenewals(); }
function getRenewalApplications() { return getTradeRenewals(); }

function getTradeLicenseDetails(query) {
  var sheet = getSheet('TradeLicense');
  var data = sheet.getDataRange().getValues();
  var q = toEnglishDigit(query || '').toUpperCase();

  for (var i = 1; i < data.length; i++) {
    var appId = (data[i][0] || '').toString().toUpperCase();
    var licNo = (data[i][1] || '').toString().toUpperCase();
    var mob = (data[i][9] || '').toString().toUpperCase();
    var org = (data[i][3] || '').toString().toUpperCase();

    if (appId.indexOf(q) > -1 || licNo.indexOf(q) > -1 || mob.indexOf(q) > -1 || org.indexOf(q) > -1) {
      return {
        found: true,
        appId: data[i][0],
        licNo: toEnglishDigit(data[i][1]),
        receiptNo: toEnglishDigit(data[i][2]),
        orgName: data[i][3],
        ownerName: data[i][4],
        fatherName: data[i][5],
        motherName: data[i][6],
        nid: toEnglishDigit(data[i][7]),
        dob: formatBanglaDate(data[i][8]),
        mobile: toEnglishDigit(data[i][9]),
        ownerAddress: data[i][10],
        category: data[i][11],
        bizDetails: data[i][12],
        bizAddress: data[i][13],
        bizStartDate: formatBanglaDate(data[i][14]),
        fiscalYear: data[i][15] || '২০২৫-২০২৬',
        capital: data[i][16],
        totalFee: toEnglishDigit(data[i][21]),
        applyDate: formatBanglaDate(data[i][22]),
        signatoryRole: data[i][23] || 'চেয়ারম্যান',
        status: data[i][24] || 'Pending'
      };
    }
  }
  return { found: false };
}

function getFamilyDetails(appId) {
  try {
    var q = toEnglishDigit(appId || '').toUpperCase();
    var sheet = getFamilyCertSheet();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var rowAppId = (data[i][0] || '').toString().toUpperCase();
      if (rowAppId === q || rowAppId.indexOf(q) > -1) {
        var members = [];
        try { members = JSON.parse(data[i][12] || '[]'); } catch (err) { members = []; }
        var storedType = (data[i][24] || '').toString().trim();
        var certificateType = rowAppId.indexOf('AUL-UW-') === 0 ? 'উত্তরাধিকারী সনদ' : (storedType || 'পারিবারিক সনদ');

        return {
          found: true,
          appId: data[i][0],
          type: certificateType,
          serviceType: certificateType,
          name: data[i][1],
          nid: toEnglishDigit(data[i][2] || ''),
          fatherName: data[i][3] || '',
          motherName: data[i][4] || '',
          mobile: toEnglishDigit(data[i][5] || ''),
          wardNo: toEnglishDigit(data[i][6] || ''),
          village: data[i][7] || '',
          postOffice: data[i][8] || 'আউলিয়াপুর ময়দান',
          status: data[i][9] || 'Pending',
          applyDate: formatBanglaDate(data[i][10]),
          signatoryRole: data[i][11] || 'চেয়ারম্যান',
          members: members,
          deceasedName: data[i][13] || ''
        };
      }
    }
    return { found: false };
  } catch (e) {
    return { found: false, error: e.toString() };
  }
}

function saveEditedApplication(data) {
  try {
    if (!data || !data.appId) return { success: false, error: 'অ্যাপ আইডি অনুপস্থিত।' };
    var appId = (data.appId || '').toString().trim();
    var type = (data.type || data.serviceType || data.certificateType || '').toString().trim();
    if (type === 'ওয়ারিশান সনদ' || appId.indexOf('AUL-WAR-') === 0) return updateWarishanData(data);
    return updateApplicationData(data);
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteApplication(appId) {
  var rawQuery = (appId || "").toString().trim();
  var queryEng = toEnglishDigit(rawQuery).toUpperCase();
  var sheetsToSearch = ['Citizenship', 'Applications', 'FamilyCert', 'Warishan', 'TradeLicense'];
  for (var s = 0; s < sheetsToSearch.length; s++) {
    var sheet = getSheet(sheetsToSearch[s]);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowA = (data[i][0] || "").toString().trim();
      var rowAEng = toEnglishDigit(rowA).toUpperCase();
      if (rowAEng === queryEng || rowAEng.indexOf(queryEng) > -1) {
        sheet.deleteRow(i + 1);
        deleteSearchIndexRecord(rawQuery);
        syncSearchIndexFromSheets();
        return { success: true };
      }
    }
  }
  deleteSearchIndexRecord(rawQuery);
  syncSearchIndexFromSheets();
  return { success: false };
}

function getMasterDashboardStats() {
  var citApps = getCitizenshipApps();
  var famApps = getFamilyApps();
  var warApps = getWarishanApps();
  var genApps = getAllApplications();
  var tradeList = getTradeLicenses();
  var taxList = getTaxPayers();

  var totalAll = citApps.length + famApps.length + warApps.length + genApps.length + tradeList.length;

  function countStats(arr) {
    var approved = 0; var pending = 0;
    for(var i=0; i<arr.length; i++) {
      var st = (arr[i].status || '').trim();
      if(st === 'Approved' || st === 'Active') approved++;
      else if(st === 'Pending') pending++;
    }
    return { total: arr.length, approved: approved, pending: pending };
  }

  var citStat = countStats(citApps);
  var famStat = countStats(famApps);
  var warStat = countStats(warApps);
  var genStat = countStats(genApps);
  var tradeStat = countStats(tradeList);

  return {
    totalApps: toBanglaDigit(totalAll),
    totalApprovedApps: toBanglaDigit(citStat.approved + famStat.approved + warStat.approved + genStat.approved + tradeStat.approved),
    totalPendingApps: toBanglaDigit(citStat.pending + famStat.pending + warStat.pending + genStat.pending + tradeStat.pending),
    citTotal: toBanglaDigit(citStat.total),
    citApproved: toBanglaDigit(citStat.approved),
    citPending: toBanglaDigit(citStat.pending),
    famTotal: toBanglaDigit(famStat.total),
    famApproved: toBanglaDigit(famStat.approved),
    famPending: toBanglaDigit(famStat.pending),
    warTotal: toBanglaDigit(warStat.total),
    warApproved: toBanglaDigit(warStat.approved),
    warPending: toBanglaDigit(warStat.pending),
    genTotal: toBanglaDigit(genStat.total),
    genApproved: toBanglaDigit(genStat.approved),
    genPending: toBanglaDigit(genStat.pending),
    tradeTotal: toBanglaDigit(tradeStat.total),
    tradeActive: toBanglaDigit(tradeStat.approved),
    tradePending: toBanglaDigit(tradeStat.pending),
    taxTotal: toBanglaDigit(taxList.length)
  };
}

function getAllApplications() {
  var sheet = getSheet('Applications');
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    result.push({
      appId: data[i][0], type: data[i][1] || 'সাধারণ প্রত্যয়ন', name: data[i][2], nid: toEnglishDigit(data[i][3]),
      fatherName: data[i][4], motherName: data[i][5], mobile: toEnglishDigit(data[i][9] || data[i][8]),
      wardNo: toEnglishDigit(data[i][10] || data[i][9]), village: data[i][11] || data[i][10],
      postOffice: data[i][12] || 'আউলিয়াপুর ময়দান', status: data[i][14] || data[i][13] || 'Pending',
      applyDate: formatBanglaDate(data[i][15] || data[i][14])
    });
  }
  return result;
}

function getFamilyApps() {
  var result = [];
  var sheet = getFamilyCertSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var appId = data[i][0].toString().trim();
    var storedType = (data[i][24] || '').toString().trim();
    var certificateType = appId.indexOf('AUL-UW-') === 0 ? 'উত্তরাধিকারী সনদ' : (storedType || 'পারিবারিক সনদ');

    result.push({
      appId: appId,
      type: certificateType,
      serviceType: certificateType,
      name: data[i][1] || '',
      applicantName: data[i][1] || '',
      nid: toEnglishDigit(data[i][2] || ''),
      fatherName: data[i][3] || '',
      motherName: data[i][4] || '',
      mobile: toEnglishDigit(data[i][5] || ''),
      wardNo: toEnglishDigit(data[i][6] || ''),
      village: data[i][7] || '',
      postOffice: data[i][8] || 'আউলিয়াপুর ময়দান',
      status: data[i][9] || 'Pending',
      applyDate: formatBanglaDate(data[i][10]),
      signatoryRole: data[i][11] || 'চেয়ারম্যান'
    });
  }
  return result;
}

function getTaxPayers() {
  var sheet = getSheet('TaxPayers');
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      holdingNo: toEnglishDigit(data[i][0]), name: data[i][1], nid: toEnglishDigit(data[i][2]),
      fatherName: data[i][3], mobile: toEnglishDigit(data[i][4]), wardNo: toEnglishDigit(data[i][5]),
      village: data[i][6], houseType: data[i][7], annualTax: toEnglishDigit(data[i][8])
    });
  }
  return list;
}

var ADMIN_MENU_LIST = [
  'dashTab', 'citTab', 'tradeTab', 'famTab', 'warTab', 'appsTab',
  'taxTab', 'accountsTab', 'bankTab', 'beneficiaryTab', 'projectTab',
  'reportTab', 'upSettingsTab', 'settingsTab'
];

function normalizeAdminPermissions(input, role) {
  var raw = (input && typeof input === 'object') ? input : {};
  var allowedMenus = [];

  if (Array.isArray(raw.allowedMenus)) {
    allowedMenus = raw.allowedMenus.slice();
  } else if (Array.isArray(raw.menus)) {
    allowedMenus = raw.menus.filter(function(item) {
      return item && item.enabled !== false;
    }).map(function(item) {
      return item.code || item.name || item.id || '';
    }).filter(Boolean);
  }

  var isSuper = String(role || '').toLowerCase() === 'superadmin';

  if (isSuper) {
    return {
      canApprove: true,
      canReject: true,
      canEdit: true,
      canDelete: true,
      allowedMenus: ADMIN_MENU_LIST.slice(),
      photoUrl: raw.photoUrl || raw.photo || ''
    };
  }

  if (!allowedMenus.length && Object.keys(raw).length === 0) {
    allowedMenus = ['dashTab'];
  }

  return {
    canApprove: raw.canApprove === true || raw.approve === true,
    canReject: raw.canReject === true || raw.reject === true,
    canEdit: raw.canEdit === true || raw.edit === true,
    canDelete: raw.canDelete === true || raw.delete === true,
    allowedMenus: allowedMenus,
    photoUrl: raw.photoUrl || raw.photo || ''
  };
}

function adminLogin(username, password) {
  var u = (username || "").toString().trim();
  var p = (password || "").toString().trim();
  try {
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var sheetUser = (data[i][0] || "").toString().trim();
      var sheetPass = (data[i][1] || "").toString().trim();
      if (sheetUser.toLowerCase() === u.toLowerCase() && sheetPass === p) {
        var parsed = {};
        try { parsed = JSON.parse(data[i][4] || '{}'); } catch (e) {}
        var perms = normalizeAdminPermissions(parsed, data[i][2] || 'Admin');
        var photoUrl = parsed.photoUrl || parsed.photo || (data[i][5] ? data[i][5].toString() : '') || '';

        return {
          success: true,
          username: sheetUser,
          role: data[i][2] || 'Admin',
          name: data[i][3] || sheetUser,
          photoUrl: photoUrl || 'https://upload.wikimedia.org/wikipedia/commons/7/7c/Profile_avatar_placeholder_large.png',
          permissions: perms,
          allowedMenus: perms.allowedMenus
        };
      }
    }
  } catch (e) {}

  if ((u.toLowerCase() === 'admin' || u.toLowerCase() === 'superadmin') && p === '123456') {
    var role = u.toLowerCase() === 'superadmin' ? 'SuperAdmin' : 'Admin';
    var defaultPerms = normalizeAdminPermissions({}, role);
    return {
      success: true,
      username: u,
      role: role,
      name: role === 'SuperAdmin' ? 'অ্যাড. মোঃ হুমায়ুন কবির (চেয়ারম্যান)' : 'অ্যাডমিন',
      photoUrl: 'https://upload.wikimedia.org/wikipedia/commons/7/7c/Profile_avatar_placeholder_large.png',
      permissions: defaultPerms,
      allowedMenus: defaultPerms.allowedMenus
    };
  }
  return { success: false, message: 'ইউজারনেম বা পাসওয়ার্ড সঠিক নয়!' };
}

function getAdminUsers() {
  var sheet = getSheet('Users');
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var perms = {};
    var photoUrl = '';
    try {
      perms = JSON.parse(data[i][4] || '{}');
      photoUrl = perms.photoUrl || perms.photo || (data[i][5] ? data[i][5].toString() : '') || '';
    } catch (e) {}
    var normalized = normalizeAdminPermissions(perms, data[i][2] || 'Admin');
    users.push({
      username: data[i][0],
      role: data[i][2] || 'Admin',
      name: data[i][3] || '',
      photoUrl: photoUrl || 'https://upload.wikimedia.org/wikipedia/commons/7/7c/Profile_avatar_placeholder_large.png',
      permissions: normalized,
      allowedMenus: normalized.allowedMenus
    });
  }
  return users;
}

function getAllUsers() { return getAdminUsers(); }

function saveNewUser(uData) {
  try {
    var sheet = getSheet('Users');
    var role = uData.role || 'Admin';
    var permissionInput = (uData && uData.permissions) ? uData.permissions : {};
    if (!permissionInput.allowedMenus && uData.allowedMenus) {
      permissionInput.allowedMenus = uData.allowedMenus;
    }
    var permsPayload = normalizeAdminPermissions(permissionInput, role);
    permsPayload.photoUrl = uData.photoUrl || uData.photo || permsPayload.photoUrl || '';

    sheet.appendRow([
      uData.username,
      uData.password,
      role,
      uData.name || '',
      JSON.stringify(permsPayload),
      permsPayload.photoUrl || ''
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updateUser(uData) {
  try {
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() !== String(uData.originalUsername || uData.username).trim().toLowerCase()) continue;

      var oldPermissions = {};
      try { oldPermissions = JSON.parse(data[i][4] || '{}'); } catch (e) {}
      var incoming = uData.permissions || {};
      if (!incoming.allowedMenus && uData.allowedMenus) incoming.allowedMenus = uData.allowedMenus;
      var role = uData.role || data[i][2] || 'Admin';
      var permissions = normalizeAdminPermissions(incoming, role);
      permissions.photoUrl = uData.photoUrl || uData.photo || oldPermissions.photoUrl || permissions.photoUrl || '';

      sheet.getRange(i + 1, 1).setValue(uData.username || data[i][0]);
      if (uData.password && uData.password.trim()) sheet.getRange(i + 1, 2).setValue(uData.password.trim());
      sheet.getRange(i + 1, 3).setValue(role);
      sheet.getRange(i + 1, 4).setValue(uData.name || data[i][3]);
      sheet.getRange(i + 1, 5).setValue(JSON.stringify(permissions));
      if (sheet.getLastColumn() >= 6) sheet.getRange(i + 1, 6).setValue(permissions.photoUrl || '');
      return { success: true };
    }
    return { success: false, error: 'এডমিন পাওয়া যায়নি!' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteAdminUser(uname) {
  var sheet = getSheet('Users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(uname).trim().toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

function deleteUser(uname) { return deleteAdminUser(uname); }

function getUPOfficialSettings() {
  var defaults = {
    chairman: 'অ্যাড. মোঃ হুমায়ুন কবির',
    panelChairman: 'মোঃ আবদুস সালাম মৃধা',
    secretary: 'মোঃ মোতাহার উদ্দিন',
    chairmanEn: 'Adv. Md. Humayun Kabir',
    panelChairmanEn: 'Md. Abdus Salam Mridha',
    secretaryEn: 'Md. Motahar Uddin'
  };
  try {
    var props = PropertiesService.getScriptProperties();
    var saved = props.getProperty('UP_OFFICIAL_SETTINGS');
    if (!saved) return defaults;
    return Object.assign({}, defaults, JSON.parse(saved));
  } catch (e) {
    return defaults;
  }
}

function saveUPOfficialSettings(data) {
  var settings = Object.assign({}, getUPOfficialSettings(), data || {});
  PropertiesService.getScriptProperties().setProperty('UP_OFFICIAL_SETTINGS', JSON.stringify(settings));
  return { success: true, data: settings };
}

function updateTradeLicenseData(ed) {
  try {
    var sheet = getSheet('TradeLicense');
    var data = sheet.getDataRange().getValues();
    var q = toEnglishDigit((ed && ed.appId) || '').toUpperCase();
    var qLic = toEnglishDigit((ed && ed.licNo) || '').toUpperCase();

    for (var i = 1; i < data.length; i++) {
      var appId = (data[i][0] || '').toString().toUpperCase();
      var licNo = (data[i][1] || '').toString().toUpperCase();
      if ((q && (appId === q || appId.indexOf(q) > -1)) || (qLic && (licNo === qLic || licNo.indexOf(qLic) > -1))) {
        if (ed && ed.orgName !== undefined) sheet.getRange(i + 1, 4).setValue(ed.orgName);
        if (ed && ed.ownerName !== undefined) sheet.getRange(i + 1, 5).setValue(ed.ownerName);
        if (ed && ed.fatherName !== undefined) sheet.getRange(i + 1, 6).setValue(ed.fatherName);
        if (ed && ed.motherName !== undefined) sheet.getRange(i + 1, 7).setValue(ed.motherName);
        if (ed && ed.nid !== undefined) sheet.getRange(i + 1, 8).setValue("'" + toEnglishDigit(ed.nid || ''));
        if (ed && ed.mobile !== undefined) sheet.getRange(i + 1, 10).setValue("'" + toEnglishDigit(ed.mobile || ''));
        if (ed && ed.ownerAddress !== undefined) sheet.getRange(i + 1, 11).setValue(ed.ownerAddress);
        if (ed && ed.category !== undefined) sheet.getRange(i + 1, 12).setValue(ed.category);
        if (ed && ed.bizDetails !== undefined) sheet.getRange(i + 1, 13).setValue(ed.bizDetails || '');
        if (ed && ed.bizAddress !== undefined) sheet.getRange(i + 1, 14).setValue(ed.bizAddress);
        if (ed && ed.bizStartDate !== undefined) sheet.getRange(i + 1, 15).setValue(formatBanglaDate(ed.bizStartDate));
        if (ed && ed.fiscalYear !== undefined) sheet.getRange(i + 1, 16).setValue(ed.fiscalYear);
        if (ed && ed.commTax !== undefined) sheet.getRange(i + 1, 19).setValue(ed.commTax || '0');
        if (ed && ed.signTax !== undefined) sheet.getRange(i + 1, 20).setValue(ed.signTax || '0');
        if (ed && ed.totalFee !== undefined) sheet.getRange(i + 1, 22).setValue(toEnglishDigit(ed.totalFee || '0'));
        if (ed && ed.signatoryRole !== undefined) sheet.getRange(i + 1, 24).setValue(ed.signatoryRole);
        return { success: true };
      }
    }
    return { success: false, error: 'ট্রেড লাইসেন্স তথ্য পাওয়া যায়নি!' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

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
    var result;

    if (action === 'adminLogin') {
      result = adminLogin(data.username, data.password);
    } else if (action === 'updateUser') {
      result = updateUser(data);
    } else if (action === 'submitCitizenshipDirect') {
      result = submitCitizenshipDirect(data);
    } else if (action === 'submitTradeLicenseApplication') {
      result = submitTradeLicenseApplication(data);
    } else if (action === 'submitTradeRenewalApplication') {
      result = submitTradeRenewalApplication(data);
    } else if (action === 'getTradeLicenseDetails') {
      result = getTradeLicenseDetails(data);
    } else if (action === 'updateTradeLicenseData') {
      result = updateTradeLicenseData(data);
    } else if (action === 'submitFamilyDirect') {
      result = submitFamilyDirect(data);
    } else if (action === 'getFamilyDetails') {
      result = getFamilyDetails(data);
    } else if (action === 'submitWarishanApplication') {
      result = submitWarishanApplication(data);
    } else if (action === 'getWarishanDetails') {
      result = getWarishanDetails(data);
    } else if (action === 'submitApplication') {
      result = submitApplication(data);
    } else if (action === 'trackApplication') {
      result = trackApplication(data);
    } else if (action === 'getMasterDashboardStats') {
      result = getMasterDashboardStats();
    } else if (action === 'getCitizenshipApps') {
      result = getCitizenshipApps();
    } else if (action === 'getTradeLicenses') {
      result = getTradeLicenses();
    } else if (action === 'getTradeRenewals' || action === 'getTradeLicenseRenewals' || action === 'getRenewalApplications') {
      result = getTradeRenewals();
    } else if (action === 'getFamilyApps') {
      result = getFamilyApps();
    } else if (action === 'getWarishanApps') {
      result = getWarishanApps();
    } else if (action === 'getAllApplications') {
      result = getAllApplications();
    } else if (action === 'getTaxPayers') {
      result = getTaxPayers();
    } else if (action === 'getAllUsers' || action === 'getAdminUsers') {
      result = getAdminUsers();
    } else if (action === 'saveNewUser' || action === 'saveAdminUser') {
      result = saveNewUser(data);
    } else if (action === 'deleteUser' || action === 'deleteAdminUser') {
      result = deleteAdminUser(data || data.username || data.uname || data);
    } else if (action === 'updateAppStatus') {
      result = updateAppStatus(data.appId, data.status);
    } else if (action === 'deleteApplication') {
      result = deleteApplication(data);
    } else if (action === 'updateWarishanData') {
      result = updateWarishanData(data);
    } else if (action === 'saveEditedApplication') {
      result = saveEditedApplication(data);
    } else if (action === 'saveUPSettings') {
      result = saveUPOfficialSettings(data);
    } else if (action === 'getUPSettings') {
      result = getUPOfficialSettings();
    } else {
      if (typeof this[action] === 'function') {
        result = this[action](data);
      } else {
        result = { success: false, error: 'ফাংশন খুঁজে পাওয়া যায়নি: ' + action };
      }
    }

    return ContentService.createTextOutput(JSON.stringify(result || { success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    initMasterDatabase();
    return ContentService.createTextOutput(JSON.stringify({ status: "API is active and running!" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}