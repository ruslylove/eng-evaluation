/**
 * Serves the HTML page.
 */
function doGet() {
  var userEmail = Session.getActiveUser().getEmail();
  
  // If the email is blank or does not belong to kmutnb.ac.th (case-insensitive)
  if (!userEmail || !/@.*kmutnb\.ac\.th$/i.test(userEmail)) {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
      '  <meta charset="UTF-8">' +
      '  <meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '  <title>เข้าสู่ระบบด้วยบัญชีมหาวิทยาลัย — คณะวิศวกรรมศาสตร์ มจพ.</title>' +
      '  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600&display=swap" rel="stylesheet">' +
      '  <style>' +
      '    body { font-family: "Sarabun", sans-serif; background-color: #F4F6FA; color: #1A2535; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; padding: 20px; box-sizing: border-box; }' +
      '    .error-card { background: #fff; padding: 40px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); text-align: center; max-width: 480px; width: 100%; border-top: 4px solid #921823; }' +
      '    .icon { font-size: 56px; margin-bottom: 20px; color: #921823; }' +
      '    h2 { margin-top: 0; margin-bottom: 16px; color: #590F16; font-size: 20px; font-weight: 600; }' +
      '    p { color: #556376; font-size: 14px; line-height: 1.6; margin-bottom: 24px; }' +
      '    .email-info { background: #FAF0F1; border: 1px solid #F5C0BC; border-radius: 6px; padding: 12px; margin-bottom: 24px; font-size: 13px; font-weight: 500; color: #B52A2A; text-align: left; word-break: break-all; }' +
      '    .btn { display: inline-block; padding: 10px 24px; color: #fff; background: #921823; border-radius: 4px; text-decoration: none; font-weight: 500; font-size: 14px; transition: background .2s; }' +
      '    .btn:hover { background: #B52635; }' +
      '  </style>' +
      '</head>' +
      '<body>' +
      '  <div class="error-card">' +
      '    <div class="icon">🔒</div>' +
      '    <h2>กรุณาเข้าสู่ระบบด้วยบัญชีมหาวิทยาลัย</h2>' +
      '    <p>ระบบนี้เปิดให้ใช้งานเฉพาะบุคลากรคณะวิศวกรรมศาสตร์ มจพ. เท่านั้น กรุณาลงชื่อเข้าใช้งานด้วยบัญชี Google ของทางมหาวิทยาลัย</p>' +
      '    <div class="email-info">' +
      '      <strong>อีเมลที่พบในปัจจุบัน:</strong> ' + (userEmail ? userEmail : 'ไม่ระบุตัวตน / บัญชีภายนอก (Gmail)') + '<br>' +
      '      <strong>โดเมนที่ได้รับอนุญาต:</strong> @eng.kmutnb.ac.th หรือ @kmutnb.ac.th' +
      '    </div>' +
      '    <a href="https://accounts.google.com/Logout" target="_top" class="btn">สลับบัญชี Google (Switch Account)</a>' +
      '  </div>' +
      '</body>' +
      '</html>'
    )
    .setTitle('กรุณาเข้าสู่ระบบด้วยบัญชีมหาวิทยาลัย — คณะวิศวกรรมศาสตร์ มจพ.')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ยื่นขอค่าตอบแทนพิเศษ — คณะวิศวกรรมศาสตร์ มจพ.')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Gets the staff profile of the logged-in user via Google SSO.
 */
function getStaffProfile() {
  var userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    throw new Error('ไม่สามารถระบุอีเมลผู้ใช้งานได้จาก Google SSO. กรุณาเข้าสู่ระบบ Google Account');
  }
  
  var sheetId = '1Ljo361CBa1p9d79V4dxOGkttubCo2TnN2eiHwSAKFkk';
  var ss = SpreadsheetApp.openById(sheetId);
  
  // 1. Get Staff sheet
  var staffSheet = ss.getSheetByName('Staff');
  if (!staffSheet) {
    throw new Error('ไม่พบแท็บ "Staff" ใน Google Sheet');
  }
  
  var staffData = staffSheet.getDataRange().getValues();
  if (staffData.length <= 1) {
    return null;
  }
  
  var staffHeaders = staffData[0];
  var emailIdx = staffHeaders.indexOf('Email');
  
  if (emailIdx === -1) {
    throw new Error('ไม่พบคอลัมน์ "Email" ในแท็บ Staff');
  }
  
  // Search row matching the email
  var matchedStaffRow = null;
  for (var i = 1; i < staffData.length; i++) {
    var row = staffData[i];
    var rowEmail = String(row[emailIdx]).trim().toLowerCase();
    if (rowEmail === userEmail.trim().toLowerCase()) {
      matchedStaffRow = row;
      break;
    }
  }
  
  if (!matchedStaffRow) {
    return null;
  }
  
  // Build staff object
  var staff = {};
  for (var j = 0; j < staffHeaders.length; j++) {
    var header = staffHeaders[j];
    staff[header] = matchedStaffRow[j];
  }
  
  // 2. Link Div_ID to Division sheet if Div_ID exists
  var divId = staff['Div_ID'];
  if (divId !== undefined && divId !== null && divId !== '') {
    var divSheet = ss.getSheetByName('Division');
    if (divSheet) {
      var divData = divSheet.getDataRange().getValues();
      if (divData.length > 1) {
        var divHeaders = divData[0];
        var divIdIdxDiv = divHeaders.indexOf('Div_ID');
        var nameIdxDiv = divHeaders.indexOf('Name');
        var shortNameIdxDiv = divHeaders.indexOf('Short_Name');
        var engNameIdxDiv = divHeaders.indexOf('Eng_Name');
        
        if (divIdIdxDiv !== -1) {
          // Find matching division
          for (var k = 1; k < divData.length; k++) {
            var divRow = divData[k];
            if (String(divRow[divIdIdxDiv]).trim() === String(divId).trim()) {
              if (nameIdxDiv !== -1) staff['Div_Name_Th'] = divRow[nameIdxDiv];
              if (shortNameIdxDiv !== -1) staff['Div_Short_Name'] = divRow[shortNameIdxDiv];
              if (engNameIdxDiv !== -1) staff['Div_Name_En'] = divRow[engNameIdxDiv];
              break;
            }
          }
        }
      }
    }
  }
  
  // Look up department head name
  var deptName = staff['Div_Name_Th'] || '';
  var matchedDivId = staff['Div_ID'] || '';
  staff['Dept_Head_Th'] = getDeptHeadName(deptName, matchedDivId);
  
  return staff;
}

/**
 * Saves a new submission to a separate Google Sheet.
 * Creates sheets and headers if they do not exist.
 * 
 * @param {Object} data The submission data from the client
 */
function submitApplication(data) {
  var submissionSheetId = '1Nyp_b-tPRxSWSE7H_WaiiFrIxjxqqQmrA0ooV1kmaZ8';
  var folderId = '1eFqtjXxYhLJyDi_ZZ0R-_B_K1_k7XEMS';
  
  var ss;
  try {
    ss = SpreadsheetApp.openById(submissionSheetId);
  } catch (e) {
    throw new Error('ไม่สามารถเปิด Google Sheet สำหรับบันทึกคำขอได้ กรุณาตรวจสอบสิทธิ์การเข้าถึงหรือ ID ของไฟล์: ' + e.message);
  }
  
  var timestamp = new Date();
  var email = Session.getActiveUser().getEmail() || data.email;
  
  // 1. Check for existing submission to prevent accidental duplicates/overwrites
  var mainSheet = ss.getSheetByName('Submissions');
  if (!mainSheet) {
    mainSheet = ss.insertSheet('Submissions');
    mainSheet.appendRow([
      'Timestamp', 'Email', 'Name', 'Academic_Title', 'Department', 
      'Staff_Type', 'Appoint_Date', 'Field_Group', 'Fiscal_Year', 
      'Round', 'Total_Score', 'Pass_Status', 'Role_Type'
    ]);
    mainSheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#f3f3f3');
  } else if (mainSheet.getLastColumn() === 12) {
    // If the sheet already exists but only has 12 columns, add 'Role_Type' as column 13
    mainSheet.getRange(1, 13).setValue('Role_Type').setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  var existingRowIdx = -1;
  var mainData = mainSheet.getDataRange().getValues();
  for (var r = 1; r < mainData.length; r++) {
    var rowEmail = mainData[r][1]; // Column 2: Email
    var rowFY = String(mainData[r][8]); // Column 9: Fiscal_Year
    var rowRound = String(mainData[r][9]); // Column 10: Round
    
    if (String(rowEmail).toLowerCase().trim() === String(email).toLowerCase().trim() &&
        String(rowFY).trim() === String(data.fiscalYear).trim() &&
        String(rowRound).trim() === String(data.round).trim()) {
      existingRowIdx = r + 1; // 1-indexed row number
      break;
    }
  }

  if (existingRowIdx !== -1) {
    if (data.overwrite !== true) {
      return {
        success: false,
        duplicate: true,
        message: 'ท่านได้เคยยื่นข้อมูลสำหรับ ปีงบประมาณ พ.ศ. ' + data.fiscalYear + ' รอบที่ ' + data.round + ' ไว้แล้วในระบบ ต้องการยื่นข้อมูลใหม่เพื่อเขียนทับข้อมูลเดิมใช่หรือไม่?'
      };
    } else {
      // Overwrite: Delete previous submission metadata row
      mainSheet.deleteRow(existingRowIdx);
      
      // Also delete from Publications sheet
      var pubSheet = ss.getSheetByName('Publications');
      if (pubSheet) {
        var pubData = pubSheet.getDataRange().getValues();
        for (var pIdx = pubData.length - 1; pIdx >= 1; pIdx--) {
          var pEmail = pubData[pIdx][1];
          var pFY = String(pubData[pIdx][2]);
          var pRound = String(pubData[pIdx][3]);
          
          if (String(pEmail).toLowerCase().trim() === String(email).toLowerCase().trim() &&
              String(pFY).trim() === String(data.fiscalYear).trim() &&
              String(pRound).trim() === String(data.round).trim()) {
            pubSheet.deleteRow(pIdx + 1);
          }
        }
      }
    }
  }
  
  mainSheet.appendRow([
    timestamp,
    email,
    data.fullNameTh,
    data.position,
    data.dept,
    data.staffType,
    data.appointDate,
    data.fieldGroup,
    data.fiscalYear,
    data.round,
    data.totalScore,
    data.passStatus ? 'ผ่านเกณฑ์' : 'ไม่ผ่านเกณฑ์',
    data.isAdmin ? 'ผู้บริหาร' : 'อาจารย์ประจำ'
  ]);
  
  // 2. Save individual publications
  var pubSheet = ss.getSheetByName('Publications');
  if (!pubSheet) {
    pubSheet = ss.insertSheet('Publications');
    pubSheet.appendRow([
      'Timestamp', 'Email', 'Fiscal_Year', 'Round', 'Title', 
      'Section', 'Database_Tier', 'Role', 'Venue', 'Volume_Issue', 
      'Publish_Year', 'DOI', 'Attachments'
    ]);
    pubSheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  if (data.pubs && data.pubs.length > 0) {
    var rootFolder;
    var targetFolder;
    try {
      rootFolder = DriveApp.getFolderById(folderId);
      targetFolder = rootFolder;
      
      if (rootFolder) {
        var yearFolderName = 'ปีงบประมาณ ' + (data.fiscalYear || '2568');
        var yearFolder = getOrCreateSubfolder_(rootFolder, yearFolderName);
        
        var roundFolderName = 'รอบที่ ' + (data.round || '1');
        var roundFolder = getOrCreateSubfolder_(yearFolder, roundFolderName);
        
        var indFolderName = (data.fullNameTh || email.split('@')[0]).trim();
        targetFolder = getOrCreateSubfolder_(roundFolder, indFolderName);
      }
    } catch (e) {
      // Fallback to root Folder
    }
    
    for (var i = 0; i < data.pubs.length; i++) {
      var p = data.pubs[i];
      var driveUrls = [];
      
      // Upload files associated with this publication
      if (p.files && targetFolder) {
        var fileKeys = Object.keys(p.files);
        for (var f = 0; f < fileKeys.length; f++) {
          var key = fileKeys[f];
          var fObj = p.files[key];
          if (fObj && fObj.base64) {
            try {
              var decoded = Utilities.base64Decode(fObj.base64);
              var sanitizedEmail = email.split('@')[0];
              var fileName = sanitizedEmail + '_FY' + data.fiscalYear + '_R' + data.round + '_pub' + (i+1) + '_' + key + '_' + fObj.name;
              var blob = Utilities.newBlob(decoded, fObj.mimeType, fileName);
              var driveFile = targetFolder.createFile(blob);
              driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              driveUrls.push({ label: key, url: driveFile.getUrl() });
            } catch (err) {
              driveUrls.push({ label: key + ' (Error)', url: '' });
            }
          }
        }
      }
      
      pubSheet.appendRow([
        timestamp,
        email,
        data.fiscalYear,
        data.round,
        p.title,
        p.section,
        p.tier,
        p.role ? p.role.join(', ') : '',
        p.venue,
        p.volume,
        p.year,
        p.doi,
        '' // Placeholder for links
      ]);
      
      var lastRow = pubSheet.getLastRow();
      setCellRichTextLinks_(pubSheet, lastRow, 13, driveUrls);
    }
  }
  
  return {
    success: true,
    message: 'ยื่นคำขอสำเร็จและบันทึกข้อมูลลงระบบเรียบร้อยแล้ว'
  };
}

/**
 * Generates a DOCX memo from a Google Doc template and returns the download data.
 * 
 * @param {Object} data The submission/preview data
 * @return {Object} Base64 data of the generated DOCX file for client-side download
 */
function generateDocxMemo(data) {
  var templateId = '1RDRb9Wgxvr9M9_D9Eb9mSxewq3w3c7N1wpSGz4Ujj3Q';
  
  var tempFolder = DriveApp.getRootFolder();
  var copyName = 'บันทึกข้อความ_' + (data.fullNameTh || 'User') + '_ปี_' + data.fiscalYear;
  var docFile = DriveApp.getFileById(templateId).makeCopy(copyName, tempFolder);
  var docId = docFile.getId();
  
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  
  // Replace signature block with full academic name (inside parentheses)
  var fullAcadName = getFullAcademicName(data.fullNameTh, data.position, data.email);
  body.replaceText('\\(\\s*\\{\\{FULL_NAME\\}\\}\\s*\\)', '( ' + (fullAcadName || '') + ' )');
  body.replaceText('{{SIGN_NAME}}', fullAcadName || '');
  
  // Replace remaining placeholders (body uses นาย/นาง/นางสาว)
  body.replaceText('{{FULL_NAME}}', data.fullNameTh || '');
  body.replaceText('{{POSITION}}', data.position || '');
  body.replaceText('{{DEPT}}', data.dept || '');
  body.replaceText('{{STAFF_TYPE}}', data.staffType || '');
  body.replaceText('{{APPOINT_DATE}}', data.appointDate || '');
  // Target Subject line to append the submitter's academic name specifically
  var paragraphs = body.getParagraphs();
  for (var i = 0; i < paragraphs.length; i++) {
    var p = paragraphs[i];
    var text = p.getText();
    if (text.indexOf('เรื่อง') !== -1 && text.indexOf('ขอเสนอผลงาน') !== -1) {
      p.replaceText('\\{\\{ROUND\\}\\}', (data.round || '') + ' ของ ' + (fullAcadName || ''));
    }
  }
  var tables = body.getTables();
  for (var t = 0; t < tables.length; t++) {
    var tbl = tables[t];
    for (var r = 0; r < tbl.getNumRows(); r++) {
      var row = tbl.getRow(r);
      for (var c = 0; c < row.getNumCells(); c++) {
        var cell = row.getCell(c);
        var text = cell.getText();
        if (text.indexOf('เรื่อง') !== -1 && text.indexOf('ขอเสนอผลงาน') !== -1) {
          cell.replaceText('\\{\\{ROUND\\}\\}', (data.round || '') + ' ของ ' + (fullAcadName || ''));
        }
      }
    }
  }

  body.replaceText('{{FISCAL_YEAR}}', data.fiscalYear || '');
  body.replaceText('{{ROUND}}', data.round || '');
  
  // Replace department head name placeholder
  var applicantEmail = data.email || Session.getActiveUser().getEmail() || '';
  var divId = getDivIdByEmail(applicantEmail);
  var deptHead = getDeptHeadName(data.dept, divId);
  body.replaceText('{{DEPT_HEAD}}', deptHead || '');
  
  // Format and replace Thai Date (e.g. 7 กรกฎาคม 2569)
  var thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];
  var today = new Date();
  var thaiDateStr = today.getDate() + ' ' + thaiMonths[today.getMonth()] + ' ' + (today.getFullYear() + 543);
  body.replaceText('{{DATE}}', thaiDateStr);
  
  // Populate publications table in Document (robust search)
  var tables = body.getTables();
  var pubTable = null;
  for (var t = 0; t < tables.length; t++) {
    var tbl = tables[t];
    if (tbl.getNumRows() > 0) {
      var firstRowText = tbl.getRow(0).getText();
      if (firstRowText.indexOf('ลำดับ') !== -1 || firstRowText.indexOf('ผลงาน') !== -1 || firstRowText.indexOf('ชื่อ') !== -1) {
        pubTable = tbl;
        break;
      }
    }
  }
  if (!pubTable && tables.length > 0) {
    pubTable = tables[tables.length - 1]; // Fallback to last table
  }
  
  if (pubTable) {
    // 1. Check and remove Author column if it exists in the header row
    var authorColIdx = -1;
    if (pubTable.getNumRows() > 0) {
      var headerRow = pubTable.getRow(0);
      for (var c = 0; c < headerRow.getNumCells(); c++) {
        var cellText = headerRow.getCell(c).getText();
        if (cellText.indexOf('ผู้แต่ง') !== -1 || cellText.indexOf('ผู้เขียน') !== -1) {
          authorColIdx = c;
          break;
        }
      }
      if (authorColIdx !== -1) {
        pubTable.removeColumn(authorColIdx);
      }
    }
    
    // 2. Add 'คะแนน (%)' column to header row if not present
    if (pubTable.getNumRows() > 0) {
      var headerRow = pubTable.getRow(0);
      var hasScoreCol = false;
      for (var c = 0; c < headerRow.getNumCells(); c++) {
        if (headerRow.getCell(c).getText().indexOf('คะแนน') !== -1) {
          hasScoreCol = true;
          break;
        }
      }
      if (!hasScoreCol) {
        var scoreHeader = headerRow.appendTableCell('คะแนน (%)');
        scoreHeader.setBold(true);
        scoreHeader.setFontSize(16);
        scoreHeader.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    }
    
    // 3. Append rows dynamically with Score Column
    if (data.pubs && data.pubs.length > 0) {
      for (var i = 0; i < data.pubs.length; i++) {
        var p = data.pubs[i];
        var newRow = pubTable.appendTableRow();
        
        newRow.appendTableCell(String(i + 1));
        newRow.appendTableCell(p.title || '');
        newRow.appendTableCell(p.tier || '');
        newRow.appendTableCell(p.role ? (Array.isArray(p.role) ? p.role.join(', ') : p.role) : '');
        newRow.appendTableCell(p.score || '0%');
        
        // Apply cell styles
        for (var c = 0; c < newRow.getNumCells(); c++) {
          var cell = newRow.getCell(c);
          cell.setFontSize(16);
          cell.setBold(false); // Explicitly remove bold styling inherited from the header
          if (c === 0 || c === 2 || c === 3 || c === 4) {
            cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          }
        }
      }
    }
    
    // 4. Append summary row at the bottom of the table
    var summaryRow = pubTable.appendTableRow();
    summaryRow.appendTableCell('');
    summaryRow.appendTableCell('');
    summaryRow.appendTableCell('');
    var labelCell = summaryRow.appendTableCell('คะแนนสะสมรวม:');
    labelCell.setFontSize(16);
    labelCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    
    var passStr = data.passStatus || 'ไม่ผ่านเกณฑ์';
    var totalPctStr = (data.totalPct !== undefined ? data.totalPct : '0') + '%';
    var totalCell = summaryRow.appendTableCell(totalPctStr + ' (' + passStr + ')');
    totalCell.setFontSize(16);
    totalCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }
  
  doc.saveAndClose();
  
  // Export as DOCX
  var url = 'https://docs.google.com/feeds/download/documents/export/Export?id=' + docId + '&exportFormat=docx';
  var response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    docFile.setTrashed(true);
    throw new Error('ไม่สามารถแปลงเอกสารเป็น DOCX ได้: ' + response.getContentText());
  }
  
  var blob = response.getBlob();
  var base64 = Utilities.base64Encode(blob.getBytes());
  
  // Clean up temporary file from Drive
  docFile.setTrashed(true);
  
  return {
    fileName: copyName + '.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    base64: base64
  };
}

function getOrCreateSubfolder_(parentFolder, name) {
  var folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(name);
  }
}

function checkSubmissionStatus(email, fiscalYear, round) {
  var submissionSheetId = '1Nyp_b-tPRxSWSE7H_WaiiFrIxjxqqQmrA0ooV1kmaZ8';
  var ss;
  try {
    ss = SpreadsheetApp.openById(submissionSheetId);
  } catch (e) {
    return { submitted: false };
  }
  
  var mainSheet = ss.getSheetByName('Submissions');
  if (!mainSheet) return { submitted: false };
  
  var mainData = mainSheet.getDataRange().getValues();
  for (var r = 1; r < mainData.length; r++) {
    var rowEmail = mainData[r][1]; // Column 2: Email
    var rowFY = String(mainData[r][8]); // Column 9: Fiscal_Year
    var rowRound = String(mainData[r][9]); // Column 10: Round
    
    if (String(rowEmail).toLowerCase().trim() === String(email || '').toLowerCase().trim() &&
        String(rowFY).trim() === String(fiscalYear).trim() &&
        String(rowRound).trim() === String(round).trim()) {
      return {
        submitted: true,
        timestamp: mainData[r][0] instanceof Date ? mainData[r][0].toISOString() : String(mainData[r][0]), // Column 1: Timestamp
        totalScore: mainData[r][10], // Column 11: Total_Score
        passStatus: mainData[r][11] // Column 12: Pass_Status
      };
    }
  }
  return { submitted: false };
}

function setCellRichTextLinks_(sheet, row, col, urls) {
  var range = sheet.getRange(row, col);
  if (!urls || urls.length === 0) {
    range.setValue('ไม่มีไฟล์แนบ');
    return;
  }
  
  var textParts = [];
  var linkRanges = [];
  var currentLength = 0;
  
  for (var i = 0; i < urls.length; i++) {
    var label = urls[i].label || 'ไฟล์แนบ';
    var url = urls[i].url;
    var separator = (i > 0 ? ', ' : '');
    
    var startIdx = currentLength + separator.length;
    var endIdx = startIdx + label.length;
    
    textParts.push(separator + label);
    currentLength += separator.length + label.length;
    
    if (url) {
      linkRanges.push({ start: startIdx, end: endIdx, url: url });
    }
  }
  
  var fullText = textParts.join('');
  var richText = SpreadsheetApp.newRichTextValue().setText(fullText);
  for (var j = 0; j < linkRanges.length; j++) {
    var lr = linkRanges[j];
    richText.setLinkUrl(lr.start, lr.end, lr.url);
  }
  range.setRichTextValue(richText.build());
}

function getUserSubmissionHistory(email) {
  var activeEmail = email || Session.getActiveUser().getEmail();
  if (!activeEmail) throw new Error('ไม่พบข้อมูลอีเมลผู้ใช้');
  
  var submissionSheetId = '1Nyp_b-tPRxSWSE7H_WaiiFrIxjxqqQmrA0ooV1kmaZ8';
  var ss;
  try {
    ss = SpreadsheetApp.openById(submissionSheetId);
  } catch (e) {
    throw new Error('ไม่สามารถเข้าถึงฐานข้อมูลสเปรดชีต (ID: ' + submissionSheetId + ') ได้: ' + e.message);
  }
  
  var mainSheet = ss.getSheetByName('Submissions');
  if (!mainSheet) {
    throw new Error('ไม่พบแผ่นงาน "Submissions" ในสเปรดชีตฐานข้อมูล');
  }
  
  var mainData = mainSheet.getDataRange().getValues();
  var history = [];
  
  // Columns: 'Timestamp', 'Email', 'Name', 'Academic_Title', 'Department', 
  // 'Staff_Type', 'Appoint_Date', 'Field_Group', 'Fiscal_Year', 
  // 'Round', 'Total_Score', 'Pass_Status'
  for (var r = 1; r < mainData.length; r++) {
    var rowEmail = mainData[r][1]; // Column 2: Email
    if (String(rowEmail).toLowerCase().trim() === String(activeEmail).toLowerCase().trim()) {
      var ts = mainData[r][0];
      var tsStr = ts instanceof Date ? ts.toISOString() : String(ts || '');
      var appoint = mainData[r][6];
      var appointStr = appoint instanceof Date ? appoint.toISOString().split('T')[0] : String(appoint || '');
      history.push({
        timestamp: tsStr,
        email: rowEmail,
        fullNameTh: mainData[r][2],
        position: mainData[r][3],
        dept: mainData[r][4],
        staffType: mainData[r][5],
        appointDate: appointStr,
        fieldGroup: mainData[r][7],
        fiscalYear: String(mainData[r][8] || '').trim(),
        round: String(mainData[r][9] || '').trim(),
        totalScore: mainData[r][10],
        passStatus: mainData[r][11]
      });
    }
  }
  
  // Sort by timestamp descending
  history.sort(function(a, b) {
    return new Date(b.timestamp) - new Date(a.timestamp);
  });
  
  return history;
}

function getSubmissionDetails(email, fiscalYear, round) {
  var activeEmail = email || Session.getActiveUser().getEmail();
  if (!activeEmail) throw new Error('ไม่พบข้อมูลอีเมลผู้ใช้');
  
  var submissionSheetId = '1Nyp_b-tPRxSWSE7H_WaiiFrIxjxqqQmrA0ooV1kmaZ8';
  var ss;
  try {
    ss = SpreadsheetApp.openById(submissionSheetId);
  } catch (e) {
    throw new Error('ไม่สามารถเข้าถึงฐานข้อมูลสเปรดชีต (ID: ' + submissionSheetId + ') ได้: ' + e.message);
  }
  
  var pubSheet = ss.getSheetByName('Publications');
  if (!pubSheet) {
    throw new Error('ไม่พบแผ่นงาน "Publications" ในสเปรดชีตฐานข้อมูล');
  }
  
  var pubData = pubSheet.getDataRange().getValues();
  var pubs = [];
  
  // Columns: Timestamp, Email, Fiscal_Year, Round, Title, Section, Tier, Role, Venue, Volume, Year, DOI, Drive_URLs
  var queryFYVal = parseInt(fiscalYear, 10);
  var queryRoundVal = parseInt(round, 10);
  
  for (var r = 1; r < pubData.length; r++) {
    var rowEmail = pubData[r][1];
    var rowFYVal = parseInt(pubData[r][2], 10);
    var rowRoundVal = parseInt(pubData[r][3], 10);
    
    if (String(rowEmail).toLowerCase().trim() === String(activeEmail).toLowerCase().trim() &&
        rowFYVal === queryFYVal &&
        rowRoundVal === queryRoundVal) {
      
      var range = pubSheet.getRange(r + 1, 13); // Column 13: Drive_URLs
      var richText = range.getRichTextValue();
      var driveUrls = [];
      
      if (richText) {
        var text = richText.getText();
        var runs = richText.getRuns();
        for (var i = 0; i < runs.length; i++) {
          var run = runs[i];
          var url = run.getLinkUrl();
          var runText = run.getText().replace(/^[,\s]+|[,\s]+$/g, '').trim();
          if (url && runText) {
            driveUrls.push({ label: runText, url: url });
          }
        }
        
        if (driveUrls.length === 0 && text && text !== 'ไม่มีไฟล์แนบ') {
          driveUrls.push({ label: 'ไฟล์แนบ', url: text });
        }
      } else {
        var val = String(pubData[r][12] || '');
        if (val && val !== 'ไม่มีไฟล์แนบ') {
          driveUrls.push({ label: 'ไฟล์แนบ', url: val });
        }
      }
      
      pubs.push({
        title: pubData[r][4],
        section: pubData[r][5],
        tier: pubData[r][6],
        role: pubData[r][7] ? String(pubData[r][7]).split(', ') : [],
        venue: pubData[r][8],
        volume: pubData[r][9],
        year: pubData[r][10],
        doi: pubData[r][11],
        driveUrls: driveUrls
      });
    }
  }
  
  return pubs;
}

function generateDocxMemoForSubmission(email, fiscalYear, round) {
  var activeEmail = email || Session.getActiveUser().getEmail();
  if (!activeEmail) throw new Error('ไม่พบข้อมูลผู้ใช้งาน');
  
  var submissionSheetId = '1Nyp_b-tPRxSWSE7H_WaiiFrIxjxqqQmrA0ooV1kmaZ8';
  var ss = SpreadsheetApp.openById(submissionSheetId);
  var mainSheet = ss.getSheetByName('Submissions');
  if (!mainSheet) throw new Error('ไม่พบแผ่นงาน Submissions');
  
  var mainData = mainSheet.getDataRange().getValues();
  var subRecord = null;
  var queryFYVal = parseInt(fiscalYear, 10);
  var queryRoundVal = parseInt(round, 10);
  
  var totalScore = 0;
  var passStatus = 'ไม่ผ่านเกณฑ์';
  for (var r = 1; r < mainData.length; r++) {
    var rowEmail = mainData[r][1];
    var rowFYVal = parseInt(mainData[r][8], 10);
    var rowRoundVal = parseInt(mainData[r][9], 10);
    
    if (String(rowEmail).toLowerCase().trim() === String(activeEmail).toLowerCase().trim() &&
        rowFYVal === queryFYVal &&
        rowRoundVal === queryRoundVal) {
      var appoint = mainData[r][6];
      var appointStr = appoint instanceof Date ? appoint.toISOString().split('T')[0] : String(appoint || '');
      var fieldGroup = String(mainData[r][7] || 'sci_tech').trim();
      totalScore = parseFloat(mainData[r][10] || 0);
      passStatus = String(mainData[r][11] || 'ไม่ผ่านเกณฑ์').trim();
      
      var roleType = mainData[r].length > 12 ? String(mainData[r][12] || '').trim() : '';
      
      subRecord = {
        email: activeEmail,
        fullNameTh: mainData[r][2],
        position: mainData[r][3],
        dept: mainData[r][4],
        staffType: mainData[r][5],
        appointDate: formatThaiDate(appointStr),
        fieldGroup: fieldGroup,
        fiscalYear: String(mainData[r][8] || '').trim(),
        round: String(mainData[r][9] || '').trim(),
        totalPct: Math.round(totalScore * 100),
        passStatus: passStatus,
        roleType: roleType
      };
      break;
    }
  }
  
  if (!subRecord) {
    throw new Error('ไม่พบข้อมูลคำขอนี้ในระบบ');
  }
  
  subRecord.pubs = getSubmissionDetails(activeEmail, fiscalYear, round);
  
  // Use saved roleType directly if available, otherwise fall back to deduction
  var isAdmin = false;
  if (subRecord.roleType === 'ผู้บริหาร') {
    isAdmin = true;
  } else if (subRecord.roleType === 'อาจารย์ประจำ') {
    isAdmin = false;
  } else {
    isAdmin = determineIsAdmin(subRecord.pubs, subRecord.position, subRecord.fieldGroup, totalScore);
  }
  
  // Calculate publication scores
  for (var i = 0; i < subRecord.pubs.length; i++) {
    var p = subRecord.pubs[i];
    var scoreVal = getPubScore(p, subRecord.position, isAdmin, subRecord.fieldGroup);
    p.score = scoreVal + '%';
  }
  
  return generateDocxMemo(subRecord);
}

function getSubmissionDebugInfo(email) {
  var activeEmail = email || Session.getActiveUser().getEmail();
  var submissionSheetId = '1Nyp_b-tPRxSWSE7H_WaiiFrIxjxqqQmrA0ooV1kmaZ8';
  var info = {
    queriedEmail: activeEmail,
    sessionEmail: Session.getActiveUser().getEmail(),
    effectiveEmail: Session.getEffectiveUser().getEmail(),
    sheetExists: false,
    submissionsCount: 0,
    matchedCount: 0,
    allEmailsInSheet: [],
    pubSheetExists: false,
    pubCount: 0,
    pubMatchedCount: 0,
    pubSamples: []
  };
  
  try {
    var ss = SpreadsheetApp.openById(submissionSheetId);
    var mainSheet = ss.getSheetByName('Submissions');
    if (mainSheet) {
      info.sheetExists = true;
      var mainData = mainSheet.getDataRange().getValues();
      info.submissionsCount = Math.max(0, mainData.length - 1);
      
      var emails = {};
      for (var r = 1; r < mainData.length; r++) {
        var rowEmail = String(mainData[r][1]).toLowerCase().trim();
        emails[rowEmail] = (emails[rowEmail] || 0) + 1;
        if (rowEmail === String(activeEmail).toLowerCase().trim()) {
          info.matchedCount++;
        }
      }
      info.allEmailsInSheet = Object.keys(emails);
    }
    
    var pubSheet = ss.getSheetByName('Publications');
    if (pubSheet) {
      info.pubSheetExists = true;
      var pubData = pubSheet.getDataRange().getValues();
      info.pubCount = Math.max(0, pubData.length - 1);
      
      for (var p = 1; p < pubData.length; p++) {
        var pEmail = String(pubData[p][1]).toLowerCase().trim();
        var pFY = String(pubData[p][2]).trim();
        var pRound = String(pubData[p][3]).trim();
        
        if (pEmail === String(activeEmail).toLowerCase().trim()) {
          info.pubMatchedCount++;
          if (info.pubSamples.length < 3) {
            info.pubSamples.push({
              title: pubData[p][4],
              fy: pFY,
              round: pRound,
              email: pEmail
            });
          }
        }
      }
    }
  } catch (e) {
    info.error = e.message;
  }
  return info;
}

function formatThaiDate(dateStr) {
  if (!dateStr || dateStr === '—' || dateStr === '') return '—';
  var parts = dateStr.split('-');
  if (parts.length !== 3) return dateStr;

  var year = parseInt(parts[0], 10);
  var monthIdx = parseInt(parts[1], 10) - 1;
  var day = parseInt(parts[2], 10);

  var thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
  ];

  return day + ' ' + thaiMonths[monthIdx] + ' ' + (year + 543);
}

function getCritKeyFromTierTh(tierTh) {
  if (!tierTh) return null;
  var t = String(tierTh).trim();
  if (t.indexOf('Scopus Q1') !== -1 || t.indexOf('Scopus Q2') !== -1 || t.indexOf('Scopus Q3') !== -1 || t.indexOf('Scopus Q4') !== -1) {
    return 'scopus';
  }
  if (t.indexOf('Review Article') !== -1 || t.indexOf('Book Chapter') !== -1) {
    return 'review_article';
  }
  if (t.indexOf('ประชุมวิชาการ') !== -1 || t.indexOf('Proceedings') !== -1 || t.indexOf('full_proc') !== -1) {
    return 'full_proc';
  }
  if (t.indexOf('TCI กลุ่ม 1') !== -1 || t.indexOf('TCI 1') !== -1) {
    return 'tci_1';
  }
  if (t.indexOf('อนุสิทธิบัตร') !== -1) {
    return 'petty_pat';
  }
  if (t.indexOf('สิทธิบัตร') !== -1) {
    return 'patent';
  }
  if (t.indexOf('งานสร้างสรรค์/รางวัล (นานาชาติ)') !== -1 || t.indexOf('ศิลปกรรม/รางวัล (นานาชาติ)') !== -1) {
    return 'artwork_intl';
  }
  if (t.indexOf('งานสร้างสรรค์/รางวัล (ชาติ)') !== -1 || t.indexOf('ศิลปกรรม/รางวัล (ชาติ)') !== -1) {
    return 'artwork_nat';
  }
  return null;
}

function getQRankFromTierTh(tierTh) {
  var t = String(tierTh).trim();
  if (t.indexOf('Scopus Q1') !== -1) return 4;
  if (t.indexOf('Scopus Q2') !== -1) return 3;
  if (t.indexOf('Scopus Q3') !== -1) return 2;
  if (t.indexOf('Scopus Q4') !== -1) return 1;
  return 0;
}

function minQRank(minQ) {
  if (!minQ || minQ === 'any') return 0;
  var n = parseInt(minQ.replace('Q', '').replace('+', ''));
  return 5 - n; // Q1->4, Q2->3, Q3->2, Q4->1
}

var CRITERIA = {
  sci_tech: {
    scopus: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['Q4+', 1, 1], 'รองศาสตราจารย์': ['Q3+', 2, 1], 'ศาสตราจารย์': ['Q2+', 3, 1], 'ศาสตราจารย์ต่อเวลา': ['Q2+', 5, null] },
    review_article: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 1, 1], 'รองศาสตราจารย์': ['any', 2, 1], 'ศาสตราจารย์': ['any', 3, 1], 'ศาสตราจารย์ต่อเวลา': null },
    full_proc: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 2, 1], 'รองศาสตราจารย์': ['any', 4, 2], 'ศาสตราจารย์': ['any', 6, 3], 'ศาสตราจารย์ต่อเวลา': null },
    patent: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 1, 1], 'รองศาสตราจารย์': ['any', 2, 1], 'ศาสตราจารย์': ['any', 3, 1], 'ศาสตราจารย์ต่อเวลา': null },
    petty_pat: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 2, 1], 'รองศาสตราจารย์': ['any', 4, 2], 'ศาสตราจารย์': ['any', 6, 3], 'ศาสตราจารย์ต่อเวลา': null },
  },
  biz_social: {
    scopus: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['Q4+', 1, 1], 'รองศาสตราจารย์': ['Q3+', 2, 1], 'ศาสตราจารย์': ['Q2+', 3, 1], 'ศาสตราจารย์ต่อเวลา': ['Q2+', 5, null] },
    review_article: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 1, 1], 'รองศาสตราจารย์': ['any', 2, 1], 'ศาสตราจารย์': ['any', 3, 1], 'ศาสตราจารย์ต่อเวลา': null },
    full_proc: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 2, 1], 'รองศาสตราจารย์': ['any', 4, 2], 'ศาสตราจารย์': ['any', 6, 3], 'ศาสตราจารย์ต่อเวลา': null },
    artwork_intl: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 1, 1], 'รองศาสตราจารย์': ['any', 2, 1], 'ศาสตราจารย์': ['any', 3, 1], 'ศาสตราจารย์ต่อเวลา': null },
    artwork_nat: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 2, 1], 'รองศาสตราจารย์': ['any', 4, 2], 'ศาสตราจารย์': ['any', 6, 2], 'ศาสตราจารย์ต่อเวลา': null },
    tci_1: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 2, 1], 'รองศาสตราจารย์': ['any', 4, 2], 'ศาสตราจารย์': null, 'ศาสตราจารย์ต่อเวลา': null },
    patent: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 1, 1], 'รองศาสตราจารย์': ['any', 2, 1], 'ศาสตราจารย์': ['any', 3, 1], 'ศาสตราจารย์ต่อเวลา': null },
    petty_pat: { 'อาจารย์': ['any', 1, 1], 'ผู้ช่วยศาสตราจารย์': ['any', 2, 1], 'รองศาสตราจารย์': ['any', 4, 2], 'ศาสตราจารย์': ['any', 6, 3], 'ศาสตราจารย์ต่อเวลา': null },
  },
};

function getPubScore(p, position, isAdmin, fieldGroup) {
  if (!p.tier || !p.role || !p.role.length) return 0;
  var critKey = getCritKeyFromTierTh(p.tier);
  if (!critKey) return 0;

  var fg = CRITERIA[fieldGroup] || CRITERIA.sci_tech;
  var typeCrit = fg[critKey];
  if (!typeCrit) return 0;

  var posCrit = typeCrit[position];
  if (!posCrit) return 0;

  if (critKey === 'scopus') {
    var need = minQRank(posCrit[0]);
    var have = getQRankFromTierTh(p.tier);
    if (have < need) return 0;
  }

  var target = isAdmin ? (posCrit[2] !== null && posCrit[2] !== undefined ? posCrit[2] : posCrit[1]) : posCrit[1];
  if (!target) return 0;

  return Math.round(100 / target);
}

function calculateTotalScore(pubs, position, isAdmin, fieldGroup) {
  if (!pubs || pubs.length === 0) return 0;
  
  var fg = CRITERIA[fieldGroup] || CRITERIA.sci_tech;
  var grouped = {};
  
  pubs.forEach(function (p) {
    if (!p.tier || !p.role || !p.role.length) return;
    var critKey = getCritKeyFromTierTh(p.tier);
    if (!critKey) return;
    
    var typeCrit = fg[critKey];
    if (!typeCrit) return;
    
    var posCrit = typeCrit[position];
    if (!posCrit) return;
    
    if (critKey === 'scopus') {
      var need = minQRank(posCrit[0]);
      var have = getQRankFromTierTh(p.tier);
      if (have < need) return;
    }
    
    if (!grouped[critKey]) grouped[critKey] = [];
    grouped[critKey].push(p);
  });
  
  var totalScore = 0;
  Object.keys(grouped).forEach(function (critKey) {
    var items = grouped[critKey];
    var typeCrit = fg[critKey];
    var posCrit = typeCrit[position];
    if (!posCrit) return;
    
    var target = isAdmin ? (posCrit[2] !== null && posCrit[2] !== undefined ? posCrit[2] : posCrit[1]) : posCrit[1];
    var contrib = items.length / target;
    totalScore += contrib;
  });
  
  return totalScore;
}

function determineIsAdmin(pubs, position, fieldGroup, savedTotalScore) {
  var scoreWithIsAdminFalse = calculateTotalScore(pubs, position, false, fieldGroup);
  if (Math.abs(scoreWithIsAdminFalse - savedTotalScore) < 0.01) {
    return false;
  }
  var scoreWithIsAdminTrue = calculateTotalScore(pubs, position, true, fieldGroup);
  if (Math.abs(scoreWithIsAdminTrue - savedTotalScore) < 0.01) {
    return true;
  }
  return false;
}

function getDivIdByEmail(email) {
  if (!email) return '';
  var sheetId = '1Ljo361CBa1p9d79V4dxOGkttubCo2TnN2eiHwSAKFkk';
  var ss = SpreadsheetApp.openById(sheetId);
  
  // Find Staff sheet robustly
  var sheets = ss.getSheets();
  var staffSheet = null;
  for (var s = 0; s < sheets.length; s++) {
    var sName = sheets[s].getName().trim().toLowerCase();
    if (sName === 'staff' || sName.indexOf('staff') !== -1) {
      staffSheet = sheets[s];
      break;
    }
  }
  if (!staffSheet) return '';
  
  var staffData = staffSheet.getDataRange().getValues();
  if (staffData.length <= 1) return '';
  
  var headers = staffData[0];
  
  function findColIndex(list, terms) {
    for (var i = 0; i < list.length; i++) {
      var val = String(list[i]).trim().toLowerCase();
      if (!val) continue;
      for (var j = 0; j < terms.length; j++) {
        var term = terms[j].toLowerCase();
        if (val === term || val.indexOf(term) !== -1 || term.indexOf(val) !== -1) {
          return i;
        }
      }
    }
    return -1;
  }
  
  var emailIdx = findColIndex(headers, ['email', 'อีเมล', 'mail']);
  var divIdIdx = findColIndex(headers, ['div_id', 'divid', 'div', 'รหัสภาควิชา', 'รหัสภาค', 'รหัสหน่วยงาน']);
  
  if (emailIdx === -1 || divIdIdx === -1) return '';
  
  var qEmail = String(email).trim().toLowerCase();
  for (var i = 1; i < staffData.length; i++) {
    var row = staffData[i];
    if (String(row[emailIdx]).trim().toLowerCase() === qEmail) {
      return String(row[divIdIdx]).trim();
    }
  }
  return '';
}

function getDeptHeadName(deptName, divId) {
  var sheetId = '1Ljo361CBa1p9d79V4dxOGkttubCo2TnN2eiHwSAKFkk';
  var ss = SpreadsheetApp.openById(sheetId);
  
  // 1. Find Division sheet
  var divSheet = ss.getSheetByName('Division') || ss.getSheetByName('division');
  if (!divSheet) return '';
  
  var divData = divSheet.getDataRange().getValues();
  if (divData.length <= 1) return '';
  
  var divHeaders = divData[0];
  
  function findColIndex(list, terms) {
    for (var i = 0; i < list.length; i++) {
      var val = String(list[i]).trim().toLowerCase();
      if (!val) continue;
      for (var j = 0; j < terms.length; j++) {
        var term = terms[j].toLowerCase();
        if (val === term || val.indexOf(term) !== -1 || term.indexOf(val) !== -1) {
          return i;
        }
      }
    }
    return -1;
  }
  
  var divIdIdx = findColIndex(divHeaders, ['div_id', 'divid', 'div', 'รหัสภาควิชา', 'รหัสภาค', 'รหัสหน่วยงาน']);
  var deptNameIdx = findColIndex(divHeaders, ['div_name_th', 'dept', 'department', 'div_name', 'divname', 'ชื่อภาควิชา', 'ภาควิชา']);
  var headEmailIdx = findColIndex(divHeaders, ['head_of_department', 'headofdepartment', 'head_email', 'heademail', 'หัวหน้าภาควิชา']);
  
  if (divIdIdx === -1 && deptNameIdx === -1) return '';
  if (headEmailIdx === -1) return '';
  
  function normalizeDept(name) {
    if (!name) return '';
    return String(name)
      .replace(/\s+/g, '')
      .replace(/ภาควิชา/g, '')
      .replace(/สาขาวิชา/g, '')
      .replace(/วิศวกรรม/g, 'วิศว')
      .trim()
      .toLowerCase();
  }
  
  function normalizeId(idVal) {
    if (!idVal) return '';
    var s = String(idVal).trim();
    if (s.indexOf('.') !== -1) {
      s = s.split('.')[0];
    }
    return s;
  }
  
  var qDept = normalizeDept(deptName);
  var qDivId = normalizeId(divId);
  var headEmail = '';
  
  // Find head's email in Division sheet
  for (var i = 1; i < divData.length; i++) {
    var row = divData[i];
    
    // Match by Div_ID
    if (divIdIdx !== -1 && qDivId) {
      var rowDivId = normalizeId(row[divIdIdx]);
      if (rowDivId === qDivId) {
        headEmail = String(row[headEmailIdx]).trim();
        break;
      }
    }
    
    // Match by Dept name
    if (deptNameIdx !== -1 && qDept) {
      var rowDept = normalizeDept(row[deptNameIdx]);
      if (rowDept && (qDept.indexOf(rowDept) !== -1 || rowDept.indexOf(qDept) !== -1)) {
        headEmail = String(row[headEmailIdx]).trim();
        break;
      }
    }
  }
  
  if (!headEmail) return '';
  
  // 2. Find Staff sheet robustly
  var sheets = ss.getSheets();
  var staffSheet = null;
  for (var s = 0; s < sheets.length; s++) {
    var sName = sheets[s].getName().trim().toLowerCase();
    if (sName === 'staff' || sName.indexOf('staff') !== -1) {
      staffSheet = sheets[s];
      break;
    }
  }
  if (!staffSheet) return '';
  
  var staffData = staffSheet.getDataRange().getValues();
  if (staffData.length <= 1) return '';
  
  var staffHeaders = staffData[0];
  var emailIdx = findColIndex(staffHeaders, ['email', 'อีเมล', 'mail']);
  var titleIdx = findColIndex(staffHeaders, ['title', 'คำนำหน้า']);
  var acadTitleIdx = findColIndex(staffHeaders, ['academic_title', 'คำนำหน้าทางวิชาการ', 'academictitle']);
  var fullnameIdx = findColIndex(staffHeaders, ['fullname', 'fullname_th', 'ชื่อ', 'name']);
  var surnameIdx = findColIndex(staffHeaders, ['surname', 'surname_th', 'นามสกุล']);
  var posIdx = findColIndex(staffHeaders, ['position', 'ตำแหน่ง', 'academic_position']);
  
  if (emailIdx === -1 || fullnameIdx === -1) return '';
  
  var qHeadEmail = headEmail.toLowerCase();
  for (var i = 1; i < staffData.length; i++) {
    var row = staffData[i];
    var rowEmail = String(row[emailIdx]).trim().toLowerCase();
    
    if (rowEmail === qHeadEmail) {
      // Reconstruct the full academic name
      var title = titleIdx !== -1 ? String(row[titleIdx]).trim() : '';
      var acadTitle = acadTitleIdx !== -1 ? String(row[acadTitleIdx]).trim() : '';
      var fullname = String(row[fullnameIdx]).trim();
      var surname = surnameIdx !== -1 ? String(row[surnameIdx]).trim() : '';
      var position = posIdx !== -1 ? String(row[posIdx]).trim() : '';
      
      if (acadTitle === 'NULL' || acadTitle === '-') acadTitle = '';
      if (title === 'NULL' || title === '-') title = '';
      
      var cleanTitle = title;
      if (/(ดร)/.test(acadTitle) || /(ดร)/.test(title)) {
        cleanTitle = 'ดร.';
      }
      
      var baseName = (cleanTitle && cleanTitle !== 'นาย' && cleanTitle !== 'นาง' && cleanTitle !== 'นางสาว' ? cleanTitle + ' ' : '') + fullname + ' ' + surname;
      baseName = baseName.replace(/\s+/g, ' ').trim();
      
      var fullAcadName = getFullAcademicName(baseName, position, headEmail);
      return expandAcademicTitle(fullAcadName);
    }
  }
  
  return '';
}

function expandAcademicTitle(nameStr) {
  if (!nameStr) return '';
  var name = String(nameStr).trim();
  
  name = name.replace(/^ศ\.ดร\./g, 'ศาสตราจารย์ ดร.');
  name = name.replace(/^รศ\.ดร\./g, 'รองศาสตราจารย์ ดร.');
  name = name.replace(/^ผศ\.ดร\./g, 'ผู้ช่วยศาสตราจารย์ ดร.');
  name = name.replace(/^ศ\./g, 'ศาสตราจารย์');
  name = name.replace(/^รศ\./g, 'รองศาสตราจารย์');
  name = name.replace(/^ผศ\./g, 'ผู้ช่วยศาสตราจารย์');
  name = name.replace(/^อ\.ดร\./g, 'อาจารย์ ดร.');
  name = name.replace(/^อ\./g, 'อาจารย์');
  
  return name;
}

function getFullAcademicName(fullNameTh, position, email) {
  if (!fullNameTh) return '';
  var name = String(fullNameTh).trim();
  
  var hasDr = /ดร\.?/.test(name);
  
  if (!hasDr && email) {
    try {
      var sheetId = '1Ljo361CBa1p9d79V4dxOGkttubCo2TnN2eiHwSAKFkk';
      var ss = SpreadsheetApp.openById(sheetId);
      var sheets = ss.getSheets();
      var staffSheet = null;
      for (var s = 0; s < sheets.length; s++) {
        var sName = sheets[s].getName().trim().toLowerCase();
        if (sName === 'staff' || sName.indexOf('staff') !== -1) {
          staffSheet = sheets[s];
          break;
        }
      }
      if (staffSheet) {
        var staffData = staffSheet.getDataRange().getValues();
        var headers = staffData[0];
        var emailIdx = -1;
        var titleIdx = -1;
        var acadTitleIdx = -1;
        for (var c = 0; c < headers.length; c++) {
          var h = String(headers[c]).trim().toLowerCase();
          if (h === 'email' || h === 'อีเมล' || h === 'mail') emailIdx = c;
          if (h === 'title' || h === 'คำนำหน้า') titleIdx = c;
          if (h === 'academic_title' || h === 'academicTitle' || h === 'คำนำหน้าทางวิชาการ' || h === 'academictitle') acadTitleIdx = c;
        }
        if (emailIdx !== -1) {
          var qEmail = String(email).trim().toLowerCase();
          for (var i = 1; i < staffData.length; i++) {
            var row = staffData[i];
            if (String(row[emailIdx]).trim().toLowerCase() === qEmail) {
              var titleVal = titleIdx !== -1 ? String(row[titleIdx]) : '';
              var acadTitleVal = acadTitleIdx !== -1 ? String(row[acadTitleIdx]) : '';
              if (/ดร/.test(titleVal) || /ดร/.test(acadTitleVal)) {
                hasDr = true;
              }
              break;
            }
          }
        }
      }
    } catch (e) {
      // Ignore lookup error and fall back
    }
  }
  
  var cleanName = name.replace(/^(นาย|นางสาว|นาง|ดร\.?|ผศ\.?|รศ\.?|ศ\.?|ผู้ช่วยศาสตราจารย์|รองศาสตราจารย์|ศาสตราจารย์|อาจารย์)\s*/g, '');
  cleanName = cleanName.replace(/^(ดร\.?)\s*/g, '').trim();
  
  var pos = String(position || '').trim();
  if (pos === 'ผู้ช่วยศาสตราจารย์' || pos === 'รองศาสตราจารย์' || pos === 'ศาสตราจารย์' || pos === 'ศาสตราจารย์ต่อเวลา') {
    var displayPos = pos === 'ศาสตราจารย์ต่อเวลา' ? 'ศาสตราจารย์' : pos;
    if (hasDr) {
      return displayPos + ' ดร.' + cleanName;
    } else {
      return displayPos + cleanName;
    }
  } else if (pos === 'อาจารย์') {
    if (hasDr) {
      return 'อาจารย์ ดร.' + cleanName;
    } else {
      return 'อาจารย์' + cleanName;
    }
  }
  
  return name;
}

