/**
 * Serves the HTML page.
 */
function doGet() {
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
      'Round', 'Total_Score', 'Pass_Status'
    ]);
    mainSheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#f3f3f3');
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
    data.passStatus ? 'ผ่านเกณฑ์' : 'ไม่ผ่านเกณฑ์'
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
              driveUrls.push(key + ': ' + driveFile.getUrl());
            } catch (err) {
              driveUrls.push(key + ': Error (' + err.message + ')');
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
        driveUrls.length > 0 ? driveUrls.join('\n') : 'ไม่มีไฟล์แนบ'
      ]);
    }
  }
  
  // 3. Send email confirmation to the submitter
  sendConfirmationEmail_(email, data.fullNameTh, data.fiscalYear, data.round, data.totalScore, data.passStatus);

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
  
  // Replace simple placeholders
  body.replaceText('{{FULL_NAME}}', data.fullNameTh || '');
  body.replaceText('{{POSITION}}', data.position || '');
  body.replaceText('{{DEPT}}', data.dept || '');
  body.replaceText('{{STAFF_TYPE}}', data.staffType || '');
  body.replaceText('{{APPOINT_DATE}}', data.appointDate || '');
  body.replaceText('{{FISCAL_YEAR}}', data.fiscalYear || '');
  body.replaceText('{{ROUND}}', data.round || '');
  
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

function sendConfirmationEmail_(email, name, fiscalYear, round, totalScore, passStatus) {
  var subject = 'ยืนยันการรับข้อมูลคำเสนอผลงานทางวิชาการ ประจำปีงบประมาณ พ.ศ. ' + fiscalYear + ' (รอบที่ ' + round + ')';
  
  var statusText = passStatus ? 'ผ่านเกณฑ์ภาระงานสะสมทางวิชาการ' : 'ไม่ผ่านเกณฑ์ภาระงานสะสมทางวิชาการ';
  var bodyTh = '<div style="font-family: \'Sarabun\', sans-serif; font-size: 14px; line-height: 1.6; color: #333;">'
    + '<p>เรียน คุณ ' + name + ',</p>'
    + '<p>ระบบขอแจ้งให้ทราบว่าได้รับข้อมูลการเสนอผลงานทางวิชาการเพื่อประกอบการขอรับเงินค่าตอบแทนพิเศษของท่าน เรียบร้อยแล้ว โดยมีข้อมูลเบื้องต้นดังนี้:</p>'
    + '<table style="border-collapse: collapse; width: 100%; margin: 15px 0; font-size: 14px;">'
    + '<tr><td style="padding: 6px; font-weight: bold; width: 150px; border-bottom: 1px solid #eee;">ปีงบประมาณ</td><td style="padding: 6px; border-bottom: 1px solid #eee;">' + fiscalYear + '</td></tr>'
    + '<tr><td style="padding: 6px; font-weight: bold; border-bottom: 1px solid #eee;">รอบที่</td><td style="padding: 6px; border-bottom: 1px solid #eee;">' + round + '</td></tr>'
    + '<tr><td style="padding: 6px; font-weight: bold; border-bottom: 1px solid #eee;">คะแนนสะสมรวม</td><td style="padding: 6px; border-bottom: 1px solid #eee;">' + Math.round(totalScore * 100) + '%</td></tr>'
    + '<tr><td style="padding: 6px; font-weight: bold; border-bottom: 1px solid #eee;">ผลการประเมินเบื้องต้น</td><td style="padding: 6px; border-bottom: 1px solid #eee; font-weight: bold; color: ' + (passStatus ? '#2e7d32' : '#c62828') + ';">' + statusText + '</td></tr>'
    + '</table>'
    + '<p>ขณะนี้คำขอของท่านอยู่ในระหว่างรอการตรวจสอบความถูกต้องโดยหัวหน้าภาควิชาและคณะต่อไป</p>'
    + '<p style="margin-top: 25px;">ขอแสดงความนับถือ,<br>คณะวิศวกรรมศาสตร์ มหาวิทยาลัยเทคโนโลยีพระจอมเกล้าพระนครเหนือ</p>'
    + '</div>';
    
  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: bodyTh
    });
  } catch (e) {
    console.error('Failed to send confirmation email: ' + e.message);
  }
}
