// =======================================================================
// SENARAI PENTADBIR KINI DIURUSKAN SECARA DINAMIK MELALUI GOOGLE SHEET
// Skrip akan membaca tab "Senarai Admin". Jika tiada, ia akan dicipta secara auto.
// =======================================================================

// Mengambil ID secara selamat dari tetapan cloud Google
// Sila pastikan SPREADSHEET_ID, FOLDER_SIJIL_ID dan TEMPLATE_ID telah ditambah 
// dalam Project Settings > Script Properties di Google Apps Script
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const FOLDER_SIJIL_ID = PropertiesService.getScriptProperties().getProperty('FOLDER_SIJIL_ID');
const TEMPLATE_ID = PropertiesService.getScriptProperties().getProperty('TEMPLATE_ID');

/**
 * 1. WEB APP SETUP
 * Memaparkan halaman utama (Index.html) semasa dibuka untuk semua warga UniSZA.
 */
function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Kursus Research Methodology')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    try {
      return HtmlService.createTemplateFromFile('Index.html')
        .evaluate()
        .setTitle('Kursus Research Methodology')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (err2) {
      return HtmlService.createHtmlOutput("<p>Ralat kritikal: Fail Index gagal dimuatkan.</p>");
    }
  }
}

/**
 * 2. SPREADSHEET MENU 
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏆 Certificate Mail Merge')
    .addItem('1. Run Test Mode (Test Rows Only)', 'runTestModeAlert')
    .addItem('2. Run Actual Mode (Real Students Only)', 'runActualModeAlert')
    .addSeparator()
    .addItem('3. 🔄 Sync Maklumat Lanjutan (Dinamik)', 'showDynamicSyncModal')
    .addToUi();
}

function runTestModeAlert() {
  const res = processCertificates(true);
  SpreadsheetApp.getUi().alert(`[Mod Ujian] Selesai.\nDihantar: ${res.success}\nRalat: ${res.error}`);
}

function runActualModeAlert() {
  const res = processCertificates(false);
  SpreadsheetApp.getUi().alert(`[Mod Sebenar] Selesai.\nDihantar: ${res.success}\nRalat: ${res.error}`);
}

/**
 * =======================================================================
 * FUNGSI PENGURUSAN ADMIN DINAMIK (DARI SHEET "Senarai Admin")
 * =======================================================================
 */
function getAdminList() {
  // Gunakan ID selamat untuk mendapatkan spreadsheet
  if (!SPREADSHEET_ID) {
      throw new Error("SPREADSHEET_ID tiada dalam Script Properties!");
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Senarai Admin");
  
  if (!sheet) {
    sheet = ss.insertSheet("Senarai Admin");
    sheet.appendRow(["Email Pentadbir"]);
    sheet.getRange("A1").setFontWeight("bold").setBackground("#dbeafe");
    
    const currentUser = Session.getEffectiveUser().getEmail();
    if (currentUser) {
      sheet.appendRow([currentUser]);
    }
    sheet.autoResizeColumn(1);
  }
  
  const data = sheet.getDataRange().getValues();
  const admins = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      admins.push(data[i][0].toString().trim().toLowerCase());
    }
  }
  return admins;
}

/**
 * =======================================================================
 * FUNGSI SISTEM IDENTITI & KESELAMATAN (OAUTH GOOGLE)
 * =======================================================================
 */
function verifyUserAccess() {
  const email = Session.getActiveUser().getEmail();
  
  if (!email) {
    return { success: false, email: "Tidak Log Masuk", role: 'guest' };
  }
  
  const allowedAdmins = getAdminList();
  const isAdmin = allowedAdmins.indexOf(email.toLowerCase().trim()) > -1;
  
  return {
    success: true,
    email: email,
    role: isAdmin ? 'admin' : 'staff' 
  };
}

function isCurrentUserAdmin() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return false;
  
  const allowedAdmins = getAdminList();
  return allowedAdmins.indexOf(email.toLowerCase().trim()) > -1;
}

/**
 * =======================================================================
 * FUNGSI WEB TUNJANG (DIPANGGIL NATIVELY OLEH google.script.run)
 * =======================================================================
 */
function runTestModeWeb() {
  if (!isCurrentUserAdmin()) throw new Error("Akses dinafikan.");
  return processCertificates(true);
}

function runActualModeWeb(selectedMatrics) {
  if (!isCurrentUserAdmin()) throw new Error("Akses dinafikan.");
  return processCertificates(false, selectedMatrics);
}

function updateStudentWeb(studentData) {
  if (!isCurrentUserAdmin()) throw new Error("Akses dinafikan.");
  return updateStudentData(studentData);
}

function getAdminsWeb() {
  if (!isCurrentUserAdmin()) throw new Error("Akses dinafikan.");
  return getAdminList();
}

function addAdminWeb(newEmail) {
  if (!isCurrentUserAdmin()) throw new Error("Akses dinafikan.");
  newEmail = newEmail.toString().trim().toLowerCase();
  
  if (!newEmail || !newEmail.includes('@')) throw new Error("Format emel tidak sah.");
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Senarai Admin");
  const currentAdmins = getAdminList();
  
  if (currentAdmins.includes(newEmail)) {
    throw new Error("Emel sudah wujud.");
  }
  
  sheet.appendRow([newEmail]);
  return getAdminList();
}

function removeAdminWeb(emailToRemove) {
  if (!isCurrentUserAdmin()) throw new Error("Akses dinafikan.");
  emailToRemove = emailToRemove.toString().trim().toLowerCase();
  
  const currentUser = Session.getActiveUser().getEmail().toLowerCase().trim();
  if (emailToRemove === currentUser) {
    throw new Error("Tidak boleh buang diri sendiri.");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Senarai Admin");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toLowerCase() === emailToRemove) {
      sheet.deleteRow(i + 1); 
      break;
    }
  }
  
  return getAdminList();
}


/**
 * HELPER FUNCTION: FUNGSI KEBAL UNTUK CARI NAMA HEADER
 */
const cleanStr = (s) => s ? s.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase() : '';
const getColIndex = (headers, names) => {
  const nameArr = Array.isArray(names) ? names : [names];
  return headers.findIndex(h => nameArr.some(n => cleanStr(h) === cleanStr(n)));
};

/**
 * FUNGSI: PAPARKAN BORANG SYNC & LAKSANAKAN SYNC DINAMIK
 */
function showDynamicSyncModal() {
  const html = HtmlService.createTemplateFromFile('SyncModal')
    .evaluate()
    .setWidth(600)
    .setHeight(680)
    .setTitle('Tetapan Sync Maklumat Pelajar');
  SpreadsheetApp.getUi().showModalDialog(html, '🔄 Konfigurasi Sync Data Dinamik');
}

function executeDynamicSync(config) {
  try {
    const urlMatch = config.url.match(/[-\w]{25,}/);
    if (!urlMatch) throw new Error("URL Spreadsheet tidak sah.");
    const sourceId = urlMatch[0];

    const sourceDoc = SpreadsheetApp.openById(sourceId);
    const sourceSheet = sourceDoc.getSheetByName(config.sheetName);
    if (!sourceSheet) throw new Error(`Tab "${config.sheetName}" tidak dijumpai.`);

    const sourceData = sourceSheet.getDataRange().getDisplayValues();
    const sourceHeaders = sourceData[0];

    const srcMatricIdx = getColIndex(sourceHeaders, config.sourceMatricHeader);
    if (srcMatricIdx === -1) throw new Error(`Kolum rujukan Matrik tidak dijumpai.`);

    const validMappings = [];
    for (let mapping of config.mappings) {
      const srcIdx = getColIndex(sourceHeaders, mapping.sourceHeader);
      if (srcIdx !== -1) {
        mapping.srcIdx = srcIdx;
        validMappings.push(mapping);
      } else {
        throw new Error(`Kolum sumber "${mapping.sourceHeader}" tidak dijumpai.`);
      }
    }

    const sourceMap = {};
    for (let i = 1; i < sourceData.length; i++) {
      const matric = sourceData[i][srcMatricIdx].toString().trim().toUpperCase();
      if (matric) {
        sourceMap[matric] = sourceData[i];
      }
    }

    const targetSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("RMC");
    if (!targetSheet) throw new Error("Tab 'RMC' tidak dijumpai!");
    
    const targetData = targetSheet.getDataRange().getDisplayValues();
    const targetHeaders = targetData[0];

    const tgtMatricIdx = getColIndex(targetHeaders, ['<<Matric>>', 'Matric']);
    for (let mapping of validMappings) {
      const tgtIdx = getColIndex(targetHeaders, mapping.destHeader);
      if (tgtIdx !== -1) mapping.tgtIdx = tgtIdx;
    }

    let updateCount = 0;
    for (let i = 1; i < targetData.length; i++) {
      const tgtMatric = targetData[i][tgtMatricIdx].toString().trim().toUpperCase();
      if (tgtMatric && sourceMap[tgtMatric]) {
        const srcRow = sourceMap[tgtMatric];
        for (let mapping of validMappings) {
          if(mapping.tgtIdx !== undefined) {
             targetSheet.getRange(i + 1, mapping.tgtIdx + 1).setValue(srcRow[mapping.srcIdx]);
          }
        }
        updateCount++;
      }
    }
    return { success: true, count: updateCount };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * 3. DATA RETRIEVAL & UPDATE LOGIC (TETAP KE TAB "RMC")
 */
function getStudentsData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("RMC");
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];

  const noIdx = getColIndex(headers, 'NO');
  const matricIdx = getColIndex(headers, ['<<Matric>>', 'Matric', 'No. Matrik']);
  const nameIdx = getColIndex(headers, ['<<StudentName>>', 'StudentName', 'Nama']);
  const emailIdx = getColIndex(headers, ['Student Email Address', 'Email']);
  const certIdIdx = getColIndex(headers, ['<<CertID>>', 'CertID']);
  const statusIdx = getColIndex(headers, ['<<Email>>', 'Email Status']);
  
  const catIdx = getColIndex(headers, 'CATEGORY');
  const examIdx = getColIndex(headers, ['Exam Marks (40%)', 'Exam Marks 40']);
  const assignIdx = getColIndex(headers, ['Assignment Marks (60%)', 'Assignment Marks 60']);
  const totalIdx = getColIndex(headers, ['Total Marks (100%)', 'Total Marks 100']);
  const statusSijilIdx = getColIndex(headers, 'Status Sijil');
  
  const faculty1Idx = getColIndex(headers, ['FACULTY (full name)', '<<Faculty>>', 'Faculty']);
  const faculty2Idx = getColIndex(headers, ['FACULTY 2']);
  const supervisorIdx = getColIndex(headers, ['<<MainSupervisor>>', 'Main Supervisor']);
  const gradCoordIdx = getColIndex(headers, ['<<GraduateCoordinator>>', 'Graduate Coordinator']);
  const facultyPicIdx = getColIndex(headers, ['<<FacultyPIC>>', 'Faculty PIC']);
  const rmcSessionIdx = getColIndex(headers, ['<<RMCsession>>', 'RMCsession']);

  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let facultyVal = faculty1Idx > -1 ? row[faculty1Idx].toString().trim() : '';
    if (!facultyVal && faculty2Idx > -1) facultyVal = row[faculty2Idx].toString().trim();

    let totalMarksVal = totalIdx > -1 ? row[totalIdx].toString().trim() : '';
    let statusSijilVal = statusSijilIdx > -1 ? row[statusSijilIdx].toString().trim() : '';
    
    if (totalMarksVal !== '') {
      let marks = parseFloat(totalMarksVal);
      if (!isNaN(marks)) statusSijilVal = marks >= 40 ? 'Lulus' : 'Gagal';
    }

    records.push({
      no: row[noIdx] || '',
      matric: row[matricIdx] || '',
      nama: row[nameIdx] || '',
      emel: row[emailIdx] || '',
      certId: row[certIdIdx] || '',
      status: row[statusIdx] || '',
      category: catIdx > -1 ? row[catIdx] : '',
      examMarks: examIdx > -1 ? row[examIdx] : '',
      assignMarks: assignIdx > -1 ? row[assignIdx] : '',
      totalMarks: totalMarksVal,
      statusSijil: statusSijilVal,
      faculty: facultyVal,
      supervisor: supervisorIdx > -1 ? row[supervisorIdx] : '',
      gradCoord: gradCoordIdx > -1 ? row[gradCoordIdx] : '',
      facultyPic: facultyPicIdx > -1 ? row[facultyPicIdx] : '',
      rmcSession: rmcSessionIdx > -1 ? row[rmcSessionIdx] : '' 
    });
  }
  return records;
}

function updateStudentData(studentData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("RMC");
  if (!sheet) return {success: false, error: 'Tab "RMC" tidak dijumpai.'};
  
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  
  const noIdx = getColIndex(headers, 'NO');
  const matricIdx = getColIndex(headers, ['<<Matric>>', 'Matric']);
  const nameIdx = getColIndex(headers, ['<<StudentName>>', 'StudentName']);
  const emailIdx = getColIndex(headers, 'Student Email Address');
  const certIdIdx = getColIndex(headers, ['<<CertID>>', 'CertID']);
  const statusIdx = getColIndex(headers, ['<<Email>>', 'Email Status']);
  const rmcSessionIdx = getColIndex(headers, ['<<RMCsession>>', 'RMCsession']);
  const faculty1Idx = getColIndex(headers, ['FACULTY (full name)']);
  const faculty2Idx = getColIndex(headers, ['FACULTY 2']);
  const facultyIdx = faculty1Idx !== -1 ? faculty1Idx : faculty2Idx;
  
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][noIdx] === studentData.no) { rowIndex = i + 1; break; }
  }
  
  if (rowIndex === -1) return {success: false, error: 'Rekod tidak dijumpai.'};
  
  if (matricIdx > -1) sheet.getRange(rowIndex, matricIdx + 1).setValue(studentData.matric);
  if (nameIdx > -1) sheet.getRange(rowIndex, nameIdx + 1).setValue(studentData.nama);
  if (emailIdx > -1) sheet.getRange(rowIndex, emailIdx + 1).setValue(studentData.emel);
  if (certIdIdx > -1) sheet.getRange(rowIndex, certIdIdx + 1).setValue(studentData.certId);
  if (statusIdx > -1) sheet.getRange(rowIndex, statusIdx + 1).setValue(studentData.status);
  if (rmcSessionIdx > -1) sheet.getRange(rowIndex, rmcSessionIdx + 1).setValue(studentData.rmcSession);
  if (facultyIdx > -1 && studentData.faculty !== undefined) sheet.getRange(rowIndex, facultyIdx + 1).setValue(studentData.faculty);
  
  return {success: true};
}

/**
 * 4. CERTIFICATE GENERATION & CUSTOM EMAIL HTML LOGIC
 */
function processCertificates(isTestMode, selectedMatrics) {
  if (!TEMPLATE_ID || !FOLDER_SIJIL_ID || !SPREADSHEET_ID) {
    return { success: 0, error: 1, message: "Sistem belum dikonfigurasi. Sila semak Script Properties." };
  }
  
  const FROM_EMAIL = PropertiesService.getScriptProperties().getProperty('FROM_EMAIL') || 'no-reply@unisza.edu.my'; 
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("RMC");
  if (!sheet) return { success: 0, error: 0, message: "Ralat: Tab 'RMC' tidak dijumpai." };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; 

  const noIdx = getColIndex(headers, 'NO');
  const matricIdx = getColIndex(headers, ['<<Matric>>', 'Matric']);
  const nameIdx = getColIndex(headers, ['<<StudentName>>', 'StudentName']);
  const studentEmailIdx = getColIndex(headers, 'Student Email Address');
  const certIdIdx = getColIndex(headers, ['<<CertID>>', 'CertID']);
  const emailStatusIdx = getColIndex(headers, ['<<Email>>', 'Email Status']); 
  
  const rmcSessionIdx = getColIndex(headers, ['<<RMCsession>>', 'RMCsession']);
  const faculty1Idx = getColIndex(headers, ['FACULTY (full name)']);
  const faculty2Idx = getColIndex(headers, ['FACULTY 2']);
  const supervisorIdx = getColIndex(headers, ['<<MainSupervisor>>', 'Main Supervisor']);
  const gradCoordIdx = getColIndex(headers, ['<<GraduateCoordinator>>', 'Graduate Coordinator']);
  const facultyPicIdx = getColIndex(headers, ['<<FacultyPIC>>', 'Faculty PIC']);
  const totalIdx = getColIndex(headers, ['Total Marks (100%)', 'Total Marks 100']);
  const statusSijilIdx = getColIndex(headers, ['Status Sijil']);

  const svEmailIdx = getColIndex(headers, 'Main Supervisor Email');
  const gradCoordEmailIdx = getColIndex(headers, 'Graduate Coordinator Email');
  const facultyPicEmailIdx = getColIndex(headers, 'Faculty PIC Email');
  const secEmailIdx = getColIndex(headers, 'SECONDARY EMAIL ADDRESS');
  
  if (noIdx === -1 || matricIdx === -1 || nameIdx === -1 || studentEmailIdx === -1 || certIdIdx === -1) {
    return { success: 0, error: 0, message: "Ralat kolum utama tiada." };
  }
  
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const targetFolder = DriveApp.getFolderById(FOLDER_SIJIL_ID);
  
  let successCount = 0;
  let errorCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const noValue = row[noIdx] !== undefined ? row[noIdx].toString().trim() : '';
    const matric = row[matricIdx] !== undefined ? row[matricIdx].toString().trim() : '';
    const name = row[nameIdx] !== undefined ? row[nameIdx].toString().trim() : '';
    const certId = row[certIdIdx] !== undefined ? row[certIdIdx].toString().trim() : '';
    const studentEmail = row[studentEmailIdx] !== undefined ? row[studentEmailIdx].toString().trim() : '';
    
    // Elakkan crash jika kolum status emel tak wujud, sekadar abaikan
    const emailStatus = (emailStatusIdx !== -1 && row[emailStatusIdx] !== undefined) ? row[emailStatusIdx].toString().trim() : '';
    
    const rmcSession = rmcSessionIdx !== -1 ? row[rmcSessionIdx].toString().trim() : '';
    let faculty = faculty1Idx !== -1 ? row[faculty1Idx].toString().trim() : '';
    if (!faculty && faculty2Idx !== -1) faculty = row[faculty2Idx].toString().trim();
    
    const supervisor = supervisorIdx !== -1 ? row[supervisorIdx].toString().trim() : '';
    const gradCoord = gradCoordIdx !== -1 ? row[gradCoordIdx].toString().trim() : '';
    const facultyPic = facultyPicIdx !== -1 ? row[facultyPicIdx].toString().trim() : '';
    const totalMarks = totalIdx !== -1 ? row[totalIdx].toString().trim() : '';
    let statusSijil = statusSijilIdx !== -1 ? row[statusSijilIdx].toString().trim() : '';
    
    if (totalMarks !== '') {
      let marks = parseFloat(totalMarks);
      if (!isNaN(marks)) {
         statusSijil = marks >= 40 ? 'Lulus' : 'Gagal';
         if (statusSijilIdx !== -1 && row[statusSijilIdx] !== statusSijil) {
            sheet.getRange(i + 1, statusSijilIdx + 1).setValue(statusSijil);
         }
      }
    }

    const svEmail = svEmailIdx !== -1 ? row[svEmailIdx].toString().trim() : '';
    const gradCoordEmail = gradCoordEmailIdx !== -1 ? row[gradCoordEmailIdx].toString().trim() : '';
    const facultyPicEmail = facultyPicEmailIdx !== -1 ? row[facultyPicEmailIdx].toString().trim() : '';
    const secEmail = secEmailIdx !== -1 ? row[secEmailIdx].toString().trim() : '';

    if (!name || !studentEmail) continue;
    
    const isTestRow = noValue.toLowerCase().startsWith('test');
    if (isTestMode) {
      if (!isTestRow) continue; 
    } else {
      if (isTestRow) continue;  
      if (emailStatus === 'Emailed') continue; 
      if (selectedMatrics !== undefined && selectedMatrics !== null && !selectedMatrics.includes(matric)) continue;
    }
    
    let tempCopy = null;
    try {
      tempCopy = templateFile.makeCopy(`Temp_${name}`, targetFolder);
      const docCopy = DocumentApp.openById(tempCopy.getId());
      const body = docCopy.getBody();
      
      body.replaceText('<<MATRIC>>', matric);
      body.replaceText('<<Matric>>', matric); 
      body.replaceText('<<CertID>>', certId);
      body.replaceText('<<StudentName>>', name);
      
      const nameLocation = body.findText(name);
      if (nameLocation) {
        const textElement = nameLocation.getElement().asText();
        let calculatedFontSize = name.length <= 15 ? 38 : (name.length <= 22 ? 32 : (name.length <= 32 ? 24 : 18));
        textElement.setFontSize(nameLocation.getStartOffset(), nameLocation.getEndOffsetInclusive(), calculatedFontSize);
      }
      docCopy.saveAndClose();
      
      const pdfBlob = tempCopy.getAs(MimeType.PDF);
      pdfBlob.setName(`${matric}.pdf`);
      
      targetFolder.createFile(pdfBlob);
      
      tempCopy.setTrashed(true);
      tempCopy = null; 
      
      const subjectPrefix = isTestMode ? '[TEST RUN] ' : '';
      const emailSubject = `${subjectPrefix}Certificate of Completion | Research Methodology Course (${matric})`;
      
      const emailHtmlBody = `
        <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6; max-width: 650px;">
          <p>Dear <strong>${name}</strong>,</p>
          <p>Congratulations on successfully completing the Research Methodology Course (RMC).<br>Please find attached your Certificate of Completion for your records.</p>
          <h3 style="color: #1e3a8a; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px; margin-top: 30px;">Participant Details</h3>
          <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr><td style="padding: 10px; border: 1px solid #e2e8f0; width: 35%; font-weight: bold; background-color: #f8fafc;">Name</td><td style="padding: 10px; border: 1px solid #e2e8f0;">${name}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #e2e8f0; font-weight: bold; background-color: #f8fafc;">Matric No</td><td style="padding: 10px; border: 1px solid #e2e8f0;">${matric}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #e2e8f0; font-weight: bold; background-color: #f8fafc;">Course</td><td style="padding: 10px; border: 1px solid #e2e8f0;">${rmcSession || '-'}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #e2e8f0; font-weight: bold; background-color: #f8fafc;">Supervisor</td><td style="padding: 10px; border: 1px solid #e2e8f0;">${supervisor || '-'}</td></tr>
          </table>
          <h3 style="color: #1e3a8a; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px; margin-top: 30px;">Faculty Details</h3>
          <table style="width: 100%; border-collapse: collapse; margin-bottom: 25px;">
            <tr><td style="padding: 10px; border: 1px solid #e2e8f0; width: 35%; font-weight: bold; background-color: #f8fafc;">Faculty</td><td style="padding: 10px; border: 1px solid #e2e8f0;">${faculty || '-'}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #e2e8f0; font-weight: bold; background-color: #f8fafc;">Graduate Coordinator</td><td style="padding: 10px; border: 1px solid #e2e8f0;">${gradCoord || '-'}</td></tr>
            <tr><td style="padding: 10px; border: 1px solid #e2e8f0; font-weight: bold; background-color: #f8fafc;">Faculty PIC</td><td style="padding: 10px; border: 1px solid #e2e8f0;">${facultyPic || '-'}</td></tr>
          </table>
          <p>Thank you for your participation and commitment throughout the course. We hope the knowledge and skills gained will support your postgraduate studies and future research activities.</p>
          <p style="margin-top: 30px;">Best regards,<br><strong>Dr. Fairuz Binti Mohd Nasir</strong><br>Course Coordinator<br>Research Methodology Course (PPS10203)<br>Graduate School<br>Universiti Sultan Zainal Abidin (UniSZA)</p>
        </div>
      `;

      const ccEmails = [];
      if (svEmail) ccEmails.push(svEmail);
      if (gradCoordEmail) ccEmails.push(gradCoordEmail);
      if (facultyPicEmail) ccEmails.push(facultyPicEmail);
      if (secEmail) ccEmails.push(secEmail);
      
      const emailOptions = {
        from: FROM_EMAIL,
        name: 'Academic Division, UniSZA Graduate School',
        htmlBody: emailHtmlBody,
        bcc: FROM_EMAIL,
        attachments: [pdfBlob]
      };
      if (ccEmails.length > 0) emailOptions.cc = ccEmails.join(',');
      
      GmailApp.sendEmail(studentEmail, emailSubject, "Sila gunakan paparan HTML.", emailOptions);
      if (emailStatusIdx !== -1) {
          sheet.getRange(i + 1, emailStatusIdx + 1).setValue(isTestMode ? 'Test Emailed' : 'Emailed');
      }
      successCount++;
    } catch (error) {
      errorCount++;
      if (tempCopy) try { tempCopy.setTrashed(true); } catch(e) {}
      if (emailStatusIdx !== -1) {
          sheet.getRange(i + 1, emailStatusIdx + 1).setValue(`Error: ${error.toString()}`);
      }
    }
  }
  return { success: successCount, error: errorCount, message: "Done" };
}

/**
 * PDF RETRIEVAL HELPER UNTUK INDEX.HTML
 * Fungsi ini membenarkan carian ID fail PDF di dalam Google Drive berdasarkan nama Matrik
 */
function getStudentPdfId(matric) {
    if (!FOLDER_SIJIL_ID) return null;
    
    // Pastikan input selamat dan seragam
    const safeMatric = matric.toString().replace(/[^a-zA-Z0-9]/g, '');
    const targetFolder = DriveApp.getFolderById(FOLDER_SIJIL_ID);
    const files = targetFolder.searchFiles(`title = '${safeMatric}.pdf'`);
    
    if (files.hasNext()) {
        return files.next().getId();
    }
    return null;
}