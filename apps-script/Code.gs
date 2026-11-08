// =====================================================
// API WRAPPER FOR VERCEL FRONTEND
// =====================================================
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonOutput_({ success: false, message: 'Invalid JSON body.' }); }

  const action = body.action;
  const args = body.args || [];
  let result;

  try {
    switch (action) {
      case 'registerStudent': result = registerStudent(args[0]); break;
      case 'loginStudent': result = loginStudent(args[0], args[1]); break;
      case 'studentGetGradeContent': result = studentGetGradeContent(args[0]); break;
      
      case 'adminLogin': result = adminLogin(args[0], args[1]); break;
      case 'adminLogout': result = adminLogout(args[0]); break;
      case 'adminGetDashboard': result = adminGetDashboard(args[0]); break;
      case 'adminSetStudentStatus': result = adminSetStudentStatus(args[0], args[1], args[2], args[3]); break;
      case 'adminSetAllStudentAccess': result = adminSetAllStudentAccess(args[0], args[1]); break;
      case 'adminDeleteStudent': result = adminDeleteStudent(args[0], args[1]); break;
      case 'adminChangeStudentPassword': result = adminChangeStudentPassword(args[0], args[1], args[2]); break;
      
      case 'adminGetGradeContent': result = adminGetGradeContent(args[0]); break;
      case 'adminSaveGradeContent': result = adminSaveGradeContent(args[0], args[1]); break;
      case 'adminDeleteGradeContent': result = adminDeleteGradeContent(args[0], args[1]); break;
      
      case 'getWebsiteContent': result = getWebsiteContent(); break;
      case 'adminGetWebsiteContent': result = adminGetWebsiteContent(args[0]); break;
      case 'adminSaveWebsiteContent': result = adminSaveWebsiteContent(args[0], args[1]); break;
      case 'adminDeleteWebsiteContent': result = adminDeleteWebsiteContent(args[0], args[1]); break;

      default:
        result = { success: false, message: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, message: String(err && err.message ? err.message : err) };
  }

  return jsonOutput_(result);
}

function doGet(e) {
  return jsonOutput_({ ok: true, message: 'GEN Z SCIENCE API is running.' });
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// PASTE YOUR ORIGINAL CODE.GS FUNCTIONS BELOW THIS LINE
// =====================================================

// (ඔයාගේ පරණ Code.gs එකේ තියෙන අනිත් හැම ෆන්ක්ෂන් එකක්ම මේකට යටින් පේස්ට් කරන්න)
// උදා: setupRegistrationDatabase, registerStudent, loginStudent...
function doGet(e) {
  var page = e && e.parameter ? String(e.parameter.page || '') : '';
  var file = page === 'admin' ? 'Admin' : 'index';
  var title = page === 'admin' ? 'GEN Z SCIENCE | ADMIN PANEL' : 'GEN Z SCIENCE | SARINDU AKALANKA';
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Run once from the Apps Script editor to create/check the registration database.
function setupRegistrationDatabase() {
  var properties = PropertiesService.getScriptProperties();
  var savedId = properties.getProperty('GENZ_REGISTRATION_SPREADSHEET_ID');
  var spreadsheet;

  if (savedId) {
    spreadsheet = SpreadsheetApp.openById(savedId);
  } else {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) spreadsheet = SpreadsheetApp.create('GEN Z SCIENCE - Student Registrations');
    properties.setProperty('GENZ_REGISTRATION_SPREADSHEET_ID', spreadsheet.getId());
  }

  var sheet = spreadsheet.getSheetByName('Registrations');
  if (!sheet) sheet = spreadsheet.insertSheet('Registrations');
  ensureRegistrationHeaders_(sheet);

  Logger.log('Registration database ready: ' + spreadsheet.getUrl());
  return spreadsheet.getUrl();
}

function ensureRegistrationHeaders_(sheet) {
  var headers = [
    'Timestamp', 'First Name', 'Last Name', 'WhatsApp Number', 'Grade',
    'District', 'Address', 'Date of Birth', 'Email Address',
    'Password Salt', 'Password Hash', 'Payment Status', 'Access Status'
  ];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return;
  }
  var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  // Migrate the earlier 10-column layout by inserting Grade after WhatsApp.
  if (existing[3] === 'WhatsApp Number' && existing[4] === 'District') {
    sheet.insertColumnBefore(5);
    sheet.getRange(1, 5).setValue('Grade');
    existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
  }
  headers.forEach(function(header, index) {
    if (!existing[index]) sheet.getRange(1, index + 1).setValue(header);
  });
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  var dataLastRow = sheet.getLastRow();
  if (dataLastRow >= 2) {
    var accessValues = sheet.getRange(2, 13, dataLastRow - 1, 1).getValues();
    var needsAccessUpdate = false;
    accessValues.forEach(function(r) { if (!r[0]) { r[0] = 'Active'; needsAccessUpdate = true; } });
    if (needsAccessUpdate) sheet.getRange(2, 13, dataLastRow - 1, 1).setValues(accessValues);
  }
}

// Registration endpoint. Passwords are stored as salted hashes, never as plain text.
function registerStudent(data) {
  if (!data || typeof data !== 'object') throw new Error('Invalid registration details.');

  var fields = ['firstName', 'lastName', 'whatsapp', 'grade', 'district', 'address', 'dateOfBirth', 'email', 'password'];
  fields.forEach(function(field) {
    if (data[field] === undefined || data[field] === null || String(data[field]).trim() === '') {
      throw new Error('Please complete all required fields.');
    }
  });

  var firstName = cleanRegistrationValue_(data.firstName, 60);
  var lastName = cleanRegistrationValue_(data.lastName, 60);
  var whatsapp = String(data.whatsapp).trim().replace(/[\r\n]/g, '');
  var grade = String(data.grade || '').trim();
  if (!/^(6|7|8|9|10|11)$/.test(grade)) throw new Error('Please select a valid grade (6 to 11).');
  var district = cleanRegistrationValue_(data.district, 40);
  var address = cleanRegistrationValue_(data.address, 250);
  var dateOfBirth = cleanRegistrationValue_(data.dateOfBirth, 10);
  var email = String(data.email).trim().toLowerCase();
  var password = String(data.password);

  if (!/^[0-9+()\s-]{7,20}$/.test(whatsapp) || normalizePhone_(whatsapp).length < 9) {
    throw new Error('Please enter a valid WhatsApp number.');
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 120) {
    throw new Error('Please enter a valid email address.');
  }
  if (password.length < 8 || password.length > 128) {
    throw new Error('Password must contain 8 to 128 characters.');
  }

  var dobParts = dateOfBirth.split('-');
  var dobYear = Number(dobParts[0]);
  var dobMonth = Number(dobParts[1]);
  var dobDay = Number(dobParts[2]);
  var dobCheck = new Date(dobYear, dobMonth - 1, dobDay);
  var todayText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateOfBirth) ||
      dobCheck.getFullYear() !== dobYear || dobCheck.getMonth() !== dobMonth - 1 ||
      dobCheck.getDate() !== dobDay || dateOfBirth > todayText) {
    throw new Error('Please enter a valid date of birth.');
  }

  var validDistricts = [
    'Ampara', 'Anuradhapura', 'Badulla', 'Batticaloa', 'Colombo', 'Galle',
    'Gampaha', 'Hambantota', 'Jaffna', 'Kalutara', 'Kandy', 'Kegalle',
    'Kilinochchi', 'Kurunegala', 'Mannar', 'Matale', 'Matara', 'Monaragala',
    'Mullaitivu', 'Nuwara Eliya', 'Polonnaruwa', 'Puttalam', 'Ratnapura',
    'Trincomalee', 'Vavuniya'
  ];
  if (validDistricts.indexOf(district) === -1) throw new Error('Please select a valid district.');

  var sheet = getRegistrationSheet_();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var rows = sheet.getDataRange().getDisplayValues();
    var emailLower = email.toLowerCase();
    var phoneKey = normalizePhone_(whatsapp);
    for (var i = 1; i < rows.length; i++) {
      var existingEmail = String(rows[i][8] || '').trim().toLowerCase();
      var existingPhone = normalizePhone_(rows[i][3] || '');
      if (existingEmail && existingEmail === emailLower) {
        throw new Error('This email address is already registered. Please login instead.');
      }
      if (existingPhone && existingPhone === phoneKey) {
        throw new Error('This WhatsApp number is already registered. Please login instead.');
      }
    }

    var salt = Utilities.getUuid() + Utilities.getUuid();
    var passwordHash = hashStudentPassword_(password, salt);
    var row = sheet.getLastRow() + 1;
    // Format the WhatsApp column as text BEFORE writing, so the first 0 is preserved.
    sheet.getRange(row, 4).setNumberFormat('@');
    sheet.getRange(row, 1, 1, 13).setNumberFormat('@');
    sheet.getRange(row, 1, 1, 13).setValues([[
      new Date(), firstName, lastName, whatsapp, grade, district, address,
      dateOfBirth, email, salt, passwordHash, 'Pending', 'Active'
    ]]);
    sheet.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(row, 4).setNumberFormat('@');
  } finally {
    lock.releaseLock();
  }

  return { success: true, message: 'Registration saved successfully.' };
}

// Login accepts either the registered email address or WhatsApp number.
function loginStudent(identifier, password) {
  identifier = String(identifier || '').trim();
  password = String(password || '');
  if (!identifier || !password) throw new Error('Enter your email/WhatsApp number and password.');

  var sheet = getRegistrationSheet_();
  var rows = sheet.getDataRange().getDisplayValues();
  var isEmail = identifier.indexOf('@') !== -1;
  var wantedEmail = identifier.toLowerCase();
  var wantedPhone = normalizePhone_(identifier);

  for (var i = 1; i < rows.length; i++) {
    var rowEmail = String(rows[i][8] || '').trim().toLowerCase();
    var rowPhone = normalizePhone_(rows[i][3] || '');
    var matches = isEmail ? (rowEmail === wantedEmail) : (wantedPhone.length >= 9 && rowPhone === wantedPhone);
    if (!matches) continue;

    var salt = String(rows[i][9] || '');
    var savedHash = String(rows[i][10] || '');
    // Old registrations created before password login was added have no password hash.
    if (!salt || !savedHash) {
      throw new Error('This account was registered before password login was added. Please contact the class administrator to activate login.');
    }
    var suppliedHash = hashStudentPassword_(password, salt);
    if (constantTimeEquals_(savedHash, suppliedHash)) {
      var grade = String(rows[i][4] || '').trim();
      if (!/^(6|7|8|9|10|11)$/.test(grade)) {
        throw new Error('Your account does not have a valid Grade saved. Please contact the class administrator.');
      }
      var paymentStatus = String(rows[i][11] || 'Pending').trim().toLowerCase();
      var accessStatus = String(rows[i][12] || 'Active').trim().toLowerCase();
      if (paymentStatus !== 'approved') {
        return { success: false, code: 'PAYMENT_PENDING', message: paymentStatus === 'rejected'
          ? 'Your payment has not been approved. Please contact the class administrator.'
          : 'Login successful, but your payment has not been approved yet. Please complete the payment and wait for admin approval.' };
      }
      if (accessStatus !== 'active') {
        return { success: false, code: 'DISCONNECTED', message: 'Your account access is currently disconnected. Please contact the class administrator.' };
      }
      var token = Utilities.getUuid() + Utilities.getUuid();
      CacheService.getScriptCache().put('GENZ_STUDENT_SESSION_' + token, JSON.stringify({
        rowId: i + 1, email: rowEmail, grade: grade, firstName: String(rows[i][1] || 'Student')
      }), 21600);
      return { success: true, firstName: String(rows[i][1] || 'Student'), grade: grade, token: token, expiresIn: 21600 };
    }
    break;
  }
  throw new Error('Login failed. Check your email/WhatsApp number and password.');
}


// -----------------------------------------------------------------------------
 // STUDENT DASHBOARD: grade-specific content, protected by a server-side session
 // -----------------------------------------------------------------------------
function getGradeContentSheet_() {
  var sheet = getRegistrationSheet_().getParent().getSheetByName('GradeContent');
  if (!sheet) sheet = getRegistrationSheet_().getParent().insertSheet('GradeContent');
  var headers = ['ID', 'Grade', 'Type', 'Title', 'Description', 'URL', 'Created At', 'Updated At', 'Drive File ID'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else {
    var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
    headers.forEach(function(h, i) { if (!existing[i]) sheet.getRange(1, i + 1).setValue(h); });
  }
  return sheet;
}

function setupGradeContent() {
  var sheet = getGradeContentSheet_();
  return 'Grade-specific content sheet is ready: ' + sheet.getName();
}

function requireStudent_(token) {
  token = String(token || '');
  if (token.length < 30) throw new Error('Student session expired. Please login again.');
  var raw = CacheService.getScriptCache().get('GENZ_STUDENT_SESSION_' + token);
  if (!raw) throw new Error('Student session expired. Please login again.');
  var session = JSON.parse(raw);
  var rowId = Number(session.rowId);
  var sheet = getRegistrationSheet_();
  if (!Number.isInteger(rowId) || rowId < 2 || rowId > sheet.getLastRow()) {
    CacheService.getScriptCache().remove('GENZ_STUDENT_SESSION_' + token);
    throw new Error('Student session expired. Please login again.');
  }
  var identity = sheet.getRange(rowId, 9, 1, 5).getDisplayValues()[0];
  if (String(identity[0] || '').trim().toLowerCase() !== String(session.email || '').trim().toLowerCase()) {
    CacheService.getScriptCache().remove('GENZ_STUDENT_SESSION_' + token);
    throw new Error('Student session expired. Please login again.');
  }
  var status = sheet.getRange(rowId, 12, 1, 2).getDisplayValues()[0];
  if (String(status[0] || 'Pending').toLowerCase() !== 'approved') {
    CacheService.getScriptCache().remove('GENZ_STUDENT_SESSION_' + token);
    throw new Error('Your payment is not approved. Please contact the class administrator.');
  }
  if (String(status[1] || 'Active').toLowerCase() !== 'active') {
    CacheService.getScriptCache().remove('GENZ_STUDENT_SESSION_' + token);
    throw new Error('Your account access has been disconnected by the administrator.');
  }
  return session;
}

function studentGetGradeContent(token) {
  var session = requireStudent_(token);
  var sheet = getGradeContentSheet_();
  var lastRow = sheet.getLastRow();
  var result = { grade: session.grade, firstName: session.firstName, liveClasses: [], studyMaterials: [], classDetails: [], videoLessons: [] };
  if (lastRow < 2) return result;
  var rows = sheet.getRange(2, 1, lastRow - 1, 9).getDisplayValues();
  var keys = {
    'live': 'liveClasses',
    'material': 'studyMaterials',
    'class': 'classDetails',
    'video': 'videoLessons'
  };
  rows.forEach(function(r) {
    if (String(r[1] || '').trim() !== session.grade) return;
    var key = keys[String(r[2] || '').trim().toLowerCase()];
    if (!key) return;
    result[key].push({
      id: String(r[0] || ''),
      title: String(r[3] || ''),
      description: String(r[4] || ''),
      url: String(r[5] || ''),
      fileId: String(r[8] || '')
    });
  });
  return result;
}

function adminSaveGradeContent(token, data) {
  requireAdmin_(token);
  if (!data || typeof data !== 'object') throw new Error('Invalid content details.');
  var grade = String(data.grade || '').trim();
  var type = String(data.type || '').trim().toLowerCase();
  var allowed = ['live', 'material', 'class', 'video'];
  if (!/^(6|7|8|9|10|11)$/.test(grade)) throw new Error('Select a valid Grade.');
  if (allowed.indexOf(type) === -1) throw new Error('Select a valid content type.');
  var title = cleanWebsiteText_(data.title, 120);
  var description = cleanWebsiteText_(data.description, 1500);
  var url = String(data.url || '').trim();
  var fileId = String(data.fileId || '').trim();
  if (!title) throw new Error('Enter a title.');
  if ((type === 'live' || type === 'video') && (!url || !/^(https:\/\/|http:\/\/)/i.test(url))) throw new Error('Enter a valid Zoom or YouTube link.');
  if (type === 'material' && !data.fileBase64 && !fileId) throw new Error('Upload a PDF study material.');
  if (type === 'class') url = '';
  if (data.fileBase64) {
    var mimeType = String(data.fileMimeType || '');
    if (mimeType !== 'application/pdf') throw new Error('Only PDF files are allowed.');
    var encoded = String(data.fileBase64);
    if (encoded.length > 11200000) throw new Error('PDF must be smaller than 8 MB.');
    var bytes = Utilities.base64Decode(encoded);
    if (bytes.length > 8 * 1024 * 1024) throw new Error('PDF must be smaller than 8 MB.');
    var fileName = String(data.fileName || 'study-material.pdf').replace(/[^a-zA-Z0-9._-]/g, '_').substring(0, 100);
    if (!/\.pdf$/i.test(fileName)) fileName += '.pdf';
    var folder = getGradeMaterialsFolder_();
    var file = folder.createFile(Utilities.newBlob(bytes, 'application/pdf', fileName));
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingError) {
      file.setTrashed(true);
      throw new Error('Google Drive cannot share this PDF with students. Check Drive sharing settings.');
    }
    if (fileId && fileId !== file.getId()) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (ignored) {} }
    fileId = file.getId();
    url = 'https://drive.google.com/uc?export=download&id=' + fileId;
  }
  var sheet = getGradeContentSheet_();
  var id = String(data.id || '').trim();
  var row = 0;
  if (id && sheet.getLastRow() >= 2) {
    var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) { row = i + 2; break; }
  }
  if (id && !row) throw new Error('Content item not found.');
  if (!id) { id = Utilities.getUuid(); row = sheet.getLastRow() + 1; }
  var created = sheet.getRange(row, 7).getValue() || new Date();
  sheet.getRange(row, 1, 1, 9).setValues([[id, grade, type, title, description, url, created, new Date(), fileId]]);
  return { success: true, id: id };
}

function adminDeleteGradeContent(token, id) {
  requireAdmin_(token);
  id = String(id || '').trim();
  var sheet = getGradeContentSheet_();
  if (!id || sheet.getLastRow() < 2) throw new Error('Content item not found.');
  var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      var fileId = String(sheet.getRange(i + 2, 9).getDisplayValue() || '');
      sheet.deleteRow(i + 2);
      if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (ignored) {} }
      return { success: true };
    }
  }
  throw new Error('Content item not found.');
}


function adminGetGradeContent(token) {
  requireAdmin_(token);
  var sheet = getGradeContentSheet_();
  var result = { live: [], material: [], class: [], video: [] };
  if (sheet.getLastRow() < 2) return result;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getDisplayValues();
  rows.forEach(function(r) {
    var type = String(r[2] || '').trim().toLowerCase();
    if (!result[type]) return;
    result[type].push({ id: String(r[0] || ''), grade: String(r[1] || ''), type: type,
      title: String(r[3] || ''), description: String(r[4] || ''), url: String(r[5] || ''), fileId: String(r[8] || '') });
  });
  return result;
}

function getGradeMaterialsFolder_() {
  var folders = DriveApp.getFoldersByName('GEN Z SCIENCE Study Materials');
  return folders.hasNext() ? folders.next() : DriveApp.createFolder('GEN Z SCIENCE Study Materials');
}


function getRegistrationSheet_() {
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('GENZ_REGISTRATION_SPREADSHEET_ID');
  if (!spreadsheetId) throw new Error('Registration database is not set up. The website owner must run setupRegistrationDatabase() once.');
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName('Registrations');
  if (!sheet) throw new Error('Registration sheet was not found.');
  ensureRegistrationHeaders_(sheet);
  return sheet;
}

function normalizePhone_(value) {
  var digits = String(value || '').replace(/\D/g, '');
  if (digits.indexOf('0094') === 0) digits = digits.substring(4);
  else if (digits.indexOf('94') === 0 && digits.length >= 11) digits = digits.substring(2);
  if (digits.length === 10 && digits.charAt(0) === '0') digits = digits.substring(1);
  return digits;
}

function hashStudentPassword_(password, salt) {
  // Salted SHA-256 with repeated rounds; raw passwords are never written to the Sheet.
  var value = salt + '|' + password;
  for (var i = 0; i < 1500; i++) {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
    value = Utilities.base64EncodeWebSafe(bytes);
  }
  return value;
}

function constantTimeEquals_(a, b) {
  a = String(a || '');
  b = String(b || '');
  if (a.length !== b.length) return false;
  var result = 0;
  for (var i = 0; i < a.length; i++) result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return result === 0;
}

function cleanRegistrationValue_(value, maxLength) {
  var text = String(value == null ? '' : value).trim();
  if (text.length > maxLength) text = text.substring(0, maxLength);
  if (/^[=+@]/.test(text)) text = "'" + text;
  return text;
}

// -----------------------------------------------------------------------------
// ADMIN PANEL SECURITY + DATA API
// -----------------------------------------------------------------------------
// Before first use, add these Script Properties in Project Settings:
// GENZ_ADMIN_USERNAME = your chosen admin username
// GENZ_ADMIN_PASSWORD = a strong temporary password (8+ characters)
// Then run setupAdminAccountOnce() manually from the Apps Script editor.
// The function hashes the password and deletes both temporary properties.

function setupAdminAccountOnce() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('GENZ_ADMIN_PASSWORD_HASH')) {
    throw new Error('Admin account is already configured.');
  }
  var username = String(props.getProperty('GENZ_ADMIN_USERNAME') || '').trim();
  var password = String(props.getProperty('GENZ_ADMIN_PASSWORD') || '');
  if (!username || username.length < 3 || !password || password.length < 10) {
    throw new Error('Set GENZ_ADMIN_USERNAME and GENZ_ADMIN_PASSWORD (at least 10 characters) in Script Properties first.');
  }
  var salt = Utilities.getUuid() + Utilities.getUuid();
  props.setProperties({
    GENZ_ADMIN_USERNAME: username,
    GENZ_ADMIN_SALT: salt,
    GENZ_ADMIN_PASSWORD_HASH: hashStudentPassword_(password, salt)
  });
  props.deleteProperty('GENZ_ADMIN_PASSWORD');
  return 'Admin account configured. The temporary password property has been deleted.';
}

function adminLogin(username, password) {
  username = String(username || '').trim();
  password = String(password || '');
  var props = PropertiesService.getScriptProperties();
  var savedUser = String(props.getProperty('GENZ_ADMIN_USERNAME') || '');
  var salt = String(props.getProperty('GENZ_ADMIN_SALT') || '');
  var savedHash = String(props.getProperty('GENZ_ADMIN_PASSWORD_HASH') || '');
  if (!savedUser || !salt || !savedHash) throw new Error('Admin account is not configured yet.');
  if (username !== savedUser || !constantTimeEquals_(hashStudentPassword_(password, salt), savedHash)) {
    throw new Error('Incorrect username or password.');
  }
  var token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('GENZ_ADMIN_SESSION_' + token, '1', 21600);
  return { success: true, token: token, expiresIn: 21600 };
}

function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove('GENZ_ADMIN_SESSION_' + String(token));
  return { success: true };
}

function requireAdmin_(token) {
  token = String(token || '');
  if (token.length < 30 || CacheService.getScriptCache().get('GENZ_ADMIN_SESSION_' + token) !== '1') {
    throw new Error('Admin session expired. Please login again.');
  }
}

function adminGetDashboard(token) {
  requireAdmin_(token);
  var sheet = getRegistrationSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0, students: [], districts: [], recent: [] };
  var values = sheet.getRange(2, 1, lastRow - 1, 13).getDisplayValues();
  var students = [];
  var districtCounts = {};
  values.forEach(function(r, index) {
    if (!r[1] && !r[2] && !r[8]) return;
    var student = {
      id: index + 2,
      timestamp: String(r[0] || ''),
      firstName: String(r[1] || ''),
      lastName: String(r[2] || ''),
      whatsapp: String(r[3] || ''),
      grade: String(r[4] || ''),
      district: String(r[5] || ''),
      address: String(r[6] || ''),
      dateOfBirth: String(r[7] || ''),
      email: String(r[8] || ''),
      paymentStatus: String(r[11] || 'Pending'),
      accessStatus: String(r[12] || 'Active')
    };
    students.push(student);
    var district = student.district || 'Not specified';
    districtCounts[district] = (districtCounts[district] || 0) + 1;
  });
  var districts = Object.keys(districtCounts).map(function(name) {
    return { name: name, count: districtCounts[name] };
  }).sort(function(a,b) { return b.count - a.count; });
  return { total: students.length, students: students.reverse(), districts: districts, recent: students.slice(0, 5) };
}

function adminSetStudentStatus(token, rowId, statusType, statusValue) {
  requireAdmin_(token);
  var row = Number(rowId);
  var sheet = getRegistrationSheet_();
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) throw new Error('Student record not found. Refresh and try again.');
  statusType = String(statusType || '').toLowerCase();
  statusValue = String(statusValue || '').trim();
  if (statusType === 'payment') {
    if (['Approved', 'Pending', 'Rejected'].indexOf(statusValue) === -1) throw new Error('Invalid payment status.');
    sheet.getRange(row, 12).setValue(statusValue);
  } else if (statusType === 'access') {
    if (['Active', 'Disconnected'].indexOf(statusValue) === -1) throw new Error('Invalid access status.');
    sheet.getRange(row, 13).setValue(statusValue);
  } else {
    throw new Error('Invalid status type.');
  }
  return { success: true };
}

function adminSetAllStudentAccess(token, statusValue) {
  requireAdmin_(token);
  statusValue = String(statusValue || '').trim();
  if (['Active', 'Disconnected'].indexOf(statusValue) === -1) throw new Error('Invalid access status.');
  var sheet = getRegistrationSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.getRange(2, 13, lastRow - 1, 1).setValue(statusValue);
  return { success: true, count: Math.max(0, lastRow - 1), status: statusValue };
}

function adminDeleteStudent(token, rowId) {
  requireAdmin_(token);
  var row = Number(rowId);
  var sheet = getRegistrationSheet_();
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) throw new Error('Student record not found. Refresh and try again.');
  sheet.deleteRow(row);
  return { success: true };
}

function adminChangeStudentPassword(token, rowId, newPassword) {
  requireAdmin_(token);
  var row = Number(rowId);
  newPassword = String(newPassword || '');
  var sheet = getRegistrationSheet_();
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) throw new Error('Student record not found. Refresh and try again.');
  if (newPassword.length < 8 || newPassword.length > 128) throw new Error('Password must contain 8 to 128 characters.');
  var salt = Utilities.getUuid() + Utilities.getUuid();
  sheet.getRange(row, 10, 1, 2).setValues([[salt, hashStudentPassword_(newPassword, salt)]]);
  return { success: true, message: 'Student password updated.' };
}


// -----------------------------------------------------------------------------
// WEBSITE CONTENT MANAGEMENT (Featured Classes + Video Learning Packs)
// -----------------------------------------------------------------------------
function getWebsiteContentSheet_() {
  var registrationSheet = getRegistrationSheet_();
  var spreadsheet = registrationSheet.getParent();
  var sheet = spreadsheet.getSheetByName('WebsiteContent');
  if (!sheet) sheet = spreadsheet.insertSheet('WebsiteContent');
  var headers = ['ID', 'Type', 'Title', 'Description', 'Price', 'Image URL', 'Button URL', 'Created At', 'Updated At', 'Drive File ID'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else {
    var existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getDisplayValues()[0];
    headers.forEach(function(header, index) {
      if (!existing[index]) sheet.getRange(1, index + 1).setValue(header);
    });
  }
  return sheet;
}

// Run once from the Apps Script editor to create the content sheet and image folder
// and authorize the required Google Sheets / Google Drive permissions.
function setupWebsiteContent() {
  var sheet = getWebsiteContentSheet_();
  var folder = getWebsiteImagesFolder_();
  return 'Website content is ready. Sheet: ' + sheet.getName() + '. Image folder: ' + folder.getName();
}

// Public website read endpoint. Only public display fields are returned.
function getWebsiteContent() {
  var sheet = getWebsiteContentSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { classes: [], packs: [] };
  var rows = sheet.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
  var result = { classes: [], packs: [] };
  rows.forEach(function(row) {
    var type = String(row[1] || '');
    if (type !== 'class' && type !== 'pack') return;
    var item = {
      id: String(row[0] || ''),
      type: type,
      title: String(row[2] || ''),
      description: String(row[3] || ''),
      price: String(row[4] || ''),
      imageUrl: String(row[5] || ''),
      buttonUrl: String(row[6] || '')
    };
    result[type === 'class' ? 'classes' : 'packs'].push(item);
  });
  return result;
}

function adminGetWebsiteContent(token) {
  requireAdmin_(token);
  return getWebsiteContent();
}

function adminSaveWebsiteContent(token, data) {
  requireAdmin_(token);
  if (!data || typeof data !== 'object') throw new Error('Invalid content details.');
  var type = String(data.type || '');
  if (type !== 'class' && type !== 'pack') throw new Error('Select a valid content section.');
  var title = cleanWebsiteText_(data.title, 100);
  var description = cleanWebsiteText_(data.description, 1200);
  if (!title || !description) throw new Error('Please enter a title and description.');

  var price = '';
  if (type === 'pack') {
    price = String(data.price == null ? '' : data.price).trim();
    if (!/^\d+(\.\d{1,2})?$/.test(price) || Number(price) < 0 || Number(price) > 100000000) {
      throw new Error('Enter a valid Video Learning Pack price.');
    }
  }

  var imageUrl = String(data.imageUrl || '').trim();
  var driveFileId = String(data.driveFileId || '').trim();
  if (data.imageBase64) {
    var mimeType = String(data.imageMimeType || '');
    if (!/^image\/(png|jpeg|jpg|webp|gif)$/.test(mimeType)) throw new Error('Upload a JPG, PNG, WEBP or GIF image.');
    var encoded = String(data.imageBase64);
    if (encoded.length > 4200000) throw new Error('Image is too large. Please upload an image smaller than 3 MB.');
    var bytes = Utilities.base64Decode(encoded);
    var fileName = String(data.imageFileName || 'website-image').replace(/[^a-zA-Z0-9._-]/g, '_').substring(0, 100);
    var folder = getWebsiteImagesFolder_();
    var file = folder.createFile(Utilities.newBlob(bytes, mimeType, fileName));
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingError) {
      file.setTrashed(true);
      throw new Error('Google Drive cannot make this image viewable on the public website. Check your Drive sharing settings and try again.');
    }
    imageUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200';
    driveFileId = file.getId();
  }
  if (!imageUrl) throw new Error('Please upload an image or enter an image URL.');
  if (!/^https:\/\//i.test(imageUrl)) throw new Error('Image URL must start with https://.');

  var buttonUrl = String(data.buttonUrl || '').trim();
  if (buttonUrl && !/^(https:\/\/|mailto:|tel:|#|\?)/i.test(buttonUrl)) {
    throw new Error('Enter a valid class button link.');
  }
  if (type === 'pack') buttonUrl = '';

  var sheet = getWebsiteContentSheet_();
  var id = String(data.id || '').trim();
  var now = new Date();
  var rowNumber = 0;
  if (id) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === id) { rowNumber = i + 2; break; }
      }
    }
    if (!rowNumber) throw new Error('This content item was not found. Refresh and try again.');
  } else {
    id = Utilities.getUuid();
    rowNumber = sheet.getLastRow() + 1;
  }

  var createdAt = rowNumber > 2 || (id && data.id) ? String(sheet.getRange(rowNumber, 8).getValue() || '') : '';
  if (!createdAt) createdAt = now;
  sheet.getRange(rowNumber, 1, 1, 10).setValues([[
    id, type, title, description, price, imageUrl, buttonUrl,
    createdAt, now, driveFileId
  ]]);
  return { success: true, id: id, imageUrl: imageUrl };
}

function adminDeleteWebsiteContent(token, contentId) {
  requireAdmin_(token);
  var id = String(contentId || '').trim();
  if (!id) throw new Error('Content item not found.');
  var sheet = getWebsiteContentSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('Content item not found.');
  var rows = sheet.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      var fileId = String(rows[i][9] || '');
      sheet.deleteRow(i + 2);
      if (fileId) {
        try { DriveApp.getFileById(fileId).setTrashed(true); } catch (ignored) {}
      }
      return { success: true };
    }
  }
  throw new Error('Content item not found. Refresh and try again.');
}

function getWebsiteImagesFolder_() {
  var folders = DriveApp.getFoldersByName('GEN Z SCIENCE Website Images');
  return folders.hasNext() ? folders.next() : DriveApp.createFolder('GEN Z SCIENCE Website Images');
}

function cleanWebsiteText_(value, maxLength) {
  var text = String(value == null ? '' : value).trim();
  if (text.length > maxLength) text = text.substring(0, maxLength);
  return text;
}
