// ==========================================
// Configuration
// ==========================================
const SHEET_COUNSELING = 'Counseling_Raw';
const SHEET_DC = 'DC_Raw';
const SHEET_PE = 'PE_Raw';
const SHEET_REPORT = 'Report Summary [new generated]';
const SHEET_HM = 'HM_Time_Raw';
const SHEET_HM_REPORT = 'HM Report Summary [new generated]';
const SHEET_FY_SUMMARY = 'Fiscal Year Summary [new generated]';
const SHEET_AUTH_USERS = 'Auth_Users';
const SHEET_AUTH_ROLES = 'Auth_Roles';

const AUTH_PAGES = ['dashboard', 'counseling', 'count', 'hm', 'report', 'users'];

const HM_HEADERS = [
  'Timestamp', 'DischargeDate', 'AN', 'DrugCount', 'Ward',
  'Step01_DoctorOrderTime', 'Step01_DoctorOrderDate',
  'Step02_NurseReceiveTime', 'Step02_NurseReceiveDate',
  'Step03_PharmCheckHMStart',
  'Step04_PharmConsultStart', 'Step04_PharmConsultEnd',
  'Step05_PharmEditStart', 'Step05_PharmEditEnd',
  'Step06_PharmVerifyTime',
  'Step07_DispenseStart', 'Step07_DispenseEnd',
  'Step08_PharmCheckStart', 'Step08_PharmCheckEnd',
  'Step09_WardChartTime', 'Step09_WardChartDate',
  'Step10_PharmCheck2Start', 'Step10_PharmCheck2End',
  'Step11_LatePickupTime', 'Step12_Remarks'
];

// Column mapping for counseling checkboxes (0-indexed from column D)
const COUNSELING_FIELDS = [
  // 5.1 ยาสูดพ่นทางปาก
  { key: 'Cat_5_1_InhalerOral_Ventolin_MDI', label: 'Ventolin MDI', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Berodual_MDI', label: 'Berodual MDI', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Budesonide_MDI', label: 'Budesonide MDI', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Seretide_Accuhaler', label: 'Seretide Accuhaler', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Symbicort_Turbuhaler', label: 'Symbicort Turbuhaler', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Seretide_Evohaler', label: 'Seretide Evohaler', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Spiriva_Handihaler', label: 'Spiriva Handihaler', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Spiolto_Respimat', label: 'Spiolto Respimat', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Spiriva_Respimat', label: 'Spiriva Respimat', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Anoro_Elipta', label: 'Anoro Elipta', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Trelegy_Ellipta', label: 'Trelegy Ellipta', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  { key: 'Cat_5_1_InhalerOral_Spacer', label: 'via Spacer', category: '5.1', subgroup: 'ยาสูดพ่นทางปาก' },
  // 5.1 ยาพ่นจมูก
  { key: 'Cat_5_1_NasalSpray_Rhinocort', label: 'Rhinocort Nasal Spray', category: '5.1', subgroup: 'ยาพ่นจมูก' },
  { key: 'Cat_5_1_NasalSpray_Avamys', label: 'Avamys Nasal Spray', category: '5.1', subgroup: 'ยาพ่นจมูก' },
  // 5.1 standalone nasal items
  { key: 'Cat_5_1_NasalDrop', label: 'ยาหยอดจมูก', category: '5.1', subgroup: 'ยาพ่นจมูก' },
  { key: 'Cat_5_1_NasalWash', label: 'ล้างจมูก', category: '5.1', subgroup: 'ยาพ่นจมูก' },
  { key: 'Cat_5_1_Miacalcic_Nasal_Spray', label: 'Miacalcic Nasal Spray', category: '5.1', subgroup: 'ยาพ่นจมูก' },
  // 5.1 ยาฉีดเบาหวาน
  { key: 'Cat_5_1_DiabetesInj_Insulin_Syringe', label: 'Insulin Syringe', category: '5.1', subgroup: 'ยาฉีดเบาหวาน' },
  { key: 'Cat_5_1_DiabetesInj_Novomix_Penfilled', label: 'Novomix Penfilled', category: '5.1', subgroup: 'ยาฉีดเบาหวาน' },
  { key: 'Cat_5_1_DiabetesInj_Insulin_Glargine', label: 'Insulin Glargine', category: '5.1', subgroup: 'ยาฉีดเบาหวาน' },
  { key: 'Cat_5_1_DiabetesInj_Gensulin_Penfilled', label: 'Gensulin Penfilled', category: '5.1', subgroup: 'ยาฉีดเบาหวาน' },
  { key: 'Cat_5_1_DiabetesInj_Semaglutide', label: 'Semaglutide', category: '5.1', subgroup: 'ยาฉีดเบาหวาน' },
  // 5.1 ยาเทคนิคพิเศษอื่น ๆ
  { key: 'Cat_5_1_SpecialOther_ChildLiquidMix', label: 'การผสมยาน้ำเด็ก', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  { key: 'Cat_5_1_SpecialOther_EyeDrop', label: 'ยาหยอดตา', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  { key: 'Cat_5_1_SpecialOther_EyeOintment', label: 'ยาป้ายตา', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  { key: 'Cat_5_1_SpecialOther_Sublingual', label: 'ยาอมใต้ลิ้น', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  { key: 'Cat_5_1_SpecialOther_RectalSupp', label: 'ยาเหน็บทวาร/สวนทวาร', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  { key: 'Cat_5_1_SpecialOther_VaginalSupp', label: 'ยาเหน็บช่องคลอด', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  { key: 'Cat_5_1_SpecialOther_EarDrop', label: 'ยาหยอดหู', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  { key: 'Cat_5_1_SpecialOther_Enoxaparin', label: 'Enoxaparin', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  { key: 'Cat_5_1_SpecialOther_Fentanyl_Patch', label: 'Fentanyl patch', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  { key: 'Cat_5_1_SpecialOther_Alendronate', label: 'Alendronate', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  // 5.2 - 5.6
  { key: 'Cat_5_2_Warfarin', label: '2. แนะนำการใช้ยา Warfarin', category: '5.2', subgroup: null },
  { key: 'Cat_5_3_TB', label: '3. แนะนำการใช้ยาในผู้ป่วย Tuberculosis', category: '5.3', subgroup: null },
  { key: 'Cat_5_4_Myanmar_Label', label: '4. แนะนำการใช้ยาในผู้ป่วยเมียนมาร์', category: '5.4', subgroup: null },
  { key: 'Cat_5_5_Stroke_Case', label: '5. แนะนำการใช้ยาในผู้ป่วยโรคหลอดเลือดสมองรายใหม่', category: '5.5', subgroup: null },
  { key: 'Cat_5_6_SJS_TEN_Risk', label: '6. แนะนำการใช้ยาที่มีความเสี่ยงต่อการเกิด SCARs ในผู้ป่วยรายใหม่', category: '5.6', subgroup: null },
  // 5.6 ยาที่เสี่ยงต่อการเกิด SCARs
  { key: 'Cat_5_6_SCARs_Allopurinol', label: 'Allopurinol', category: '5.6', subgroup: 'ยาที่เสี่ยงต่อการเกิด SCARs' },
  { key: 'Cat_5_6_SCARs_TMP_SMX', label: 'Trimethoprim + Sulfamethoxazole (Bactrim)', category: '5.6', subgroup: 'ยาที่เสี่ยงต่อการเกิด SCARs' },
  { key: 'Cat_5_6_SCARs_Carbamazepine', label: 'Carbamazepine', category: '5.6', subgroup: 'ยาที่เสี่ยงต่อการเกิด SCARs' },
  { key: 'Cat_5_6_SCARs_Phenobarbital', label: 'Phenobarbital', category: '5.6', subgroup: 'ยาที่เสี่ยงต่อการเกิด SCARs' },
  { key: 'Cat_5_6_SCARs_Phenytoin', label: 'Phenytoin', category: '5.6', subgroup: 'ยาที่เสี่ยงต่อการเกิด SCARs' },
  { key: 'Cat_5_6_SCARs_Nevirapine', label: 'Nevirapine', category: '5.6', subgroup: 'ยาที่เสี่ยงต่อการเกิด SCARs' },
  // 5.7 - 5.8
  { key: 'Cat_5_7_ARV', label: '7. แนะนำการใช้ยาในผู้ป่วย ARV', category: '5.7', subgroup: null },
  { key: 'Cat_5_8_Other', label: '8. อื่น ๆ', category: '5.8', subgroup: null },
  { key: 'Cat_5_8_Other_Text', label: 'รายละเอียด อื่น ๆ', category: '5.8', subgroup: null, type: 'text' },
  // Appended at end to preserve existing column positions
  { key: 'Cat_5_1_SpecialOther_Omeprazole_NG', label: 'Omeprazole via NG Tube', category: '5.1', subgroup: 'ยาเทคนิคพิเศษอื่น ๆ' },
  { key: 'Cat_5_6_SCARs_StartDate', label: 'วันที่เริ่มยา', category: '5.6', subgroup: null, type: 'date' },
];

// ==========================================
// Web App Entry Point
// ==========================================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.userAuth = JSON.stringify(getUserAuth());
  return template.evaluate()
    .setTitle('Discharge Counseling Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// Authorization
// ==========================================
function getUserAuth() {
  // Get current user email
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email) {
    try { email = Session.getEffectiveUser().getEmail(); } catch (e) {}
  }
  if (!email) {
    try {
      var token = ScriptApp.getOAuthToken();
      var res = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (res.getResponseCode() === 200) {
        email = JSON.parse(res.getContentText()).email || '';
      }
    } catch (e) {}
  }
  email = (email || '').toLowerCase().trim();

  if (!email) {
    return { authorized: false, email: '', reason: 'ไม่สามารถระบุอีเมลผู้ใช้ได้' };
  }

  const ss = getSpreadsheet();

  // Look up user in Auth_Users sheet
  const usersSheet = ss.getSheetByName(SHEET_AUTH_USERS);
  if (!usersSheet || usersSheet.getLastRow() <= 1) {
    return { authorized: false, email: email, reason: 'ไม่พบตาราง Auth_Users' };
  }
  const usersData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 2).getValues();
  let role = null;
  for (let i = 0; i < usersData.length; i++) {
    if (String(usersData[i][0]).toLowerCase().trim() === email) {
      role = String(usersData[i][1]).trim();
      break;
    }
  }
  if (!role) {
    return { authorized: false, email: email, reason: 'ไม่พบอีเมลนี้ในระบบ' };
  }

  // Look up role permissions in Auth_Roles sheet
  const rolesSheet = ss.getSheetByName(SHEET_AUTH_ROLES);
  if (!rolesSheet || rolesSheet.getLastRow() <= 1) {
    return { authorized: false };
  }
  const rolesHeaders = rolesSheet.getRange(1, 1, 1, rolesSheet.getLastColumn()).getValues()[0];
  const rolesData = rolesSheet.getRange(2, 1, rolesSheet.getLastRow() - 1, rolesSheet.getLastColumn()).getValues();

  let permissions = {};
  AUTH_PAGES.forEach(function(p) { permissions[p] = false; });

  for (let i = 0; i < rolesData.length; i++) {
    if (String(rolesData[i][0]).trim() === role) {
      for (let j = 1; j < rolesHeaders.length; j++) {
        const pageKey = String(rolesHeaders[j]).trim().toLowerCase();
        if (permissions.hasOwnProperty(pageKey)) {
          permissions[pageKey] = rolesData[i][j] === true || String(rolesData[i][j]).toUpperCase() === 'TRUE';
        }
      }
      break;
    }
  }

  // Admin always has access to user management (even if column missing from sheet)
  if (role === 'Admin') {
    permissions.users = true;
  }

  return { authorized: true, email: email, role: role, permissions: permissions };
}

function checkPermission_(pageId) {
  const auth = getUserAuth();
  if (!auth.authorized) {
    throw new Error('ไม่มีสิทธิ์เข้าถึง');
  }
  if (!auth.permissions[pageId]) {
    throw new Error('ไม่มีสิทธิ์เข้าถึงหน้า ' + pageId);
  }
  return auth;
}

// ==========================================
// User Management (Admin only)
// ==========================================
function getAuthUsers() {
  checkPermission_('users');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_AUTH_USERS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  return data.map(function(row) {
    return { email: String(row[0]).trim(), role: String(row[1]).trim() };
  });
}

function getAuthRoles() {
  checkPermission_('users');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_AUTH_ROLES);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return data.map(function(row) { return String(row[0]).trim(); });
}

function saveAuthUser(email, role) {
  checkPermission_('users');
  email = (email || '').toLowerCase().trim();
  role = (role || '').trim();
  if (!email) return { success: false, error: 'กรุณากรอกอีเมล' };
  if (!role) return { success: false, error: 'กรุณาเลือก Role' };

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_AUTH_USERS);
  if (!sheet) return { success: false, error: 'ไม่พบตาราง Auth_Users' };

  // Check if email already exists — update role
  if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === email) {
        sheet.getRange(i + 2, 2).setValue(role);
        // Clear cache for this user
        return { success: true, message: 'อัพเดท Role ของ ' + email + ' เป็น ' + role };
      }
    }
  }

  // New user
  sheet.appendRow([email, role]);
  return { success: true, message: 'เพิ่มผู้ใช้ ' + email + ' เป็น ' + role };
}

function deleteAuthUser(email) {
  checkPermission_('users');
  email = (email || '').toLowerCase().trim();
  if (!email) return { success: false, error: 'กรุณาระบุอีเมล' };

  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_AUTH_USERS);
  if (!sheet || sheet.getLastRow() <= 1) return { success: false, error: 'ไม่พบข้อมูล' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === email) {
      sheet.deleteRow(i + 2);
      return { success: true, message: 'ลบผู้ใช้ ' + email + ' สำเร็จ' };
    }
  }
  return { success: false, error: 'ไม่พบอีเมลนี้' };
}

// ==========================================
// Sheet Helpers
// ==========================================
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(name, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sheet;
}

// ==========================================
// Column letter helper (1-based: 1=A, 27=AA)
// ==========================================
function colLetter(colNum) {
  let letter = '';
  while (colNum > 0) {
    colNum--;
    letter = String.fromCharCode(65 + (colNum % 26)) + letter;
    colNum = Math.floor(colNum / 26);
  }
  return letter;
}

// ==========================================
// Setup (run once)
// ==========================================
function setupSheets() {
  // Counseling_Raw headers
  const counselingHeaders = ['Timestamp', 'CounselingDate', 'AN'];
  COUNSELING_FIELDS.forEach(f => counselingHeaders.push(f.key));
  const cSheet = getOrCreateSheet(SHEET_COUNSELING, counselingHeaders);
  // Format AN column as plain text to preserve leading zeros
  cSheet.getRange('C:C').setNumberFormat('@');
  // Update headers if sheet already existed (new columns added)
  if (cSheet.getLastColumn() < counselingHeaders.length || cSheet.getLastColumn() > counselingHeaders.length) {
    cSheet.getRange(1, 1, 1, counselingHeaders.length).setValues([counselingHeaders]);
    cSheet.getRange(1, 1, 1, counselingHeaders.length).setFontWeight('bold');
  }

  // DC_Raw headers
  getOrCreateSheet(SHEET_DC, ['Timestamp', 'DischargeDate', 'DischargeCount']);

  // PE_Raw headers
  getOrCreateSheet(SHEET_PE, ['Timestamp', 'ErrorDate', 'ErrorCount']);

  // HM_Time_Raw headers
  const hmSheet = getOrCreateSheet(SHEET_HM, HM_HEADERS);
  hmSheet.getRange('C:C').setNumberFormat('@'); // AN as text
  if (hmSheet.getLastColumn() !== HM_HEADERS.length) {
    hmSheet.getRange(1, 1, 1, HM_HEADERS.length).setValues([HM_HEADERS]);
    hmSheet.getRange(1, 1, 1, HM_HEADERS.length).setFontWeight('bold');
  }

  // Auth_Users sheet
  const authUsersHeaders = ['Email', 'Role'];
  const authUsersSheet = getOrCreateSheet(SHEET_AUTH_USERS, authUsersHeaders);
  authUsersSheet.getRange(1, 1, 1, authUsersHeaders.length).setFontWeight('bold');
  authUsersSheet.setColumnWidth(1, 300);
  authUsersSheet.setColumnWidth(2, 150);

  // Auth_Roles sheet with default rows
  const authRolesHeaders = ['Role', 'dashboard', 'counseling', 'count', 'hm', 'report', 'users'];
  const authRolesSheet = getOrCreateSheet(SHEET_AUTH_ROLES, authRolesHeaders);
  authRolesSheet.getRange(1, 1, 1, authRolesHeaders.length).setFontWeight('bold');
  // Only populate default rows if sheet is empty (just headers)
  if (authRolesSheet.getLastRow() <= 1) {
    authRolesSheet.appendRow(['Admin', true, true, true, true, true, true]);
    authRolesSheet.appendRow(['Editor', true, true, true, true, true, false]);
    authRolesSheet.appendRow(['Viewer', true, false, false, false, true, false]);
  }

  // Report Summary
  setupReportSummary();

  // Fiscal Year Summary
  setupFiscalYearSummary();

  // HM Report Summary (formula-driven HM Time report)
  setupHMReportSummary();
}

// ==========================================
// Setup Report Summary with formulas
// ==========================================
function setupReportSummary() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_REPORT);

  // Delete and recreate to ensure clean state
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet(SHEET_REPORT);

  // Date range formula parts (reference $B$1 as the selected month date)
  const START = 'DATE(YEAR($B$1),MONTH($B$1),1)';
  const END = 'EOMONTH($B$1,0)';
  const DATE_FILTER = 'Counseling_Raw!B:B,">="&' + START + ',Counseling_Raw!B:B,"<="&' + END;

  // Helper: COUNTIFS formula for a counseling checkbox column
  function drugFormula(fieldIndex) {
    const col = colLetter(fieldIndex + 4); // D=4 for index 0
    return '=COUNTIFS(' + DATE_FILTER + ',Counseling_Raw!' + col + ':' + col + ',TRUE)';
  }

  const RESULT_TEXT = 'ผู้ป่วยเข้าใจปฏิบัติได้ถูกต้อง';

  // Build rows: [label, value_or_formula, result_text]
  const rows = [];
  let currentRow = 1;

  // Row 1: Month selector
  rows.push(['ประจำเดือน', new Date(new Date().getFullYear(), new Date().getMonth(), 1), '']);
  currentRow++;

  // Row 2: empty
  rows.push(['', '', '']);
  currentRow++;

  // Row 3: Headers
  rows.push(['รายละเอียด', 'Discharge Counseling (ครั้ง)', 'ผลลัพธ์']);
  currentRow++;

  // Track which rows are section headers / sub-headers for formatting
  const sectionHeaderRows = [];
  const subHeaderRows = [];
  const formulaRows = []; // [row, formula]

  // --- Helper to build a category with sub-items (used for 5.1 and 5.6) ---
  function buildCategoryWithItems(category, headerLabel) {
    sectionHeaderRows.push(currentRow);
    rows.push([headerLabel, '', RESULT_TEXT]);
    const sectionHeaderRow = currentRow;
    currentRow++;

    const itemRows = [];
    let lastSub = '';
    COUNSELING_FIELDS.forEach((field, idx) => {
      if (field.category !== category) return;
      if (field.subgroup === null) return; // skip top-level entry (e.g. Cat_5_6_SJS_TEN_Risk)
      if (field.type === 'text') return; // skip text fields

      // Sub-group header
      if (field.subgroup && field.subgroup !== lastSub) {
        subHeaderRows.push(currentRow);
        rows.push([field.subgroup, '', '']);
        currentRow++;
        lastSub = field.subgroup;
      }

      itemRows.push(currentRow);
      formulaRows.push([currentRow, drugFormula(idx)]);
      rows.push([field.label, '', '']);
      currentRow++;
    });

    return { sectionHeaderRow, itemRows };
  }

  // 5.1 Section
  const sec51 = buildCategoryWithItems('5.1', '1. แนะนำการใช้ยาเทคนิคพิเศษ');

  // 5.2 - 5.5 (standalone categories)
  const topLevelLabels = {
    '5.2': '2. แนะนำการใช้ยา Warfarin',
    '5.3': '3. แนะนำการใช้ยาในผู้ป่วย Tuberculosis',
    '5.4': '4. แนะนำการใช้ยาในผู้ป่วยเมียนมาร์',
    '5.5': '5. แนะนำการใช้ยาในผู้ป่วยโรคหลอดเลือดสมองรายใหม่'
  };

  // Track all top-level category header rows for total SUM
  const allCategoryHeaderRows = [sec51.sectionHeaderRow];

  ['5.2', '5.3', '5.4', '5.5'].forEach(cat => {
    const fieldIdx = COUNSELING_FIELDS.findIndex(f => f.category === cat && f.subgroup === null);
    if (fieldIdx >= 0) {
      sectionHeaderRows.push(currentRow);
      allCategoryHeaderRows.push(currentRow);
      formulaRows.push([currentRow, drugFormula(fieldIdx)]);
      rows.push([topLevelLabels[cat], '', RESULT_TEXT]);
      currentRow++;
    }
  });

  // 5.6 Section (with sub-items like 5.1)
  const sec56 = buildCategoryWithItems('5.6', '6. แนะนำการใช้ยาที่มีความเสี่ยงต่อการเกิด SCARs ในผู้ป่วยรายใหม่');
  allCategoryHeaderRows.push(sec56.sectionHeaderRow);

  // 5.7 ARV
  const field57Idx = COUNSELING_FIELDS.findIndex(f => f.category === '5.7');
  if (field57Idx >= 0) {
    sectionHeaderRows.push(currentRow);
    allCategoryHeaderRows.push(currentRow);
    formulaRows.push([currentRow, drugFormula(field57Idx)]);
    rows.push(['7. แนะนำการใช้ยาในผู้ป่วย ARV', '', RESULT_TEXT]);
    currentRow++;
  }

  // 5.8 อื่น ๆ
  const field58Idx = COUNSELING_FIELDS.findIndex(f => f.key === 'Cat_5_8_Other');
  if (field58Idx >= 0) {
    sectionHeaderRows.push(currentRow);
    allCategoryHeaderRows.push(currentRow);
    formulaRows.push([currentRow, drugFormula(field58Idx)]);
    rows.push(['8. อื่น ๆ', '', '']);
    currentRow++;
  }

  // Empty row
  rows.push(['', '', '']);
  currentRow++;

  // Summary rows
  const totalSessionsRow = currentRow;
  rows.push(['รวมทั้งหมด (ครั้ง)', '', '']);
  currentRow++;

  const totalPatientsRow = currentRow;
  rows.push(['รวมทั้งหมด (ราย)', 'นับจาก AN ที่ลงประจำเดือนที่เลือก', '']);
  currentRow++;

  const totalDCRow = currentRow;
  rows.push(['จำนวนผู้ป่วยกลับบ้านทั้งหมด', 'จำนวนผู้ป่วยที่ Discharge ประจำเดือนที่เลือก', '']);
  currentRow++;

  const totalPERow = currentRow;
  rows.push(['Prescribing Error จากการทำ MR ในผู้ป่วยที่ Discharge', 'จำนวน Prescribing Error ประจำเดือนที่เลือก', '']);
  currentRow++;

  const pePctRow = currentRow;
  rows.push(['ร้อยละของ Prescribing Error จากการทำ MR', 'Percentage of Prescribing Error from total Discharge', '']);
  currentRow++;

  // Write all labels and static values
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);

  // Set B1 as date format
  sheet.getRange(1, 2).setNumberFormat('MMMM yyyy');

  // Set formulas for drug/category counts
  formulaRows.forEach(([row, formula]) => {
    sheet.getRange(row, 2).setFormula(formula);
  });

  // 5.1 header: SUM of all individual 5.1 item rows
  if (sec51.itemRows.length > 0) {
    const sumRefs51 = sec51.itemRows.map(r => 'B' + r).join(',');
    sheet.getRange(sec51.sectionHeaderRow, 2).setFormula('=SUM(' + sumRefs51 + ')');
  }

  // 5.6 header: SUM of all individual 5.6 sub-item rows (like 5.1)
  if (sec56.itemRows.length > 0) {
    const sumRefs56 = sec56.itemRows.map(r => 'B' + r).join(',');
    sheet.getRange(sec56.sectionHeaderRow, 2).setFormula('=SUM(' + sumRefs56 + ')');
  }

  // Total (ครั้ง) = SUM of all category header rows (5.1 + 5.2 + ... + 5.8)
  const totalSumRefs = allCategoryHeaderRows.map(r => 'B' + r).join(',');
  sheet.getRange(totalSessionsRow, 2).setFormula('=SUM(' + totalSumRefs + ')');

  // Total (ราย) = count of AN entries in the month (with upsert, each AN per day = 1 row)
  sheet.getRange(totalPatientsRow, 2).setFormula(
    '=COUNTIFS(Counseling_Raw!B2:B,">="&' + START + ',Counseling_Raw!B2:B,"<="&' + END + ')'
  );

  // Total discharges = SUM of DischargeCount from DC_Raw in selected month
  sheet.getRange(totalDCRow, 2).setFormula(
    '=IFERROR(SUMIFS(DC_Raw!C2:C,DC_Raw!B2:B,">="&' + START + ',DC_Raw!B2:B,"<="&' + END + '),0)'
  );

  // Total PE = SUM of ErrorCount from PE_Raw in selected month
  sheet.getRange(totalPERow, 2).setFormula(
    '=IFERROR(SUMIFS(PE_Raw!C2:C,PE_Raw!B2:B,">="&' + START + ',PE_Raw!B2:B,"<="&' + END + '),0)'
  );

  // PE percentage
  sheet.getRange(pePctRow, 2).setFormula(
    '=IFERROR(B' + totalPERow + '/B' + totalDCRow + '*100,0)'
  );
  sheet.getRange(pePctRow, 2).setNumberFormat('0.00');

  // ==========================================
  // Formatting
  // ==========================================

  // Title row
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setFontSize(12);
  sheet.getRange(1, 2).setBackground('#FFF3D6');

  // Header row
  sheet.getRange(3, 1, 1, 3).setFontWeight('bold').setBackground('#A7B5FE').setFontColor('#FFFFFF');

  // Section headers (bold, colored background)
  sectionHeaderRows.forEach(r => {
    sheet.getRange(r, 1, 1, 3).setFontWeight('bold').setBackground('#E0F0F3');
  });

  // Sub-group headers (semi-bold, light indent)
  subHeaderRows.forEach(r => {
    sheet.getRange(r, 1).setFontWeight('bold').setFontColor('#7BB8C0');
  });

  // Summary rows (bold)
  [totalSessionsRow, totalPatientsRow, totalDCRow, totalPERow, pePctRow].forEach(r => {
    sheet.getRange(r, 1, 1, 3).setFontWeight('bold');
  });
  sheet.getRange(totalSessionsRow, 1, 1, 3).setBackground('#FFF3D6');
  // Highlight summary B column descriptions
  [totalPatientsRow, totalDCRow, totalPERow, pePctRow].forEach(r => {
    sheet.getRange(r, 2).setBackground('#FFFF00');
  });

  // Column widths
  sheet.setColumnWidth(1, 400);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 250);

  // Center column B values (data rows only, not summary)
  const dataRowCount = totalSessionsRow - 4; // rows from 4 to totalSessionsRow-1
  if (dataRowCount > 0) {
    sheet.getRange(4, 2, dataRowCount, 1).setHorizontalAlignment('center');
  }
}

// ==========================================
// Setup Fiscal Year Summary with monthly columns
// ==========================================
function setupFiscalYearSummary() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_FY_SUMMARY);

  // Delete and recreate to ensure clean state
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  sheet = ss.insertSheet(SHEET_FY_SUMMARY);

  // --- Compute current fiscal year (BE) ---
  const now = new Date();
  const currentBEYear = now.getFullYear() + 543;
  const defaultFY = now.getMonth() >= 9 ? currentBEYear + 1 : currentBEYear; // Oct+ = next FY

  // --- Month configuration (Thai fiscal year: Oct-Sep) ---
  const FY_MONTHS = [
    { month: 10, yearOffset: -544, thAbbr: 'ต.ค.' },
    { month: 11, yearOffset: -544, thAbbr: 'พ.ย.' },
    { month: 12, yearOffset: -544, thAbbr: 'ธ.ค.' },
    { month: 1,  yearOffset: -543, thAbbr: 'ม.ค.' },
    { month: 2,  yearOffset: -543, thAbbr: 'ก.พ.' },
    { month: 3,  yearOffset: -543, thAbbr: 'มี.ค.' },
    { month: 4,  yearOffset: -543, thAbbr: 'เม.ย.' },
    { month: 5,  yearOffset: -543, thAbbr: 'พ.ค.' },
    { month: 6,  yearOffset: -543, thAbbr: 'มิ.ย.' },
    { month: 7,  yearOffset: -543, thAbbr: 'ก.ค.' },
    { month: 8,  yearOffset: -543, thAbbr: 'ส.ค.' },
    { month: 9,  yearOffset: -543, thAbbr: 'ก.ย.' },
  ];

  // --- Compute column letters for counseling fields ---
  // Cat 5.1: all fields with category '5.1' (excluding text type)
  const cat51Cols = [];
  COUNSELING_FIELDS.forEach(function(f, idx) {
    if (f.category === '5.1' && f.type !== 'text') {
      cat51Cols.push(colLetter(idx + 4));
    }
  });

  // Cat 5.6: sub-item fields only (exclude top-level Cat_5_6_SJS_TEN_Risk)
  const cat56Cols = [];
  COUNSELING_FIELDS.forEach(function(f, idx) {
    if (f.category === '5.6' && f.subgroup !== null && f.type !== 'text') {
      cat56Cols.push(colLetter(idx + 4));
    }
  });

  // Single-field categories: find their column letters
  function singleFieldCol(key) {
    var idx = COUNSELING_FIELDS.findIndex(function(f) { return f.key === key; });
    return idx >= 0 ? colLetter(idx + 4) : null;
  }
  const col52 = singleFieldCol('Cat_5_2_Warfarin');
  const col54 = singleFieldCol('Cat_5_4_Myanmar_Label');
  const col53 = singleFieldCol('Cat_5_3_TB');
  const col57 = singleFieldCol('Cat_5_7_ARV');
  const col55 = singleFieldCol('Cat_5_5_Stroke_Case');

  // --- Formula builders ---
  // Single-field COUNTIFS for a month column
  function singleFieldFormula(monthCol, rawCol) {
    return '=COUNTIFS(Counseling_Raw!$B$2:$B,">="&' + monthCol + '$3,'
         + 'Counseling_Raw!$B$2:$B,"<="&EOMONTH(' + monthCol + '$3,0),'
         + 'Counseling_Raw!$' + rawCol + '$2:$' + rawCol + ',TRUE)';
  }

  // Multi-field SUMPRODUCT for a month column with OR across multiple columns
  function multiFieldFormula(monthCol, rawCols) {
    var dateCond = '(Counseling_Raw!$B$2:$B>=' + monthCol + '$3)*'
                 + '(Counseling_Raw!$B$2:$B<=EOMONTH(' + monthCol + '$3,0))';
    var orParts = rawCols.map(function(rc) {
      return '(Counseling_Raw!$' + rc + '$2:$' + rc + '=TRUE)';
    }).join('+');
    return '=SUMPRODUCT(' + dateCond + '*(' + orParts + '))';
  }

  // Unique patient count for a month column using SUMPRODUCT to count distinct AN
  function uniquePatientFormula(monthCol) {
    return '=SUMPRODUCT(IFERROR('
         + '(Counseling_Raw!$B$2:$B>=' + monthCol + '$3)*'
         + '(Counseling_Raw!$B$2:$B<=EOMONTH(' + monthCol + '$3,0))*'
         + '(Counseling_Raw!$C$2:$C<>"")*'
         + '(1/COUNTIFS('
         + 'Counseling_Raw!$C$2:$C,Counseling_Raw!$C$2:$C,'
         + 'Counseling_Raw!$B$2:$B,">="&' + monthCol + '$3,'
         + 'Counseling_Raw!$B$2:$B,"<="&EOMONTH(' + monthCol + '$3,0)'
         + ')),0))';
  }

  // DC total for a month column
  function dcFormula(monthCol) {
    return '=IFERROR(SUMIFS(DC_Raw!$C$2:$C,'
         + 'DC_Raw!$B$2:$B,">="&' + monthCol + '$3,'
         + 'DC_Raw!$B$2:$B,"<="&EOMONTH(' + monthCol + '$3,0)),0)';
  }

  // PE total for a month column
  function peFormula(monthCol) {
    return '=IFERROR(SUMIFS(PE_Raw!$C$2:$C,'
         + 'PE_Raw!$B$2:$B,">="&' + monthCol + '$3,'
         + 'PE_Raw!$B$2:$B,"<="&EOMONTH(' + monthCol + '$3,0)),0)';
  }

  // ==========================================
  // Row 1: Title
  // ==========================================
  sheet.getRange('A1').setValue('จำนวนผู้ป่วยที่ได้รับการให้คำปรึกษาเรื่องยาผู้ป่วยก่อนกลับบ้าน (Discharge Counseling)');

  // ==========================================
  // Row 2: Fiscal year selector + explanation
  // ==========================================
  var fyOptions = [];
  for (var y = 2567; y <= 2580; y++) fyOptions.push('ปีงบประมาณ ' + y);
  sheet.getRange('A2').setValue('ปีงบประมาณ ' + defaultFY);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(fyOptions)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('A2').setDataValidation(rule);
  sheet.getRange('C2').setFormula('=$A$2&" = ตุลาคม "&($A$3-1)&" - กันยายน "&$A$3');

  // ==========================================
  // Row 3: Hidden helper - A3 extracts fiscal year number, B3-M3 month start dates
  // ==========================================
  sheet.getRange('A3').setFormula('=VALUE(RIGHT(A2,4))');
  for (var i = 0; i < 12; i++) {
    var m = FY_MONTHS[i];
    sheet.getRange(3, i + 2).setFormula('=DATE($A$3+(' + m.yearOffset + '),' + m.month + ',1)');
  }
  sheet.hideRows(3);

  // ==========================================
  // Row 4: Column headers
  // ==========================================
  sheet.getRange('A4').setValue('กลุ่มเป้าหมาย');
  for (var i = 0; i < 12; i++) {
    var m = FY_MONTHS[i];
    var beYearRef = m.month >= 10 ? '$A$3-1' : '$A$3';
    sheet.getRange(4, i + 2).setFormula('="' + m.thAbbr + '-"&TEXT(MOD(' + beYearRef + ',100),"00")');
  }
  sheet.getRange(4, 14).setValue('รวม');
  sheet.getRange(4, 15).setValue('เฉลี่ยต่อเดือน');

  // ==========================================
  // Rows 5-11: Category data rows
  // ==========================================
  var categoryRows = [
    { row: 5, label: '1. ผู้ป่วยรายใหม่ที่ได้รับยาเทคนิคพิเศษ', type: 'multi', cols: cat51Cols },
    { row: 6, label: '2. ผู้ป่วยที่ได้รับยา Warfarin', type: 'single', col: col52 },
    { row: 7, label: '3. ผู้ป่วยต่างชาติ ได้แก่ ผู้ป่วยเมียนมาร์', type: 'single', col: col54 },
    { row: 8, label: '4. ผู้ป่วยในกลุ่มวัณโรค (Tuberculosis)', type: 'single', col: col53 },
    { row: 9, label: '5. ผู้ป่วยในกลุ่มที่ได้รับยา ARV', type: 'single', col: col57 },
    { row: 10, label: '6. ผู้ป่วยโรคหลอดเลือดสมอง (Stroke)', type: 'single', col: col55 },
    { row: 11, label: '7. ผู้ป่วยรายใหม่ที่ได้รับยาที่เสี่ยงต่อการเกิด Severe Cutaneous Adverse Reactions (SCARs)', type: 'multi', cols: cat56Cols },
  ];

  categoryRows.forEach(function(cat) {
    sheet.getRange(cat.row, 1).setValue(cat.label);
    for (var c = 0; c < 12; c++) {
      var monthCol = colLetter(c + 2);
      if (cat.type === 'single') {
        sheet.getRange(cat.row, c + 2).setFormula(singleFieldFormula(monthCol, cat.col));
      } else {
        sheet.getRange(cat.row, c + 2).setFormula(multiFieldFormula(monthCol, cat.cols));
      }
    }
    // Column N: Total
    sheet.getRange(cat.row, 14).setFormula('=SUM(B' + cat.row + ':M' + cat.row + ')');
    // Column O: Average
    sheet.getRange(cat.row, 15).setFormula('=N' + cat.row + '/12');
  });

  // ==========================================
  // Row 12: จำนวนครั้ง (total sessions)
  // ==========================================
  sheet.getRange(12, 1).setValue('จำนวนครั้ง');
  for (var c = 0; c < 12; c++) {
    var mc = colLetter(c + 2);
    sheet.getRange(12, c + 2).setFormula('=SUM(' + mc + '5:' + mc + '11)');
  }
  sheet.getRange(12, 14).setFormula('=SUM(B12:M12)');
  sheet.getRange(12, 15).setFormula('=N12/12');

  // ==========================================
  // Row 13: จำนวนผู้ป่วย (distinct AN)
  // ==========================================
  sheet.getRange(13, 1).setValue('จำนวนผู้ป่วย (ราย)');
  for (var c = 0; c < 12; c++) {
    var mc = colLetter(c + 2);
    sheet.getRange(13, c + 2).setFormula(uniquePatientFormula(mc));
  }
  sheet.getRange(13, 14).setFormula('=SUM(B13:M13)');
  sheet.getRange(13, 15).setFormula('=N13/12');

  // ==========================================
  // Row 14: ผู้ป่วยกลับบ้านทั้งหมด (DC_Raw)
  // ==========================================
  sheet.getRange(14, 1).setValue('ผู้ป่วยกลับบ้านทั้งหมด (ราย)');
  for (var c = 0; c < 12; c++) {
    var mc = colLetter(c + 2);
    sheet.getRange(14, c + 2).setFormula(dcFormula(mc));
  }
  sheet.getRange(14, 14).setFormula('=SUM(B14:M14)');
  sheet.getRange(14, 15).setFormula('=N14/12');

  // ==========================================
  // Row 15: Prescribing Error (PE_Raw)
  // ==========================================
  sheet.getRange(15, 1).setValue('จำนวน Prescribing Error ในผู้ป่วยกลับบ้าน');
  for (var c = 0; c < 12; c++) {
    var mc = colLetter(c + 2);
    sheet.getRange(15, c + 2).setFormula(peFormula(mc));
  }
  sheet.getRange(15, 14).setFormula('=SUM(B15:M15)');
  sheet.getRange(15, 15).setFormula('=N15/12');

  // ==========================================
  // Row 16: PE rate per 1,000
  // ==========================================
  sheet.getRange(16, 1).setValue('อัตราการเกิด Prescribing Error ในผู้ป่วยกลับบ้าน (อัตราต่อ 1,000 ผู้ป่วยจำหน่าย)');
  for (var c = 0; c < 12; c++) {
    var mc = colLetter(c + 2);
    sheet.getRange(16, c + 2).setFormula('=IFERROR(' + mc + '15/' + mc + '14*1000,0)');
  }
  sheet.getRange(16, 14).setFormula('=IFERROR(N15/N14*1000,0)');
  sheet.getRange(16, 15).setFormula('=IFERROR(O15/O14*1000,0)');
  sheet.getRange(16, 2, 1, 14).setNumberFormat('0.00');

  // ==========================================
  // Formatting
  // ==========================================

  // Row 1: Title
  sheet.getRange(1, 1, 1, 15).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  // Row 2: FY selector
  sheet.getRange('A2').setFontWeight('bold').setBackground('#666666').setFontColor('#FFFFFF');

  // Row 4: Headers
  sheet.getRange(4, 1, 1, 15).setFontWeight('bold').setBackground('#A7B5FE').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  sheet.getRange(4, 1).setHorizontalAlignment('left');

  // Category rows (5-11): light background for labels
  sheet.getRange(5, 1, 7, 1).setBackground('#E0F0F3');

  // Row 12: Total sessions (bold, honey background)
  sheet.getRange(12, 1, 1, 15).setFontWeight('bold').setBackground('#FFF3D6');

  // Rows 13-16: Summary rows (bold)
  sheet.getRange(13, 1, 4, 15).setFontWeight('bold');

  // Center-align data columns (B-O, rows 5-16)
  sheet.getRange(5, 2, 12, 14).setHorizontalAlignment('center');

  // Column widths
  sheet.setColumnWidth(1, 350);
  for (var c = 2; c <= 13; c++) {
    sheet.setColumnWidth(c, 100);
  }
  sheet.setColumnWidth(14, 80);
  sheet.setColumnWidth(15, 130);

  // Freeze header row and label column
  sheet.setFrozenRows(4);
  sheet.setFrozenColumns(1);
}

// ==========================================
// Get field config (for client)
// ==========================================
function getCounselingFieldConfig() {
  checkPermission_('counseling');
  return COUNSELING_FIELDS;
}

// ==========================================
// Submit Counseling Data
// ==========================================
function submitCounselingData(formData) {
  checkPermission_('counseling');
  if (!formData.an || !/^\d{9}$/.test(formData.an)) {
    return { success: false, error: 'AN ต้องเป็นตัวเลข 9 หลัก' };
  }
  if (!formData.counselingDate) {
    return { success: false, error: 'กรุณาเลือกวันที่' };
  }

  const sheet = getOrCreateSheet(SHEET_COUNSELING);
  const data = sheet.getDataRange().getValues();
  const targetDate = new Date(formData.counselingDate).toDateString();
  const targetAN = String(formData.an);

  // Check for existing row with same AN + same date
  let existingRowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]) === targetAN && data[i][1] && new Date(data[i][1]).toDateString() === targetDate) {
      existingRowIdx = i + 1; // 1-indexed sheet row
      break;
    }
  }

  if (existingRowIdx > 0 && !formData.forceUpsert) {
    return { success: false, confirmUpsert: true, message: 'AN ' + targetAN + ' มีข้อมูลวันที่เดียวกันแล้ว ต้องการแทนที่ด้วยข้อมูลล่าสุดหรือไม่?' };
  }

  const row = [new Date(), new Date(formData.counselingDate), formData.an];
  COUNSELING_FIELDS.forEach(field => {
    if (field.type === 'text') {
      row.push(formData.freeText || '');
    } else if (field.type === 'date') {
      const v = formData.dates && formData.dates[field.key];
      row.push(v ? new Date(v) : '');
    } else {
      row.push(formData.checkboxes[field.key] === true);
    }
  });

  if (existingRowIdx > 0) {
    // Update existing row
    sheet.getRange(existingRowIdx, 1, 1, row.length).setValues([row]);
    sheet.getRange(existingRowIdx, 3).setNumberFormat('@');
    return { success: true, message: 'อัพเดทข้อมูล Counseling สำเร็จ' };
  } else {
    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 3).setNumberFormat('@');
    return { success: true, message: 'บันทึกข้อมูล Counseling สำเร็จ' };
  }
}

// ==========================================
// Submit DC Data (upsert by date)
// ==========================================
function submitDCData(formData) {
  checkPermission_('count');
  if (!formData.dischargeDate) {
    return { success: false, error: 'กรุณาเลือกวันที่' };
  }
  const count = parseInt(formData.dischargeCount, 10);
  if (isNaN(count) || count < 0) {
    return { success: false, error: 'จำนวนต้องเป็นตัวเลขที่ไม่ติดลบ' };
  }

  const sheet = getOrCreateSheet(SHEET_DC);
  const data = sheet.getDataRange().getValues();
  const targetDate = new Date(formData.dischargeDate).toDateString();

  // Search for existing row with same date
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && new Date(data[i][1]).toDateString() === targetDate) {
      sheet.getRange(i + 1, 1).setValue(new Date()); // update timestamp
      sheet.getRange(i + 1, 3).setValue(count);
      return { success: true, message: 'อัพเดทข้อมูล Discharge สำเร็จ' };
    }
  }

  sheet.appendRow([new Date(), new Date(formData.dischargeDate), count]);
  return { success: true, message: 'บันทึกข้อมูล Discharge สำเร็จ' };
}

// ==========================================
// Submit PE Data (upsert by date)
// ==========================================
function submitPEData(formData) {
  checkPermission_('count');
  if (!formData.errorDate) {
    return { success: false, error: 'กรุณาเลือกวันที่' };
  }
  const count = parseInt(formData.errorCount, 10);
  if (isNaN(count) || count < 0) {
    return { success: false, error: 'จำนวนต้องเป็นตัวเลขที่ไม่ติดลบ' };
  }

  const sheet = getOrCreateSheet(SHEET_PE);
  const data = sheet.getDataRange().getValues();
  const targetDate = new Date(formData.errorDate).toDateString();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && new Date(data[i][1]).toDateString() === targetDate) {
      sheet.getRange(i + 1, 1).setValue(new Date());
      sheet.getRange(i + 1, 3).setValue(count);
      return { success: true, message: 'อัพเดทข้อมูล Prescribing Error สำเร็จ' };
    }
  }

  sheet.appendRow([new Date(), new Date(formData.errorDate), count]);
  return { success: true, message: 'บันทึกข้อมูล Prescribing Error สำเร็จ' };
}

// ==========================================
// Submit HM Time Data (upsert by AN + date)
// ==========================================
function submitHMData(formData) {
  checkPermission_('hm');
  if (!formData.an || !/^\d{9}$/.test(formData.an)) {
    return { success: false, error: 'AN ต้องเป็นตัวเลข 9 หลัก' };
  }
  if (!formData.dischargeDate) {
    return { success: false, error: 'กรุณาเลือกวันที่ Discharge' };
  }
  const drugCount = parseInt(formData.drugCount, 10);
  if (isNaN(drugCount) || drugCount < 1) {
    return { success: false, error: 'กรุณากรอกจำนวนยาที่ถูกต้อง' };
  }
  if (!formData.ward) {
    return { success: false, error: 'กรุณาเลือก Ward' };
  }

  const sheet = getOrCreateSheet(SHEET_HM, HM_HEADERS);
  const data = sheet.getDataRange().getValues();
  const targetDate = new Date(formData.dischargeDate).toDateString();
  const targetAN = String(formData.an);

  // Check for existing row with same AN + same date
  let existingRowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]) === targetAN && data[i][1] && new Date(data[i][1]).toDateString() === targetDate) {
      existingRowIdx = i + 1;
      break;
    }
  }

  if (existingRowIdx > 0 && !formData.forceUpsert) {
    return { success: false, confirmUpsert: true, message: 'AN ' + targetAN + ' มีข้อมูลวันที่เดียวกันแล้ว ต้องการแทนที่ด้วยข้อมูลล่าสุดหรือไม่?' };
  }

  const step09Date = formData.step09Date ? new Date(formData.step09Date) : '';

  const row = [
    new Date(),
    new Date(formData.dischargeDate),
    formData.an,
    drugCount,
    formData.ward,
    formData.step01 || '',
    formData.step01Date ? new Date(formData.step01Date) : '',
    formData.step02 || '',
    formData.step02Date ? new Date(formData.step02Date) : '',
    formData.step03 || '',
    formData.step04Start || '',
    formData.step04End || '',
    formData.step05Start || '',
    formData.step05End || '',
    formData.step06 || '',
    formData.step07Start || '',
    formData.step07End || '',
    formData.step08Start || '',
    formData.step08End || '',
    formData.step09Time || '',
    step09Date,
    formData.step10Start || '',
    formData.step10End || '',
    formData.step11 || '',
    formData.remarks || ''
  ];

  if (existingRowIdx > 0) {
    sheet.getRange(existingRowIdx, 1, 1, row.length).setValues([row]);
    sheet.getRange(existingRowIdx, 3).setNumberFormat('@');
    return { success: true, message: 'อัพเดทข้อมูล HM Time สำเร็จ' };
  } else {
    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 3).setNumberFormat('@');
    return { success: true, message: 'บันทึกข้อมูล HM Time สำเร็จ' };
  }
}

// ==========================================
// Pre-check functions for DC/PE duplicate confirmation
// ==========================================
function checkExistingDC(dischargeDate) {
  checkPermission_('count');
  const sheet = getSpreadsheet().getSheetByName(SHEET_DC);
  if (!sheet || sheet.getLastRow() <= 1) return { exists: false };
  const data = sheet.getDataRange().getValues();
  const targetDate = new Date(dischargeDate).toDateString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && new Date(data[i][1]).toDateString() === targetDate) {
      return { exists: true, currentCount: data[i][2] || 0 };
    }
  }
  return { exists: false };
}

function checkExistingPE(errorDate) {
  checkPermission_('count');
  const sheet = getSpreadsheet().getSheetByName(SHEET_PE);
  if (!sheet || sheet.getLastRow() <= 1) return { exists: false };
  const data = sheet.getDataRange().getValues();
  const targetDate = new Date(errorDate).toDateString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && new Date(data[i][1]).toDateString() === targetDate) {
      return { exists: true, currentCount: data[i][2] || 0 };
    }
  }
  return { exists: false };
}

// ==========================================
// Dashboard Data Functions
// ==========================================
function getAvailableMonths() {
  // Needed by both Dashboard and Report pages
  const auth = getUserAuth();
  if (!auth.authorized || (!auth.permissions.dashboard && !auth.permissions.report)) {
    throw new Error('ไม่มีสิทธิ์เข้าถึง');
  }
  const ss = getSpreadsheet();
  const months = new Set();

  // From Counseling_Raw
  const cSheet = ss.getSheetByName(SHEET_COUNSELING);
  if (cSheet && cSheet.getLastRow() > 1) {
    const cData = cSheet.getRange(2, 2, cSheet.getLastRow() - 1, 1).getValues();
    cData.forEach(row => {
      if (row[0]) {
        const d = new Date(row[0]);
        months.add(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM'));
      }
    });
  }

  // From DC_Raw
  const dSheet = ss.getSheetByName(SHEET_DC);
  if (dSheet && dSheet.getLastRow() > 1) {
    const dData = dSheet.getRange(2, 2, dSheet.getLastRow() - 1, 1).getValues();
    dData.forEach(row => {
      if (row[0]) {
        const d = new Date(row[0]);
        months.add(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM'));
      }
    });
  }

  // From PE_Raw
  const pSheet = ss.getSheetByName(SHEET_PE);
  if (pSheet && pSheet.getLastRow() > 1) {
    const pData = pSheet.getRange(2, 2, pSheet.getLastRow() - 1, 1).getValues();
    pData.forEach(row => {
      if (row[0]) {
        const d = new Date(row[0]);
        months.add(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM'));
      }
    });
  }

  return Array.from(months).sort().reverse();
}

function getCounselingSummaryByMonth(yearMonth) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_COUNSELING);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { totalSessions: 0, totalUniquePatients: 0, topLevelCounts: {}, detailedItems: {}, allItemCounts: {}, scarsItems: {}, scarsPatientRows: [] };
  }

  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();

  // Step 1: Filter rows for the month and deduplicate by (date+AN) keeping latest timestamp
  const dedupMap = {}; // key: "dateStr|AN" -> { rowIdx, timestamp }
  let totalRecords = 0; // total rows in the month (before dedup)
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][1];
    if (!rowDate) continue;
    const rowYM = Utilities.formatDate(new Date(rowDate), tz, 'yyyy-MM');
    if (rowYM !== yearMonth) continue;

    totalRecords++;
    const dateStr = new Date(rowDate).toDateString();
    const an = String(data[i][2]);
    const key = dateStr + '|' + an;
    const ts = new Date(data[i][0]).getTime();

    if (!dedupMap[key] || ts > dedupMap[key].timestamp) {
      dedupMap[key] = { rowIdx: i, timestamp: ts };
    }
  }

  // Step 2: Aggregate from deduplicated rows only
  const topLevelCounts = { '5.1': 0, '5.2': 0, '5.3': 0, '5.4': 0, '5.5': 0, '5.6': 0, '5.7': 0, '5.8': 0 };
  const detailedItems = {};
  const allItemCounts = {};
  const scarsItems = {};
  const scarsPatientRows = [];
  const startDateFieldIdx = COUNSELING_FIELDS.findIndex(function(f) { return f.key === 'Cat_5_6_SCARs_StartDate'; });

  // Initialize
  COUNSELING_FIELDS.forEach(f => {
    if (f.type === 'text' || f.type === 'date') return;
    allItemCounts[f.key] = 0;
    if (f.category === '5.1') detailedItems[f.label] = 0;
    if (f.category === '5.6' && f.subgroup) scarsItems[f.label] = 0;
  });

  const dedupKeys = Object.keys(dedupMap);

  // totalSessions = total records (rows) in the month from Counseling_Raw
  // totalUniquePatients = distinct AN values in the month
  const uniqueANs = new Set();
  dedupKeys.forEach(function(key) { uniqueANs.add(key.split('|')[1]); });
  const totalUniquePatients = uniqueANs.size;

  dedupKeys.forEach(key => {
    const i = dedupMap[key].rowIdx;
    let has51 = false;
    let has56 = false;

    COUNSELING_FIELDS.forEach((field, idx) => {
      if (field.type === 'text' || field.type === 'date') return;
      const colIdx = idx + 3;
      if (data[i][colIdx] === true) {
        allItemCounts[field.key]++;

        if (field.category === '5.1') {
          has51 = true;
          detailedItems[field.label] = (detailedItems[field.label] || 0) + 1;
        } else if (field.category === '5.6' && field.subgroup) {
          has56 = true;
          scarsItems[field.label] = (scarsItems[field.label] || 0) + 1;
        } else if (field.category === '5.6' && field.subgroup === null) {
          // Skip top-level 5.6 entry (auto-derived, not user-checked)
        } else if (field.subgroup === null) {
          topLevelCounts[field.category]++;
        }
      }
    });

    if (has51) topLevelCounts['5.1']++;
    if (has56) {
      topLevelCounts['5.6']++;
      const rawDate = data[i][1];
      const dateStr = Utilities.formatDate(new Date(rawDate), tz, 'yyyy-MM-dd');
      const an = String(data[i][2]);
      const drugs = [];
      COUNSELING_FIELDS.forEach(function(field, idx) {
        if (field.category === '5.6' && field.subgroup && data[i][idx + 3] === true) {
          drugs.push(field.label);
        }
      });
      let startDateStr = '';
      if (startDateFieldIdx >= 0) {
        const rawStart = data[i][startDateFieldIdx + 3];
        if (rawStart instanceof Date && !isNaN(rawStart.getTime())) {
          startDateStr = Utilities.formatDate(rawStart, tz, 'yyyy-MM-dd');
        } else if (rawStart) {
          const d = new Date(rawStart);
          if (!isNaN(d.getTime())) startDateStr = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        }
      }
      scarsPatientRows.push({ rowNumber: i + 1, date: dateStr, startDate: startDateStr, an: an, drugs: drugs });
    }
  });

  scarsPatientRows.sort(function(a, b) { return a.date.localeCompare(b.date); });

  return { totalSessions: totalRecords, totalUniquePatients: totalUniquePatients, topLevelCounts: topLevelCounts, detailedItems: detailedItems, allItemCounts: allItemCounts, scarsItems: scarsItems, scarsPatientRows: scarsPatientRows };
}

function getDCDataByMonth(yearMonth) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DC);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const result = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    const d = new Date(data[i][1]);
    const rowYM = Utilities.formatDate(d, tz, 'yyyy-MM');
    if (rowYM !== yearMonth) continue;
    result.push({
      date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
      day: d.getDate(),
      count: data[i][2] || 0
    });
  }

  return result.sort((a, b) => a.date.localeCompare(b.date));
}

function getPEDataLast6Months() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PE);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const monthMap = {};

  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    const d = new Date(data[i][1]);
    const ym = Utilities.formatDate(d, tz, 'yyyy-MM');
    monthMap[ym] = (monthMap[ym] || 0) + (data[i][2] || 0);
  }

  const allMonths = Object.keys(monthMap).sort().reverse();
  const last6 = allMonths.slice(0, 6).reverse();

  return last6.map(ym => ({ month: ym, totalErrors: monthMap[ym] }));
}

function getDashboardData(yearMonth) {
  checkPermission_('dashboard');
  return {
    counseling: getCounselingSummaryByMonth(yearMonth),
    dc: getDCDataByMonth(yearMonth),
    pe: getPEDataLast6Months(),
    months: getAvailableMonths(),
    fieldConfig: COUNSELING_FIELDS.filter(f => f.type !== 'text' && f.type !== 'date')
  };
}

// Update only the Cat_5_6_SCARs_StartDate cell of a Counseling_Raw row.
function updateSCARsStartDate(rowNumber, isoDate) {
  checkPermission_('counseling');
  const row = parseInt(rowNumber, 10);
  if (!row || row < 2) return { success: false, error: 'Invalid row' };
  const sheet = getOrCreateSheet(SHEET_COUNSELING);
  if (row > sheet.getLastRow()) return { success: false, error: 'Row not found' };

  const fieldIdx = COUNSELING_FIELDS.findIndex(function(f) { return f.key === 'Cat_5_6_SCARs_StartDate'; });
  if (fieldIdx < 0) return { success: false, error: 'Field not configured' };
  const colNumber = fieldIdx + 4; // A=1, B=CounselingDate, C=AN, D+=fields

  const cell = sheet.getRange(row, colNumber);
  if (!isoDate) {
    cell.clearContent();
  } else {
    const d = new Date(isoDate);
    if (isNaN(d.getTime())) return { success: false, error: 'Invalid date' };
    cell.setValue(d);
  }
  return { success: true, message: 'อัพเดทวันที่เริ่มยาสำเร็จ' };
}

// ==========================================
// Report Summary Page Data
// ==========================================
function getReportSummaryData(yearMonth) {
  checkPermission_('report');
  const counseling = getCounselingSummaryByMonth(yearMonth);
  const dc = getDCDataByMonth(yearMonth);
  const tz = Session.getScriptTimeZone();

  // DC total = SUM of discharge counts for the month
  const dcTotal = dc.reduce(function(sum, d) { return sum + d.count; }, 0);

  // PE total = SUM of error counts for the month (direct query, not last-6-months)
  let peTotal = 0;
  const peSheet = getSpreadsheet().getSheetByName(SHEET_PE);
  if (peSheet && peSheet.getLastRow() > 1) {
    const peData = peSheet.getDataRange().getValues();
    for (let i = 1; i < peData.length; i++) {
      if (!peData[i][1]) continue;
      const ym = Utilities.formatDate(new Date(peData[i][1]), tz, 'yyyy-MM');
      if (ym === yearMonth) peTotal += (peData[i][2] || 0);
    }
  }

  const pePercentage = dcTotal > 0 ? (peTotal / dcTotal * 100) : 0;

  return {
    counseling: counseling,
    dcTotal: dcTotal,
    peTotal: peTotal,
    pePercentage: pePercentage
  };
}

// ==========================================
// HM Time Report Data
// ==========================================
function timeValueToMinutes_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (v instanceof Date) {
    return v.getHours() * 60 + v.getMinutes() + v.getSeconds() / 60;
  }
  const s = String(v).trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;
  const h = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  const ss = m[3] ? parseInt(m[3], 10) : 0;
  if (isNaN(h) || isNaN(mm)) return null;
  return h * 60 + mm + ss / 60;
}

function timeDiffMin_(startVal, endVal) {
  const a = timeValueToMinutes_(startVal);
  const b = timeValueToMinutes_(endVal);
  if (a === null || b === null) return null;
  const d = b - a;
  if (d < 0) return null;
  return d;
}

function getHMTimeReportByMonth(yearMonth) {
  checkPermission_('report');
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_HM);
  const tz = Session.getScriptTimeZone();

  const empty = {
    yearMonth: yearMonth,
    dayCount: 0,
    durationsMin: { step01: null, step02: null, step03: null, step04: null, step05: null, sumSteps: null, sumSteps1to4: null, step09: null },
    avgDrugCount: null,
    perDrugMin: { step02: null, step03: null, step04: null, step05: null },
    rowCount: 0,
    rowsUsedPerMetric: { step01: 0, step02: 0, step03: 0, step04: 0, step05: 0, step09: 0, drugCount: 0 }
  };

  if (!sheet || sheet.getLastRow() <= 1) return empty;

  const data = sheet.getDataRange().getValues();
  const H = HM_HEADERS;
  const col = {};
  H.forEach(function(name, idx) { col[name] = idx; });

  const distinctDates = new Set();
  const sums = { step01: 0, step02: 0, step03: 0, step04: 0, step05: 0, step09: 0, gap89: 0, drugCount: 0 };
  const counts = { step01: 0, step02: 0, step03: 0, step04: 0, step05: 0, step09: 0, gap89: 0, drugCount: 0 };
  let rowsInMonth = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dischargeDate = row[col['DischargeDate']];
    if (!dischargeDate) continue;
    const d = dischargeDate instanceof Date ? dischargeDate : new Date(dischargeDate);
    if (isNaN(d.getTime())) continue;
    const ym = Utilities.formatDate(d, tz, 'yyyy-MM');
    if (ym !== yearMonth) continue;

    rowsInMonth++;
    distinctDates.add(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));

    const dc = row[col['DrugCount']];
    if (typeof dc === 'number' && !isNaN(dc)) {
      sums.drugCount += dc;
      counts.drugCount++;
    } else if (dc !== '' && dc !== null && dc !== undefined) {
      const n = parseFloat(dc);
      if (!isNaN(n)) { sums.drugCount += n; counts.drugCount++; }
    }

    const s01 = timeDiffMin_(row[col['Step02_NurseReceiveTime']], row[col['Step03_PharmCheckHMStart']]);
    if (s01 !== null) { sums.step01 += s01; counts.step01++; }

    const verifySpan = timeDiffMin_(row[col['Step03_PharmCheckHMStart']], row[col['Step06_PharmVerifyTime']]);
    if (verifySpan !== null) {
      const consultEdit = timeDiffMin_(row[col['Step04_PharmConsultStart']], row[col['Step05_PharmEditEnd']]);
      const pause = consultEdit !== null ? consultEdit : 0;
      sums.step02 += verifySpan - pause;
      counts.step02++;
    }

    const s03 = timeDiffMin_(row[col['Step07_DispenseStart']], row[col['Step07_DispenseEnd']]);
    if (s03 !== null) { sums.step03 += s03; counts.step03++; }

    const s04 = timeDiffMin_(row[col['Step08_PharmCheckStart']], row[col['Step08_PharmCheckEnd']]);
    if (s04 !== null) { sums.step04 += s04; counts.step04++; }

    const p10 = timeDiffMin_(row[col['Step10_PharmCheck2Start']], row[col['Step10_PharmCheck2End']]);
    const pLate = timeDiffMin_(row[col['Step10_PharmCheck2Start']], row[col['Step11_LatePickupTime']]);
    if (p10 !== null || pLate !== null) {
      sums.step05 += (p10 || 0) + (pLate || 0);
      counts.step05++;
    }

    const s09 = timeDiffMin_(row[col['Step02_NurseReceiveTime']], row[col['Step09_WardChartTime']]);
    if (s09 !== null) { sums.step09 += s09; counts.step09++; }

    const g89 = timeDiffMin_(row[col['Step08_PharmCheckEnd']], row[col['Step09_WardChartTime']]);
    if (g89 !== null) { sums.gap89 += g89; counts.gap89++; }
  }

  function avg(key) { return counts[key] > 0 ? sums[key] / counts[key] : null; }

  const step01 = avg('step01');
  const step02 = avg('step02');
  const step03 = avg('step03');
  const step04 = avg('step04');
  const step05 = avg('step05');
  const step09 = avg('step09');
  const avgDrugCount = avg('drugCount');

  const gap89 = avg('gap89');
  const stepVals = [step01, step02, step03, step04, step05];
  const sumSteps = stepVals.every(function(v) { return v !== null; })
    ? stepVals.reduce(function(a, b) { return a + b; }, 0) + (gap89 || 0)
    : null;
  const steps1to4 = [step01, step02, step03, step04];
  const sumSteps1to4 = steps1to4.every(function(v) { return v !== null; })
    ? steps1to4.reduce(function(a, b) { return a + b; }, 0)
    : null;

  function perDrug(v) {
    if (v === null || avgDrugCount === null || avgDrugCount === 0) return null;
    return v / avgDrugCount;
  }

  return {
    yearMonth: yearMonth,
    dayCount: distinctDates.size,
    durationsMin: { step01: step01, step02: step02, step03: step03, step04: step04, step05: step05, sumSteps: sumSteps, sumSteps1to4: sumSteps1to4, step09: step09 },
    avgDrugCount: avgDrugCount,
    perDrugMin: { step02: perDrug(step02), step03: perDrug(step03), step04: perDrug(step04), step05: perDrug(step05) },
    rowCount: rowsInMonth,
    rowsUsedPerMetric: counts
  };
}

// ==========================================
// Setup HM Report Summary (formula-driven mirror of HM Time web report)
// ==========================================
function setupHMReportSummary() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_HM_REPORT);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(SHEET_HM_REPORT);

  // Column letter for each HM_HEADERS field (1-based)
  function hmCol(name) {
    const idx = HM_HEADERS.indexOf(name);
    if (idx < 0) throw new Error('Missing HM header: ' + name);
    return colLetter(idx + 1);
  }

  const RAW = 'HM_Time_Raw';
  const C_DATE = hmCol('DischargeDate');                  // B (canonical visit date)
  const C_S02T = hmCol('Step02_NurseReceiveTime');        // H
  const C_S03  = hmCol('Step03_PharmCheckHMStart');       // J
  const C_S04S = hmCol('Step04_PharmConsultStart');       // K
  const C_S05E = hmCol('Step05_PharmEditEnd');            // N
  const C_S06  = hmCol('Step06_PharmVerifyTime');         // O
  const C_S07S = hmCol('Step07_DispenseStart');           // P
  const C_S07E = hmCol('Step07_DispenseEnd');             // Q
  const C_S08S = hmCol('Step08_PharmCheckStart');         // R
  const C_S08E = hmCol('Step08_PharmCheckEnd');           // S
  const C_S09T = hmCol('Step09_WardChartTime');           // T
  const C_S10S = hmCol('Step10_PharmCheck2Start');        // V
  const C_S10E = hmCol('Step10_PharmCheck2End');          // W
  const C_S11  = hmCol('Step11_LatePickupTime');          // X
  const C_DRUG = hmCol('DrugCount');                      // D

  // Month selector cell (B2). START/END expressions reference $B$2.
  // END is start of next month (exclusive) so any time-of-day on the last
  // day of the selected month is included via "<" comparisons.
  const START = 'DATE(YEAR($B$2),MONTH($B$2),1)';
  const END   = 'EOMONTH($B$2,0)+1';

  // Range references (used in SUMPRODUCT over HM_Time_Raw rows 2..end)
  function rng(col) { return RAW + '!' + col + '2:' + col; }

  // Per-cell "time to number of days since midnight" expression that works
  // whether the cell stores a text like "HH:MM" or a numeric time value.
  // Evaluates to 0 when cell is blank (caller must guard with a nonblank mask).
  function tv(col) {
    const r = rng(col);
    return 'IFERROR(TIMEVALUE(' + r + '),IFERROR(' + r + '*1,0))';
  }

  // Date-in-month mask
  const MASK_MONTH = '(' + rng(C_DATE) + '>=' + START + ')*(' + rng(C_DATE) + '<' + END + ')';

  // Generic avg duration (end - start) in minutes, filtered by month,
  // requires both endpoints non-blank. Returns "-" if no rows.
  function avgDuration(startCol, endCol) {
    const mask = MASK_MONTH +
      '*(' + rng(startCol) + '<>"")*(' + rng(endCol) + '<>"")';
    const num = 'SUMPRODUCT(' + mask + '*((' + tv(endCol) + ')-(' + tv(startCol) + ')))*1440';
    const den = 'SUMPRODUCT(' + mask + ')';
    return '=IFERROR(' + num + '/' + den + ',"-")';
  }

  // O2: (Step06 - Step03) - (Step05End - Step04Start when both present)
  function avgO2() {
    const maskBase = MASK_MONTH +
      '*(' + rng(C_S03) + '<>"")*(' + rng(C_S06) + '<>"")';
    const verifySpan = '((' + tv(C_S06) + ')-(' + tv(C_S03) + '))';
    // consult pause only when both consultStart and editEnd exist
    const pauseMask = '((' + rng(C_S04S) + '<>"")*(' + rng(C_S05E) + '<>""))';
    const pause = pauseMask + '*((' + tv(C_S05E) + ')-(' + tv(C_S04S) + '))';
    const num = 'SUMPRODUCT(' + maskBase + '*(' + verifySpan + '-' + pause + '))*1440';
    const den = 'SUMPRODUCT(' + maskBase + ')';
    return '=IFERROR(' + num + '/' + den + ',"-")';
  }

  // O5: (Step10End - Step10Start) + (Step11 - Step10Start), missing parts = 0.
  // Row counts iff Step10Start is present AND at least one of the two diffs is >= 0
  // (matches timeDiffMin_ which returns null on negative diffs, mirroring the JS logic).
  function avgO5() {
    const diff10 = '((' + tv(C_S10E) + ')-(' + tv(C_S10S) + '))';
    const diffLate = '((' + tv(C_S11) + ')-(' + tv(C_S10S) + '))';
    const ok10 = '(' + rng(C_S10E) + '<>"")*(' + diff10 + '>=0)';
    const okLate = '(' + rng(C_S11) + '<>"")*(' + diffLate + '>=0)';
    const anyOk = '(((' + ok10 + ')+(' + okLate + '))>0)';
    const mask = MASK_MONTH + '*(' + rng(C_S10S) + '<>"")*' + anyOk;
    const contrib = '((' + ok10 + ')*' + diff10 + '+(' + okLate + ')*' + diffLate + ')';
    const num = 'SUMPRODUCT(' + mask + '*' + contrib + ')*1440';
    const den = 'SUMPRODUCT(' + mask + ')';
    return '=IFERROR(' + num + '/' + den + ',"-")';
  }

  // --- Build header / info block ---
  // Row 1: title spanning A:E
  sheet.getRange('A1').setValue('ระยะเวลาเฉลี่ย กระบวนการยากลับบ้าน');
  sheet.getRange('A1:E1').merge();

  // Row 2: ประจำเดือน (month selector in B2)
  sheet.getRange('A2').setValue('ประจำเดือน');
  sheet.getRange('B2').setValue(new Date(new Date().getFullYear(), new Date().getMonth(), 1));
  sheet.getRange('B2').setNumberFormat('MMMM yyyy');

  // Row 3: จำนวนวัน
  sheet.getRange('A3').setValue('จำนวนวัน');
  sheet.getRange('B3').setFormula(
    '=IFERROR(COUNTA(UNIQUE(FILTER(INT(' + rng(C_DATE) + '),' +
      rng(C_DATE) + '>=' + START + ',' + rng(C_DATE) + '<' + END + '))),0) & " วัน"'
  );

  // Row 4: จำนวนผู้ป่วยที่เก็บข้อมูล (count of records with DischargeDate in selected month)
  sheet.getRange('A4').setValue('จำนวนผู้ป่วยที่เก็บข้อมูล');
  sheet.getRange('B4').setFormula(
    '=COUNTIFS(' + RAW + '!' + C_DATE + ':' + C_DATE + ',">="&' + START + ',' +
    RAW + '!' + C_DATE + ':' + C_DATE + ',"<"&' + END + ') & " ราย"'
  );

  // Row 5: จำนวนขนานยากลับบ้านเฉลี่ย (moved from table into info block)
  sheet.getRange('A5').setValue('จำนวนขนานยากลับบ้านเฉลี่ย');
  sheet.getRange('B5').setFormula(
    '=IFERROR(ROUND(AVERAGEIFS(' + RAW + '!' + C_DRUG + ':' + C_DRUG + ',' +
    RAW + '!' + C_DATE + ':' + C_DATE + ',">="&' + START + ',' +
    RAW + '!' + C_DATE + ':' + C_DATE + ',"<"&' + END + '),2),"-")'
  );

  // Row 7: table header (row 6 is now blank spacer)
  const header = ['รายละเอียด', 'เกณฑ์/เป้าหมาย (นาที)', 'ระยะเวลาเฉลี่ย (นาที)', 'ระยะเวลาเฉลี่ย/ขนานยา (นาที/ขนานยา)', 'สถานะ'];
  sheet.getRange(7, 1, 1, header.length).setValues([header]);

  // Rows 8..15
  const ROW_O1 = 8;
  const ROW_O2 = 9;
  const ROW_O3 = 10;
  const ROW_O4 = 11;
  const ROW_O5 = 12;
  const ROW_SUM_1_4 = 13;
  const ROW_SUM_1_5 = 14;
  const ROW_DRUG_CELL = '$B$5'; // drug avg moved to info block
  const ROW_S09 = 15;

  const rowSpec = [
    { row: ROW_O1, label: '01 กระบวนการ - ก่อนตรวจสอบรายการยากลับบ้าน [Ward รับคำสั่ง - เริ่มตรวจสอบ]',
      formula: avgDuration(C_S02T, C_S03), perDrug: false },
    { row: ROW_O2, label: '02 กระบวนการ - คัดกรองคำสั่งใช้ยากลับบ้าน',
      formula: avgO2(), perDrug: true },
    { row: ROW_O3, label: '03 กระบวนการ - จัดยา',
      formula: avgDuration(C_S07S, C_S07E), perDrug: true },
    { row: ROW_O4, label: '04 กระบวนการ - ตรวจสอบความถูกต้องของรายการยากลับบ้าน',
      formula: avgDuration(C_S08S, C_S08E), perDrug: true },
    { row: ROW_O5, label: '05 กระบวนการ - ผู้ป่วยติดต่อรับยา - จ่ายยากลับบ้าน',
      formula: avgO5(), perDrug: true }
  ];

  // Columns (new layout): A=label, B=target, C=avg, D=perDrug, E=status
  rowSpec.forEach(function(s) {
    sheet.getRange(s.row, 1).setValue(s.label);
    sheet.getRange(s.row, 3).setFormula(s.formula);
    if (s.perDrug) {
      sheet.getRange(s.row, 4).setFormula('=IFERROR(IF(ISNUMBER(C' + s.row + '),C' + s.row + '/' + ROW_DRUG_CELL + ',"-"),"-")');
    }
  });

  // Row 12: Sum O1..O4 + target + pass/fail tag
  sheet.getRange(ROW_SUM_1_4, 1).setValue('ระยะเวลากระบวนการยากลับบ้าน [ตั้งแต่คำสั่งใช้ยากลับบ้านแสดงบน Dashboard ถึง เภสัชกรตรวจสอบความถูกต้องของรายการยากลับบ้าน]');
  sheet.getRange(ROW_SUM_1_4, 2).setValue(45);
  sheet.getRange(ROW_SUM_1_4, 3).setFormula(
    '=IFERROR(IF(AND(ISNUMBER(C' + ROW_O1 + '),ISNUMBER(C' + ROW_O2 + '),ISNUMBER(C' + ROW_O3 + '),ISNUMBER(C' + ROW_O4 + ')),' +
    'C' + ROW_O1 + '+C' + ROW_O2 + '+C' + ROW_O3 + '+C' + ROW_O4 + ',"-"),"-")'
  );
  sheet.getRange(ROW_SUM_1_4, 5).setFormula(
    '=IF(ISNUMBER(C' + ROW_SUM_1_4 + '),IF(C' + ROW_SUM_1_4 + '<=B' + ROW_SUM_1_4 + ',"ผ่านเกณฑ์","ไม่ผ่านเกณฑ์"),"")'
  );

  // Row 13: Sum O1..O5 + gap (Step09 - Step08)
  sheet.getRange(ROW_SUM_1_5, 1).setValue('ระยะเวลากระบวนการยากลับบ้าน ห้องยาผู้ป่วยใน ทั้งหมด [ตั้งแต่คำสั่งใช้ยากลับบ้านแสดงบน Dashboard ถึง เภสัชกรจ่ายยากลับบ้านให้ผู้ป่วยเสร็จสิ้น]');
  // Gap (Step09_WardChartTime - Step08_PharmCheckEnd) average — matches JS timeDiffMin_:
  // include row only if BOTH cells parse to a valid (non-zero) time AND diff >= 0.
  // Treats blank/0/unparseable as missing (per-row contribution = 0, row excluded from denom).
  function avgGap89() {
    const startV = tv(C_S08E);
    const endV = tv(C_S09T);
    const diff = '((' + endV + ')-(' + startV + '))';
    const validStart = '(' + rng(C_S08E) + '<>"")*((' + startV + ')>0)';
    const validEnd   = '(' + rng(C_S09T) + '<>"")*((' + endV + ')>0)';
    const validDiff  = '(' + diff + '>=0)';
    const mask = MASK_MONTH + '*' + validStart + '*' + validEnd + '*' + validDiff;
    const num = 'SUMPRODUCT(' + mask + '*' + diff + ')*1440';
    const den = 'SUMPRODUCT(' + mask + ')';
    return 'IFERROR(' + num + '/' + den + ',0)';
  }
  const gap89Expr = avgGap89();
  sheet.getRange(ROW_SUM_1_5, 3).setFormula(
    '=IFERROR(IF(AND(ISNUMBER(C' + ROW_O1 + '),ISNUMBER(C' + ROW_O2 + '),ISNUMBER(C' + ROW_O3 + '),ISNUMBER(C' + ROW_O4 + '),ISNUMBER(C' + ROW_O5 + ')),' +
    'C' + ROW_O1 + '+C' + ROW_O2 + '+C' + ROW_O3 + '+C' + ROW_O4 + '+C' + ROW_O5 + '+' + gap89Expr + ',"-"),"-")'
  );

  // Row 15: Step09 extra
  sheet.getRange(ROW_S09, 1).setValue('ระยะเวลาตั้งแต่ Ward รับคำสั่งยากลับบ้าน - ส่ง Chart และ คืนยา');
  sheet.getRange(ROW_S09, 3).setFormula(avgDuration(C_S02T, C_S09T));

  // --- Formatting ---
  sheet.setColumnWidth(1, 680);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 220);

  // Title row (merged A1:E1)
  sheet.getRange('A1:E1').setBackground('#7BB8C0').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');

  // Info block
  sheet.getRange('A2:A5').setFontWeight('bold').setBackground('#EAEEFF');
  sheet.getRange('B2:B5').setFontWeight('bold');

  // Header row
  sheet.getRange(7, 1, 1, 5).setBackground('#A7B5FE').setFontColor('#ffffff').setFontWeight('bold');

  // Data rows number formatting (columns C and D)
  sheet.getRange(ROW_O1, 3, ROW_S09 - ROW_O1 + 1, 2).setNumberFormat('0.00');

  // Center-align numeric columns (B target, C avg, D per-drug, E status) in data rows
  sheet.getRange(ROW_O1, 2, ROW_S09 - ROW_O1 + 1, 4).setHorizontalAlignment('center');

  // Sum rows highlight
  sheet.getRange(ROW_SUM_1_4, 1, 1, 5).setBackground('#FEB737').setFontWeight('bold').setFontColor('#ffffff');
  sheet.getRange(ROW_SUM_1_5, 1, 1, 5).setBackground('#FEB737').setFontWeight('bold').setFontColor('#ffffff');

  // Wrap for long label column and status column
  sheet.getRange('A1:A' + ROW_S09).setWrap(true).setVerticalAlignment('middle');
  sheet.getRange('E1:E' + ROW_S09).setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');

  // Conditional formatting on E12 (pass/fail tag)
  const tagRange = sheet.getRange(ROW_SUM_1_4, 5);
  const passRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('ผ่านเกณฑ์')
    .setBackground('#2d6a4f').setFontColor('#ffffff').setBold(true)
    .setRanges([tagRange])
    .build();
  const failRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('ไม่ผ่านเกณฑ์')
    .setBackground('#c1121f').setFontColor('#ffffff').setBold(true)
    .setRanges([tagRange])
    .build();
  sheet.setConditionalFormatRules([passRule, failRule]);

  sheet.setFrozenRows(7);
}
