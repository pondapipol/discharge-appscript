// ==========================================
// Configuration
// ==========================================
const SHEET_COUNSELING = 'Counseling_Raw';
const SHEET_DC = 'DC_Raw';
const SHEET_PE = 'PE_Raw';
const SHEET_REPORT = 'Report Summary';

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
  { key: 'Cat_5_2_Warfarin', label: '5.2 แนะนำการใช้ยา Warfarin', category: '5.2', subgroup: null },
  { key: 'Cat_5_3_Chronic_TB_ARV', label: '5.3 แนะนำการใช้ยาผู้ป่วยโรคเรื้อรัง (TB/ARV)', category: '5.3', subgroup: null },
  { key: 'Cat_5_4_Myanmar_Label', label: '5.4 ผู้ป่วยพม่า ออกฉลากภาษาพม่า', category: '5.4', subgroup: null },
  { key: 'Cat_5_5_Stroke_Case', label: '5.5 แนะนำ ผู้ป่วย Stroke case', category: '5.5', subgroup: null },
  { key: 'Cat_5_6_SJS_TEN_Risk', label: '5.6 แนะนำ ผู้ป่วยที่ได้รับยาที่เสี่ยงต่อการเกิด SJS/TEN', category: '5.6', subgroup: null },
];

// ==========================================
// Web App Entry Point
// ==========================================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('Discharge Counseling Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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
// Setup (run once)
// ==========================================
function setupSheets() {
  // Counseling_Raw headers
  const counselingHeaders = ['Timestamp', 'CounselingDate', 'AN'];
  COUNSELING_FIELDS.forEach(f => counselingHeaders.push(f.key));
  getOrCreateSheet(SHEET_COUNSELING, counselingHeaders);

  // DC_Raw headers
  getOrCreateSheet(SHEET_DC, ['Timestamp', 'DischargeDate', 'DischargeCount']);

  // PE_Raw headers
  getOrCreateSheet(SHEET_PE, ['Timestamp', 'ErrorDate', 'ErrorCount']);

  // Report Summary - preserve existing
  getOrCreateSheet(SHEET_REPORT, []);
}

// ==========================================
// Get field config (for client)
// ==========================================
function getCounselingFieldConfig() {
  return COUNSELING_FIELDS;
}

// ==========================================
// Submit Counseling Data
// ==========================================
function submitCounselingData(formData) {
  if (!formData.an || !/^\d{9}$/.test(formData.an)) {
    return { success: false, error: 'AN ต้องเป็นตัวเลข 9 หลัก' };
  }
  if (!formData.counselingDate) {
    return { success: false, error: 'กรุณาเลือกวันที่' };
  }

  const sheet = getOrCreateSheet(SHEET_COUNSELING);
  const row = [new Date(), new Date(formData.counselingDate), formData.an];

  COUNSELING_FIELDS.forEach(field => {
    row.push(formData.checkboxes[field.key] === true);
  });

  sheet.appendRow(row);

  // Format AN column as plain text to preserve leading zeros
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 3).setNumberFormat('@');

  return { success: true, message: 'บันทึกข้อมูล Counseling สำเร็จ' };
}

// ==========================================
// Submit DC Data (upsert by date)
// ==========================================
function submitDCData(formData) {
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
// Dashboard Data Functions
// ==========================================
function getAvailableMonths() {
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
    return { totalSessions: 0, topLevelCounts: {}, detailedItems: {} };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tz = Session.getScriptTimeZone();

  const topLevelCounts = { '5.1': 0, '5.2': 0, '5.3': 0, '5.4': 0, '5.5': 0, '5.6': 0 };
  const detailedItems = {};
  let totalSessions = 0;

  // Initialize detailed items for 5.1
  COUNSELING_FIELDS.forEach(f => {
    if (f.category === '5.1') {
      detailedItems[f.label] = 0;
    }
  });

  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][1];
    if (!rowDate) continue;
    const rowYM = Utilities.formatDate(new Date(rowDate), tz, 'yyyy-MM');
    if (rowYM !== yearMonth) continue;

    totalSessions++;
    let has51 = false;

    COUNSELING_FIELDS.forEach((field, idx) => {
      const colIdx = idx + 3; // offset for Timestamp, Date, AN
      if (data[i][colIdx] === true) {
        if (field.category === '5.1') {
          has51 = true;
          detailedItems[field.label] = (detailedItems[field.label] || 0) + 1;
        } else {
          topLevelCounts[field.category]++;
        }
      }
    });

    if (has51) topLevelCounts['5.1']++;
  }

  return { totalSessions, topLevelCounts, detailedItems };
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
  return {
    counseling: getCounselingSummaryByMonth(yearMonth),
    dc: getDCDataByMonth(yearMonth),
    pe: getPEDataLast6Months(),
    months: getAvailableMonths()
  };
}
