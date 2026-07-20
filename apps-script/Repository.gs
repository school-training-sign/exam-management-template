function getSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const id = properties.getProperty(APP.PROPERTY_SHEET_ID);
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error("연결된 통합문서를 찾지 못했습니다.");
  properties.setProperty(APP.PROPERTY_SHEET_ID, active.getId());
  return active;
}

function ensureSheets_() {
  const spreadsheet = getSpreadsheet_();
  Object.keys(HEADERS).forEach((name) => {
    let sheet = spreadsheet.getSheetByName(name);
    if (!sheet) sheet = spreadsheet.insertSheet(name);
    const headers = HEADERS[name];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground("#0d2343")
        .setFontColor("#ffffff")
        .setFontWeight("bold");
    }
  });
  let guide = spreadsheet.getSheetByName(SHEETS.GUIDE);
  if (!guide) guide = spreadsheet.insertSheet(SHEETS.GUIDE, 0);
  if (guide.getLastRow() === 0) {
    guide.getRange("A1:B8").setValues([
      ["정기고사 관리 시스템", "배포용 원본"],
      ["1", "상단 메뉴 ‘정기고사 관리 → 초기 설정’을 실행합니다."],
      ["2", "학교명, 학교코드, 관리자 암호를 입력합니다."],
      ["3", "Apps Script를 웹 앱으로 배포합니다. 실행 사용자는 본인, 접근 권한은 모든 사용자로 설정합니다."],
      ["4", "웹 앱 주소를 GitHub Pages 저장소 public/config.js의 API_URL에 입력합니다."],
      ["보안", "학교코드와 관리자 암호의 원문은 시트에 저장되지 않습니다."],
      ["데이터", "학생 이름은 학생 시트에만 있으며 감사로그에는 남지 않습니다."],
      ["버전", APP.VERSION],
    ]);
    guide.getRange("A1:B1").merge().setBackground("#0d2343").setFontColor("#ffffff").setFontWeight("bold");
    guide.setColumnWidth(1, 100);
    guide.setColumnWidth(2, 620);
    guide.getRange("A1:B8").setWrap(true).setVerticalAlignment("middle");
  }
}

function normalizeCell_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, APP.TIMEZONE, "yyyy-MM-dd");
  return value;
}

function getRecords_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error("'" + sheetName + "' 시트를 찾지 못했습니다.");
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS[sheetName].length);
  if (lastRow < 2) return [];
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map(String);
  return values.slice(1).filter((row) => row.some((value) => value !== "")).map((row) => {
    const record = {};
    headers.forEach((header, index) => {
      if (header) record[header] = normalizeCell_(row[index]);
    });
    return record;
  });
}

function replaceRecords_(sheetName, records) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  const headers = HEADERS[sheetName];
  if (!sheet) throw new Error("'" + sheetName + "' 시트를 찾지 못했습니다.");
  const oldRows = Math.max(0, sheet.getLastRow() - 1);
  if (oldRows) sheet.getRange(2, 1, oldRows, Math.max(sheet.getLastColumn(), headers.length)).clearContent();
  if (!records.length) return;
  const values = records.map((record) => headers.map((header) => record[header] == null ? "" : record[header]));
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
}

function appendAudit_(role, action, targetType, targetId, count, success, detail) {
  const sheet = getSpreadsheet_().getSheetByName(SHEETS.AUDIT);
  sheet.appendRow([
    createId_("audit"),
    nowIso_(),
    role || "",
    action || "",
    targetType || "",
    targetId || "",
    Number(count) || 0,
    Boolean(success),
    String(detail || "").slice(0, 250),
  ]);
}

function getSettings_() {
  const result = {};
  getRecords_(SHEETS.SETTINGS).forEach((row) => { result[String(row.key)] = row.value; });
  return result;
}

function setSetting_(key, value) {
  const now = nowIso_();
  const rows = getRecords_(SHEETS.SETTINGS).filter((item) => String(item.key) !== String(key));
  rows.push({ key: key, value: value, updated_at: now });
  replaceRecords_(SHEETS.SETTINGS, rows);
}

function parseJson_(value, fallback) {
  if (!value) return fallback;
  try { return JSON.parse(String(value)); } catch (error) { return fallback; }
}

function serializeSeat_(chart) {
  return {
    id: chart.id || createId_("seat"),
    mode: chart.mode || "separate",
    exam_date: chart.exam_date || "",
    period: Number(chart.period) || 1,
    grade: Number(chart.grade) || "",
    class_id: chart.class_id || "",
    subject_name: chart.subject_name || "",
    room_name: chart.room_name || "",
    rows: Number(chart.rows) || 5,
    cols: Number(chart.cols) || 6,
    start_side: chart.start_side || "window",
    disabled_seats_json: JSON.stringify(chart.disabled_seats || []),
    examinee_ids_json: JSON.stringify(chart.examinee_ids || []),
    assignment_json: JSON.stringify(chart.assignment || []),
    updated_at: nowIso_(),
  };
}

function deserializeSeat_(row) {
  return {
    id: row.id,
    mode: row.mode,
    exam_date: row.exam_date,
    period: Number(row.period),
    grade: Number(row.grade) || "",
    class_id: row.class_id,
    subject_name: row.subject_name,
    room_name: row.room_name,
    rows: Number(row.rows),
    cols: Number(row.cols),
    start_side: row.start_side,
    disabled_seats: parseJson_(row.disabled_seats_json, []),
    examinee_ids: parseJson_(row.examinee_ids_json, []),
    assignment: parseJson_(row.assignment_json, []),
    updated_at: row.updated_at,
  };
}

function teacherBootstrap_() {
  return {
    settings: getSettings_(),
    classes: getRecords_(SHEETS.CLASSES).filter((item) => item.active !== false && String(item.active).toLowerCase() !== "false"),
    exam_dates: getRecords_(SHEETS.EXAM_DATES).filter((item) => item.active !== false && String(item.active).toLowerCase() !== "false"),
  };
}

function adminBootstrap_() {
  const result = teacherBootstrap_();
  result.students = getRecords_(SHEETS.STUDENTS).filter((item) => item.active !== false && String(item.active).toLowerCase() !== "false");
  result.timetable = getRecords_(SHEETS.TIMETABLE);
  result.enrollments = getRecords_(SHEETS.ENROLLMENTS);
  result.seat_charts = getRecords_(SHEETS.SEATING).map(deserializeSeat_).reverse();
  return result;
}

function decorateAbsences_(rows) {
  const students = new Map(getRecords_(SHEETS.STUDENTS).map((item) => [String(item.id), item]));
  const classes = new Map(getRecords_(SHEETS.CLASSES).map((item) => [String(item.id), item]));
  return rows.map((item) => {
    const student = students.get(String(item.student_id)) || {};
    const classInfo = classes.get(String(item.class_id)) || {};
    return Object.assign({}, item, {
      grade: Number(classInfo.grade) || "",
      class_num: Number(classInfo.class_num) || "",
      class_label: classInfo.id ? classInfo.grade + "학년 " + classInfo.class_num + "반" : "",
      student_number: Number(student.number) || "",
      student_name: student.name || "",
    });
  });
}
