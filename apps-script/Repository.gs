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
  Object.keys(HEADERS).forEach((name) => ensureTableSheet_(spreadsheet, name, HEADERS[name]));

  let guide = spreadsheet.getSheetByName(SHEETS.GUIDE);
  if (!guide) guide = spreadsheet.insertSheet(SHEETS.GUIDE, 0);
  if (guide.getLastRow() === 0) {
    guide.getRange("A1:B8").setValues([
      ["정기고사 관리 시스템", "배포용 원본"],
      ["1", "상단 메뉴 '정기고사 관리 > 초기 설정'을 실행합니다."],
      ["2", "학교명, 학교코드, 관리자 암호를 입력합니다."],
      ["3", "Apps Script를 웹 앱으로 배포합니다. 실행 사용자는 본인, 접근 권한은 조직 정책이 허용하는 범위로 설정합니다."],
      ["4", "웹 앱 주소를 GitHub Pages 공개 설정의 API_URL에 입력합니다."],
      ["보안", "학교코드와 관리자 암호의 원문은 시트에 저장되지 않습니다."],
      ["데이터", "학생 이름은 학생 시트에만 저장되며 감사로그에는 남기지 않습니다."],
      ["버전", APP.VERSION],
    ]);
    guide.getRange("A1:B1").merge().setBackground("#0d2343").setFontColor("#ffffff").setFontWeight("bold");
    guide.setColumnWidth(1, 100);
    guide.setColumnWidth(2, 620);
    guide.getRange("A1:B8").setWrap(true).setVerticalAlignment("middle");
  }

  const settings = getRecords_(SHEETS.SETTINGS);
  const schema = settings.find((item) => String(item.key) === "schema_version");
  if (!schema || String(schema.value) !== APP.SCHEMA_VERSION) {
    setSetting_("schema_version", APP.SCHEMA_VERSION);
  }
}

function ensureTableSheet_(spreadsheet, name, requiredHeaders) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);

  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  } else {
    const width = Math.max(1, sheet.getLastColumn());
    const existing = sheet.getRange(1, 1, 1, width).getValues()[0].map((value) => String(value || "").trim());
    const missing = requiredHeaders.filter((header) => existing.indexOf(header) === -1);
    if (missing.length) {
      sheet.getRange(1, width + 1, 1, missing.length).setValues([missing]);
    }
  }

  const headerWidth = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headerWidth)
    .setBackground("#0d2343")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
}

function normalizeCell_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, APP.TIMEZONE, "yyyy-MM-dd");
  return value;
}

function dataCacheKey_(sheetName) {
  return "records:" + sheetName;
}

function getRecords_(sheetName) {
  if (CACHEABLE_SHEETS.has(sheetName)) {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(dataCacheKey_(sheetName));
    if (cached) {
      try { return JSON.parse(cached); } catch (ignored) {}
    }
    const records = getRecordsFresh_(sheetName);
    try {
      cache.put(dataCacheKey_(sheetName), JSON.stringify(records), APP.DATA_CACHE_TTL_SECONDS);
    } catch (ignored) {}
    return records;
  }
  return getRecordsFresh_(sheetName);
}

function getRecordsFresh_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error("'" + sheetName + "' 시트를 찾지 못했습니다.");
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), (HEADERS[sheetName] || []).length);
  if (lastRow < 2) return [];
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map((value) => String(value || "").trim());
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
  if (!sheet) throw new Error("'" + sheetName + "' 시트를 찾지 못했습니다.");
  const canonicalHeaders = HEADERS[sheetName] || [];
  const lastColumn = Math.max(sheet.getLastColumn(), canonicalHeaders.length);
  const actualHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map((value) => String(value || "").trim());
  const oldRows = Math.max(0, sheet.getLastRow() - 1);
  if (oldRows) sheet.getRange(2, 1, oldRows, lastColumn).clearContent();
  if (records.length) {
    const values = records.map((record) => actualHeaders.map((header) => {
      if (!header) return "";
      return record[header] == null ? "" : record[header];
    }));
    sheet.getRange(2, 1, values.length, actualHeaders.length).setValues(values);
  }
  invalidateDataCaches_([sheetName]);
}

function invalidateDataCaches_(sheetNames) {
  const cache = CacheService.getScriptCache();
  const names = (sheetNames || Array.from(CACHEABLE_SHEETS)).filter((name, index, all) => all.indexOf(name) === index);
  const keys = names.filter((name) => CACHEABLE_SHEETS.has(name)).map(dataCacheKey_);
  if (keys.length) cache.removeAll(keys);
}

function sanitizeAuditDetail_(success, detail) {
  if (success) return "OK";
  const code = String(detail || "REQUEST_FAILED").toUpperCase().replace(/[^A-Z0-9_:-]/g, "");
  return (code || "REQUEST_FAILED").slice(0, 80);
}

function appendAudit_(role, action, targetType, targetId, count, success, detail) {
  const sheet = getSpreadsheet_().getSheetByName(SHEETS.AUDIT);
  sheet.appendRow([
    createId_("audit"),
    nowIso_(),
    String(role || "").slice(0, 20),
    String(action || "").slice(0, 60),
    String(targetType || "").slice(0, 30),
    String(targetId || "").slice(0, 100),
    Number(count) || 0,
    Boolean(success),
    sanitizeAuditDetail_(Boolean(success), detail),
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
  if (value == null || value === "") return fallback;
  if (Array.isArray(value) || (typeof value === "object" && value !== null)) return value;
  try { return JSON.parse(String(value)); } catch (error) { return fallback; }
}

function normalizeClassIds_(value) {
  let values = value;
  if (typeof value === "string") {
    const parsed = parseJson_(value, null);
    values = Array.isArray(parsed) ? parsed : value.split(/[|,]/);
  }
  if (!Array.isArray(values)) values = [];
  return values.map((item) => String(item || "").trim()).filter(Boolean)
    .filter((item, index, all) => all.indexOf(item) === index);
}

function isActive_(item) {
  return item && item.active !== false && String(item.active).toLowerCase() !== "false";
}

function serializeSeat_(chart) {
  return {
    id: chart.id || createId_("seat"),
    mode: chart.mode === "own" ? "own" : "separate",
    exam_date: String(chart.exam_date || ""),
    period: Number(chart.period) || 1,
    grade: Number(chart.grade) || "",
    class_id: String(chart.class_id || ""),
    subject_name: String(chart.subject_name || "").slice(0, 80),
    room_name: String(chart.room_name || "").slice(0, 80),
    rows: Number(chart.rows) || 6,
    cols: Number(chart.cols) || 5,
    start_side: chart.start_side === "aisle" ? "aisle" : "window",
    seat_order: chart.seat_order === "column" ? "column" : "row",
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
    seat_order: row.seat_order === "column" ? "column" : "row",
    disabled_seats: parseJson_(row.disabled_seats_json, []),
    examinee_ids: parseJson_(row.examinee_ids_json, []),
    assignment: parseJson_(row.assignment_json, []),
    updated_at: row.updated_at,
  };
}

function seatNaturalKey_(chart) {
  return [
    chart.mode || "separate",
    chart.exam_date || "",
    Number(chart.period) || 1,
    Number(chart.grade) || "",
    chart.class_id || "",
    chart.subject_name || "",
    chart.room_name || "",
  ].join("|");
}

function teacherBootstrap_() {
  return {
    settings: getSettings_(),
    classes: getRecords_(SHEETS.CLASSES).filter(isActive_),
    exam_dates: getRecords_(SHEETS.EXAM_DATES).filter(isActive_),
  };
}

function adminBootstrap_() {
  const result = teacherBootstrap_();
  result.students = getRecords_(SHEETS.STUDENTS).filter(isActive_);
  result.timetable = getRecords_(SHEETS.TIMETABLE).map((item) => Object.assign({}, item, {
    class_ids: normalizeClassIds_(item.class_ids),
  }));
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
      student_active: isActive_(student),
    });
  });
}

function decorateCompletions_(rows) {
  const classes = new Map(getRecords_(SHEETS.CLASSES).map((item) => [String(item.id), item]));
  return rows.map((item) => {
    const classInfo = classes.get(String(item.class_id)) || {};
    return Object.assign({}, item, {
      grade: Number(classInfo.grade) || "",
      class_num: Number(classInfo.class_num) || "",
      class_label: classInfo.id ? classInfo.grade + "학년 " + classInfo.class_num + "반" : "",
      revision: Number(item.revision) || 0,
    });
  });
}
