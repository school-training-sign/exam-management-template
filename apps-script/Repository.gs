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
  const guideText = guide.getLastRow() ? guide.getDataRange().getDisplayValues().flat().join(" ") : "";
  if (!guideText.includes(APP.SYSTEM_NAME) || !guideText.includes(APP.VERSION) || guideText.includes(["학교", "코드"].join("")) || guideText.includes("PIN 6자리")) {
    const oldGuideRange = guide.getDataRange();
    oldGuideRange.breakApart();
    oldGuideRange.clearContent();
    guide.getRange("A1:B8").setValues([
      [APP.SYSTEM_NAME, "학교 운영 시트"],
      ["1", "상단 메뉴 '정기고사 관리 > 초기 설정'을 실행합니다."],
      ["2", "학교명과 4자 이상의 관리자 암호를 입력합니다."],
      ["3", "Apps Script를 웹 앱으로 배포합니다. 실행 사용자는 본인, 접근 권한은 조직 정책이 허용하는 범위로 설정합니다."],
      ["4", "웹 앱 주소를 GitHub Pages 공개 설정의 API_URL에 입력합니다."],
      ["접속", "웹 화면의 관리자 설정으로 들어가 접속 이름과 숫자 4~6자리 개인 PIN을 등록합니다."],
      ["데이터", "학생 이름은 학생 시트에만 저장되며 감사로그에는 남기지 않습니다."],
      ["버전", APP.VERSION],
    ]);
    guide.getRange("A1:B1").merge().setBackground("#0d2343").setFontColor("#ffffff").setFontWeight("bold");
    guide.setColumnWidth(1, 100);
    guide.setColumnWidth(2, 620);
    guide.getRange("A1:B8").setWrap(true).setVerticalAlignment("middle");
  }

  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty(APP.PROPERTY_ACCESS_USER_PEPPER)) {
    properties.setProperty(APP.PROPERTY_ACCESS_USER_PEPPER, randomToken_() + randomToken_());
  }
  if (!properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_PEPPER)) {
    properties.setProperty(APP.PROPERTY_PACKAGING_INVITE_PEPPER, randomToken_() + randomToken_());
  }
  if (!properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_VERSION)) {
    properties.setProperty(APP.PROPERTY_PACKAGING_INVITE_VERSION, "1");
  }
  properties.deleteProperty(["SCHOOL", "CODE", "SALT"].join("_"));
  properties.deleteProperty(["SCHOOL", "CODE", "HASH"].join("_"));

  const settings = getRecords_(SHEETS.SETTINGS);
  const schema = settings.find((item) => String(item.key) === "schema_version");
  if (!schema || String(schema.value) !== APP.SCHEMA_VERSION) {
    setSetting_("schema_version", APP.SCHEMA_VERSION);
  }
  const version = settings.find((item) => String(item.key) === "version");
  if (!version || String(version.value) !== APP.VERSION) {
    setSetting_("version", APP.VERSION);
  }
  const appName = settings.find((item) => String(item.key) === "app_name");
  if (!appName || String(appName.value) !== APP.SYSTEM_NAME) {
    setSetting_("app_name", APP.SYSTEM_NAME);
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

function appendAudit_(role, action, targetType, targetId, count, success, detail, actorId) {
  const sheet = getSpreadsheet_().getSheetByName(SHEETS.AUDIT);
  const width = Math.max(sheet.getLastColumn(), HEADERS[SHEETS.AUDIT].length);
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0].map((value) => String(value || "").trim());
  const record = {
    id: createId_("audit"),
    occurred_at: nowIso_(),
    role: String(role || "").slice(0, 20),
    actor_id: String(actorId || "").slice(0, 100),
    action: String(action || "").slice(0, 60),
    target_type: String(targetType || "").slice(0, 30),
    target_id: String(targetId || "").slice(0, 100),
    count: Number(count) || 0,
    success: Boolean(success),
    detail: sanitizeAuditDetail_(Boolean(success), detail),
  };
  sheet.appendRow(headers.map((header) => Object.prototype.hasOwnProperty.call(record, header) ? record[header] : ""));
}

function getSettings_() {
  const result = {};
  getRecords_(SHEETS.SETTINGS).forEach((row) => { result[String(row.key)] = row.value; });
  return result;
}

function getPublicSettings_() {
  const settings = getSettings_();
  return ["school_name", "app_name", "timezone", "version", "schema_version"].reduce((result, key) => {
    if (Object.prototype.hasOwnProperty.call(settings, key)) result[key] = settings[key];
    return result;
  }, {});
}

function setSetting_(key, value) {
  const now = nowIso_();
  const rows = getRecordsFresh_(SHEETS.SETTINGS).filter((item) => String(item.key) !== String(key));
  rows.push({ key: key, value: value, updated_at: now });
  replaceRecords_(SHEETS.SETTINGS, rows);
}

function accessUserSettingKey_(id) {
  return APP.ACCESS_USER_SETTING_PREFIX + String(id || "").trim();
}

function normalizeStoredAccessUser_(row) {
  if (!row || !String(row.key || "").startsWith(APP.ACCESS_USER_SETTING_PREFIX)) return null;
  const parsed = parseJson_(row.value, null);
  if (!parsed || Array.isArray(parsed) || typeof parsed !== "object") return null;
  const id = String(row.key).slice(APP.ACCESS_USER_SETTING_PREFIX.length).trim();
  const loginName = canonicalLoginName_(parsed.login_name);
  if (!id || !loginName || !parsed.pin_salt || !parsed.pin_hash) return null;
  return {
    id: id,
    login_name: loginName,
    normalized_login_name: loginLookupKey_(loginName),
    active: parsed.active === true,
    session_version: Math.max(1, Number(parsed.session_version) || 1),
    pin_salt: String(parsed.pin_salt),
    pin_hash: String(parsed.pin_hash),
    created_at: String(parsed.created_at || row.updated_at || ""),
    updated_at: String(parsed.updated_at || row.updated_at || ""),
  };
}

function getAccessUsers_() {
  return getRecords_(SHEETS.SETTINGS)
    .map(normalizeStoredAccessUser_)
    .filter(Boolean);
}

function findAccessUserById_(id) {
  const wanted = String(id || "").trim();
  return getAccessUsers_().find((user) => user.id === wanted) || null;
}

function findAccessUserByLogin_(loginName) {
  const lookup = loginLookupKey_(loginName);
  return getAccessUsers_().find((user) => user.normalized_login_name === lookup) || null;
}

function safeAccessUser_(user) {
  return {
    id: String(user.id),
    login_name: String(user.login_name),
    active: user.active === true,
    updated_at: String(user.updated_at || ""),
  };
}

function accessUsersForAdmin_() {
  return getAccessUsers_()
    .map(safeAccessUser_)
    .sort((left, right) => left.login_name.localeCompare(right.login_name, "ko"));
}

function writeAccessUser_(user) {
  const stored = {
    id: String(user.id),
    login_name: canonicalLoginName_(user.login_name),
    normalized_login_name: loginLookupKey_(user.login_name),
    active: user.active === true,
    session_version: Math.max(1, Number(user.session_version) || 1),
    pin_salt: String(user.pin_salt),
    pin_hash: String(user.pin_hash),
    created_at: String(user.created_at || nowIso_()),
    updated_at: String(user.updated_at || nowIso_()),
  };
  setSetting_(accessUserSettingKey_(stored.id), JSON.stringify(stored));
  return stored;
}

function parseJson_(value, fallback) {
  if (value == null || value === "") return fallback;
  if (Array.isArray(value) || (typeof value === "object" && value !== null)) return value;
  try { return JSON.parse(String(value)); } catch (error) { return fallback; }
}

function isValidIsoDate_(value) {
  const text = String(value || "").trim();
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!match) return false;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function normalizeStoredSubjectCatalog_(value) {
  const parsed = parseJson_(value, {});
  if (!parsed || Array.isArray(parsed) || typeof parsed !== "object") return {};
  const result = {};
  Object.keys(parsed).forEach((key) => {
    const grade = Number(key);
    if (!Number.isInteger(grade) || grade < 1 || grade > 9 || !Array.isArray(parsed[key])) return;
    const seen = new Set();
    const subjects = [];
    parsed[key].forEach((item) => {
      const subject = typeof item === "string" ? item.trim() : "";
      if (!subject || subject.length > 80 || seen.has(subject) || subjects.length >= 200) return;
      seen.add(subject);
      subjects.push(subject);
    });
    result[String(grade)] = subjects;
  });
  return result;
}

function normalizeStoredExamNotice_(value, schoolName) {
  const parsedValue = parseJson_(value, {});
  const parsed = parsedValue && !Array.isArray(parsedValue) && typeof parsedValue === "object" ? parsedValue : {};
  const academicYear = Number(parsed.year || parsed.academic_year);
  const semester = Number(parsed.semester);
  const notes = Array.isArray(parsed.notes)
    ? parsed.notes.map((item) => typeof item === "string" ? item.trim() : "")
      .filter((item) => item && item.length <= 200)
      .slice(0, 20)
    : [];
  const defaultIssuer = String(schoolName || "").trim();
  return {
    year: Number.isInteger(academicYear) && academicYear >= 1900 && academicYear <= 2200 ? academicYear : "",
    semester: semester === 1 || semester === 2 ? semester : "",
    exam_name: String(parsed.exam_name || "").trim().slice(0, 80),
    greeting: String(parsed.greeting || "").trim().slice(0, 1500),
    notes: notes,
    issue_date: isValidIsoDate_(parsed.issue_date)
      ? String(parsed.issue_date).trim()
      : "",
    issuer: String(parsed.issuer || "").trim().slice(0, 100) || (defaultIssuer ? defaultIssuer + "장" : ""),
  };
}

function isValidClockTime_(value) {
  const match = /^(\d{2}):(\d{2})$/.exec(String(value || "").trim());
  return Boolean(match && Number(match[1]) >= 0 && Number(match[1]) <= 23 && Number(match[2]) >= 0 && Number(match[2]) <= 59);
}

function isValidDateTime_(value) {
  const text = String(value || "").trim();
  if (!text || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d{1,3})?)?(?:Z|[+-]\d{2}:?\d{2})?$/.test(text)) return false;
  return Number.isFinite(new Date(text).getTime());
}

function defaultExamPackagingRows_() {
  return [
    ["slot_0830", "slot", "1회", "08:30", "08:55", "", 10],
    ["slot_0900", "slot", "2회", "09:00", "09:25", "", 20],
    ["slot_0930", "slot", "3회", "09:30", "09:55", "", 30],
    ["slot_1000", "slot", "4회", "10:00", "10:25", "", 40],
    ["slot_1030", "slot", "5회", "10:30", "10:55", "", 50],
    ["slot_1100", "slot", "6회", "11:00", "11:25", "", 60],
    ["slot_1130", "slot", "7회", "11:30", "11:55", "", 70],
    ["break_lunch", "break", "점심시간", "12:00", "13:00", "과목 배정 불가", 80],
    ["slot_1300", "slot", "8회", "13:00", "13:25", "", 90],
    ["slot_1330", "slot", "9회", "13:30", "13:55", "", 100],
    ["slot_1400", "slot", "10회", "14:00", "14:25", "", 110],
    ["slot_1430", "slot", "11회", "14:30", "14:55", "", 120],
    ["slot_1500", "slot", "12회", "15:00", "15:25", "", 130],
    ["slot_1530", "slot", "13회", "15:30", "15:55", "", 140],
    ["slot_1600", "slot", "14회", "16:00", "16:25", "", 150],
  ].map((row) => ({
    id: row[0], kind: row[1], period_label: row[2], start_time: row[3], end_time: row[4], note: row[5], sort_order: row[6],
  }));
}

function defaultExamPackagingConfig_() {
  return {
    revision: 0,
    input_deadline: "",
    capacity: 3,
    packaging_dates: [],
    rows: defaultExamPackagingRows_(),
    staff_assignments: [],
  };
}

function normalizeStoredExamPackagingConfig_(value) {
  const fallback = defaultExamPackagingConfig_();
  const parsedValue = parseJson_(value, {});
  const parsed = parsedValue && !Array.isArray(parsedValue) && typeof parsedValue === "object" ? parsedValue : {};
  const dates = Array.isArray(parsed.packaging_dates)
    ? parsed.packaging_dates.map((item) => String(item || "").trim()).filter(isValidIsoDate_)
      .filter((item, index, all) => all.indexOf(item) === index).slice(0, 31)
    : [];
  const rowIds = new Set();
  const rows = (Array.isArray(parsed.rows) ? parsed.rows : fallback.rows).map((item, index) => {
    const id = String(item && item.id || "").trim().slice(0, 80);
    const kind = item && item.kind === "break" ? "break" : "slot";
    const startTime = String(item && item.start_time || "").trim();
    const endTime = String(item && item.end_time || "").trim();
    if (!id || rowIds.has(id) || !isValidClockTime_(startTime) || !isValidClockTime_(endTime) || startTime >= endTime) return null;
    rowIds.add(id);
    return {
      id: id,
      kind: kind,
      period_label: String(item.period_label || "").trim().slice(0, 40),
      start_time: startTime,
      end_time: endTime,
      note: String(item.note || "").trim().slice(0, 100),
      sort_order: Number.isFinite(Number(item.sort_order)) ? Number(item.sort_order) : (index + 1) * 10,
    };
  }).filter(Boolean).slice(0, 60).sort((left, right) => left.sort_order - right.sort_order || left.start_time.localeCompare(right.start_time));
  const staffSeen = new Set();
  const staffAssignments = (Array.isArray(parsed.staff_assignments) ? parsed.staff_assignments : []).map((item) => {
    const packagingDate = String(item && item.packaging_date || "").trim();
    const slotId = String(item && item.slot_id || "").trim();
    const key = packagingDate + "|" + slotId;
    if (dates.indexOf(packagingDate) === -1 || !rowIds.has(slotId) || staffSeen.has(key)) return null;
    staffSeen.add(key);
    return { packaging_date: packagingDate, slot_id: slotId, staff_name: String(item.staff_name || "").trim().slice(0, 100) };
  }).filter(Boolean);
  return {
    revision: Number.isInteger(Number(parsed.revision)) && Number(parsed.revision) >= 0 ? Number(parsed.revision) : 0,
    input_deadline: isValidDateTime_(parsed.input_deadline) ? String(parsed.input_deadline).trim() : "",
    capacity: Number.isInteger(Number(parsed.capacity)) && Number(parsed.capacity) >= 1 && Number(parsed.capacity) <= 20 ? Number(parsed.capacity) : 3,
    packaging_dates: dates,
    rows: rows,
    staff_assignments: staffAssignments,
  };
}

function getExamPackagingConfig_(fresh) {
  const settings = (fresh ? getRecordsFresh_(SHEETS.SETTINGS) : getRecords_(SHEETS.SETTINGS));
  const row = settings.find((item) => String(item.key) === "exam_packaging_config_json");
  return normalizeStoredExamPackagingConfig_(row ? row.value : null);
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
    settings: getPublicSettings_(),
    classes: getRecords_(SHEETS.CLASSES).filter(isActive_),
    exam_dates: getRecords_(SHEETS.EXAM_DATES).filter(isActive_),
  };
}

function adminBootstrap_() {
  const result = teacherBootstrap_();
  const privateSettings = getSettings_();
  result.subject_catalog = normalizeStoredSubjectCatalog_(privateSettings.subject_catalog_json);
  result.exam_notice = normalizeStoredExamNotice_(privateSettings.exam_notice_json, result.settings.school_name);
  result.exam_packaging_config = getExamPackagingConfig_();
  result.exam_packaging_invite_status = examPackagingInviteStatus_();
  result.access_users = accessUsersForAdmin_();
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
