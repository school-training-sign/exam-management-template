function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("정기고사 관리")
    .addItem("초기 설정", "initializeExamSystem")
    .addItem("관리자 암호 변경", "changeAdminPassword")
    .addSeparator()
    .addItem("웹 앱 배포 안내", "showDeploymentGuide")
    .addToUi();
}

function initializeExamSystem() {
  const ui = SpreadsheetApp.getUi();
  ensureSheets_();
  const schoolNamePrompt = ui.prompt("초기 설정", "학교명을 입력하세요.", ui.ButtonSet.OK_CANCEL);
  if (schoolNamePrompt.getSelectedButton() !== ui.Button.OK) return;
  const schoolCodePrompt = ui.prompt("초기 설정", "학교코드를 입력하세요. 원문은 저장되지 않습니다.", ui.ButtonSet.OK_CANCEL);
  if (schoolCodePrompt.getSelectedButton() !== ui.Button.OK) return;
  const passwordPrompt = ui.prompt("초기 설정", "관리자 암호를 입력하세요. 10자 이상을 권장합니다.", ui.ButtonSet.OK_CANCEL);
  if (passwordPrompt.getSelectedButton() !== ui.Button.OK) return;

  const schoolName = String(schoolNamePrompt.getResponseText() || "").trim();
  const schoolCode = String(schoolCodePrompt.getResponseText() || "").trim();
  const password = String(passwordPrompt.getResponseText() || "");
  if (!schoolName || !schoolCode || password.length < 8) {
    ui.alert("학교명·학교코드를 입력하고 관리자 암호는 8자 이상으로 설정하세요.");
    return;
  }

  const properties = PropertiesService.getScriptProperties();
  const schoolSalt = randomToken_();
  const passwordSalt = randomToken_();
  properties.setProperties({
    [APP.PROPERTY_SHEET_ID]: SpreadsheetApp.getActiveSpreadsheet().getId(),
    [APP.PROPERTY_SCHOOL_SALT]: schoolSalt,
    [APP.PROPERTY_SCHOOL_HASH]: hashSecret_(schoolCode, schoolSalt),
    [APP.PROPERTY_ADMIN_SALT]: passwordSalt,
    [APP.PROPERTY_ADMIN_HASH]: hashSecret_(password, passwordSalt),
  }, false);

  setSetting_("school_name", schoolName);
  setSetting_("app_name", "정기고사 관리 시스템");
  setSetting_("timezone", APP.TIMEZONE);
  setSetting_("version", APP.VERSION);
  appendAudit_("admin", "initialize", "system", "", 1, true, "초기 설정 완료");
  ui.alert("초기 설정을 완료했습니다. 이제 Apps Script를 웹 앱으로 배포하세요.");
}

function changeAdminPassword() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty(APP.PROPERTY_ADMIN_HASH)) {
    ui.alert("초기 설정을 먼저 실행하세요.");
    return;
  }
  const prompt = ui.prompt("관리자 암호 변경", "새 관리자 암호를 입력하세요.", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  const password = String(prompt.getResponseText() || "");
  if (password.length < 8) {
    ui.alert("관리자 암호는 8자 이상으로 설정하세요.");
    return;
  }
  const salt = randomToken_();
  properties.setProperty(APP.PROPERTY_ADMIN_SALT, salt);
  properties.setProperty(APP.PROPERTY_ADMIN_HASH, hashSecret_(password, salt));
  CacheService.getScriptCache().removeAll(["admin-session"]);
  appendAudit_("admin", "change_password", "system", "", 1, true, "관리자 암호 변경");
  ui.alert("관리자 암호를 변경했습니다. 기존 브라우저 세션은 만료됩니다.");
}

function showDeploymentGuide() {
  SpreadsheetApp.getUi().alert(
    "1. 확장 프로그램 → Apps Script를 엽니다.\n" +
    "2. 배포 → 새 배포 → 웹 앱을 선택합니다.\n" +
    "3. 실행 사용자는 본인, 액세스 권한은 모든 사용자로 설정합니다.\n" +
    "4. 배포 URL을 프런트엔드 public/config.js의 API_URL에 입력합니다."
  );
}

function doGet() {
  return jsonOutput_({ ok: true, data: { service: "exam-management", version: APP.VERSION }, error: null });
}

function doPost(event) {
  let request = {};
  try {
    request = JSON.parse(event && event.postData ? event.postData.contents : "{}");
    const data = dispatch_(request);
    return jsonOutput_({ ok: true, data: data, error: null });
  } catch (error) {
    try {
      appendAudit_("unknown", request.action || "unknown", "api", "", 0, false, error.message || String(error));
    } catch (ignored) {}
    return jsonOutput_({
      ok: false,
      data: null,
      error: { code: error.code || "REQUEST_FAILED", message: error.message || "요청을 처리하지 못했습니다." },
    });
  }
}

function dispatch_(request) {
  const action = String(request.action || "");
  if (!action) throw apiError_("INVALID_ACTION", "작업 이름이 없습니다.");
  if (action === "school_login") return schoolLogin_(request);
  if (action === "logout") return logout_(request);

  const teacher = requireSession_(request.school_session, "teacher");
  if (action === "admin_login") return adminLogin_(request, teacher);
  if (ADMIN_ACTIONS.has(action)) requireSession_(request.admin_session, "admin");

  if (MUTATION_ACTIONS.has(action)) {
    return withMutation_(request, ADMIN_ACTIONS.has(action) ? "admin" : "teacher", () => route_(action, request));
  }
  return route_(action, request);
}

function route_(action, request) {
  switch (action) {
    case "get_bootstrap": return teacherBootstrap_();
    case "get_admin_bootstrap": return adminBootstrap_();
    case "get_absence_context": return getAbsenceContext_(request);
    case "set_absence": return setAbsence_(request);
    case "complete_input": return completeInput_(request);
    case "get_hq_status": return getHqStatus_(request);
    case "get_period_report": return getPeriodReport_(request);
    case "save_exam_dates": return saveExamDates_(request);
    case "save_class": return saveClass_(request);
    case "delete_class": return deleteClass_(request);
    case "replace_students": return replaceStudents_(request);
    case "import_students": return importStudents_(request);
    case "save_timetable": return saveTimetable_(request);
    case "save_enrollments": return saveEnrollments_(request);
    case "save_seat_chart": return saveSeatChart_(request);
    case "save_seat_charts_batch": return saveSeatChartsBatch_(request);
    case "delete_seat_chart": return deleteSeatChart_(request);
    case "cleanup": return cleanup_(request);
    default: throw apiError_("INVALID_ACTION", "지원하지 않는 작업입니다: " + action);
  }
}

function schoolLogin_(request) {
  const properties = PropertiesService.getScriptProperties();
  const salt = properties.getProperty(APP.PROPERTY_SCHOOL_SALT);
  const expected = properties.getProperty(APP.PROPERTY_SCHOOL_HASH);
  if (!salt || !expected) throw apiError_("NOT_INITIALIZED", "학교 시트에서 초기 설정을 먼저 실행하세요.");
  const actual = hashSecret_(String(request.school_code || "").trim(), salt);
  if (!constantTimeEqual_(actual, expected)) throw apiError_("INVALID_SCHOOL_CODE", "학교코드가 올바르지 않습니다.");
  const token = createSession_("teacher");
  appendAudit_("teacher", "school_login", "session", "", 1, true, "교사 세션 발급");
  return { session: token, bootstrap: teacherBootstrap_() };
}

function adminLogin_(request) {
  const properties = PropertiesService.getScriptProperties();
  const salt = properties.getProperty(APP.PROPERTY_ADMIN_SALT);
  const expected = properties.getProperty(APP.PROPERTY_ADMIN_HASH);
  const actual = hashSecret_(String(request.password || ""), salt || "");
  if (!salt || !expected || !constantTimeEqual_(actual, expected)) {
    throw apiError_("INVALID_ADMIN_PASSWORD", "관리자 암호가 올바르지 않습니다.");
  }
  const token = createSession_("admin");
  appendAudit_("admin", "admin_login", "session", "", 1, true, "관리자 세션 발급");
  return { admin_session: token, bootstrap: adminBootstrap_() };
}

function logout_(request) {
  const cache = CacheService.getScriptCache();
  if (request.school_session) cache.remove("session:" + request.school_session);
  if (request.admin_session) cache.remove("session:" + request.admin_session);
  return { logged_out: true };
}

function getAbsenceContext_(request) {
  const students = getRecords_(SHEETS.STUDENTS)
    .filter((item) => String(item.class_id) === String(request.class_id))
    .sort((left, right) => Number(left.number) - Number(right.number));
  const absences = getRecords_(SHEETS.ABSENCES).filter((item) =>
    item.exam_date === request.exam_date &&
    Number(item.period) === Number(request.period) &&
    String(item.class_id) === String(request.class_id)
  );
  const completed = getRecords_(SHEETS.COMPLETIONS).some((item) =>
    item.exam_date === request.exam_date &&
    Number(item.period) === Number(request.period) &&
    String(item.class_id) === String(request.class_id)
  );
  return { students: students, absences: absences, completed: completed };
}

function setAbsence_(request) {
  validateRequired_(request, ["exam_date", "period", "class_id", "student_id"]);
  const key = request.exam_date + "|" + Number(request.period) + "|" + request.student_id;
  const now = nowIso_();
  let rows = getRecords_(SHEETS.ABSENCES).filter((item) => String(item.key) !== key);
  if (request.absent) {
    rows.push({
      key: key,
      id: createId_("absence"),
      exam_date: request.exam_date,
      period: Number(request.period),
      class_id: request.class_id,
      student_id: request.student_id,
      reason: String(request.reason || "미인정").slice(0, 20),
      reason_detail: String(request.reason_detail || "").slice(0, 100),
      updated_at: now,
    });
  }
  replaceRecords_(SHEETS.ABSENCES, rows);
  let completions = getRecords_(SHEETS.COMPLETIONS);
  const completionKey = request.exam_date + "|" + Number(request.period) + "|" + request.class_id;
  completions = completions.filter((item) => String(item.key) !== completionKey);
  replaceRecords_(SHEETS.COMPLETIONS, completions);
  return { saved: true };
}

function completeInput_(request) {
  validateRequired_(request, ["exam_date", "period", "class_id"]);
  const key = request.exam_date + "|" + Number(request.period) + "|" + request.class_id;
  let rows = getRecords_(SHEETS.COMPLETIONS).filter((item) => String(item.key) !== key);
  rows.push({ key: key, id: createId_("completion"), exam_date: request.exam_date, period: Number(request.period), class_id: request.class_id, completed_at: nowIso_(), updated_at: nowIso_() });
  replaceRecords_(SHEETS.COMPLETIONS, rows);
  return { completed: true };
}

function getHqStatus_(request) {
  let absences = decorateAbsences_(getRecords_(SHEETS.ABSENCES));
  absences = absences.filter((item) =>
    (!request.exam_date || item.exam_date === request.exam_date) &&
    (!request.period || Number(item.period) === Number(request.period)) &&
    (!request.grade || Number(item.grade) === Number(request.grade)) &&
    (!request.class_id || String(item.class_id) === String(request.class_id))
  );
  return { absences: absences, completions: getRecords_(SHEETS.COMPLETIONS) };
}

function getPeriodReport_(request) {
  let absences = decorateAbsences_(getRecords_(SHEETS.ABSENCES));
  absences = absences.filter((item) =>
    (!request.start_date || item.exam_date >= request.start_date) &&
    (!request.end_date || item.exam_date <= request.end_date) &&
    (!request.grade || Number(item.grade) === Number(request.grade)) &&
    (!request.class_id || String(item.class_id) === String(request.class_id))
  );
  return { absences: absences };
}

function saveExamDates_(request) {
  const now = nowIso_();
  const rows = (request.exam_dates || []).map((item) => ({
    id: item.id || createId_("date"), exam_date: item.exam_date, label: String(item.label || "").slice(0, 50), active: item.active !== false, updated_at: now,
  })).filter((item) => /^\d{4}-\d{2}-\d{2}$/.test(item.exam_date));
  replaceRecords_(SHEETS.EXAM_DATES, rows);
  return adminBootstrap_();
}

function saveClass_(request) {
  const incoming = request.class_item || {};
  if (!Number.isInteger(Number(incoming.grade)) || !Number.isInteger(Number(incoming.class_num))) throw apiError_("INVALID_CLASS", "학년과 반은 정수여야 합니다.");
  const now = nowIso_();
  let rows = getRecords_(SHEETS.CLASSES);
  if (incoming.id) rows = rows.filter((item) => String(item.id) !== String(incoming.id));
  else rows = rows.filter((item) => !(Number(item.grade) === Number(incoming.grade) && Number(item.class_num) === Number(incoming.class_num)));
  rows.push({ id: incoming.id || createId_("class"), grade: Number(incoming.grade), class_num: Number(incoming.class_num), active: incoming.active !== false, updated_at: now });
  rows.sort((a, b) => Number(a.grade) - Number(b.grade) || Number(a.class_num) - Number(b.class_num));
  replaceRecords_(SHEETS.CLASSES, rows);
  return adminBootstrap_();
}

function deleteClass_(request) {
  validateRequired_(request, ["class_id"]);
  replaceRecords_(SHEETS.CLASSES, getRecords_(SHEETS.CLASSES).filter((item) => String(item.id) !== String(request.class_id)));
  replaceRecords_(SHEETS.STUDENTS, getRecords_(SHEETS.STUDENTS).filter((item) => String(item.class_id) !== String(request.class_id)));
  replaceRecords_(SHEETS.ENROLLMENTS, getRecords_(SHEETS.ENROLLMENTS).filter((item) => String(item.class_id) !== String(request.class_id)));
  return adminBootstrap_();
}

function replaceStudents_(request) {
  validateRequired_(request, ["class_id"]);
  const now = nowIso_();
  let rows = getRecords_(SHEETS.STUDENTS).filter((item) => String(item.class_id) !== String(request.class_id));
  const seen = new Set();
  (request.students || []).forEach((item) => {
    const number = Number(item.number);
    const name = String(item.name || "").trim().slice(0, 50);
    if (!Number.isInteger(number) || number < 1 || !name || seen.has(number)) return;
    seen.add(number);
    rows.push({ id: item.id || createId_("student"), class_id: request.class_id, number: number, name: name, active: true, updated_at: now });
  });
  replaceRecords_(SHEETS.STUDENTS, rows);
  return adminBootstrap_();
}

function importStudents_(request) {
  const now = nowIso_();
  let classes = getRecords_(SHEETS.CLASSES);
  let students = getRecords_(SHEETS.STUDENTS);
  (request.rows || []).forEach((row) => {
    const grade = Number(row.grade);
    const classNum = Number(row.class_num);
    const number = Number(row.number);
    const name = String(row.name || "").trim().slice(0, 50);
    if (![grade, classNum, number].every(Number.isInteger) || !name) return;
    let classInfo = classes.find((item) => Number(item.grade) === grade && Number(item.class_num) === classNum);
    if (!classInfo) {
      classInfo = { id: createId_("class"), grade: grade, class_num: classNum, active: true, updated_at: now };
      classes.push(classInfo);
    }
    students = students.filter((item) => !(String(item.class_id) === String(classInfo.id) && Number(item.number) === number));
    students.push({ id: createId_("student"), class_id: classInfo.id, number: number, name: name, active: true, updated_at: now });
  });
  classes.sort((a, b) => Number(a.grade) - Number(b.grade) || Number(a.class_num) - Number(b.class_num));
  replaceRecords_(SHEETS.CLASSES, classes);
  replaceRecords_(SHEETS.STUDENTS, students);
  return adminBootstrap_();
}

function saveTimetable_(request) {
  const now = nowIso_();
  const rows = (request.timetable || []).map((item) => Object.assign({}, item, {
    id: item.id || createId_("timetable"), grade: Number(item.grade), period: Number(item.period), subject_type: item.subject_type === "elective" ? "elective" : "common", updated_at: now,
  })).filter((item) => item.exam_date && item.subject_name && Number.isInteger(item.grade) && Number.isInteger(item.period));
  replaceRecords_(SHEETS.TIMETABLE, rows);
  return adminBootstrap_();
}

function saveEnrollments_(request) {
  const grade = Number(request.grade);
  let rows = getRecords_(SHEETS.ENROLLMENTS).filter((item) => Number(item.grade) !== grade);
  const now = nowIso_();
  rows = rows.concat((request.enrollments || []).map((item) => ({
    id: item.id || createId_("enrollment"), grade: grade, class_id: item.class_id, student_id: item.student_id, subject_name: String(item.subject_name || "").slice(0, 80), room_name: String(item.room_name || "").slice(0, 80), updated_at: now,
  })).filter((item) => item.class_id && item.student_id && item.subject_name));
  replaceRecords_(SHEETS.ENROLLMENTS, rows);
  return adminBootstrap_();
}

function saveSeatChart_(request) {
  const row = serializeSeat_(request.chart || {});
  let rows = getRecords_(SHEETS.SEATING).filter((item) => String(item.id) !== String(row.id));
  rows.push(row);
  replaceRecords_(SHEETS.SEATING, rows);
  return { chart: deserializeSeat_(row), seat_charts: rows.map(deserializeSeat_).reverse() };
}

function saveSeatChartsBatch_(request) {
  let rows = getRecords_(SHEETS.SEATING);
  (request.charts || []).forEach((chart) => rows.push(serializeSeat_(chart)));
  replaceRecords_(SHEETS.SEATING, rows);
  return { seat_charts: rows.map(deserializeSeat_).reverse() };
}

function deleteSeatChart_(request) {
  const rows = getRecords_(SHEETS.SEATING).filter((item) => String(item.id) !== String(request.id));
  replaceRecords_(SHEETS.SEATING, rows);
  return { seat_charts: rows.map(deserializeSeat_).reverse() };
}

function cleanup_(request) {
  const cutoff = String(request.cutoff_date || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(cutoff)) throw apiError_("INVALID_DATE", "정리 기준일이 올바르지 않습니다.");
  [SHEETS.EXAM_DATES, SHEETS.TIMETABLE, SHEETS.ABSENCES, SHEETS.COMPLETIONS, SHEETS.SEATING].forEach((name) => {
    replaceRecords_(name, getRecords_(name).filter((item) => !item.exam_date || item.exam_date >= cutoff));
  });
  return adminBootstrap_();
}

function withMutation_(request, role, operation) {
  const requestId = String(request.request_id || "");
  const cache = CacheService.getScriptCache();
  const idempotencyKey = requestId ? "request:" + requestId : "";
  if (idempotencyKey) {
    const cached = cache.get(idempotencyKey);
    if (cached) return JSON.parse(cached);
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(28000)) throw apiError_("BUSY", "동시 저장 요청이 많습니다. 잠시 후 자동으로 다시 시도합니다.");
  try {
    if (idempotencyKey) {
      const cachedAfterLock = cache.get(idempotencyKey);
      if (cachedAfterLock) return JSON.parse(cachedAfterLock);
    }
    const result = operation();
    if (idempotencyKey) cache.put(idempotencyKey, JSON.stringify(result), APP.IDEMPOTENCY_TTL_SECONDS);
    appendAudit_(role, request.action, "records", "", inferCount_(request), true, "완료");
    return result;
  } catch (error) {
    appendAudit_(role, request.action, "records", "", 0, false, error.message || String(error));
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function inferCount_(request) {
  if (Array.isArray(request.rows)) return request.rows.length;
  if (Array.isArray(request.students)) return request.students.length;
  if (Array.isArray(request.timetable)) return request.timetable.length;
  if (Array.isArray(request.enrollments)) return request.enrollments.length;
  if (Array.isArray(request.charts)) return request.charts.length;
  return 1;
}

function createSession_(role) {
  const token = randomToken_() + randomToken_();
  const ttl = role === "admin" ? APP.ADMIN_TTL_SECONDS : APP.TEACHER_TTL_SECONDS;
  CacheService.getScriptCache().put("session:" + token, JSON.stringify({ role: role, created_at: nowIso_() }), ttl);
  return token;
}

function requireSession_(token, role) {
  if (!token) throw apiError_("SESSION_REQUIRED", "로그인이 필요합니다.");
  const cache = CacheService.getScriptCache();
  const key = "session:" + token;
  const raw = cache.get(key);
  if (!raw) throw apiError_("SESSION_EXPIRED", "세션이 만료되었습니다. 다시 로그인하세요.");
  const session = JSON.parse(raw);
  if (session.role !== role) throw apiError_("FORBIDDEN", "이 작업에 필요한 권한이 없습니다.");
  const ttl = role === "admin" ? APP.ADMIN_TTL_SECONDS : APP.TEACHER_TTL_SECONDS;
  cache.put(key, raw, ttl);
  return session;
}

function hashSecret_(secret, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + "\u001f" + String(secret), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, "");
}

function constantTimeEqual_(left, right) {
  left = String(left || "");
  right = String(right || "");
  let mismatch = left.length ^ right.length;
  const length = Math.max(left.length, right.length);
  for (let index = 0; index < length; index += 1) mismatch |= (left.charCodeAt(index) || 0) ^ (right.charCodeAt(index) || 0);
  return mismatch === 0;
}

function randomToken_() {
  return Utilities.getUuid().replace(/-/g, "");
}

function createId_(prefix) {
  return prefix + "_" + Utilities.getUuid();
}

function nowIso_() {
  return Utilities.formatDate(new Date(), APP.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function validateRequired_(object, keys) {
  keys.forEach((key) => {
    if (object[key] === "" || object[key] == null) throw apiError_("INVALID_INPUT", key + " 값이 필요합니다.");
  });
}

function apiError_(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
