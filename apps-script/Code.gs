function onOpen() {
  ensureSheets_();
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
    ui.alert("학교명과 학교코드를 입력하고 관리자 암호를 8자 이상으로 설정하세요.");
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
    [APP.PROPERTY_ADMIN_SESSION_VERSION]: "1",
  }, false);

  setSetting_("school_name", schoolName);
  setSetting_("app_name", "정기고사 관리 시스템");
  setSetting_("timezone", APP.TIMEZONE);
  setSetting_("version", APP.VERSION);
  setSetting_("schema_version", APP.SCHEMA_VERSION);
  safeAppendAudit_("admin", "initialize", "system", "", 1, true, "OK");
  ui.alert("초기 설정을 마쳤습니다. 이제 Apps Script를 웹 앱으로 배포하세요.");
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
    ui.alert("관리자 암호를 8자 이상으로 설정하세요.");
    return;
  }
  const salt = randomToken_();
  const nextVersion = Number(properties.getProperty(APP.PROPERTY_ADMIN_SESSION_VERSION) || 1) + 1;
  properties.setProperty(APP.PROPERTY_ADMIN_SALT, salt);
  properties.setProperty(APP.PROPERTY_ADMIN_HASH, hashSecret_(password, salt));
  properties.setProperty(APP.PROPERTY_ADMIN_SESSION_VERSION, String(nextVersion));
  safeAppendAudit_("admin", "change_password", "system", "", 1, true, "OK");
  ui.alert("관리자 암호를 바꿨습니다. 기존 관리자 세션은 모두 만료됩니다.");
}

function showDeploymentGuide() {
  SpreadsheetApp.getUi().alert(
    "1. 확장 프로그램 > Apps Script를 엽니다.\n" +
    "2. 배포 > 새 배포 > 웹 앱을 선택합니다.\n" +
    "3. 실행 사용자는 본인, 접근 권한은 조직 정책이 허용하는 범위로 설정합니다.\n" +
    "4. 배포 URL을 GitHub Pages 공개 설정의 API_URL에 입력합니다."
  );
}

function doGet() {
  return jsonOutput_({ ok: true, data: { service: "exam-management", version: APP.VERSION }, error: null });
}

function doPost(event) {
  let request = {};
  try {
    request = JSON.parse(event && event.postData ? event.postData.contents : "{}");
    ensureSchemaOnRequest_();
    const data = dispatch_(request);
    return jsonOutput_({ ok: true, data: data, error: null });
  } catch (error) {
    if (!error.auditLogged) {
      try {
        appendAudit_("unknown", request.action || "unknown", "api", "", 0, false, error.code || "REQUEST_FAILED");
        error.auditLogged = true;
      } catch (ignored) {}
    }
    return jsonOutput_({
      ok: false,
      data: null,
      error: {
        code: error.code || "REQUEST_FAILED",
        message: error.message || "요청을 처리하지 못했습니다.",
        details: error.details || null,
      },
    });
  }
}

function ensureSchemaOnRequest_() {
  const cache = CacheService.getScriptCache();
  const marker = "schema-ready:" + APP.SCHEMA_VERSION;
  if (cache.get(marker)) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw apiError_("BUSY", "초기 데이터 구조를 확인하는 중입니다. 잠시 후 다시 시도하세요.");
  try {
    if (!cache.get(marker)) {
      ensureSheets_();
      cache.put(marker, "1", APP.DATA_CACHE_TTL_SECONDS);
    }
  } finally {
    lock.releaseLock();
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
    case "submit_absence_input": return submitAbsenceInput_(request);
    case "get_hq_status": return getHqStatus_(request);
    case "get_period_report": return getPeriodReport_(request);
    case "save_exam_dates": return saveExamDates_(request);
    case "delete_exam_date": return deleteExamDate_(request);
    case "save_class": return saveClass_(request);
    case "delete_class": return deleteClass_(request);
    case "replace_students": return replaceStudents_(request);
    case "import_students": return importStudents_(request);
    case "save_student": return saveStudent_(request);
    case "delete_student": return deleteStudent_(request);
    case "save_timetable": return saveTimetable_(request);
    case "delete_timetable": return deleteTimetable_(request);
    case "save_enrollments": return saveEnrollments_(request);
    case "save_seat_chart": return saveSeatChart_(request);
    case "save_seat_charts_batch": return saveSeatChartsBatch_(request);
    case "delete_seat_chart": return deleteSeatChart_(request);
    case "cleanup": return cleanup_(request);
    default: throw apiError_("INVALID_ACTION", "지원하지 않는 작업입니다: " + action);
  }
}

function schoolLogin_(request) {
  const secret = String(request.school_code || "").trim();
  const rateKeys = loginRateKeys_("school", request);
  assertLoginAllowed_(rateKeys);
  const properties = PropertiesService.getScriptProperties();
  const salt = properties.getProperty(APP.PROPERTY_SCHOOL_SALT);
  const expected = properties.getProperty(APP.PROPERTY_SCHOOL_HASH);
  if (!salt || !expected) throw loggedApiError_("teacher", "school_login", "NOT_INITIALIZED", "학교 시트에서 초기 설정을 먼저 실행하세요.");
  const actual = hashSecret_(secret, salt);
  if (!constantTimeEqual_(actual, expected)) {
    recordLoginFailure_(rateKeys);
    throw loggedApiError_("teacher", "school_login", "INVALID_SCHOOL_CODE", "학교코드가 올바르지 않습니다.");
  }
  clearLoginFailures_(rateKeys);
  const token = createSession_("teacher");
  safeAppendAudit_("teacher", "school_login", "session", "", 1, true, "OK");
  return { session: token, bootstrap: teacherBootstrap_() };
}

function adminLogin_(request) {
  const password = String(request.password || "");
  const rateKeys = loginRateKeys_("admin", request);
  assertLoginAllowed_(rateKeys);
  const properties = PropertiesService.getScriptProperties();
  const salt = properties.getProperty(APP.PROPERTY_ADMIN_SALT);
  const expected = properties.getProperty(APP.PROPERTY_ADMIN_HASH);
  const actual = hashSecret_(password, salt || "");
  if (!salt || !expected || !constantTimeEqual_(actual, expected)) {
    recordLoginFailure_(rateKeys);
    throw loggedApiError_("admin", "admin_login", "INVALID_ADMIN_PASSWORD", "관리자 암호가 올바르지 않습니다.");
  }
  clearLoginFailures_(rateKeys);
  const token = createSession_("admin");
  safeAppendAudit_("admin", "admin_login", "session", "", 1, true, "OK");
  return { admin_session: token, bootstrap: adminBootstrap_() };
}

function logout_(request) {
  const cache = CacheService.getScriptCache();
  if (request.school_session) cache.remove("session:" + request.school_session);
  if (request.admin_session) cache.remove("session:" + request.admin_session);
  return { logged_out: true };
}

function completionKey_(examDate, period, classId) {
  return [examDate, Number(period), classId].join("|");
}

function getAbsenceContext_(request) {
  validateRequired_(request, ["exam_date", "period", "class_id"]);
  validateExamContext_(request.exam_date, request.period, request.class_id);
  const students = getRecords_(SHEETS.STUDENTS)
    .filter((item) => String(item.class_id) === String(request.class_id) && isActive_(item))
    .sort((left, right) => Number(left.number) - Number(right.number));
  const activeStudentIds = new Set(students.map((item) => String(item.id)));
  const allAbsences = getRecords_(SHEETS.ABSENCES).filter((item) =>
    item.exam_date === request.exam_date &&
    Number(item.period) === Number(request.period) &&
    String(item.class_id) === String(request.class_id)
  );
  const absences = allAbsences.filter((item) => activeStudentIds.has(String(item.student_id)));
  const key = completionKey_(request.exam_date, request.period, request.class_id);
  const state = getRecords_(SHEETS.COMPLETIONS).find((item) => String(item.key) === key);
  const submitted = Boolean(state && state.completed_at);
  return {
    students: students,
    absences: absences,
    submitted: submitted,
    completed: submitted,
    revision: state ? Number(state.revision) || 0 : 0,
    completed_at: submitted ? state.completed_at : "",
    archived_absence_count: allAbsences.length - absences.length,
  };
}

function submitAbsenceInput_(request) {
  validateRequired_(request, ["exam_date", "period", "class_id", "expected_revision"]);
  validateExamContext_(request.exam_date, request.period, request.class_id);
  if (!Array.isArray(request.absences)) throw apiError_("INVALID_ABSENCES", "결시 목록 형식이 올바르지 않습니다.");
  const expectedRevision = Number(request.expected_revision);
  if (!Number.isInteger(expectedRevision) || expectedRevision < 0) {
    throw apiError_("INVALID_REVISION", "revision 값이 올바르지 않습니다.");
  }

  const key = completionKey_(request.exam_date, request.period, request.class_id);
  const oldAbsences = getRecords_(SHEETS.ABSENCES);
  const oldCompletions = getRecords_(SHEETS.COMPLETIONS);
  const currentState = oldCompletions.find((item) => String(item.key) === key);
  const currentRevision = currentState ? Number(currentState.revision) || 0 : 0;
  if (currentRevision !== expectedRevision) {
    throw apiError_("REVISION_CONFLICT", "다른 사용자가 같은 화면을 먼저 수정했습니다. 최신 내용을 다시 불러오세요.", {
      current_revision: currentRevision,
      submitted: Boolean(currentState && currentState.completed_at),
    });
  }

  const validStudents = new Map(getRecords_(SHEETS.STUDENTS)
    .filter((item) => String(item.class_id) === String(request.class_id) && isActive_(item))
    .map((item) => [String(item.id), item]));
  const existingByStudent = new Map(oldAbsences.filter((item) =>
    item.exam_date === request.exam_date &&
    Number(item.period) === Number(request.period) &&
    String(item.class_id) === String(request.class_id)
  ).map((item) => [String(item.student_id), item]));
  const seen = new Set();
  const now = nowIso_();
  const incoming = request.absences.map((item) => {
    const studentId = String(item.student_id || "");
    if (!validStudents.has(studentId)) throw apiError_("STUDENT_NOT_FOUND", "선택한 학급에 없는 학생이 포함되어 있습니다.");
    if (seen.has(studentId)) throw apiError_("DUPLICATE_ABSENCE", "같은 학생이 결시 목록에 두 번 들어 있습니다.");
    seen.add(studentId);
    const reason = validateAbsenceReason_(item.reason, item.reason_detail);
    const reasonDetail = reason === "기타"
      ? String(item.reason_detail || "").trim().slice(0, 100)
      : "";
    const existing = existingByStudent.get(studentId);
    return {
      key: [request.exam_date, Number(request.period), request.class_id, studentId].join("|"),
      id: existing ? existing.id : createId_("absence"),
      exam_date: request.exam_date,
      period: Number(request.period),
      class_id: request.class_id,
      student_id: studentId,
      reason: reason,
      reason_detail: reasonDetail,
      updated_at: now,
    };
  });
  const preservedInactive = oldAbsences.filter((item) =>
    item.exam_date === request.exam_date &&
    Number(item.period) === Number(request.period) &&
    String(item.class_id) === String(request.class_id) &&
    !validStudents.has(String(item.student_id))
  );

  const nextAbsences = oldAbsences.filter((item) => !(
    item.exam_date === request.exam_date &&
    Number(item.period) === Number(request.period) &&
    String(item.class_id) === String(request.class_id)
  )).concat(preservedInactive, incoming);
  const nextRevision = currentRevision + 1;
  const nextCompletions = oldCompletions.filter((item) => String(item.key) !== key);
  nextCompletions.push({
    key: key,
    id: currentState ? currentState.id : createId_("completion"),
    exam_date: request.exam_date,
    period: Number(request.period),
    class_id: request.class_id,
    completed_at: now,
    updated_at: now,
    revision: nextRevision,
  });
  replaceAbsenceStateAtomically_(oldAbsences, oldCompletions, nextAbsences, nextCompletions);
  return {
    submitted: true,
    completed: true,
    revision: nextRevision,
    absence_count: incoming.length,
    archived_absence_count: preservedInactive.length,
    completed_at: now,
    absences: incoming,
  };
}

function replaceAbsenceStateAtomically_(oldAbsences, oldCompletions, nextAbsences, nextCompletions) {
  try {
    replaceRecords_(SHEETS.ABSENCES, nextAbsences);
    replaceRecords_(SHEETS.COMPLETIONS, nextCompletions);
    SpreadsheetApp.flush();
  } catch (error) {
    let recovered = true;
    try {
      replaceRecords_(SHEETS.ABSENCES, oldAbsences);
    } catch (ignored) {
      recovered = false;
    }
    try {
      replaceRecords_(SHEETS.COMPLETIONS, oldCompletions);
    } catch (ignored) {
      recovered = false;
    }
    try {
      SpreadsheetApp.flush();
    } catch (ignored) {
      recovered = false;
    }
    if (!recovered) {
      throw apiError_(
        "ATOMIC_RECOVERY_FAILED",
        "결시 제출 중 시트 복구가 완료되지 않았습니다. 고사본부에서 현재 상태를 확인한 뒤 다시 제출하세요.",
        { recovery_required: true }
      );
    }
    throw apiError_("ATOMIC_WRITE_FAILED", "결시 제출을 저장하지 못해 기존 상태로 되돌렸습니다.");
  }
}

function validateAbsenceReason_(value, detail) {
  const reason = String(value || "").trim();
  const allowed = new Set(["인정", "미인정", "질병", "기타"]);
  if (!reason) throw apiError_("ABSENCE_REASON_REQUIRED", "결시 사유를 선택하세요.");
  if (!allowed.has(reason)) throw apiError_("INVALID_ABSENCE_REASON", "결시 사유가 올바르지 않습니다.");
  if (reason === "기타" && !String(detail || "").trim()) {
    throw apiError_("ABSENCE_REASON_DETAIL_REQUIRED", "기타 사유의 상세 내용을 입력하세요.");
  }
  return reason;
}

function validateExamContext_(examDate, period, classId) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(examDate || ""))) throw apiError_("INVALID_DATE", "고사일이 올바르지 않습니다.");
  if (!getRecords_(SHEETS.EXAM_DATES).some((item) => item.exam_date === String(examDate) && isActive_(item))) {
    throw apiError_("EXAM_DATE_NOT_FOUND", "등록된 고사일을 찾지 못했습니다.");
  }
  if (!Number.isInteger(Number(period)) || Number(period) < 1) throw apiError_("INVALID_PERIOD", "교시가 올바르지 않습니다.");
  const classInfo = getRecords_(SHEETS.CLASSES).find((item) => String(item.id) === String(classId) && isActive_(item));
  if (!classInfo) throw apiError_("CLASS_NOT_FOUND", "학급을 찾지 못했습니다.");
  return classInfo;
}

function reasonCounts_(rows) {
  return rows.reduce((counts, item) => {
    const reason = String(item.reason || "");
    if (reason) counts[reason] = (counts[reason] || 0) + 1;
    return counts;
  }, {});
}

function getHqStatus_(request) {
  const matches = (item) =>
    (!request.exam_date || item.exam_date === request.exam_date) &&
    (!request.period || Number(item.period) === Number(request.period)) &&
    (!request.grade || Number(item.grade) === Number(request.grade)) &&
    (!request.class_id || String(item.class_id) === String(request.class_id));
  const absences = decorateAbsences_(getRecords_(SHEETS.ABSENCES)).filter(matches);
  const completions = decorateCompletions_(getRecords_(SHEETS.COMPLETIONS))
    .map((item) => Object.assign({}, item, { submitted: Boolean(item.completed_at) }))
    .filter(matches);
  const classes = getRecords_(SHEETS.CLASSES).filter((item) =>
    isActive_(item) &&
    (!request.grade || Number(item.grade) === Number(request.grade)) &&
    (!request.class_id || String(item.id) === String(request.class_id))
  );
  const classIds = new Set(classes.map((item) => String(item.id)));
  const activeStudents = getRecords_(SHEETS.STUDENTS).filter((item) =>
    isActive_(item) && classIds.has(String(item.class_id))
  );
  const activeStudentIds = new Set(activeStudents.map((item) => String(item.id)));
  const activeAbsences = absences.filter((item) => activeStudentIds.has(String(item.student_id)));
  const timetable = getRecords_(SHEETS.TIMETABLE).filter((item) =>
    (!request.exam_date || item.exam_date === request.exam_date) &&
    (!request.period || Number(item.period) === Number(request.period)) &&
    (!request.grade || Number(item.grade) === Number(request.grade))
  );
  const combinations = new Map();
  absences.concat(completions).concat(timetable).forEach((item) => {
    if (!item.exam_date || !item.period) return;
    const key = item.exam_date + "|" + Number(item.period);
    combinations.set(key, { exam_date: item.exam_date, period: Number(item.period) });
  });
  if (request.exam_date && request.period) {
    combinations.set(request.exam_date + "|" + Number(request.period), {
      exam_date: request.exam_date,
      period: Number(request.period),
    });
  }
  const classStatus = [];
  const periodTotals = Array.from(combinations.values()).sort(compareDatePeriod_).map((combo) => {
    const comboAbsences = absences.filter((item) => item.exam_date === combo.exam_date && Number(item.period) === combo.period);
    const comboActiveAbsences = comboAbsences.filter((item) => activeStudentIds.has(String(item.student_id)));
    const comboStates = completions.filter((item) => item.exam_date === combo.exam_date && Number(item.period) === combo.period);
    classes.forEach((classInfo) => {
      const state = comboStates.find((item) => String(item.class_id) === String(classInfo.id));
      classStatus.push({
        exam_date: combo.exam_date,
        period: combo.period,
        class_id: classInfo.id,
        grade: Number(classInfo.grade),
        class_num: Number(classInfo.class_num),
        submitted: Boolean(state && state.submitted),
        revision: state ? Number(state.revision) || 0 : 0,
        completed_at: state && state.completed_at ? state.completed_at : "",
        absence_count: comboActiveAbsences.filter((item) => String(item.class_id) === String(classInfo.id)).length,
      });
    });
    const submittedCount = comboStates.filter((item) => item.submitted && classes.some((classInfo) => String(classInfo.id) === String(item.class_id))).length;
    return {
      exam_date: combo.exam_date,
      period: combo.period,
      absence_count: comboActiveAbsences.length,
      archived_absence_count: comboAbsences.length - comboActiveAbsences.length,
      submitted_count: submittedCount,
      total_classes: classes.length,
      pending_count: Math.max(0, classes.length - submittedCount),
    };
  });
  return {
    absences: absences,
    completions: completions,
    class_status: classStatus,
    period_totals: periodTotals,
    summary: {
      enrolled: activeStudents.length,
      present: Math.max(0, activeStudents.length - activeAbsences.length),
      absent: activeAbsences.length,
      archived_absences: absences.length - activeAbsences.length,
      reasons: reasonCounts_(activeAbsences),
    },
  };
}

function getPeriodReport_(request) {
  const absences = decorateAbsences_(getRecords_(SHEETS.ABSENCES)).filter((item) =>
    (!request.start_date || item.exam_date >= request.start_date) &&
    (!request.end_date || item.exam_date <= request.end_date) &&
    (!request.exam_date || item.exam_date === request.exam_date) &&
    (!request.period || Number(item.period) === Number(request.period)) &&
    (!request.grade || Number(item.grade) === Number(request.grade)) &&
    (!request.class_id || String(item.class_id) === String(request.class_id))
  ).sort(compareDatePeriod_);
  return {
    absences: absences,
    details: absences,
    reason_summary: aggregateRows_(absences, ["reason"]),
    class_summary: aggregateRows_(absences, ["grade", "class_num", "class_id", "class_label"]),
    date_summary: aggregateRows_(absences, ["exam_date"]),
    period_summary: aggregateRows_(absences, ["exam_date", "period"]),
    total_count: absences.length,
  };
}

function aggregateRows_(rows, fields) {
  const groups = new Map();
  rows.forEach((row) => {
    const key = fields.map((field) => String(row[field] == null ? "" : row[field])).join("\u001f");
    if (!groups.has(key)) {
      const result = { count: 0 };
      fields.forEach((field) => { result[field] = row[field] == null ? "" : row[field]; });
      groups.set(key, result);
    }
    groups.get(key).count += 1;
  });
  return Array.from(groups.values()).sort((left, right) => {
    const leftKey = fields.map((field) => String(left[field])).join("|");
    const rightKey = fields.map((field) => String(right[field])).join("|");
    return leftKey.localeCompare(rightKey, "ko");
  });
}

function compareDatePeriod_(left, right) {
  return String(left.exam_date || "").localeCompare(String(right.exam_date || "")) ||
    Number(left.period || 0) - Number(right.period || 0) ||
    Number(left.grade || 0) - Number(right.grade || 0) ||
    Number(left.class_num || 0) - Number(right.class_num || 0);
}

function saveExamDates_(request) {
  const now = nowIso_();
  const oldRows = getRecords_(SHEETS.EXAM_DATES);
  const oldByDate = new Map(oldRows.map((item) => [String(item.exam_date), item]));
  const seen = new Set();
  const rows = (request.exam_dates || []).map((item) => {
    const examDate = String(item.exam_date || "");
    if (!/^\d{4}-\d{2}-\d{2}$/.test(examDate)) throw apiError_("INVALID_DATE", "고사일 형식이 올바르지 않습니다.");
    if (seen.has(examDate)) throw apiError_("DUPLICATE_DATE", "같은 고사일이 두 번 들어 있습니다.");
    seen.add(examDate);
    const existing = oldByDate.get(examDate);
    return {
      id: existing ? existing.id : item.id || createId_("date"),
      exam_date: examDate,
      label: String(item.label || "").trim().slice(0, 50),
      active: item.active !== false,
      updated_at: now,
    };
  });
  const incomingDates = new Set(rows.map((item) => String(item.exam_date)));
  if (oldRows.some((item) => isActive_(item) && !incomingDates.has(String(item.exam_date)))) {
    throw apiError_(
      "EXAM_DATE_DELETE_ACTION_REQUIRED",
      "고사일 삭제는 영향 범위를 확인할 수 있는 삭제 작업을 사용하세요."
    );
  }
  oldRows.filter((item) => !isActive_(item) && !incomingDates.has(String(item.exam_date)))
    .forEach((item) => rows.push(item));
  rows.sort((left, right) => String(left.exam_date).localeCompare(String(right.exam_date)));
  replaceRecords_(SHEETS.EXAM_DATES, rows);
  return adminBootstrap_();
}

function examDateDeleteImpact_(examDate) {
  return {
    timetable: getRecords_(SHEETS.TIMETABLE).filter((item) => item.exam_date === examDate).length,
    absences: getRecords_(SHEETS.ABSENCES).filter((item) => item.exam_date === examDate).length,
    completions: getRecords_(SHEETS.COMPLETIONS).filter((item) => item.exam_date === examDate).length,
    seat_charts: getRecords_(SHEETS.SEATING).filter((item) => item.exam_date === examDate).length,
  };
}

function deleteExamDate_(request) {
  validateRequired_(request, ["exam_date"]);
  const examDate = String(request.exam_date);
  const dates = getRecords_(SHEETS.EXAM_DATES);
  const target = dates.find((item) => item.exam_date === examDate && isActive_(item));
  if (!target) throw apiError_("EXAM_DATE_NOT_FOUND", "고사일을 찾지 못했습니다.");
  const impact = examDateDeleteImpact_(examDate);
  if (request.preview_only) return { deleted: false, preview: true, impact: impact };

  const now = nowIso_();
  replaceRecords_(SHEETS.TIMETABLE, getRecords_(SHEETS.TIMETABLE).filter((item) => item.exam_date !== examDate));
  replaceRecords_(SHEETS.ABSENCES, getRecords_(SHEETS.ABSENCES).filter((item) => item.exam_date !== examDate));
  replaceRecords_(SHEETS.COMPLETIONS, getRecords_(SHEETS.COMPLETIONS).filter((item) => item.exam_date !== examDate));
  replaceRecords_(SHEETS.SEATING, getRecords_(SHEETS.SEATING).filter((item) => item.exam_date !== examDate));
  replaceRecords_(SHEETS.EXAM_DATES, dates.map((item) =>
    item.exam_date === examDate ? Object.assign({}, item, { active: false, updated_at: now }) : item
  ));
  return Object.assign(adminBootstrap_(), { deleted: true, impact: impact });
}

function saveClass_(request) {
  const incoming = request.class_item || {};
  const grade = Number(incoming.grade);
  const classNum = Number(incoming.class_num);
  if (!Number.isInteger(grade) || grade < 1 || !Number.isInteger(classNum) || classNum < 1) {
    throw apiError_("INVALID_CLASS", "학년과 반은 1 이상의 정수여야 합니다.");
  }
  const now = nowIso_();
  let rows = getRecords_(SHEETS.CLASSES);
  const duplicate = rows.find((item) =>
    Number(item.grade) === grade &&
    Number(item.class_num) === classNum &&
    String(item.id) !== String(incoming.id || "")
  );
  if (duplicate && isActive_(duplicate)) throw apiError_("DUPLICATE_CLASS", "같은 학년과 반이 이미 있습니다.");
  const existing = incoming.id
    ? rows.find((item) => String(item.id) === String(incoming.id))
    : duplicate;
  rows = rows.filter((item) => String(item.id) !== String(existing ? existing.id : incoming.id || ""));
  const saved = {
    id: existing ? existing.id : incoming.id || createId_("class"),
    grade: grade,
    class_num: classNum,
    active: incoming.active !== false,
    updated_at: now,
  };
  rows.push(saved);
  rows.sort((a, b) => Number(a.grade) - Number(b.grade) || Number(a.class_num) - Number(b.class_num));
  replaceRecords_(SHEETS.CLASSES, rows);
  return Object.assign(adminBootstrap_(), { saved_class_id: saved.id });
}

function classDeleteImpact_(classId) {
  const studentIds = new Set(getRecords_(SHEETS.STUDENTS)
    .filter((item) => String(item.class_id) === String(classId))
    .map((item) => String(item.id)));
  const relatedSeatCharts = getRecords_(SHEETS.SEATING).filter((item) => {
    if (String(item.class_id) === String(classId)) return true;
    const chart = deserializeSeat_(item);
    return chart.examinee_ids.map(String).some((id) => studentIds.has(id)) ||
      extractAssignmentStudentIds_(chart.assignment).some((id) => studentIds.has(id));
  });
  return {
    students: studentIds.size,
    absences: getRecords_(SHEETS.ABSENCES).filter((item) => String(item.class_id) === String(classId)).length,
    completions: getRecords_(SHEETS.COMPLETIONS).filter((item) => String(item.class_id) === String(classId)).length,
    enrollments: getRecords_(SHEETS.ENROLLMENTS).filter((item) => String(item.class_id) === String(classId) || studentIds.has(String(item.student_id))).length,
    timetable: getRecords_(SHEETS.TIMETABLE).filter((item) => normalizeClassIds_(item.class_ids).indexOf(String(classId)) !== -1).length,
    seat_charts: relatedSeatCharts.length,
  };
}

function deleteClass_(request) {
  validateRequired_(request, ["class_id"]);
  const classId = String(request.class_id);
  const classes = getRecords_(SHEETS.CLASSES);
  const target = classes.find((item) => String(item.id) === classId);
  if (!target) throw apiError_("CLASS_NOT_FOUND", "학급을 찾지 못했습니다.");
  const impact = classDeleteImpact_(classId);
  if (request.preview_only) return { deleted: false, preview: true, impact: impact };

  const now = nowIso_();
  replaceRecords_(SHEETS.CLASSES, classes.map((item) =>
    String(item.id) === classId ? Object.assign({}, item, { active: false, updated_at: now }) : item
  ));
  const students = getRecords_(SHEETS.STUDENTS).map((item) =>
    String(item.class_id) === classId ? Object.assign({}, item, { active: false, updated_at: now }) : item
  );
  replaceRecords_(SHEETS.STUDENTS, students);
  const inactiveStudentIds = new Set(students.filter((item) => String(item.class_id) === classId).map((item) => String(item.id)));
  replaceRecords_(SHEETS.ENROLLMENTS, getRecords_(SHEETS.ENROLLMENTS).filter((item) =>
    String(item.class_id) !== classId && !inactiveStudentIds.has(String(item.student_id))
  ));
  const timetable = [];
  getRecords_(SHEETS.TIMETABLE).forEach((item) => {
    const classIds = normalizeClassIds_(item.class_ids);
    if (classIds.indexOf(classId) === -1) {
      timetable.push(item);
      return;
    }
    const nextIds = classIds.filter((id) => id !== classId);
    if (nextIds.length) timetable.push(Object.assign({}, item, { class_ids: JSON.stringify(nextIds), updated_at: now }));
  });
  replaceRecords_(SHEETS.TIMETABLE, timetable);
  replaceRecords_(SHEETS.SEATING, getRecords_(SHEETS.SEATING).filter((item) => {
    if (String(item.class_id) === classId) return false;
    const chart = deserializeSeat_(item);
    return !chart.examinee_ids.map(String).some((id) => inactiveStudentIds.has(id)) &&
      !extractAssignmentStudentIds_(chart.assignment).some((id) => inactiveStudentIds.has(id));
  }));
  return Object.assign(adminBootstrap_(), { deleted: true, impact: impact });
}

function replaceStudents_(request) {
  validateRequired_(request, ["class_id"]);
  const classId = String(request.class_id);
  if (!getRecords_(SHEETS.CLASSES).some((item) => String(item.id) === classId && isActive_(item))) {
    throw apiError_("CLASS_NOT_FOUND", "학급을 찾지 못했습니다.");
  }
  if (!Array.isArray(request.students)) throw apiError_("INVALID_STUDENTS", "학생 명단 형식이 올바르지 않습니다.");
  const now = nowIso_();
  const allRows = getRecords_(SHEETS.STUDENTS);
  const classRows = allRows.filter((item) => String(item.class_id) === classId);
  const byId = new Map(classRows.map((item) => [String(item.id), item]));
  const byNumber = new Map(classRows.map((item) => [Number(item.number), item]));
  const seen = new Set();
  const activeRows = request.students.map((item) => {
    const number = Number(item.number);
    const name = String(item.name || "").trim().slice(0, 50);
    if (!Number.isInteger(number) || number < 1 || !name) throw apiError_("INVALID_STUDENT", "번호와 이름을 확인하세요.");
    if (seen.has(number)) throw apiError_("DUPLICATE_STUDENT_NUMBER", "같은 번호가 두 번 들어 있습니다.");
    seen.add(number);
    const requestedById = item.id ? byId.get(String(item.id)) : null;
    if (item.id && !requestedById) throw apiError_("INVALID_STUDENT_ID", "다른 학급의 학생 ID는 사용할 수 없습니다.");
    const existing = requestedById || byNumber.get(number);
    return {
      id: existing ? existing.id : createId_("student"),
      class_id: classId,
      number: number,
      name: name,
      active: true,
      updated_at: now,
    };
  });
  const activeIds = new Set(activeRows.map((item) => String(item.id)));
  const inactiveRows = classRows.filter((item) => !activeIds.has(String(item.id))).map((item) =>
    Object.assign({}, item, { active: false, updated_at: now })
  );
  const nextRows = allRows.filter((item) => String(item.class_id) !== classId).concat(activeRows, inactiveRows);
  replaceRecords_(SHEETS.STUDENTS, nextRows);
  const inactiveIds = new Set(inactiveRows.map((item) => String(item.id)));
  if (inactiveIds.size) {
    replaceRecords_(SHEETS.ENROLLMENTS, getRecords_(SHEETS.ENROLLMENTS).filter((item) => !inactiveIds.has(String(item.student_id))));
    replaceRecords_(SHEETS.SEATING, getRecords_(SHEETS.SEATING).filter((item) => {
      const chart = deserializeSeat_(item);
      return !chart.examinee_ids.map(String).some((id) => inactiveIds.has(id)) &&
        !extractAssignmentStudentIds_(chart.assignment).some((id) => inactiveIds.has(id));
    }));
  }
  return Object.assign(adminBootstrap_(), {
    student_result: { active_count: activeRows.length, deactivated_count: inactiveRows.length, preserved_id_count: activeRows.filter((item) => byId.has(String(item.id))).length },
  });
}

function importStudents_(request) {
  if (!Array.isArray(request.rows)) throw apiError_("INVALID_STUDENTS", "학생 명단 형식이 올바르지 않습니다.");
  const now = nowIso_();
  let classes = getRecords_(SHEETS.CLASSES);
  let students = getRecords_(SHEETS.STUDENTS);
  const errors = [];
  const seen = new Set();
  request.rows.forEach((row, index) => {
    const grade = Number(row.grade);
    const classNum = Number(row.class_num);
    const number = Number(row.number);
    const name = String(row.name || "").trim().slice(0, 50);
    const rowNumber = Number(row.row_number) || index + 2;
    if (![grade, classNum, number].every(Number.isInteger) || grade < 1 || classNum < 1 || number < 1 || !name) {
      errors.push({ row_number: rowNumber, code: "INVALID_STUDENT_ROW" });
      return;
    }
    const key = [grade, classNum, number].join("|");
    if (seen.has(key)) {
      errors.push({ row_number: rowNumber, code: "DUPLICATE_STUDENT_NUMBER" });
      return;
    }
    seen.add(key);
    let classInfo = classes.find((item) => Number(item.grade) === grade && Number(item.class_num) === classNum);
    if (!classInfo) {
      classInfo = { id: createId_("class"), grade: grade, class_num: classNum, active: true, updated_at: now };
      classes.push(classInfo);
    } else if (!isActive_(classInfo)) {
      classInfo.active = true;
      classInfo.updated_at = now;
    }
    const existing = students.find((item) =>
      String(item.class_id) === String(classInfo.id) && Number(item.number) === number
    );
    const saved = {
      id: existing ? existing.id : createId_("student"),
      class_id: classInfo.id,
      number: number,
      name: name,
      active: true,
      updated_at: now,
    };
    students = students.filter((item) => String(item.id) !== String(saved.id));
    students.push(saved);
  });
  if (errors.length) {
    throw apiError_(
      "STUDENT_IMPORT_VALIDATION_FAILED",
      "학생 명단에 저장할 수 없는 행이 있습니다.",
      { errors: errors, invalid_count: errors.length }
    );
  }
  classes.sort((a, b) => Number(a.grade) - Number(b.grade) || Number(a.class_num) - Number(b.class_num));
  replaceRecords_(SHEETS.CLASSES, classes);
  replaceRecords_(SHEETS.STUDENTS, students);
  return Object.assign(adminBootstrap_(), { import_result: { saved_count: request.rows.length - errors.length, errors: errors } });
}

function saveStudent_(request) {
  const incoming = request.student || request.student_item || {};
  const classId = String(incoming.class_id || request.class_id || "");
  const number = Number(incoming.number);
  const name = String(incoming.name || "").trim().slice(0, 50);
  if (!classId || !Number.isInteger(number) || number < 1 || !name) {
    throw apiError_("INVALID_STUDENT", "학급, 번호, 이름을 확인하세요.");
  }
  if (!getRecords_(SHEETS.CLASSES).some((item) => String(item.id) === classId && isActive_(item))) {
    throw apiError_("CLASS_NOT_FOUND", "학급을 찾지 못했습니다.");
  }
  const rows = getRecords_(SHEETS.STUDENTS);
  const byId = incoming.id ? rows.find((item) => String(item.id) === String(incoming.id)) : null;
  if (incoming.id && (!byId || String(byId.class_id) !== classId)) {
    throw apiError_("INVALID_STUDENT_ID", "학생 ID와 학급이 일치하지 않습니다.");
  }
  const byNumber = rows.find((item) =>
    String(item.class_id) === classId &&
    Number(item.number) === number &&
    String(item.id) !== String(incoming.id || "") &&
    isActive_(item)
  );
  if (byNumber) throw apiError_("DUPLICATE_STUDENT_NUMBER", "같은 번호의 학생이 이미 있습니다.");
  const reusable = byId || rows.find((item) => String(item.class_id) === classId && Number(item.number) === number);
  const saved = {
    id: reusable ? reusable.id : createId_("student"),
    class_id: classId,
    number: number,
    name: name,
    active: true,
    updated_at: nowIso_(),
  };
  replaceRecords_(SHEETS.STUDENTS, rows.filter((item) => String(item.id) !== String(saved.id)).concat(saved));
  return Object.assign(adminBootstrap_(), { saved_student_id: saved.id });
}

function studentDeleteImpact_(studentId) {
  const student = getRecords_(SHEETS.STUDENTS).find((item) => String(item.id) === String(studentId));
  if (!student) throw apiError_("STUDENT_NOT_FOUND", "학생을 찾지 못했습니다.");
  const seatCharts = getRecords_(SHEETS.SEATING).filter((item) => {
    const chart = deserializeSeat_(item);
    const assignmentIds = extractAssignmentStudentIds_(chart.assignment);
    return chart.examinee_ids.map(String).indexOf(String(studentId)) !== -1 || assignmentIds.indexOf(String(studentId)) !== -1;
  }).length;
  return {
    absences: getRecords_(SHEETS.ABSENCES).filter((item) => String(item.student_id) === String(studentId)).length,
    enrollments: getRecords_(SHEETS.ENROLLMENTS).filter((item) => String(item.student_id) === String(studentId)).length,
    seat_charts: seatCharts,
  };
}

function deleteStudent_(request) {
  validateRequired_(request, ["student_id"]);
  const studentId = String(request.student_id);
  const impact = studentDeleteImpact_(studentId);
  if (request.preview_only) return { deleted: false, preview: true, impact: impact };
  const now = nowIso_();
  replaceRecords_(SHEETS.STUDENTS, getRecords_(SHEETS.STUDENTS).map((item) =>
    String(item.id) === studentId ? Object.assign({}, item, { active: false, updated_at: now }) : item
  ));
  replaceRecords_(SHEETS.ENROLLMENTS, getRecords_(SHEETS.ENROLLMENTS).filter((item) => String(item.student_id) !== studentId));
  replaceRecords_(SHEETS.SEATING, getRecords_(SHEETS.SEATING).filter((item) => {
    const chart = deserializeSeat_(item);
    return chart.examinee_ids.map(String).indexOf(studentId) === -1 &&
      extractAssignmentStudentIds_(chart.assignment).indexOf(studentId) === -1;
  }));
  return Object.assign(adminBootstrap_(), { deleted: true, impact: impact });
}

function saveTimetable_(request) {
  if (!Array.isArray(request.timetable)) throw apiError_("INVALID_TIMETABLE", "시간표 형식이 올바르지 않습니다.");
  const activeClasses = getRecords_(SHEETS.CLASSES).filter(isActive_);
  const classById = new Map(activeClasses.map((item) => [String(item.id), item]));
  const activeExamDates = new Set(getRecords_(SHEETS.EXAM_DATES).filter(isActive_).map((item) => String(item.exam_date)));
  const oldRows = getRecords_(SHEETS.TIMETABLE);
  const naturalId = new Map(oldRows.map((item) => [[item.exam_date, Number(item.grade), Number(item.period), item.subject_type, item.subject_name].join("|"), item.id]));
  const now = nowIso_();
  const seen = new Set();
  const subjectSeen = new Set();
  const rows = request.timetable.map((item) => {
    const examDate = String(item.exam_date || "");
    const grade = Number(item.grade);
    const period = Number(item.period);
    const subjectName = String(item.subject_name || "").trim().slice(0, 80);
    const subjectType = item.subject_type === "elective" ? "elective" : "common";
    if (!/^\d{4}-\d{2}-\d{2}$/.test(examDate) || !Number.isInteger(grade) || !Number.isInteger(period) || period < 1 || !subjectName) {
      throw apiError_("INVALID_TIMETABLE_ROW", "고사일, 학년, 교시, 과목명을 확인하세요.");
    }
    if (!activeExamDates.has(examDate)) throw apiError_("TIMETABLE_DATE_NOT_FOUND", "등록된 고사일을 찾지 못했습니다.");
    const classIds = normalizeClassIds_(item.class_ids);
    if (!classIds.length) throw apiError_("TIMETABLE_CLASS_REQUIRED", "시간표 적용 학급을 한 곳 이상 선택하세요.");
    classIds.forEach((classId) => {
      const classInfo = classById.get(classId);
      if (!classInfo) throw apiError_("TIMETABLE_CLASS_NOT_FOUND", "시간표 적용 학급을 찾지 못했습니다.", { class_id: classId });
      if (Number(classInfo.grade) !== grade) throw apiError_("TIMETABLE_CLASS_GRADE_MISMATCH", "시간표 학년과 적용 학급의 학년이 다릅니다.", { class_id: classId });
    });
    const naturalKey = [examDate, grade, period, subjectType, subjectName].join("|");
    if (seen.has(naturalKey)) throw apiError_("DUPLICATE_TIMETABLE_ROW", "같은 시간표 항목이 두 번 들어 있습니다.");
    seen.add(naturalKey);
    const subjectKey = [examDate, grade, period, subjectName].join("|");
    if (subjectSeen.has(subjectKey)) {
      throw apiError_("AMBIGUOUS_TIMETABLE_SUBJECT", "같은 교시에 같은 과목명을 공통·선택으로 중복 등록할 수 없습니다.");
    }
    subjectSeen.add(subjectKey);
    return {
      id: item.id || naturalId.get(naturalKey) || createId_("timetable"),
      exam_date: examDate,
      grade: grade,
      period: period,
      start_time: String(item.start_time || "").slice(0, 10),
      end_time: String(item.end_time || "").slice(0, 10),
      subject_name: subjectName,
      subject_type: subjectType,
      room_name: String(item.room_name || "").trim().slice(0, 80),
      class_ids: JSON.stringify(classIds),
      updated_at: now,
    };
  });
  const incomingIds = new Set(rows.map((item) => String(item.id)));
  if (oldRows.some((item) => !incomingIds.has(String(item.id)))) {
    throw apiError_(
      "TIMETABLE_DELETE_ACTION_REQUIRED",
      "시간표 삭제는 영향 범위를 확인할 수 있는 삭제 작업을 사용하세요."
    );
  }
  replaceRecords_(SHEETS.TIMETABLE, rows);
  return adminBootstrap_();
}

function timetableDeleteImpact_(target, timetable) {
  const remainingSubjectRows = timetable.filter((item) =>
    String(item.id) !== String(target.id) &&
    Number(item.grade) === Number(target.grade) &&
    String(item.subject_name) === String(target.subject_name) &&
    item.subject_type === "elective"
  );
  const removeEnrollments = target.subject_type === "elective" && remainingSubjectRows.length === 0;
  return {
    enrollments: removeEnrollments
      ? getRecords_(SHEETS.ENROLLMENTS).filter((item) =>
          Number(item.grade) === Number(target.grade) &&
          String(item.subject_name) === String(target.subject_name)
        ).length
      : 0,
    seat_charts: getRecords_(SHEETS.SEATING).filter((item) =>
      item.exam_date === target.exam_date &&
      Number(item.period) === Number(target.period) &&
      Number(item.grade) === Number(target.grade) &&
      String(item.subject_name) === String(target.subject_name)
    ).length,
    remove_subject_enrollments: removeEnrollments,
  };
}

function deleteTimetable_(request) {
  validateRequired_(request, ["id"]);
  const timetable = getRecords_(SHEETS.TIMETABLE);
  const target = timetable.find((item) => String(item.id) === String(request.id));
  if (!target) throw apiError_("TIMETABLE_NOT_FOUND", "시간표 항목을 찾지 못했습니다.");
  const impact = timetableDeleteImpact_(target, timetable);
  if (request.preview_only) return { deleted: false, preview: true, impact: impact };

  if (impact.remove_subject_enrollments) {
    replaceRecords_(SHEETS.ENROLLMENTS, getRecords_(SHEETS.ENROLLMENTS).filter((item) => !(
      Number(item.grade) === Number(target.grade) &&
      String(item.subject_name) === String(target.subject_name)
    )));
  }
  replaceRecords_(SHEETS.SEATING, getRecords_(SHEETS.SEATING).filter((item) => !(
    item.exam_date === target.exam_date &&
    Number(item.period) === Number(target.period) &&
    Number(item.grade) === Number(target.grade) &&
    String(item.subject_name) === String(target.subject_name)
  )));
  replaceRecords_(SHEETS.TIMETABLE, timetable.filter((item) => String(item.id) !== String(target.id)));
  return Object.assign(adminBootstrap_(), { deleted: true, impact: impact });
}

function saveEnrollments_(request) {
  const grade = Number(request.grade);
  if (!Number.isInteger(grade) || grade < 1) throw apiError_("INVALID_GRADE", "학년이 올바르지 않습니다.");
  if (!Array.isArray(request.enrollments)) throw apiError_("INVALID_ENROLLMENTS", "선택과목 명단 형식이 올바르지 않습니다.");
  const classes = getRecords_(SHEETS.CLASSES).filter((item) => Number(item.grade) === grade && isActive_(item));
  const classById = new Map(classes.map((item) => [String(item.id), item]));
  const classByNumber = new Map(classes.map((item) => [Number(item.class_num), item]));
  const students = getRecords_(SHEETS.STUDENTS).filter(isActive_);
  const studentById = new Map(students.map((item) => [String(item.id), item]));
  const studentByClassNumber = new Map(students.map((item) => [String(item.class_id) + "|" + Number(item.number), item]));
  const electiveSubjects = new Map();
  getRecords_(SHEETS.TIMETABLE)
    .filter((item) => Number(item.grade) === grade && item.subject_type === "elective")
    .forEach((item) => {
      const subjectName = String(item.subject_name);
      const classIds = electiveSubjects.get(subjectName) || new Set();
      normalizeClassIds_(item.class_ids).forEach((id) => classIds.add(String(id)));
      electiveSubjects.set(subjectName, classIds);
    });
  const oldRows = getRecords_(SHEETS.ENROLLMENTS);
  const reusable = new Map(oldRows.filter((item) => Number(item.grade) === grade)
    .map((item) => [[item.student_id, item.subject_name].join("|"), item]));
  const errors = [];
  const now = nowIso_();
  const validRows = [];
  const seen = new Set();

  request.enrollments.forEach((item, index) => {
    const rowNumber = Number(item.row_number) || index + 2;
    const classInfo = item.class_id ? classById.get(String(item.class_id)) : classByNumber.get(Number(item.class_num));
    if (!classInfo) {
      errors.push({ row_number: rowNumber, code: "CLASS_NOT_FOUND" });
      return;
    }
    const student = item.student_id
      ? studentById.get(String(item.student_id))
      : studentByClassNumber.get(String(classInfo.id) + "|" + Number(item.number));
    if (!student || String(student.class_id) !== String(classInfo.id)) {
      errors.push({ row_number: rowNumber, code: "STUDENT_NOT_FOUND" });
      return;
    }
    if (item.name && String(item.name).trim() !== String(student.name).trim()) {
      errors.push({ row_number: rowNumber, code: "STUDENT_NAME_MISMATCH" });
      return;
    }
    const subjectName = String(item.subject_name || "").trim().slice(0, 80);
    const applicableClasses = electiveSubjects.get(subjectName);
    if (!subjectName || !applicableClasses) {
      errors.push({ row_number: rowNumber, code: "SUBJECT_NOT_FOUND" });
      return;
    }
    if (!applicableClasses.has(String(classInfo.id))) {
      errors.push({ row_number: rowNumber, code: "SUBJECT_CLASS_MISMATCH" });
      return;
    }
    const duplicateKey = student.id + "|" + subjectName;
    if (seen.has(duplicateKey)) {
      errors.push({ row_number: rowNumber, code: "DUPLICATE_ENROLLMENT" });
      return;
    }
    seen.add(duplicateKey);
    const existing = reusable.get(duplicateKey);
    validRows.push({
      id: existing ? existing.id : item.id || createId_("enrollment"),
      grade: grade,
      class_id: classInfo.id,
      student_id: student.id,
      subject_name: subjectName,
      room_name: String(item.room_name || "").trim().slice(0, 80),
      updated_at: now,
    });
  });

  if (errors.length) {
    throw apiError_(
      "ENROLLMENT_VALIDATION_FAILED",
      "선택과목 명단에 저장할 수 없는 행이 있습니다.",
      { errors: errors, unmatched_count: errors.length, matched_count: validRows.length }
    );
  }
  const rows = oldRows.filter((item) => Number(item.grade) !== grade).concat(validRows);
  replaceRecords_(SHEETS.ENROLLMENTS, rows);
  return Object.assign(adminBootstrap_(), {
    enrollment_result: {
      matched_count: validRows.length,
      unmatched_count: errors.length,
      errors: errors,
    },
  });
}

function validateSeatChart_(chart) {
  const examDate = String(chart.exam_date || "");
  const period = Number(chart.period);
  const grade = Number(chart.grade);
  const mode = chart.mode === "own" ? "own" : "separate";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(examDate)) throw apiError_("INVALID_DATE", "자리배치 고사일이 올바르지 않습니다.");
  if (!getRecords_(SHEETS.EXAM_DATES).some((item) => item.exam_date === examDate && isActive_(item))) {
    throw apiError_("EXAM_DATE_NOT_FOUND", "자리배치 고사일을 찾지 못했습니다.");
  }
  if (!Number.isInteger(period) || period < 1 || !Number.isInteger(grade) || grade < 1) {
    throw apiError_("INVALID_SEAT_CONTEXT", "자리배치 학년과 교시를 확인하세요.");
  }
  const rows = Number(chart.rows) || 6;
  const cols = Number(chart.cols) || 5;
  if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1 || rows > 20 || cols > 20) {
    throw apiError_("INVALID_SEAT_SIZE", "자리배치 행과 열을 확인하세요.");
  }
  const classes = getRecords_(SHEETS.CLASSES).filter(isActive_);
  const classInfo = chart.class_id ? classes.find((item) => String(item.id) === String(chart.class_id)) : null;
  if (mode === "own" && !classInfo) throw apiError_("CLASS_NOT_FOUND", "각자교실 자리배치의 학급을 찾지 못했습니다.");
  if (classInfo && Number(classInfo.grade) !== grade) throw apiError_("SEAT_CLASS_GRADE_MISMATCH", "자리배치 학년과 학급이 일치하지 않습니다.");
  if (mode === "separate" && (!String(chart.subject_name || "").trim() || !String(chart.room_name || "").trim())) {
    throw apiError_("SEAT_SUBJECT_ROOM_REQUIRED", "별실 자리배치에는 과목과 호실이 필요합니다.");
  }
  const subjectName = String(chart.subject_name || "").trim();
  let subjectRows = [];
  if (subjectName) {
    subjectRows = getRecords_(SHEETS.TIMETABLE).filter((item) =>
      item.exam_date === examDate &&
      Number(item.period) === period &&
      Number(item.grade) === grade &&
      String(item.subject_name) === subjectName
    );
    if (!subjectRows.length) throw apiError_("SEAT_SUBJECT_NOT_FOUND", "시간표에서 자리배치 과목을 찾지 못했습니다.");
  }
  const activeStudents = getRecords_(SHEETS.STUDENTS).filter(isActive_);
  const studentById = new Map(activeStudents.map((item) => [String(item.id), item]));
  const examineeIds = (chart.examinee_ids || []).map(String);
  if (new Set(examineeIds).size !== examineeIds.length) {
    throw apiError_("DUPLICATE_SEAT_EXAMINEE", "자리배치 응시자 목록에 같은 학생이 두 번 들어 있습니다.");
  }
  const assignmentIds = extractAssignmentStudentIds_(chart.assignment || []);
  if (new Set(assignmentIds).size !== assignmentIds.length) {
    throw apiError_("DUPLICATE_SEAT_STUDENT", "같은 학생을 두 좌석에 배치할 수 없습니다.");
  }
  examineeIds.concat(assignmentIds).forEach((studentId) => {
    const student = studentById.get(studentId);
    if (!student) throw apiError_("SEAT_STUDENT_NOT_FOUND", "자리배치에 없는 학생이 포함되어 있습니다.");
    const studentClass = classes.find((item) => String(item.id) === String(student.class_id));
    if (!studentClass || Number(studentClass.grade) !== grade) throw apiError_("SEAT_STUDENT_GRADE_MISMATCH", "자리배치 학년과 학생의 학년이 다릅니다.");
    if (mode === "own" && String(student.class_id) !== String(chart.class_id)) {
      throw apiError_("SEAT_STUDENT_CLASS_MISMATCH", "각자교실에는 다른 학급 학생을 배치할 수 없습니다.");
    }
  });
  if (subjectRows.length) {
    const applicableClasses = new Set();
    subjectRows.forEach((item) => normalizeClassIds_(item.class_ids)
      .forEach((classId) => applicableClasses.add(String(classId))));
    if (mode === "own" && !applicableClasses.has(String(chart.class_id))) {
      throw apiError_("SEAT_SUBJECT_CLASS_MISMATCH", "선택한 과목이 각자교실 학급에 적용되지 않습니다.");
    }
    const elective = subjectRows.some((item) => item.subject_type === "elective");
    const enrolled = elective
      ? new Set(getRecords_(SHEETS.ENROLLMENTS)
          .filter((item) => Number(item.grade) === grade && String(item.subject_name) === subjectName)
          .map((item) => String(item.student_id)))
      : null;
    examineeIds.forEach((studentId) => {
      const student = studentById.get(studentId);
      if (!applicableClasses.has(String(student.class_id))) {
        throw apiError_("SEAT_SUBJECT_CLASS_MISMATCH", "과목 적용 학급 밖의 학생이 응시자에 포함되어 있습니다.");
      }
      if (enrolled && !enrolled.has(studentId)) {
        throw apiError_("SEAT_STUDENT_NOT_ENROLLED", "선택과목 미등록 학생이 응시자에 포함되어 있습니다.");
      }
    });
  }
  if (mode === "separate") {
    const examineeSet = new Set(examineeIds);
    assignmentIds.forEach((studentId) => {
      if (!examineeSet.has(studentId)) {
        throw apiError_("SEAT_ASSIGNMENT_NOT_EXAMINEE", "별실 배치에 응시자로 선택하지 않은 학생이 포함되어 있습니다.");
      }
    });
  }
  const disabledList = (chart.disabled_seats || []).map(String);
  const disabled = new Set(disabledList);
  if (disabled.size !== disabledList.length) {
    throw apiError_("DUPLICATE_DISABLED_SEAT", "같은 제외 좌석이 두 번 들어 있습니다.");
  }
  disabled.forEach((key) => {
    if (!validSeatKey_(key, rows, cols)) throw apiError_("INVALID_DISABLED_SEAT", "범위를 벗어난 제외 좌석이 있습니다.");
  });
  const assignment = Array.isArray(chart.assignment) ? chart.assignment : [];
  const assignedSeats = new Set();
  assignment.forEach((item) => {
    const key = String(item && typeof item === "object" ? item.key || "" : "");
    if (!validSeatKey_(key, rows, cols)) throw apiError_("INVALID_ASSIGNMENT_SEAT", "범위를 벗어난 배치 좌석이 있습니다.");
    if (disabled.has(key)) throw apiError_("ASSIGNMENT_ON_DISABLED_SEAT", "제외 좌석에 학생을 배치할 수 없습니다.");
    if (assignedSeats.has(key)) throw apiError_("DUPLICATE_ASSIGNMENT_SEAT", "같은 좌석이 배치표에 두 번 들어 있습니다.");
    assignedSeats.add(key);
  });
  const capacity = rows * cols - disabled.size;
  if (assignment.length > capacity || assignmentIds.length > capacity) {
    throw apiError_("SEAT_CAPACITY_EXCEEDED", "사용 가능한 좌석보다 배치 학생이 많습니다.");
  }
}

function validSeatKey_(key, rows, cols) {
  const match = /^(\d+)-(\d+)$/.exec(String(key || ""));
  return Boolean(match && Number(match[1]) < rows && Number(match[2]) < cols);
}

function extractAssignmentStudentIds_(assignment) {
  return (Array.isArray(assignment) ? assignment : []).map((item) => {
    if (item && typeof item === "object") return String(item.student_id || item.id || "");
    return String(item || "");
  }).filter(Boolean);
}

function saveSeatChart_(request) {
  const incoming = request.chart || {};
  validateSeatChart_(incoming);
  const row = serializeSeat_(incoming);
  let rows = getRecords_(SHEETS.SEATING);
  const naturalKey = seatNaturalKey_(row);
  const existing = rows.find((item) =>
    (row.id && String(item.id) === String(row.id)) ||
    seatNaturalKey_(item) === naturalKey
  );
  const updated = Boolean(existing);
  if (existing) row.id = existing.id;
  rows = rows.filter((item) => String(item.id) !== String(row.id) && seatNaturalKey_(item) !== naturalKey);
  rows.push(row);
  replaceRecords_(SHEETS.SEATING, rows);
  return {
    chart: deserializeSeat_(row),
    seat_charts: rows.map(deserializeSeat_).reverse(),
    created_count: updated ? 0 : 1,
    updated_count: updated ? 1 : 0,
  };
}

function saveSeatChartsBatch_(request) {
  if (!Array.isArray(request.charts)) throw apiError_("INVALID_SEAT_CHARTS", "자리배치 목록 형식이 올바르지 않습니다.");
  let rows = getRecords_(SHEETS.SEATING);
  let createdCount = 0;
  let updatedCount = 0;
  request.charts.forEach((chart) => {
    validateSeatChart_(chart);
    const row = serializeSeat_(chart);
    const naturalKey = seatNaturalKey_(row);
    const existing = rows.find((item) =>
      (row.id && String(item.id) === String(row.id)) ||
      seatNaturalKey_(item) === naturalKey
    );
    if (existing) {
      row.id = existing.id;
      updatedCount += 1;
    } else {
      createdCount += 1;
    }
    rows = rows.filter((item) => String(item.id) !== String(row.id) && seatNaturalKey_(item) !== naturalKey);
    rows.push(row);
  });
  replaceRecords_(SHEETS.SEATING, rows);
  return {
    seat_charts: rows.map(deserializeSeat_).reverse(),
    created_count: createdCount,
    updated_count: updatedCount,
  };
}

function deleteSeatChart_(request) {
  validateRequired_(request, ["id"]);
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

function safeAppendAudit_(role, action, targetType, targetId, count, success, detail) {
  try {
    appendAudit_(role, action, targetType, targetId, count, success, detail);
  } catch (error) {
    console.error("AUDIT_WRITE_FAILED", action, success ? "success" : "failure");
  }
}

function readMutationCache_(cache, key, request, role) {
  let raw = "";
  try {
    raw = cache.get(key) || "";
  } catch (ignored) {
    return null;
  }
  if (!raw) return null;
  const cached = JSON.parse(raw);
  if (!cached.__committed) return cached;
  if (request.action === "submit_absence_input") {
    return Object.assign(getAbsenceContext_(request), { idempotent_replay: true });
  }
  return Object.assign(role === "admin" ? adminBootstrap_() : teacherBootstrap_(), {
    idempotent_replay: true,
  });
}

function writeMutationCache_(cache, key, result) {
  const serialized = JSON.stringify(result);
  let value = serialized;
  try {
    if (Utilities.newBlob(serialized).getBytes().length > 80000) {
      value = JSON.stringify({ __committed: true });
    }
    cache.put(key, value, APP.IDEMPOTENCY_TTL_SECONDS);
  } catch (error) {
    try {
      cache.put(key, JSON.stringify({ __committed: true }), APP.IDEMPOTENCY_TTL_SECONDS);
    } catch (ignored) {}
  }
}

function withMutation_(request, role, operation) {
  const requestId = String(request.request_id || "");
  if (!requestId) throw apiError_("REQUEST_ID_REQUIRED", "저장 요청 식별자가 없습니다. 화면을 새로고침한 뒤 다시 시도하세요.");
  const cache = CacheService.getScriptCache();
  const idempotencyKey = ["request", role, request.action, requestId].join(":");
  const cached = readMutationCache_(cache, idempotencyKey, request, role);
  if (cached) return cached;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(28000)) throw apiError_("BUSY", "동시 저장 요청이 많습니다. 잠시 후 다시 시도하세요.");
  try {
    const cachedAfterLock = readMutationCache_(cache, idempotencyKey, request, role);
    if (cachedAfterLock) return cachedAfterLock;
    const result = operation();
    writeMutationCache_(cache, idempotencyKey, result);
    safeAppendAudit_(role, request.action, "records", "", inferCount_(request), true, "OK");
    return result;
  } catch (error) {
    if (!error.auditLogged) {
      safeAppendAudit_(role, request.action, "records", "", 0, false, error.code || "REQUEST_FAILED");
      error.auditLogged = true;
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function inferCount_(request) {
  if (Array.isArray(request.absences)) return request.absences.length;
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
  const session = { role: role, created_at: nowIso_() };
  if (role === "admin") {
    session.session_version = Number(PropertiesService.getScriptProperties().getProperty(APP.PROPERTY_ADMIN_SESSION_VERSION) || 1);
  }
  CacheService.getScriptCache().put("session:" + token, JSON.stringify(session), ttl);
  return token;
}

function requireSession_(token, role) {
  if (!token) throw apiError_("SESSION_REQUIRED", "로그인이 필요합니다.");
  const cache = CacheService.getScriptCache();
  const key = "session:" + token;
  const raw = cache.get(key);
  if (!raw) {
    throw role === "admin"
      ? apiError_("ADMIN_SESSION_EXPIRED", "관리자 세션이 만료되었습니다. 관리자 암호를 다시 입력하세요.")
      : apiError_("SESSION_EXPIRED", "세션이 만료되었습니다. 다시 로그인하세요.");
  }
  const session = JSON.parse(raw);
  if (session.role !== role) throw apiError_("FORBIDDEN", "이 작업에 필요한 권한이 없습니다.");
  if (role === "admin") {
    const currentVersion = Number(PropertiesService.getScriptProperties().getProperty(APP.PROPERTY_ADMIN_SESSION_VERSION) || 1);
    if (Number(session.session_version) !== currentVersion) {
      cache.remove(key);
      throw apiError_("ADMIN_SESSION_EXPIRED", "관리자 암호가 변경되어 관리자 세션이 만료되었습니다.");
    }
  }
  const ttl = role === "admin" ? APP.ADMIN_TTL_SECONDS : APP.TEACHER_TTL_SECONDS;
  cache.put(key, raw, ttl);
  return session;
}

function loginRateKeys_(scope, request) {
  const identity = String(request.client_id || request.client_key || "").trim() || "anonymous";
  return [
    { key: "login-rate:" + scope + ":client:" + identity.slice(0, 80), max: APP.LOGIN_RATE_MAX_ATTEMPTS },
    { key: "login-rate:" + scope + ":global", max: APP.LOGIN_RATE_GLOBAL_MAX_ATTEMPTS },
  ];
}

function assertLoginAllowed_(rateKeys) {
  const cache = CacheService.getScriptCache();
  rateKeys.forEach((item) => {
    const attempts = Number(cache.get(item.key) || 0);
    if (attempts >= item.max) {
      throw apiError_("LOGIN_RATE_LIMIT", "로그인 시도가 너무 많습니다. 5분 뒤 다시 시도하세요.");
    }
  });
}

function recordLoginFailure_(rateKeys) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  const cache = CacheService.getScriptCache();
  try {
    rateKeys.forEach((item) => {
      const attempts = Number(cache.get(item.key) || 0) + 1;
      cache.put(item.key, String(attempts), APP.LOGIN_RATE_TTL_SECONDS);
    });
  } finally {
    lock.releaseLock();
  }
}

function clearLoginFailures_(rateKeys) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  const cache = CacheService.getScriptCache();
  try {
    rateKeys.forEach((item) => cache.remove(item.key));
  } finally {
    lock.releaseLock();
  }
}

function loggedApiError_(role, action, code, message, details) {
  const error = apiError_(code, message, details);
  safeAppendAudit_(role, action, "session", "", 0, false, code);
  error.auditLogged = true;
  return error;
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

function apiError_(code, message, details) {
  const error = new Error(message);
  error.code = code;
  error.details = details || null;
  return error;
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
