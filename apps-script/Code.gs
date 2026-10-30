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
  const passwordPrompt = ui.prompt("초기 설정", "관리자 암호를 4자 이상 입력하세요.", ui.ButtonSet.OK_CANCEL);
  if (passwordPrompt.getSelectedButton() !== ui.Button.OK) return;

  const schoolName = String(schoolNamePrompt.getResponseText() || "").trim();
  const password = String(passwordPrompt.getResponseText() || "");
  if (!schoolName || password.length < 4) {
    ui.alert("학교명을 입력하고 관리자 암호를 4자 이상으로 설정하세요.");
    return;
  }

  const properties = PropertiesService.getScriptProperties();
  const passwordSalt = randomToken_();
  const nextAdminVersion = Number(properties.getProperty(APP.PROPERTY_ADMIN_SESSION_VERSION) || 0) + 1;
  properties.setProperties({
    [APP.PROPERTY_SHEET_ID]: SpreadsheetApp.getActiveSpreadsheet().getId(),
    [APP.PROPERTY_ACCESS_USER_PEPPER]: properties.getProperty(APP.PROPERTY_ACCESS_USER_PEPPER) || randomToken_() + randomToken_(),
    [APP.PROPERTY_ADMIN_SALT]: passwordSalt,
    [APP.PROPERTY_ADMIN_HASH]: hashSecret_(password, passwordSalt),
    [APP.PROPERTY_ADMIN_SESSION_VERSION]: String(nextAdminVersion),
  }, false);
  properties.deleteProperty(["SCHOOL", "CODE", "SALT"].join("_"));
  properties.deleteProperty(["SCHOOL", "CODE", "HASH"].join("_"));

  setSetting_("school_name", schoolName);
  setSetting_("app_name", APP.SYSTEM_NAME);
  setSetting_("timezone", APP.TIMEZONE);
  setSetting_("version", APP.VERSION);
  setSetting_("schema_version", APP.SCHEMA_VERSION);
  safeAppendAudit_("admin", "initialize", "system", "", 1, true, "OK", "system_admin");
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
  if (password.length < 4) {
    ui.alert("관리자 암호를 4자 이상으로 설정하세요.");
    return;
  }
  const salt = randomToken_();
  const nextVersion = Number(properties.getProperty(APP.PROPERTY_ADMIN_SESSION_VERSION) || 1) + 1;
  properties.setProperty(APP.PROPERTY_ADMIN_SALT, salt);
  properties.setProperty(APP.PROPERTY_ADMIN_HASH, hashSecret_(password, salt));
  properties.setProperty(APP.PROPERTY_ADMIN_SESSION_VERSION, String(nextVersion));
  safeAppendAudit_("admin", "change_password", "system", "", 1, true, "OK", "system_admin");
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
  // A public health check must never migrate sheets, read students, or run synchronization.
  return jsonOutput_({ ok: true, data: { service: "exam-management", version: APP.VERSION }, error: null });
}

function doPost(event) {
  let request = {};
  try {
    const body = event && event.postData ? event.postData.contents : "{}";
    if (typeof body !== "string" || body.length > 2000000) throw apiError_("INVALID_REQUEST", "요청 형식이나 크기를 확인하세요.");
    const parsed = JSON.parse(body);
    if (!parsed || Array.isArray(parsed) || typeof parsed !== "object" || typeof parsed.action !== "string" || parsed.action.length > 60) throw apiError_("INVALID_REQUEST", "올바른 작업 요청을 보내세요.");
    request = parsed;
    const data = dispatch_(request);
    return jsonOutput_({ ok: true, data: data, error: null });
  } catch (error) {
    if (!error.auditLogged && request.action !== "presence_ping") {
      try {
        auditRejectedRequest_(error.public_error ? error.code : "REQUEST_FAILED");
        error.auditLogged = true;
      } catch (ignored) {}
    }
    return jsonOutput_({
      ok: false,
      data: null,
      error: {
        code: error.public_error ? error.code : "REQUEST_FAILED",
        message: error.public_error ? error.message : "요청을 처리하지 못했습니다. 잠시 후 다시 시도하거나 관리자에게 문의하세요.",
        details: error.public_error ? error.details || null : null,
      },
    });
  }
}

function auditRejectedRequest_(code) {
  const lock = LockService.getScriptLock(), alreadyLocked = lock.hasLock();
  if (!alreadyLocked && !lock.tryLock(50)) return;
  try {
    const cache = CacheService.getScriptCache(), key = "rejected-request-audit-minute:" + Math.floor(Date.now() / 60000);
    if (!cache.get(key)) { cache.put(key, "1", 60); appendAudit_("unknown", "request_rejected", "api", "", 0, false, code); }
  } finally { if (!alreadyLocked) lock.releaseLock(); }
}

function ensureSchemaOnRequest_() {
  const cache = CacheService.getScriptCache();
  const marker = "schema-ready:" + APP.SCHEMA_VERSION;
  if (cache.get(marker)) return;
  const lock = LockService.getScriptLock();
  const alreadyLocked = lock.hasLock();
  if (!alreadyLocked && !lock.tryLock(5000)) throw apiError_("BUSY", "초기 데이터 구조를 확인하는 중입니다. 잠시 후 다시 시도하세요.");
  try {
    if (!cache.get(marker)) {
      ensureSheets_();
      cache.put(marker, "1", APP.DATA_CACHE_TTL_SECONDS);
    }
  } finally {
    if (!alreadyLocked) lock.releaseLock();
  }
}

function dispatch_(request) {
  const action = String(request.action || "");
  if (!action) throw apiError_("INVALID_ACTION", "작업 이름이 없습니다.");
  if (action === "user_login") return userLogin_(request);
  if (action === "admin_entry_login") return adminEntryLogin_(request);
  if (action === "redeem_exam_packaging_invite") return redeemExamPackagingInvite_(request);
  if (action === "logout") return logout_(request);

  if (action === "presence_ping") {
    if (request.packaging_session) {
      requirePackagingSession_(request.packaging_session, request.packaging_active === true);
      return presencePing_(request, "packaging");
    }
    requireSession_(request.user_session, "user", request.user_active === true);
    return presencePing_(request, request.scope === "exam_packaging" ? "packaging" : "user");
  }

  if (request.packaging_session) {
    const packagingSession = requirePackagingSession_(
      request.packaging_session,
      action === "save_exam_packaging_items"
    );
    if (action === "get_exam_packaging") return withDataRead_(() => getExamPackaging_(request, packagingSession));
    if (action === "save_exam_packaging_items") {
      ensureSchemaOnRequest_();
      return withPackagingMutation_(request, packagingSession, () => saveExamPackagingItems_(request, packagingSession));
    }
    throw apiError_("FORBIDDEN", "비밀 편집 링크는 고사 원안 포장 기능에만 사용할 수 있습니다.");
  }

  const user = requireSession_(request.user_session, "user");
  if (action === "admin_login") return adminLogin_(request, user);
  if (ADMIN_ACTIONS.has(action)) {
    const admin = requireSession_(request.admin_session, "admin");
    if (String(admin.actor_id || "") !== String(user.user_id || "")) {
      throw apiError_("FORBIDDEN", "관리자 세션과 사용자 세션이 일치하지 않습니다.");
    }
  }
  if (action !== "sync_academic_data") ensureSchemaOnRequest_();
  const needsCurrent = ["get_absence_context", "submit_absence_input", "get_hq_status", "get_period_report", "save_enrollments", "save_seat_chart", "save_seat_charts_batch"].includes(action);

  if (["save_class", "delete_class", "replace_students", "import_students", "save_student", "delete_student"].includes(action) && academicSourceConfig_().spreadsheet_id) {
    throw apiError_("ACADEMIC_MANAGED_RECORD", "연결된 학급·학생 명단은 학적 통합에서 수정한 뒤 동기화하세요.");
  }

  if (MUTATION_ACTIONS.has(action)) {
    return withMutation_(request, ADMIN_ACTIONS.has(action) ? "admin" : "user", user.user_id, () => { if (needsCurrent) assertAcademicCurrent_(); return route_(action, request); });
  }
  if (["get_bootstrap", "get_admin_bootstrap", "sync_academic_data"].includes(action)) return route_(action, request, user);
  return withDataRead_(() => { if (needsCurrent) assertAcademicCurrent_(); return route_(action, request, user); });
}

function route_(action, request, userSession) {
  switch (action) {
    case "get_bootstrap": syncAcademicData_(false); return bootstrapForUser_(false, userSession);
    case "get_admin_bootstrap": syncAcademicData_(false); return bootstrapForUser_(true, userSession);
    case "sync_academic_data":
      if (!request.request_id) throw apiError_("REQUEST_ID_REQUIRED", "동기화 요청 식별자가 없습니다.");
      syncAcademicData_(true, request.request_id, userSession.user_id); return bootstrapForUser_(true, userSession);
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
    case "save_subject_catalog": return saveSubjectCatalog_(request);
    case "save_exam_notice_settings": return saveExamNoticeSettings_(request);
    case "get_exam_packaging": return getExamPackaging_(request, userSession);
    case "save_exam_packaging_items": return saveExamPackagingItems_(request, userSession);
    case "save_exam_packaging_config": return saveExamPackagingConfig_(request);
    case "create_exam_packaging_invite": return createExamPackagingInvite_(request);
    case "disable_exam_packaging_invite": return disableExamPackagingInvite_();
    case "save_access_user": return saveAccessUser_(request);
    case "set_access_user_active": return setAccessUserActive_(request);
    case "reset_access_user_pin": return resetAccessUserPin_(request);
    case "save_enrollments": return saveEnrollments_(request);
    case "save_seat_chart": return saveSeatChart_(request);
    case "save_seat_charts_batch": return saveSeatChartsBatch_(request);
    case "delete_seat_chart": return deleteSeatChart_(request);
    case "cleanup": return cleanup_(request);
    default: throw apiError_("INVALID_ACTION", "지원하지 않는 작업입니다: " + action);
  }
}

function userLogin_(request) {
  const loginName = canonicalLoginName_(request.login_name).slice(0, 100);
  const pin = String(request.pin || "").trim();
  const rateKeys = loginRateKeys_("user", request, loginName);
  assertLoginAllowed_(rateKeys);
  const user = loginName && /^\d{4,6}$/.test(pin) ? findAccessUserByLogin_(loginName) : null;
  const comparisonSalt = user ? user.pin_salt : "invalid-login-salt";
  const expectedHash = user ? user.pin_hash : "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
  const actualHash = hashAccessPin_(pin, comparisonSalt);
  const valid = Boolean(user && user.active && constantTimeEqual_(actualHash, expectedHash));
  if (!valid) {
    recordLoginFailure_(rateKeys);
    throw loggedApiError_("user", "user_login", "INVALID_LOGIN", "접속 이름 또는 개인 PIN이 올바르지 않습니다.");
  }
  clearLoginFailures_(rateKeys);
  ensureSchemaOnRequest_();
  syncAcademicData_(false);
  const token = createUserSession_(user);
  const currentUser = safeAccessUser_(user);
  safeAppendAudit_("user", "user_login", "session", "", 1, true, "OK", user.id);
  return { user_session: token, current_user: currentUser, bootstrap: withDataRead_(() => bootstrapForCurrentUser_(teacherBootstrap_(), currentUser)) };
}

function verifyAdminPassword_(request, action) {
  const password = String(request.password || "");
  const rateKeys = loginRateKeys_("admin", request);
  assertLoginAllowed_(rateKeys);
  const properties = PropertiesService.getScriptProperties();
  const salt = properties.getProperty(APP.PROPERTY_ADMIN_SALT);
  const expected = properties.getProperty(APP.PROPERTY_ADMIN_HASH);
  const actual = hashSecret_(password, salt || "");
  if (!salt || !expected || !constantTimeEqual_(actual, expected)) {
    recordLoginFailure_(rateKeys);
    throw loggedApiError_("admin", action, "INVALID_ADMIN_PASSWORD", "관리자 암호가 올바르지 않습니다.");
  }
  clearLoginFailures_(rateKeys);
}

function adminEntryLogin_(request) {
  runAdminLoginPhase_("verify", () => verifyAdminPassword_(request, "admin_entry_login"));
  runAdminLoginPhase_("recovery", () => {
    if (PropertiesService.getScriptProperties().getProperty("ACADEMIC_SYNC_JOURNAL")) syncAcademicData_(true, createId_("recovery"), "system_admin");
  });
  runAdminLoginPhase_("schema", () => ensureSchemaOnRequest_());
  runAdminLoginPhase_("sync", () => syncAcademicData_(false));
  const currentUser = { id: "system_admin", login_name: "관리자", active: true, system: true };
  const userToken = runAdminLoginPhase_("user_session", () => createUserSession_(currentUser, true));
  const adminToken = runAdminLoginPhase_("admin_session", () => createAdminSession_(currentUser.id));
  safeAppendAudit_("admin", "admin_entry_login", "session", "", 1, true, "OK", currentUser.id);
  return {
    user_session: userToken,
    admin_session: adminToken,
    current_user: currentUser,
    bootstrap: runAdminLoginPhase_("bootstrap", () => withDataRead_(() => bootstrapForCurrentUser_(adminBootstrap_(), currentUser))),
  };
}

function adminLogin_(request, userSession) {
  runAdminLoginPhase_("verify", () => verifyAdminPassword_(request, "admin_login"));
  const token = runAdminLoginPhase_("admin_session", () => createAdminSession_(userSession.user_id));
  safeAppendAudit_("admin", "admin_login", "session", "", 1, true, "OK", userSession.user_id);
  return { admin_session: token, bootstrap: runAdminLoginPhase_("bootstrap", () => bootstrapForUser_(true, userSession)) };
}

function runAdminLoginPhase_(phase, operation) {
  try { return operation(); }
  catch (error) {
    // Preserve deliberate public errors (wrong password/rate limit/recovery); never expose
    // raw exception text, stack, workbook IDs, access-user records, passwords or PINs.
    if (error && error.public_error === true) throw error;
    const phases = {
      verify: ["A1", "암호 확인"], recovery: ["A2", "복구 상태 확인"],
      schema: ["A3", "초기 설정 확인"], sync: ["A4", "학적 연결 확인"],
      user_session: ["A5", "사용자 세션 발급"], admin_session: ["A6", "관리자 세션 발급"],
      bootstrap: ["A7", "관리자 화면 조회"],
    };
    const selected = Object.prototype.hasOwnProperty.call(phases, phase) ? phases[phase] : ["A0", "로그인 처리"];
    const names = { TypeError: "T", ReferenceError: "R", RangeError: "N", SyntaxError: "S", Exception: "G" };
    const errorName = String(error && error.name || "");
    const kind = Object.prototype.hasOwnProperty.call(names, errorName) ? names[errorName] : "E";
    const stack = String(error && error.stack || ""), frames = /\bat\s+([A-Za-z_$][\w$]*)\s*\(/g;
    const allowed = { ensureSheets_: "SI", ensureTableSheet_: "TH", getSpreadsheet_: "WB", getRecordsFresh_: "RR", getRecords_: "RC", withDataRead_: "LK", adminBootstrap_: "AB", teacherBootstrap_: "TB", academicSourceConfig_: "AC", academicLinkStatus_: "AS", normalizeStoredExamPackagingConfig_: "PC", getExamPackagingConfig_: "PG", readLoginRateState_: "LR", writeLoginRateState_: "LW", hashSecret_: "HS" };
    let location = "X";
    let frame;
    while ((frame = frames.exec(stack)) !== null) { if (Object.prototype.hasOwnProperty.call(allowed, frame[1])) { location = allowed[frame[1]]; break; } }
    const diagnostic = selected[0] + "-" + kind + "-" + location;
    throw apiError_("ADMIN_ENTRY_FAILED", "관리자 " + selected[1] + " 중 내부 오류가 발생했습니다. 진단번호 " + diagnostic + "를 알려주세요.", { diagnostic: diagnostic });
  }
}

function currentUserFromSession_(session) {
  if (session && session.system_admin === true && session.user_id === "system_admin") {
    return { id: "system_admin", login_name: "관리자", active: true, system: true };
  }
  const user = session ? findAccessUserById_(session.user_id) : null;
  if (!user || !user.active) throw apiError_("SESSION_EXPIRED", "접속 권한이 변경되었습니다. 다시 로그인하세요.");
  return safeAccessUser_(user);
}

function bootstrapForCurrentUser_(bootstrap, currentUser) {
  return Object.assign({}, bootstrap, { current_user: currentUser });
}

function bootstrapForUser_(admin, session) {
  return withDataRead_(() => bootstrapForCurrentUser_(admin ? adminBootstrap_() : teacherBootstrap_(), currentUserFromSession_(session)));
}

function logout_(request) {
  const cache = CacheService.getScriptCache();
  leavePresence_(request.user_session, "user");
  leavePresence_(request.user_session, "packaging");
  leavePresence_(request.packaging_session, "packaging");
  if (request.user_session) cache.remove("session:" + request.user_session);
  if (request.admin_session) cache.remove("session:" + request.admin_session);
  if (request.packaging_session) cache.remove("session:" + request.packaging_session);
  return { logged_out: true };
}

function presenceIdentity_(userSession) {
  const token = String(userSession || "").trim();
  if (!token) return "";
  return hashSecret_(token, "presence-v1").slice(0, 43);
}

function readPresenceRegistry_(cache, registryKey) {
  try {
    const parsed = JSON.parse(cache.get(registryKey || "presence:registry:v1") || "{}");
    if (!parsed || Array.isArray(parsed) || typeof parsed !== "object") return {};
    return Object.keys(parsed).reduce((result, key) => {
      const lastSeen = Number(parsed[key]);
      if (/^[A-Za-z0-9_-]{20,50}$/.test(key) && Number.isFinite(lastSeen)) result[key] = lastSeen;
      return result;
    }, {});
  } catch (ignored) {
    return {};
  }
}

function prunePresenceRegistry_(registry, nowMs) {
  const cutoff = nowMs - APP.PRESENCE_WINDOW_SECONDS * 1000;
  return Object.fromEntries(
    Object.entries(registry)
      .filter((entry) => Number(entry[1]) >= cutoff)
      .sort((left, right) => Number(right[1]) - Number(left[1]))
      .slice(0, APP.PRESENCE_MAX_ENTRIES)
  );
}

function presenceResponse_(registry, ownIdentity, stale) {
  return {
    connected: true,
    online_count: Object.keys(registry).filter((key) => key !== ownIdentity).length,
    observed_at: nowIso_(),
    active_window_seconds: APP.PRESENCE_WINDOW_SECONDS,
    approximate: true,
    stale: Boolean(stale),
  };
}

function presencePing_(request, scope) {
  const presenceScope = scope === "packaging" ? "packaging" : "user";
  const ownIdentity = presenceIdentity_(presenceScope === "packaging"
    ? (request.packaging_session || request.user_session)
    : request.user_session);
  if (!ownIdentity) throw apiError_("SESSION_REQUIRED", "로그인이 필요합니다.");
  const cache = CacheService.getScriptCache();
  const registryKey = presenceScope === "packaging" ? "presence:registry:packaging:v1" : "presence:registry:v1";
  const nowMs = Date.now();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(APP.PRESENCE_LOCK_WAIT_MS)) {
    return presenceResponse_(prunePresenceRegistry_(readPresenceRegistry_(cache, registryKey), nowMs), ownIdentity, true);
  }
  try {
    const registry = readPresenceRegistry_(cache, registryKey);
    registry[ownIdentity] = nowMs;
    const active = prunePresenceRegistry_(registry, nowMs);
    cache.put(registryKey, JSON.stringify(active), APP.PRESENCE_CACHE_TTL_SECONDS);
    return presenceResponse_(active, ownIdentity, false);
  } finally {
    lock.releaseLock();
  }
}

function leavePresence_(sessionToken, scope) {
  const ownIdentity = presenceIdentity_(sessionToken);
  if (!ownIdentity) return;
  const cache = CacheService.getScriptCache();
  const registryKey = scope === "packaging" ? "presence:registry:packaging:v1" : "presence:registry:v1";
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(APP.PRESENCE_LOCK_WAIT_MS)) return;
  try {
    const registry = readPresenceRegistry_(cache, registryKey);
    delete registry[ownIdentity];
    const active = prunePresenceRegistry_(registry, Date.now());
    if (Object.keys(active).length) {
      cache.put(registryKey, JSON.stringify(active), APP.PRESENCE_CACHE_TTL_SECONDS);
    } else {
      cache.remove(registryKey);
    }
  } finally {
    lock.releaseLock();
  }
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
      student_name_snapshot: validStudents.get(studentId).name,
      student_number_snapshot: validStudents.get(studentId).number,
      grade_snapshot: (getRecords_(SHEETS.CLASSES).find((row) => String(row.id) === String(request.class_id)) || {}).grade || "",
      class_num_snapshot: (getRecords_(SHEETS.CLASSES).find((row) => String(row.id) === String(request.class_id)) || {}).class_num || "",
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
  const timetable = getRecords_(SHEETS.TIMETABLE).filter((item) => item.exam_date === examDate);
  const packagingItems = timetable.filter(hasExamPackagingData_).length;
  return {
    timetable: timetable.length,
    packaging_items: packagingItems,
    exam_packaging: packagingItems,
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
  const relatedTimetable = getRecords_(SHEETS.TIMETABLE).filter((item) => normalizeClassIds_(item.class_ids).indexOf(String(classId)) !== -1);
  const packagingItems = relatedTimetable.filter((item) => normalizeClassIds_(item.class_ids).length === 1 && hasExamPackagingData_(item)).length;
  return {
    students: studentIds.size,
    absences: getRecords_(SHEETS.ABSENCES).filter((item) => String(item.class_id) === String(classId)).length,
    completions: getRecords_(SHEETS.COMPLETIONS).filter((item) => String(item.class_id) === String(classId)).length,
    enrollments: getRecords_(SHEETS.ENROLLMENTS).filter((item) => String(item.class_id) === String(classId) || studentIds.has(String(item.student_id))).length,
    timetable: relatedTimetable.length,
    packaging_items: packagingItems,
    exam_packaging: packagingItems,
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
  const oldById = new Map(oldRows.map((item) => [String(item.id), item]));
  const oldByNatural = new Map(oldRows.map((item) => [[item.exam_date, Number(item.grade), Number(item.period), item.subject_type, item.subject_name].join("|"), item]));
  const now = nowIso_();
  const seen = new Set();
  const subjectSeen = new Set();
  const linkedSubjects = getRecords_(SHEETS.ENROLLMENTS).filter((row) => row.source_academic_year);
  const rows = request.timetable.map((item) => {
    const examDate = String(item.exam_date || "");
    const grade = Number(item.grade);
    const period = Number(item.period);
    let subjectName = String(item.subject_name || "").trim().slice(0, 80);
    const subjectType = item.subject_type === "elective" ? "elective" : "common";
    if (subjectType === "elective") {
      const aliases = Array.from(new Set(linkedSubjects.filter((row) => Number(row.grade) === grade && academicNameKey_(row.subject_name) === academicNameKey_(subjectName)).map((row) => String(row.subject_name))));
      if (aliases.length > 1) throw apiError_("ACADEMIC_SUBJECT_AMBIGUOUS", "학적 수강 과목의 표기를 확인하세요.");
      if (aliases.length === 1) subjectName = aliases[0];
    }
    const startTime = String(item.start_time || "").trim(), endTime = String(item.end_time || "").trim();
    if ((startTime || endTime) && (!isValidClockTime_(startTime) || !isValidClockTime_(endTime) || startTime >= endTime)) throw apiError_("INVALID_TIMETABLE_TIME", "시작·종료 시각을 입력하고 종료 시각을 시작 시각보다 늦게 설정하세요.");
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
    const preserved = (item.id ? oldById.get(String(item.id)) : null) || oldByNatural.get(naturalKey) || null;
    return {
      id: item.id || (preserved && preserved.id) || createId_("timetable"),
      exam_date: examDate,
      grade: grade,
      period: period,
      start_time: startTime,
      end_time: endTime,
      subject_name: subjectName,
      subject_type: subjectType,
      room_name: String(item.room_name || "").trim().slice(0, 80),
      class_ids: JSON.stringify(classIds),
      updated_at: now,
      representative_teacher: String(preserved && preserved.representative_teacher || "").trim().slice(0, 100),
      answer_sheet_type: normalizeAnswerSheetType_(preserved && preserved.answer_sheet_type),
      packaging_date: String(preserved && preserved.packaging_date || "").trim(),
      packaging_slot_id: String(preserved && preserved.packaging_slot_id || "").trim().slice(0, 80),
      packaging_revision: Math.max(0, Number(preserved && preserved.packaging_revision) || 0),
      packaging_updated_at: String(preserved && preserved.packaging_updated_at || ""),
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

function normalizeSubjectCatalogInput_(gradeValue, subjectsValue) {
  const grade = Number(gradeValue);
  if (!Number.isInteger(grade) || grade < 1 || grade > 9) {
    throw apiError_("INVALID_GRADE", "학년을 확인하세요.");
  }
  if (!Array.isArray(subjectsValue)) {
    throw apiError_("INVALID_SUBJECTS", "교과목 목록 형식이 올바르지 않습니다.");
  }
  if (subjectsValue.length > 200) {
    throw apiError_("TOO_MANY_SUBJECTS", "교과목은 한 학년에 200개까지 등록할 수 있습니다.");
  }

  const seen = new Set();
  const subjects = subjectsValue.map((value, index) => {
    const details = { row_number: index + 2 };
    if (typeof value !== "string") {
      throw apiError_("INVALID_SUBJECT_NAME", "교과목 이름 형식이 올바르지 않습니다.", details);
    }
    const subject = value.trim();
    if (!subject || subject.length > 80) {
      throw apiError_("INVALID_SUBJECT_NAME", "교과목 이름은 1자 이상 80자 이하로 입력하세요.", details);
    }
    if (seen.has(subject)) {
      throw apiError_("DUPLICATE_SUBJECT", "중복된 교과목이 있습니다: " + subject, details);
    }
    seen.add(subject);
    return subject;
  });
  return { grade: grade, subjects: subjects };
}

function saveSubjectCatalog_(request) {
  const normalized = normalizeSubjectCatalogInput_(request.grade, request.subjects);
  const settings = getSettings_();
  const catalog = normalizeStoredSubjectCatalog_(settings.subject_catalog_json);
  catalog[String(normalized.grade)] = normalized.subjects;
  setSetting_("subject_catalog_json", JSON.stringify(catalog));
  return adminBootstrap_();
}

function normalizeExamNoticeInput_(noticeValue) {
  if (!noticeValue || Array.isArray(noticeValue) || typeof noticeValue !== "object") {
    throw apiError_("INVALID_EXAM_NOTICE", "가정통신문 설정 형식이 올바르지 않습니다.");
  }

  const year = Number(noticeValue.year);
  if (!Number.isInteger(year) || year < 1900 || year > 2200) {
    throw apiError_("INVALID_ACADEMIC_YEAR", "학년도를 확인하세요.");
  }
  const semester = Number(noticeValue.semester);
  if (semester !== 1 && semester !== 2) {
    throw apiError_("INVALID_SEMESTER", "학기는 1학기 또는 2학기로 입력하세요.");
  }

  const examName = String(noticeValue.exam_name || "").trim();
  if (!examName) throw apiError_("EXAM_NAME_REQUIRED", "고사명을 입력하세요.");
  if (examName.length > 80) throw apiError_("INVALID_EXAM_NAME", "고사명은 80자 이하로 입력하세요.");

  const greeting = String(noticeValue.greeting || "").trim();
  if (greeting.length > 1500) throw apiError_("INVALID_GREETING", "인사말은 1,500자 이하로 입력하세요.");

  if (!Array.isArray(noticeValue.notes) || noticeValue.notes.length > 20) {
    throw apiError_("INVALID_NOTICE_NOTES", "안내사항은 20개 이하의 목록으로 입력하세요.");
  }
  const notes = noticeValue.notes.map((value, index) => {
    const details = { item_number: index + 1 };
    if (typeof value !== "string") {
      throw apiError_("INVALID_NOTICE_NOTE", "안내사항 형식이 올바르지 않습니다.", details);
    }
    const note = value.trim();
    if (!note || note.length > 200) {
      throw apiError_("INVALID_NOTICE_NOTE", "안내사항은 항목마다 1자 이상 200자 이하로 입력하세요.", details);
    }
    return note;
  });

  const issueDate = String(noticeValue.issue_date || "").trim();
  if (issueDate && !isValidIsoDate_(issueDate)) {
    throw apiError_("INVALID_ISSUE_DATE", "발행일을 확인하세요.");
  }

  const issuer = String(noticeValue.issuer || "").trim();
  if (!issuer) throw apiError_("NOTICE_ISSUER_REQUIRED", "발행 명의를 입력하세요.");
  if (issuer.length > 100) throw apiError_("INVALID_NOTICE_ISSUER", "발행 명의는 100자 이하로 입력하세요.");

  return {
    year: year,
    semester: semester,
    exam_name: examName,
    greeting: greeting,
    notes: notes,
    issue_date: issueDate,
    issuer: issuer,
  };
}

function saveExamNoticeSettings_(request) {
  const notice = normalizeExamNoticeInput_(request.notice);
  setSetting_("exam_notice_json", JSON.stringify(notice));
  return adminBootstrap_();
}

function canonicalLoginName_(value) {
  let name = String(value == null ? "" : value);
  if (typeof name.normalize === "function") name = name.normalize("NFKC");
  return name.replace(/\s+/g, " ").trim();
}

function loginLookupKey_(value) {
  return canonicalLoginName_(value).toLowerCase();
}

function validateLoginName_(value) {
  const name = canonicalLoginName_(value);
  if (!name || name.length > 50 || /[\u0000-\u001f\u007f]/.test(name)) {
    throw apiError_("INVALID_LOGIN_NAME", "접속 이름은 1자 이상 50자 이하로 입력하세요.");
  }
  return name;
}

function validatePin_(value) {
  const pin = String(value == null ? "" : value).trim();
  if (!/^\d{4,6}$/.test(pin)) throw apiError_("INVALID_PIN", "개인 PIN은 숫자 4~6자리로 입력하세요.");
  return pin;
}

function hashAccessPin_(pin, salt) {
  const pepper = PropertiesService.getScriptProperties().getProperty(APP.PROPERTY_ACCESS_USER_PEPPER);
  if (!pepper) throw apiError_("NOT_INITIALIZED", "초기 설정을 다시 실행하세요.");
  return hashSecret_(String(pin) + "\u001f" + pepper, salt);
}

function createAccessUserId_() {
  return Utilities.getUuid();
}

function assertUniqueLoginName_(users, loginName, exceptId) {
  const lookup = loginLookupKey_(loginName);
  if (users.some((item) => item.id !== exceptId && item.normalized_login_name === lookup)) {
    throw apiError_("DUPLICATE_LOGIN_NAME", "이미 등록된 접속 이름입니다.");
  }
}

function saveAccessUser_(request) {
  const loginName = validateLoginName_(request.login_name);
  const users = getAccessUsers_();
  const id = String(request.id || "").trim();
  const existing = id ? users.find((item) => item.id === id) : null;
  if (id && !existing) throw apiError_("ACCESS_USER_NOT_FOUND", "접속 사용자를 찾지 못했습니다.");
  if (!existing && users.length >= 300) throw apiError_("TOO_MANY_ACCESS_USERS", "접속 사용자는 300명까지 등록할 수 있습니다.");
  assertUniqueLoginName_(users, loginName, existing ? existing.id : "");

  const now = nowIso_();
  let user = existing ? Object.assign({}, existing) : {
    id: createAccessUserId_(),
    active: true,
    session_version: 1,
    created_at: now,
  };
  const nameChanged = Boolean(existing && existing.login_name !== loginName);
  user.login_name = loginName;
  user.normalized_login_name = loginLookupKey_(loginName);
  user.updated_at = now;

  const pinProvided = request.pin != null && String(request.pin).trim() !== "";
  if (!existing || pinProvided) {
    const pin = validatePin_(request.pin);
    user.pin_salt = randomToken_();
    user.pin_hash = hashAccessPin_(pin, user.pin_salt);
  }
  if (existing && (nameChanged || pinProvided)) user.session_version += 1;
  writeAccessUser_(user);
  return adminBootstrap_();
}

function setAccessUserActive_(request) {
  validateRequired_(request, ["id"]);
  if (request.active !== true && request.active !== false) {
    throw apiError_("INVALID_ACTIVE_STATE", "활성 상태를 확인하세요.");
  }
  const user = findAccessUserById_(request.id);
  if (!user) throw apiError_("ACCESS_USER_NOT_FOUND", "접속 사용자를 찾지 못했습니다.");
  if (user.active !== request.active) {
    user.active = request.active;
    user.session_version += 1;
    user.updated_at = nowIso_();
    writeAccessUser_(user);
  }
  return adminBootstrap_();
}

function resetAccessUserPin_(request) {
  validateRequired_(request, ["id"]);
  const pin = validatePin_(request.pin);
  const user = findAccessUserById_(request.id);
  if (!user) throw apiError_("ACCESS_USER_NOT_FOUND", "접속 사용자를 찾지 못했습니다.");
  user.pin_salt = randomToken_();
  user.pin_hash = hashAccessPin_(pin, user.pin_salt);
  user.session_version += 1;
  user.updated_at = nowIso_();
  writeAccessUser_(user);
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
  const packagingItems = hasExamPackagingData_(target) ? 1 : 0;
  return {
    enrollments: removeEnrollments
      ? getRecords_(SHEETS.ENROLLMENTS).filter((item) =>
          Number(item.grade) === Number(target.grade) &&
          String(item.subject_name) === String(target.subject_name) && !item.source_academic_year
        ).length
      : 0,
    seat_charts: getRecords_(SHEETS.SEATING).filter((item) =>
      item.exam_date === target.exam_date &&
      Number(item.period) === Number(target.period) &&
      Number(item.grade) === Number(target.grade) &&
      String(item.subject_name) === String(target.subject_name)
    ).length,
    packaging_items: packagingItems,
    exam_packaging: packagingItems,
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
      String(item.subject_name) === String(target.subject_name) && !item.source_academic_year
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
  let rows;
  if (academicSourceConfig_().spreadsheet_id) {
    const changes = new Map(validRows.map((row) => [[row.student_id, row.subject_name].join("|"), row]));
    const managed = new Set(oldRows.filter((row) => row.source_academic_year).map((row) => [row.student_id, row.subject_name].join("|")));
    if (validRows.some((row) => !managed.has([row.student_id, row.subject_name].join("|")))) throw apiError_("ACADEMIC_MANAGED_RECORD", "수강 과목은 학적 통합에서 수정하세요. 고사 통합에서는 기존 수강생의 시험 호실만 입력할 수 있습니다.");
    rows = oldRows.map((row) => {
      const change = changes.get([row.student_id, row.subject_name].join("|"));
      return change ? Object.assign({}, row, { room_name: change.room_name, updated_at: now }) : row;
    });
  } else rows = oldRows.filter((item) => Number(item.grade) !== grade).concat(validRows);
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
  const seatOrder = chart.seat_order == null || chart.seat_order === ""
    ? "row"
    : String(chart.seat_order);
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
  if (!new Set(["row", "column"]).has(seatOrder)) {
    throw apiError_("INVALID_SEAT_ORDER", "자리배치 순서를 확인하세요.");
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
  const assignedSequences = new Set();
  const capacity = rows * cols - disabled.size;
  assignment.forEach((item) => {
    const key = String(item && typeof item === "object" ? item.key || "" : "");
    if (!validSeatKey_(key, rows, cols)) throw apiError_("INVALID_ASSIGNMENT_SEAT", "범위를 벗어난 배치 좌석이 있습니다.");
    if (disabled.has(key)) throw apiError_("ASSIGNMENT_ON_DISABLED_SEAT", "제외 좌석에 학생을 배치할 수 없습니다.");
    if (assignedSeats.has(key)) throw apiError_("DUPLICATE_ASSIGNMENT_SEAT", "같은 좌석이 배치표에 두 번 들어 있습니다.");
    assignedSeats.add(key);
    if (item && item.sequence != null && item.sequence !== "") {
      const sequence = Number(item.sequence);
      if (!Number.isInteger(sequence) || sequence < 1 || sequence > capacity) {
        throw apiError_("INVALID_ASSIGNMENT_SEQUENCE", "배치 순번이 좌석 범위를 벗어났습니다.");
      }
      if (assignedSequences.has(sequence)) {
        throw apiError_("DUPLICATE_ASSIGNMENT_SEQUENCE", "같은 배치 순번이 두 번 들어 있습니다.");
      }
      assignedSequences.add(sequence);
    }
  });
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

function normalizeAnswerSheetType_(value) {
  const answer = String(value || "").trim();
  return answer === "card" || answer === "a4" ? answer : "";
}

function hasExamPackagingData_(item) {
  return Boolean(
    String(item && item.representative_teacher || "").trim() ||
    normalizeAnswerSheetType_(item && item.answer_sheet_type) ||
    String(item && item.packaging_date || "").trim() ||
    String(item && item.packaging_slot_id || "").trim()
  );
}

function serializeExamPackagingItem_(row) {
  return {
    id: String(row.id || ""),
    timetable_id: String(row.id || ""),
    exam_date: String(row.exam_date || ""),
    grade: Number(row.grade) || "",
    period: Number(row.period) || "",
    start_time: String(row.start_time || ""),
    end_time: String(row.end_time || ""),
    subject_name: String(row.subject_name || ""),
    subject_type: row.subject_type === "elective" ? "elective" : "common",
    representative_teacher: String(row.representative_teacher || "").trim().slice(0, 100),
    answer_sheet_type: normalizeAnswerSheetType_(row.answer_sheet_type),
    packaging_date: String(row.packaging_date || "").trim(),
    packaging_slot_id: String(row.packaging_slot_id || "").trim().slice(0, 80),
    revision: Math.max(0, Number(row.packaging_revision) || 0),
    updated_at: String(row.packaging_updated_at || ""),
  };
}

function compareExamPackagingItems_(left, right) {
  return String(left.exam_date).localeCompare(String(right.exam_date)) ||
    Number(left.period) - Number(right.period) ||
    Number(left.grade) - Number(right.grade) ||
    String(left.subject_name).localeCompare(String(right.subject_name));
}

function examPackagingSummary_(items) {
  return (items || []).reduce((result, item) => {
    result.total += 1;
    if (item.packaging_date && item.packaging_slot_id) result.assigned += 1;
    else result.unassigned += 1;
    if (item.answer_sheet_type) result.answer_selected += 1;
    else result.answer_unselected += 1;
    return result;
  }, { total: 0, assigned: 0, unassigned: 0, answer_selected: 0, answer_unselected: 0 });
}

function examPackagingTitle_(notice) {
  if (notice && notice.year && notice.semester && notice.exam_name) {
    return notice.year + "학년도 " + notice.semester + "학기 " + notice.exam_name;
  }
  return "고사 원안 포장";
}

function isPackagingDeadlinePassed_(config, nowMs) {
  if (!config || !config.input_deadline) return false;
  const deadline = new Date(config.input_deadline).getTime();
  return Number.isFinite(deadline) && deadline <= (nowMs == null ? Date.now() : Number(nowMs));
}

function getExamPackaging_(request, authSession) {
  const settings = getPublicSettings_();
  const privateSettings = getSettings_();
  const config = getExamPackagingConfig_();
  const examDates = getRecords_(SHEETS.EXAM_DATES).filter(isActive_).map((item) => ({
    id: String(item.id || ""),
    exam_date: String(item.exam_date || ""),
    label: String(item.label || ""),
  })).sort((left, right) => left.exam_date.localeCompare(right.exam_date));
  const activeExamDates = new Set(examDates.map((item) => item.exam_date));
  const items = getRecords_(SHEETS.TIMETABLE)
    .filter((item) => activeExamDates.has(String(item.exam_date || "")))
    .map(serializeExamPackagingItem_)
    .sort(compareExamPackagingItems_);
  const notice = normalizeStoredExamNotice_(privateSettings.exam_notice_json, settings.school_name);
  const linkSession = authSession && authSession.role === "packaging";
  return {
    settings: settings,
    exam_title: examPackagingTitle_(notice),
    exam_dates: examDates,
    items: items,
    config: config,
    summary: examPackagingSummary_(items),
    read_only: Boolean(linkSession && isPackagingDeadlinePassed_(config)),
    link_expires_at: linkSession ? String(authSession.invite_expires_at || "") : "",
  };
}

function examPackagingInviteStatus_() {
  const properties = PropertiesService.getScriptProperties();
  const enabled = properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_ENABLED) === "true";
  const expiresAt = String(properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_EXPIRES_AT) || "");
  const expired = Boolean(expiresAt && new Date(expiresAt).getTime() <= Date.now());
  return {
    active: Boolean(enabled && !expired),
    enabled: enabled,
    expires_at: expiresAt,
    expired: expired,
    version: Math.max(1, Number(properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_VERSION)) || 1),
    issued: Boolean(properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_HASH)),
  };
}

function hashExamPackagingInviteToken_(token, salt, pepper) {
  return hashSecret_(String(token || ""), String(salt || "") + "\u001f" + String(pepper || ""));
}

function createExamPackagingInvite_(request) {
  const expiresAt = String(request.expires_at || "").trim();
  if (!isValidDateTime_(expiresAt) || new Date(expiresAt).getTime() <= Date.now()) {
    throw apiError_("INVALID_PACKAGING_INVITE_EXPIRY", "비밀 링크 종료일시는 현재보다 뒤여야 합니다.");
  }
  const properties = PropertiesService.getScriptProperties();
  const pepper = properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_PEPPER) || randomToken_() + randomToken_();
  const salt = randomToken_();
  const token = randomSecret256_();
  const version = Math.max(1, Number(properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_VERSION)) || 1) + 1;
  properties.setProperties({
    [APP.PROPERTY_PACKAGING_INVITE_PEPPER]: pepper,
    [APP.PROPERTY_PACKAGING_INVITE_SALT]: salt,
    [APP.PROPERTY_PACKAGING_INVITE_HASH]: hashExamPackagingInviteToken_(token, salt, pepper),
    [APP.PROPERTY_PACKAGING_INVITE_VERSION]: String(version),
    [APP.PROPERTY_PACKAGING_INVITE_EXPIRES_AT]: expiresAt,
    [APP.PROPERTY_PACKAGING_INVITE_ENABLED]: "true",
  }, false);
  return { token: token, fragment: "#packaging=" + token, expires_at: expiresAt, invite_status: examPackagingInviteStatus_() };
}

function disableExamPackagingInvite_() {
  const properties = PropertiesService.getScriptProperties();
  const version = Math.max(1, Number(properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_VERSION)) || 1) + 1;
  properties.setProperty(APP.PROPERTY_PACKAGING_INVITE_ENABLED, "false");
  properties.setProperty(APP.PROPERTY_PACKAGING_INVITE_VERSION, String(version));
  return { disabled: true, invite_status: examPackagingInviteStatus_() };
}

function createExamPackagingSession_(inviteVersion, expiresAt) {
  const token = randomSecret256_();
  const session = {
    role: "packaging",
    session_id: createId_("packaging_session"),
    invite_version: Number(inviteVersion),
    invite_expires_at: String(expiresAt || ""),
    created_at: nowIso_(),
  };
  CacheService.getScriptCache().put("session:" + token, JSON.stringify(session), APP.PACKAGING_SESSION_TTL_SECONDS);
  return { token: token, session: session };
}

function packagingLinkError_(code, message) {
  const error = loggedApiError_("packaging", "redeem_exam_packaging_invite", code, message);
  return error;
}

function redeemExamPackagingInvite_(request) {
  const token = String(request.token || "").trim();
  const rateKeys = loginRateKeys_("packaging_invite", request);
  assertLoginAllowed_(rateKeys);
  const properties = PropertiesService.getScriptProperties();
  const salt = properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_SALT) || "invalid-packaging-salt";
  const pepper = properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_PEPPER) || "invalid-packaging-pepper";
  const expected = properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_HASH) || "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
  const actual = hashExamPackagingInviteToken_(token, salt, pepper);
  if (!/^[A-Za-z0-9_-]{43}$/.test(token) || !constantTimeEqual_(actual, expected)) {
    recordLoginFailure_(rateKeys);
    throw packagingLinkError_("INVALID_PACKAGING_LINK", "비밀 편집 링크가 올바르지 않습니다.");
  }
  if (properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_ENABLED) !== "true") {
    recordLoginFailure_(rateKeys);
    throw packagingLinkError_("PACKAGING_LINK_DISABLED", "폐기된 비밀 편집 링크입니다.");
  }
  const expiresAt = String(properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_EXPIRES_AT) || "");
  if (!isValidDateTime_(expiresAt) || new Date(expiresAt).getTime() <= Date.now()) {
    recordLoginFailure_(rateKeys);
    throw packagingLinkError_("PACKAGING_LINK_EXPIRED", "비밀 편집 링크의 사용 기간이 끝났습니다.");
  }
  clearLoginFailures_(rateKeys);
  const version = Math.max(1, Number(properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_VERSION)) || 1);
  const issued = createExamPackagingSession_(version, expiresAt);
  safeAppendAudit_("packaging", "redeem_exam_packaging_invite", "session", "", 1, true, "OK", issued.session.session_id);
  return {
    packaging_session: issued.token,
    context: getExamPackaging_(request, issued.session),
  };
}

function requirePackagingSession_(token, touch) {
  if (!token) throw apiError_("PACKAGING_SESSION_REQUIRED", "고사 원안 포장 링크를 다시 여세요.");
  const cache = CacheService.getScriptCache();
  const key = "session:" + String(token);
  const raw = cache.get(key);
  if (!raw) throw apiError_("PACKAGING_SESSION_EXPIRED", "30분 동안 사용하지 않아 편집 세션이 만료되었습니다. 비밀 링크를 다시 여세요.");
  let session;
  try {
    session = JSON.parse(raw);
  } catch (ignored) {
    cache.remove(key);
    throw apiError_("PACKAGING_SESSION_EXPIRED", "편집 세션이 만료되었습니다. 비밀 링크를 다시 여세요.");
  }
  if (!session || session.role !== "packaging") {
    cache.remove(key);
    throw apiError_("FORBIDDEN", "비밀 편집 링크는 고사 원안 포장 기능에만 사용할 수 있습니다.");
  }
  const properties = PropertiesService.getScriptProperties();
  const currentVersion = Math.max(1, Number(properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_VERSION)) || 1);
  if (properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_ENABLED) !== "true" || Number(session.invite_version) !== currentVersion) {
    cache.remove(key);
    throw apiError_("PACKAGING_SESSION_EXPIRED", "비밀 편집 링크가 폐기되거나 재발급되었습니다.");
  }
  const expiresAt = String(properties.getProperty(APP.PROPERTY_PACKAGING_INVITE_EXPIRES_AT) || session.invite_expires_at || "");
  if (!isValidDateTime_(expiresAt) || new Date(expiresAt).getTime() <= Date.now()) {
    cache.remove(key);
    throw apiError_("PACKAGING_LINK_EXPIRED", "비밀 편집 링크의 사용 기간이 끝났습니다.");
  }
  session.invite_expires_at = expiresAt;
  if (touch !== false) cache.put(key, JSON.stringify(session), APP.PACKAGING_SESSION_TTL_SECONDS);
  return session;
}

function normalizeExamPackagingConfigInput_(value) {
  const input = value && !Array.isArray(value) && typeof value === "object" ? value : null;
  if (!input) throw apiError_("INVALID_PACKAGING_CONFIG", "원안 포장 설정 형식이 올바르지 않습니다.");
  const capacity = Number(input.capacity);
  if (!Number.isInteger(capacity) || capacity < 1 || capacity > 20) {
    throw apiError_("INVALID_PACKAGING_CAPACITY", "슬롯 정원은 1명 이상 20명 이하로 입력하세요.");
  }
  const deadline = String(input.input_deadline || "").trim();
  if (deadline && !isValidDateTime_(deadline)) throw apiError_("INVALID_PACKAGING_DEADLINE", "입력 마감일시를 확인하세요.");
  if (!Array.isArray(input.packaging_dates) || input.packaging_dates.length > 31) {
    throw apiError_("INVALID_PACKAGING_DATES", "포장일 목록을 확인하세요.");
  }
  const dateSeen = new Set();
  const dates = input.packaging_dates.map((item) => {
    const date = String(item || "").trim();
    if (!isValidIsoDate_(date)) throw apiError_("INVALID_PACKAGING_DATE", "포장일 형식이 올바르지 않습니다.");
    if (dateSeen.has(date)) throw apiError_("DUPLICATE_PACKAGING_DATE", "같은 포장일이 두 번 들어 있습니다.");
    dateSeen.add(date);
    return date;
  }).sort();
  if (!Array.isArray(input.rows) || input.rows.length > 60) {
    throw apiError_("INVALID_PACKAGING_ROWS", "포장 시간대 목록을 확인하세요.");
  }
  const idSeen = new Set();
  const timeSeen = new Set();
  const rows = input.rows.map((item, index) => {
    const id = String(item && item.id || "").trim();
    const kind = item && item.kind === "break" ? "break" : item && item.kind === "slot" ? "slot" : "";
    const startTime = String(item && item.start_time || "").trim();
    const endTime = String(item && item.end_time || "").trim();
    if (!/^[A-Za-z0-9_-]{1,80}$/.test(id)) throw apiError_("INVALID_PACKAGING_SLOT_ID", "시간대 ID 형식이 올바르지 않습니다.", { row_number: index + 1 });
    if (idSeen.has(id)) throw apiError_("DUPLICATE_PACKAGING_SLOT", "같은 시간대가 두 번 들어 있습니다.", { row_number: index + 1 });
    if (!kind || !isValidClockTime_(startTime) || !isValidClockTime_(endTime) || startTime >= endTime) {
      throw apiError_("INVALID_PACKAGING_TIME", "시간대 시작·종료 시각을 확인하세요.", { row_number: index + 1 });
    }
    const startMinutes = Number(startTime.slice(0, 2)) * 60 + Number(startTime.slice(3, 5));
    const endMinutes = Number(endTime.slice(0, 2)) * 60 + Number(endTime.slice(3, 5));
    if (kind === "slot" && endMinutes - startMinutes !== 25) {
      throw apiError_("INVALID_PACKAGING_SLOT_DURATION", "과목을 배정하는 시간대는 25분이어야 합니다.", { row_number: index + 1 });
    }
    const timeKey = startTime + "|" + endTime;
    if (timeSeen.has(timeKey)) throw apiError_("DUPLICATE_PACKAGING_TIME", "같은 시작·종료 시간대가 두 번 들어 있습니다.", { row_number: index + 1 });
    idSeen.add(id);
    timeSeen.add(timeKey);
    return {
      id: id,
      kind: kind,
      period_label: String(item.period_label || "").trim().slice(0, 40),
      start_time: startTime,
      end_time: endTime,
      note: String(item.note || "").trim().slice(0, 100),
      sort_order: Number.isFinite(Number(item.sort_order)) ? Number(item.sort_order) : (index + 1) * 10,
    };
  }).sort((left, right) => left.sort_order - right.sort_order || left.start_time.localeCompare(right.start_time));
  const rowsByTime = rows.slice().sort((left, right) => left.start_time.localeCompare(right.start_time) || left.end_time.localeCompare(right.end_time));
  for (let index = 1; index < rowsByTime.length; index += 1) {
    if (rowsByTime[index].start_time < rowsByTime[index - 1].end_time) {
      throw apiError_("OVERLAPPING_PACKAGING_TIME", "서로 겹치는 포장 시간대가 있습니다.");
    }
  }
  const rowIds = new Set(rows.map((item) => item.id));
  if (!Array.isArray(input.staff_assignments) || input.staff_assignments.length > 1000) {
    throw apiError_("INVALID_PACKAGING_STAFF", "교무 담당자 목록을 확인하세요.");
  }
  const staffSeen = new Set();
  const staffAssignments = input.staff_assignments.map((item, index) => {
    const packagingDate = String(item && item.packaging_date || "").trim();
    const slotId = String(item && item.slot_id || "").trim();
    const staffNameRaw = String(item && item.staff_name || "").trim();
    const key = packagingDate + "|" + slotId;
    if (!dateSeen.has(packagingDate) || !rowIds.has(slotId)) {
      throw apiError_("INVALID_PACKAGING_STAFF_SLOT", "교무 담당자의 포장일과 시간대를 확인하세요.", { row_number: index + 1 });
    }
    if (staffSeen.has(key)) throw apiError_("DUPLICATE_PACKAGING_STAFF_SLOT", "같은 시간대의 교무 담당자가 두 번 들어 있습니다.", { row_number: index + 1 });
    if (staffNameRaw.length > 100) throw apiError_("INVALID_PACKAGING_STAFF_NAME", "교무 담당자 이름은 100자 이하여야 합니다.", { row_number: index + 1 });
    staffSeen.add(key);
    return { packaging_date: packagingDate, slot_id: slotId, staff_name: staffNameRaw };
  });
  return {
    revision: 0,
    input_deadline: deadline,
    capacity: capacity,
    packaging_dates: dates,
    rows: rows,
    staff_assignments: staffAssignments,
  };
}

function examPackagingConfigImpact_(config, timetable) {
  const dates = new Set(config.packaging_dates);
  const slots = new Set(config.rows.filter((item) => item.kind === "slot").map((item) => item.id));
  const affected = (timetable || []).filter((item) => {
    const date = String(item.packaging_date || "").trim();
    const slotId = String(item.packaging_slot_id || "").trim();
    return Boolean((date || slotId) && (!date || !slotId || !dates.has(date) || !slots.has(slotId)));
  });
  return {
    affected_count: affected.length,
    assigned_items: affected.length,
    timetable_ids: affected.map((item) => String(item.id)),
  };
}

function saveExamPackagingConfig_(request) {
  const current = getExamPackagingConfig_(true);
  const expectedRevision = Number(request.expected_revision);
  if (!Number.isInteger(expectedRevision) || expectedRevision !== Number(current.revision)) {
    throw apiError_("PACKAGING_CONFIG_CONFLICT", "다른 사용자가 원안 포장 설정을 변경했습니다. 최신 자료를 다시 불러오세요.", { current_revision: current.revision });
  }
  const next = normalizeExamPackagingConfigInput_(request.config);
  next.revision = current.revision + 1;
  const timetable = getRecordsFresh_(SHEETS.TIMETABLE);
  const impact = examPackagingConfigImpact_(next, timetable);
  if (request.preview_only) return { preview: true, impact: impact, config: next };
  if (impact.assigned_items && request.force !== true) {
    throw apiError_("PACKAGING_CONFIG_IN_USE", "사용 중인 포장일 또는 시간대가 있습니다. 영향 범위를 확인한 뒤 다시 저장하세요.", impact);
  }
  if (impact.assigned_items) {
    const affectedIds = new Set(impact.timetable_ids);
    const now = nowIso_();
    replaceRecords_(SHEETS.TIMETABLE, timetable.map((item) => affectedIds.has(String(item.id)) ? Object.assign({}, item, {
      packaging_date: "",
      packaging_slot_id: "",
      packaging_revision: Math.max(0, Number(item.packaging_revision) || 0) + 1,
      packaging_updated_at: now,
    }) : item));
  }
  setSetting_("exam_packaging_config_json", JSON.stringify(next));
  return Object.assign(adminBootstrap_(), { saved: true, impact: impact, exam_packaging_config: next });
}

function saveExamPackagingItems_(request, authSession) {
  if (!Array.isArray(request.items) || request.items.length > 1000) {
    throw apiError_("INVALID_PACKAGING_ITEMS", "원안 포장 항목 형식이 올바르지 않습니다.");
  }
  const config = getExamPackagingConfig_(true);
  const expectedConfigRevision = Number(request.expected_config_revision);
  if (!Number.isInteger(expectedConfigRevision) || expectedConfigRevision !== Number(config.revision)) {
    throw apiError_("PACKAGING_CONFIG_CONFLICT", "원안 포장 설정이 변경되었습니다. 최신 자료를 다시 불러오세요.", { current_revision: config.revision });
  }
  if (authSession && authSession.role === "packaging" && isPackagingDeadlinePassed_(config)) {
    throw apiError_("PACKAGING_READ_ONLY", "입력 마감이 지나 비밀 링크에서는 읽기만 할 수 있습니다.");
  }
  const packagingDates = new Set(config.packaging_dates);
  const slotRows = new Map(config.rows.map((item) => [String(item.id), item]));
  const timetable = getRecordsFresh_(SHEETS.TIMETABLE);
  const activeExamDates = new Set(
    getRecordsFresh_(SHEETS.EXAM_DATES)
      .filter(isActive_)
      .map((item) => String(item.exam_date || ""))
  );
  const byId = new Map(timetable.map((item) => [String(item.id), item]));
  const incomingIds = new Set();
  const changedIds = new Set();
  const changedItems = [];
  const now = nowIso_();
  request.items.forEach((incoming, index) => {
    const id = String(incoming && incoming.timetable_id || "").trim();
    if (!id || incomingIds.has(id)) throw apiError_("DUPLICATE_PACKAGING_ITEM", "같은 과목이 저장 목록에 두 번 들어 있습니다.", { row_number: index + 1 });
    incomingIds.add(id);
    const existing = byId.get(id);
    if (!existing || !activeExamDates.has(String(existing.exam_date || ""))) {
      throw apiError_("TIMETABLE_NOT_FOUND", "현재 고사의 시간표에서 과목을 찾지 못했습니다.", { timetable_id: id });
    }
    const expectedRevision = Number(incoming.expected_revision);
    const currentRevision = Math.max(0, Number(existing.packaging_revision) || 0);
    if (!Number.isInteger(expectedRevision) || expectedRevision !== currentRevision) {
      throw apiError_("PACKAGING_REVISION_CONFLICT", "다른 사용자가 같은 과목을 수정했습니다. 최신 자료를 다시 불러오세요.", { timetable_id: id, current_revision: currentRevision });
    }
    const representativeRaw = String(incoming.representative_teacher || "").trim();
    if (representativeRaw.length > 100) throw apiError_("INVALID_REPRESENTATIVE_TEACHER", "대표교사 이름은 100자 이하여야 합니다.", { timetable_id: id });
    const answerRaw = String(incoming.answer_sheet_type || "").trim();
    if (answerRaw && answerRaw !== "card" && answerRaw !== "a4") throw apiError_("INVALID_ANSWER_SHEET_TYPE", "답안지 종류를 확인하세요.", { timetable_id: id });
    const packagingDate = String(incoming.packaging_date || "").trim();
    const slotId = String(incoming.packaging_slot_id || "").trim();
    if (Boolean(packagingDate) !== Boolean(slotId)) throw apiError_("INVALID_PACKAGING_ASSIGNMENT", "포장일과 시간대를 함께 선택하세요.", { timetable_id: id });
    if (packagingDate && !packagingDates.has(packagingDate)) throw apiError_("PACKAGING_DATE_NOT_FOUND", "등록된 포장일을 찾지 못했습니다.", { timetable_id: id });
    const slot = slotId ? slotRows.get(slotId) : null;
    if (slotId && !slot) throw apiError_("PACKAGING_SLOT_NOT_FOUND", "등록된 포장 시간대를 찾지 못했습니다.", { timetable_id: id });
    if (slot && slot.kind === "break") throw apiError_("PACKAGING_BREAK_NOT_ASSIGNABLE", "쉬는 시간에는 과목을 배정할 수 없습니다.", { timetable_id: id });
    const nextAnswer = normalizeAnswerSheetType_(answerRaw);
    const changed = representativeRaw !== String(existing.representative_teacher || "").trim() ||
      nextAnswer !== normalizeAnswerSheetType_(existing.answer_sheet_type) ||
      packagingDate !== String(existing.packaging_date || "").trim() ||
      slotId !== String(existing.packaging_slot_id || "").trim();
    if (!changed) return;
    existing.representative_teacher = representativeRaw;
    existing.answer_sheet_type = nextAnswer;
    existing.packaging_date = packagingDate;
    existing.packaging_slot_id = slotId;
    existing.packaging_revision = currentRevision + 1;
    existing.packaging_updated_at = now;
    changedIds.add(id);
    changedItems.push({ timetable_id: id, revision: existing.packaging_revision });
  });
  const occupancy = new Map();
  timetable.forEach((item) => {
    if (!activeExamDates.has(String(item.exam_date || ""))) return;
    const date = String(item.packaging_date || "").trim();
    const slotId = String(item.packaging_slot_id || "").trim();
    if (!date || !slotId) return;
    const key = date + "|" + slotId;
    const count = (occupancy.get(key) || 0) + 1;
    occupancy.set(key, count);
    if (count > config.capacity) {
      throw apiError_("PACKAGING_SLOT_FULL", "선택한 포장 시간대의 정원이 찼습니다.", { packaging_date: date, packaging_slot_id: slotId, capacity: config.capacity });
    }
  });
  if (changedIds.size) replaceRecords_(SHEETS.TIMETABLE, timetable);
  return Object.assign(getExamPackaging_(request, authSession), { changed_items: changedItems });
}

function cleanup_(request) {
  const cutoff = String(request.cutoff_date || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(cutoff)) throw apiError_("INVALID_DATE", "정리 기준일이 올바르지 않습니다.");
  [SHEETS.EXAM_DATES, SHEETS.TIMETABLE, SHEETS.ABSENCES, SHEETS.COMPLETIONS, SHEETS.SEATING].forEach((name) => {
    replaceRecords_(name, getRecords_(name).filter((item) => !item.exam_date || item.exam_date >= cutoff));
  });
  return adminBootstrap_();
}

function safeAppendAudit_(role, action, targetType, targetId, count, success, detail, actorId) {
  try {
    appendAudit_(role, action, targetType, targetId, count, success, detail, actorId);
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
  if (request.action === "save_exam_packaging_items") {
    return Object.assign(getExamPackaging_(request, { role: "user" }), { idempotent_replay: true });
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

function withMutation_(request, role, actorId, operation) {
  const requestId = String(request.request_id || "");
  if (!requestId) throw apiError_("REQUEST_ID_REQUIRED", "저장 요청 식별자가 없습니다. 화면을 새로고침한 뒤 다시 시도하세요.");
  const cache = CacheService.getScriptCache();
  const idempotencyKey = ["request", role, String(actorId || "unknown"), request.action, requestId].join(":");
  const cached = readMutationCache_(cache, idempotencyKey, request, role);
  if (cached) return cached;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(28000)) throw apiError_("BUSY", "동시 저장 요청이 많습니다. 잠시 후 다시 시도하세요.");
  try {
    const currentUserSession = requireSession_(request.user_session, "user", false);
    if (String(currentUserSession.user_id || "") !== String(actorId || "")) {
      throw apiError_("FORBIDDEN", "사용자 세션이 일치하지 않습니다.");
    }
    if (role === "admin") {
      const currentAdminSession = requireSession_(request.admin_session, "admin", false);
      if (String(currentAdminSession.actor_id || "") !== String(actorId || "")) {
        throw apiError_("FORBIDDEN", "관리자 세션과 사용자 세션이 일치하지 않습니다.");
      }
    }
    const cachedAfterLock = readMutationCache_(cache, idempotencyKey, request, role);
    if (cachedAfterLock) return cachedAfterLock;
    assertAcademicReadable_();
    const result = operation();
    if (request.action === "create_exam_packaging_invite") {
      cache.put(idempotencyKey, JSON.stringify({ __committed: true }), APP.IDEMPOTENCY_TTL_SECONDS);
    } else {
      writeMutationCache_(cache, idempotencyKey, result);
    }
    safeAppendAudit_(role, request.action, "records", "", inferCount_(request), true, "OK", actorId);
    return result;
  } catch (error) {
    if (!error.auditLogged) {
      safeAppendAudit_(role, request.action, "records", "", 0, false, error.code || "REQUEST_FAILED", actorId);
      error.auditLogged = true;
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function withPackagingMutation_(request, session, operation) {
  const requestId = String(request.request_id || "");
  if (!requestId) throw apiError_("REQUEST_ID_REQUIRED", "저장 요청 식별자가 없습니다. 화면을 새로고침한 뒤 다시 시도하세요.");
  const actorId = String(session.session_id || "packaging_session");
  const cache = CacheService.getScriptCache();
  const idempotencyKey = ["request", "packaging", actorId, request.action, requestId].join(":");
  const readCached = () => {
    const raw = cache.get(idempotencyKey);
    if (!raw) return null;
    const cached = JSON.parse(raw);
    return cached.__committed
      ? Object.assign(getExamPackaging_(request, session), { idempotent_replay: true })
      : cached;
  };
  const cached = readCached();
  if (cached) return cached;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(28000)) throw apiError_("BUSY", "동시 저장 요청이 많습니다. 잠시 후 다시 시도하세요.");
  try {
    const currentSession = requirePackagingSession_(request.packaging_session, false);
    if (String(currentSession.session_id || "") !== actorId) throw apiError_("FORBIDDEN", "비밀 링크 세션이 일치하지 않습니다.");
    const cachedAfterLock = readCached();
    if (cachedAfterLock) return cachedAfterLock;
    const result = operation();
    writeMutationCache_(cache, idempotencyKey, result);
    safeAppendAudit_("packaging", request.action, "records", "", inferCount_(request), true, "OK", actorId);
    return result;
  } catch (error) {
    if (!error.auditLogged) {
      safeAppendAudit_("packaging", request.action, "records", "", 0, false, error.code || "REQUEST_FAILED", actorId);
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
  if (Array.isArray(request.subjects)) return request.subjects.length;
  if (Array.isArray(request.enrollments)) return request.enrollments.length;
  if (Array.isArray(request.charts)) return request.charts.length;
  if (Array.isArray(request.items)) return request.items.length;
  return 1;
}

function createUserSession_(user, systemAdmin) {
  const token = randomToken_() + randomToken_();
  const session = {
    role: "user",
    user_id: String(user.id),
    login_name: String(user.login_name),
    session_version: Number(user.session_version) || 1,
    system_admin: systemAdmin === true,
    created_at: nowIso_(),
  };
  if (session.system_admin) {
    session.session_version = Number(PropertiesService.getScriptProperties().getProperty(APP.PROPERTY_ADMIN_SESSION_VERSION) || 1);
  }
  CacheService.getScriptCache().put("session:" + token, JSON.stringify(session), APP.USER_TTL_SECONDS);
  return token;
}

function createAdminSession_(actorId) {
  const token = randomToken_() + randomToken_();
  const session = {
    role: "admin",
    actor_id: String(actorId || ""),
    session_version: Number(PropertiesService.getScriptProperties().getProperty(APP.PROPERTY_ADMIN_SESSION_VERSION) || 1),
    created_at: nowIso_(),
  };
  CacheService.getScriptCache().put("session:" + token, JSON.stringify(session), APP.ADMIN_TTL_SECONDS);
  return token;
}

function requireSession_(token, role, touch) {
  if (!token) throw apiError_("SESSION_REQUIRED", "로그인이 필요합니다.");
  const cache = CacheService.getScriptCache();
  const key = "session:" + token;
  const raw = cache.get(key);
  if (!raw) {
    throw role === "admin"
      ? apiError_("ADMIN_SESSION_EXPIRED", "관리자 세션이 만료되었습니다. 관리자 암호를 다시 입력하세요.")
      : apiError_("SESSION_EXPIRED", "세션이 만료되었습니다. 다시 로그인하세요.");
  }
  let session;
  try {
    session = JSON.parse(raw);
  } catch (ignored) {
    cache.remove(key);
    throw role === "admin"
      ? apiError_("ADMIN_SESSION_EXPIRED", "관리자 세션이 만료되었습니다. 관리자 암호를 다시 입력하세요.")
      : apiError_("SESSION_EXPIRED", "세션이 만료되었습니다. 다시 로그인하세요.");
  }
  if (session.role !== role) {
    cache.remove(key);
    if (role === "user") throw apiError_("SESSION_EXPIRED", "이전 로그인 세션이 만료되었습니다. 다시 로그인하세요.");
    throw apiError_("FORBIDDEN", "이 작업에 필요한 권한이 없습니다.");
  }
  if (role === "admin") {
    const currentVersion = Number(PropertiesService.getScriptProperties().getProperty(APP.PROPERTY_ADMIN_SESSION_VERSION) || 1);
    if (Number(session.session_version) !== currentVersion) {
      cache.remove(key);
      throw apiError_("ADMIN_SESSION_EXPIRED", "관리자 암호가 변경되어 관리자 세션이 만료되었습니다.");
    }
  } else if (role === "user") {
    if (!session.user_id) {
      cache.remove(key);
      throw apiError_("SESSION_EXPIRED", "이전 로그인 세션이 만료되었습니다. 다시 로그인하세요.");
    }
    if (session.system_admin === true) {
      const adminVersion = Number(PropertiesService.getScriptProperties().getProperty(APP.PROPERTY_ADMIN_SESSION_VERSION) || 1);
      if (session.user_id !== "system_admin" || Number(session.session_version) !== adminVersion) {
        cache.remove(key);
        throw apiError_("SESSION_EXPIRED", "관리자 설정 세션이 만료되었습니다. 다시 로그인하세요.");
      }
    } else {
      const currentUser = findAccessUserById_(session.user_id);
      if (!currentUser || !currentUser.active || Number(currentUser.session_version) !== Number(session.session_version)) {
        cache.remove(key);
        throw apiError_("SESSION_EXPIRED", "접속 권한이 변경되어 세션이 만료되었습니다. 다시 로그인하세요.");
      }
    }
  }
  if (touch !== false) {
    const ttl = role === "admin" ? APP.ADMIN_TTL_SECONDS : APP.USER_TTL_SECONDS;
    cache.put(key, raw, ttl);
  }
  return session;
}

function loginRateKeys_(scope, request, loginName) {
  const identity = String(request.client_id || request.client_key || "").trim() || "anonymous";
  const keys = [
    {
      key: "login-rate:" + scope + ":client:" + hashSecret_(identity.slice(0, 200), "login-client-rate-v1").slice(0, 32),
      max: APP.LOGIN_RATE_MAX_ATTEMPTS,
      clear_on_success: true,
    },
  ];
  const nameKey = loginLookupKey_(loginName || "");
  if (nameKey) {
    keys.push({
      key: "login-rate:" + scope + ":name:" + hashSecret_(nameKey, "login-name-rate-v1").slice(0, 32),
      max: APP.LOGIN_RATE_MAX_ATTEMPTS,
      clear_on_success: true,
    });
  }
  keys.push({ key: "login-rate:" + scope + ":global", max: APP.LOGIN_RATE_GLOBAL_MAX_ATTEMPTS, clear_on_success: false });
  return keys;
}

function nowMillis_() {
  return Date.now();
}

function loginRateStorageKey_(cacheKey) {
  return hashSecret_(String(cacheKey || ""), "login-rate-property-v1").slice(0, 32);
}

function loginRatePropertyName_(index) {
  return APP.PROPERTY_LOGIN_RATE_STATE_PREFIX + String(index);
}

function normalizeLoginRateEntry_(value, nowMs) {
  const data = Array.isArray(value)
    ? { count: value[0], expires_at: value[1], updated_at: value[2] }
    : value;
  if (!data || typeof data !== "object") return null;
  const count = Number(data.count);
  const expiresAt = Number(data.expires_at);
  const updatedAt = Number(data.updated_at);
  if (!Number.isInteger(count) || count < 1 || count > 100000) return null;
  if (!Number.isFinite(expiresAt) || expiresAt <= nowMs) return null;
  if (!Number.isFinite(updatedAt) || updatedAt < 0) return null;
  return { count: count, expires_at: expiresAt, updated_at: updatedAt };
}

function pruneLoginRateEntries_(entries, nowMs) {
  return Object.fromEntries(
    Object.entries(entries || {})
      .map((item) => [item[0], normalizeLoginRateEntry_(item[1], nowMs)])
      .filter((item) => /^[A-Za-z0-9_-]{32}$/.test(item[0]) && item[1])
      .sort((left, right) => right[1].updated_at - left[1].updated_at)
      .slice(0, APP.LOGIN_RATE_MAX_ENTRIES)
  );
}

function readLoginRateState_(nowMs) {
  const properties = PropertiesService.getScriptProperties();
  const allProperties = properties.getProperties();
  const entries = {};
  let changed = false;
  for (let index = 0; index < APP.LOGIN_RATE_STATE_SHARDS; index += 1) {
    const raw = allProperties[loginRatePropertyName_(index)];
    if (!raw) continue;
    let shard;
    try {
      shard = JSON.parse(raw);
    } catch (ignored) {
      changed = true;
      continue;
    }
    if (!shard || Array.isArray(shard) || typeof shard !== "object") {
      changed = true;
      continue;
    }
    Object.entries(shard).forEach((item) => {
      const key = String(item[0] || "");
      const normalized = normalizeLoginRateEntry_(item[1], nowMs);
      if (!/^[A-Za-z0-9_-]{32}$/.test(key) || !normalized) {
        changed = true;
        return;
      }
      if (entries[key]) {
        changed = true;
        entries[key] = {
          count: Math.max(entries[key].count, normalized.count),
          expires_at: Math.max(entries[key].expires_at, normalized.expires_at),
          updated_at: Math.max(entries[key].updated_at, normalized.updated_at),
        };
      } else {
        entries[key] = normalized;
      }
    });
  }
  const pruned = pruneLoginRateEntries_(entries, nowMs);
  if (Object.keys(pruned).length !== Object.keys(entries).length) changed = true;
  return { entries: pruned, changed: changed };
}

function writeLoginRateState_(entries, nowMs) {
  const pruned = pruneLoginRateEntries_(entries, nowMs);
  const shards = Array.from({ length: APP.LOGIN_RATE_STATE_SHARDS }, () => ({}));
  Object.entries(pruned).forEach((item, index) => {
    const entry = item[1];
    shards[index % APP.LOGIN_RATE_STATE_SHARDS][item[0]] = [entry.count, entry.expires_at, entry.updated_at];
  });
  const payload = {};
  shards.forEach((shard, index) => {
    payload[loginRatePropertyName_(index)] = JSON.stringify(shard);
  });
  PropertiesService.getScriptProperties().setProperties(payload, false);
  return pruned;
}

function readLoginRateCacheEntry_(cache, key, nowMs) {
  let raw = "";
  try {
    raw = cache.get(key) || "";
  } catch (ignored) {
    return null;
  }
  if (!raw) return null;
  let entry;
  try {
    const parsed = JSON.parse(raw);
    entry = normalizeLoginRateEntry_(parsed, nowMs);
  } catch (ignored) {
    entry = null;
  }
  if (!entry) {
    const legacyCount = Number(raw);
    entry = Number.isInteger(legacyCount) && legacyCount > 0
      ? { count: legacyCount, expires_at: nowMs + APP.LOGIN_RATE_TTL_SECONDS * 1000, updated_at: nowMs }
      : null;
  }
  if (!entry) {
    try { cache.remove(key); } catch (ignored) {}
  }
  return entry;
}

function writeLoginRateCacheEntry_(cache, key, entry, nowMs) {
  const ttl = Math.max(1, Math.ceil((entry.expires_at - nowMs) / 1000));
  try {
    cache.put(key, JSON.stringify(entry), Math.min(APP.LOGIN_RATE_TTL_SECONDS, ttl));
  } catch (ignored) {}
}

function assertLoginAllowed_(rateKeys) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw apiError_("BUSY", "로그인 보안 상태를 확인하는 중입니다. 잠시 후 다시 시도하세요.");
  const cache = CacheService.getScriptCache();
  try {
    const nowMs = nowMillis_();
    const state = readLoginRateState_(nowMs);
    let changed = state.changed;
    let blocked = false;
    rateKeys.forEach((item) => {
      const storageKey = loginRateStorageKey_(item.key);
      const cached = readLoginRateCacheEntry_(cache, item.key, nowMs);
      const stored = state.entries[storageKey] || null;
      if (cached) {
        const merged = stored ? {
          count: Math.max(stored.count, cached.count),
          expires_at: Math.max(stored.expires_at, cached.expires_at),
          updated_at: Math.max(stored.updated_at, cached.updated_at),
        } : cached;
        if (!stored || JSON.stringify(stored) !== JSON.stringify(merged)) {
          state.entries[storageKey] = merged;
          changed = true;
        }
      }
      const effective = state.entries[storageKey] || null;
      if (effective) {
        writeLoginRateCacheEntry_(cache, item.key, effective, nowMs);
        if (effective.count >= item.max) blocked = true;
      }
    });
    if (changed) writeLoginRateState_(state.entries, nowMs);
    if (blocked) throw apiError_("LOGIN_RATE_LIMIT", "로그인 시도가 너무 많습니다. 5분 뒤 다시 시도하세요.");
  } finally {
    lock.releaseLock();
  }
}

function recordLoginFailure_(rateKeys) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw apiError_("BUSY", "로그인 보안 상태를 저장하는 중입니다. 잠시 후 다시 시도하세요.");
  const cache = CacheService.getScriptCache();
  try {
    const nowMs = nowMillis_();
    const state = readLoginRateState_(nowMs);
    rateKeys.forEach((item) => {
      const storageKey = loginRateStorageKey_(item.key);
      const cached = readLoginRateCacheEntry_(cache, item.key, nowMs);
      const stored = state.entries[storageKey] || null;
      const attempts = Math.max(stored ? stored.count : 0, cached ? cached.count : 0) + 1;
      const entry = {
        count: attempts,
        expires_at: nowMs + APP.LOGIN_RATE_TTL_SECONDS * 1000,
        updated_at: nowMs,
      };
      state.entries[storageKey] = entry;
      writeLoginRateCacheEntry_(cache, item.key, entry, nowMs);
    });
    writeLoginRateState_(state.entries, nowMs);
  } finally {
    lock.releaseLock();
  }
}

function clearLoginFailures_(rateKeys) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw apiError_("BUSY", "로그인 보안 상태를 갱신하는 중입니다. 잠시 후 다시 시도하세요.");
  const cache = CacheService.getScriptCache();
  try {
    const nowMs = nowMillis_();
    const state = readLoginRateState_(nowMs);
    let changed = state.changed;
    rateKeys.filter((item) => item.clear_on_success !== false).forEach((item) => {
      const storageKey = loginRateStorageKey_(item.key);
      if (state.entries[storageKey]) {
        delete state.entries[storageKey];
        changed = true;
      }
      try { cache.remove(item.key); } catch (ignored) {}
    });
    if (changed) writeLoginRateState_(state.entries, nowMs);
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

function randomSecret256_() {
  const entropy = randomToken_() + randomToken_() + randomToken_();
  return hashSecret_(entropy, randomToken_() + "\u001f" + String(Date.now()));
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
  error.public_error = true;
  return error;
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
