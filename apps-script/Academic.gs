/** 학적 원장은 읽기만 한다. 원장 위치는 비공개 Script Properties에 보관한다. */
function academicSourceConfig_() {
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty("ACADEMIC_SOURCE_SHEET_ID") && typeof academicDeploymentDefaults_ === "function") {
    const defaults = academicDeploymentDefaults_();
    if (defaults && defaults.spreadsheet_id) properties.setProperty("ACADEMIC_SOURCE_SHEET_ID", defaults.spreadsheet_id);
  }
  const notice = normalizeStoredExamNotice_(getSettings_().exam_notice_json, "");
  const today = Utilities.formatDate(new Date(), APP.TIMEZONE, "yyyy-MM-dd");
  const currentYear = Number(today.slice(0, 4)) - (Number(today.slice(5, 7)) < 3 ? 1 : 0);
  return { spreadsheet_id: properties.getProperty("ACADEMIC_SOURCE_SHEET_ID") || "",
    academic_year: Number(notice.year) || currentYear, semester: Number(notice.semester) || (Number(today.slice(5, 7)) >= 8 ? 2 : 1) };
}

function academicLinkStatus_() {
  const configured = Boolean(academicSourceConfig_().spreadsheet_id);
  const stored = parseJson_(getSettings_().academic_sync_status_json, {});
  return Object.assign({ configured: configured, state: configured ? "pending" : "not_configured" }, stored);
}

function readAcademicTable_(book, name, required) {
  const sheet = book.getSheetByName(name);
  if (!sheet) { if (required) throw apiError_("ACADEMIC_SOURCE_SCHEMA", "학적 원장의 필수 표를 찾지 못했습니다."); return []; }
  const values = sheet.getDataRange().getValues();
  const headers = (values.shift() || []).map(String);
  (required || []).forEach((key) => { if (headers.indexOf(key) < 0) throw apiError_("ACADEMIC_SOURCE_SCHEMA", "학적 원장의 열 구성을 확인하세요."); });
  return values.filter((row) => row.some((value) => value !== "")).map((row) => Object.fromEntries(headers.map((key, i) => [key, row[i]])));
}

function academicActive_(row) { return row.active !== false && !["false", "0"].includes(String(row.active).toLowerCase()); }
function academicNameKey_(value) { return String(value || "").normalize("NFKC").replace(/[\s·･∙]/g, "").toLowerCase(); }
function academicList_(value) {
  const parsed = typeof value === "string" ? parseJson_(value, null) : value;
  if (!Array.isArray(parsed)) throw apiError_("ACADEMIC_SOURCE_INVALID", "학적 원장의 과목 자료 형식을 확인하세요.");
  return parsed;
}

function readAcademicSnapshot_(config) {
  const book = SpreadsheetApp.openById(config.spreadsheet_id);
  if (book.getId() === getSpreadsheet_().getId()) throw apiError_("ACADEMIC_SOURCE_INVALID", "학적 원장과 고사 원장은 서로 달라야 합니다.");
  const inYear = (row) => Number(row.academic_year) === config.academic_year;
  const tables = {
    students: readAcademicTable_(book, "시스템_학생", ["student_uid", "name", "active"]),
    assignments: readAcademicTable_(book, "시스템_학급배정", ["student_uid", "academic_year", "grade", "class_no", "student_no", "active"]).filter(inYear),
    classes: readAcademicTable_(book, "시스템_학급정원", ["academic_year", "grade", "class_no"]).filter(inYear),
    subjects: readAcademicTable_(book, "시스템_과목", ["subject_id", "academic_year", "grade", "name"]).filter(inYear),
    sections: readAcademicTable_(book, "시스템_분반", ["section_id", "academic_year", "subject_id", "section_name"]).filter(inYear),
    enrollments: readAcademicTable_(book, "시스템_수강", ["student_uid", "academic_year", "subject_id"]).filter(inYear),
    courses: readAcademicTable_(book, "시스템_선택과목명단").filter(inYear),
    plans: readAcademicTable_(book, "시스템_교육과정").filter(inYear),
    section_map: parseJson_((readAcademicTable_(book, "시스템_설정", ["setting_key", "setting_value_json"]).find((row) => row.setting_key === "course_section_map_" + config.academic_year) || {}).setting_value_json, {}),
  };
  if (typeof academicFallbackPlan_ === "function") [2, 3].forEach((grade) => {
    if (!tables.plans.some((row) => Number(row.grade) === grade)) {
      const plan = academicFallbackPlan_(config.academic_year, grade);
      if (plan && plan.groups) tables.plans.push({ grade: grade, groups_json: JSON.stringify(plan.groups) });
    }
  });
  return buildAcademicSnapshot_(tables, config);
}

function buildAcademicSnapshot_(tables, config) {
  const master = new Map(tables.students.map((row) => [String(row.student_uid), row]));
  const classes = new Map(), students = [], seenUid = new Set(), seenNumber = new Set();
  function addClass(grade, classNo) {
    if (![1, 2, 3].includes(Number(grade)) || !Number.isInteger(Number(classNo)) || Number(classNo) < 1) throw apiError_("ACADEMIC_SOURCE_INVALID", "학적 원장의 학급 정보를 확인하세요.");
    const key = Number(grade) + "|" + Number(classNo);
    if (!classes.has(key)) classes.set(key, { grade: Number(grade), class_num: Number(classNo) });
    return key;
  }
  tables.classes.filter(academicActive_).forEach((row) => addClass(row.grade, row.class_no));
  tables.assignments.filter(academicActive_).forEach((row) => {
    const uid = String(row.student_uid || ""), student = master.get(uid);
    if (!student) throw apiError_("ACADEMIC_SOURCE_INVALID", "학적 원장의 학생 연결을 확인하세요.");
    if (!academicActive_(student) || (student.enrollment_status && student.enrollment_status !== "ACTIVE")) return;
    const classKey = addClass(row.grade, row.class_no), number = Number(row.student_no), name = String(student.name || "").trim();
    if (!uid || !name || !Number.isInteger(number) || number < 1 || seenUid.has(uid) || seenNumber.has(classKey + "|" + number)) throw apiError_("ACADEMIC_SOURCE_CONFLICT", "학적 명단의 중복 학생·번호를 확인하세요. 기존 고사 명단은 유지됩니다.");
    seenUid.add(uid); seenNumber.add(classKey + "|" + number);
    students.push({ source_student_uid: uid, class_key: classKey, number: number, name: name, grade: Number(row.grade) });
  });
  if (!students.length) throw apiError_("ACADEMIC_EMPTY_SOURCE", "해당 학년도에 재학생 명단이 없습니다. 기존 고사 명단은 유지됩니다.");
  const catalog = { "1": [], "2": [], "3": [] }, sourceSubjects = new Map(), courses = [], warnings = { unreviewed_choices: 0 };
  const inSemester = (value) => !value || value === "YEAR" || String(value) === String(config.semester);
  function addSubject(grade, name) {
    name = String(name || "").trim();
    if (!catalog[String(grade)] || !name || name.length > 80) throw apiError_("ACADEMIC_SUBJECT_INVALID", "학적 원장의 과목명을 확인하세요.");
    if (!catalog[String(grade)].some((old) => academicNameKey_(old) === academicNameKey_(name))) catalog[String(grade)].push(name);
    return name;
  }
  tables.subjects.filter(academicActive_).forEach((row) => {
    const semester = String(row.code || "").split(":")[0];
    sourceSubjects.set(String(row.subject_id), row);
    if (!/^[12]$/.test(semester) || inSemester(semester)) addSubject(row.grade, row.name);
  });
  tables.plans.forEach((plan) => academicList_(plan.groups_json).filter((group) => inSemester(group.semester)).forEach((group) => academicList_(group.subjects).forEach((name) => addSubject(plan.grade, name))));
  const studentByUid = new Map(students.map((row) => [row.source_student_uid, row]));
  const sections = new Map(tables.sections.filter(academicActive_).map((row) => [String(row.section_id), row]));
  const canonicalCourseUids = new Set(tables.courses.filter(academicActive_).filter((row) => academicList_(row.choices_json).some((choice) => inSemester(choice.semester))).map((row) => String(row.student_uid)));
  const courseKeys = new Set();
  function addCourse(uid, name, subjectId, sectionId) {
    const student = studentByUid.get(String(uid)); if (!student) return;
    name = addSubject(student.grade, name);
    const key = uid + "|" + academicNameKey_(name);
    if (courseKeys.has(key)) return;
    courseKeys.add(key);
    const section = sections.get(String(sectionId || ""));
    courses.push({ source_student_uid: String(uid), grade: student.grade, subject_name: name,
      source_subject_id: String(subjectId || ""), source_section_id: section ? String(section.section_id) : "", section_name: section ? String(section.section_name || "") : "" });
  }
  tables.enrollments.filter(academicActive_).filter((row) => !row.status || ["ENROLLED", "ACTIVE", "CONFIRMED"].includes(row.status)).forEach((row) => {
    const subject = sourceSubjects.get(String(row.subject_id));
    if (!subject || !academicActive_(subject)) return;
    const semester = String(subject.code || "").split(":")[0];
    if (/^[12]$/.test(semester) && !inSemester(semester)) return;
    if (canonicalCourseUids.has(String(row.student_uid))) return;
    addCourse(row.student_uid, subject.name, row.subject_id, row.section_id);
  });
  const recordUids = new Set();
  tables.courses.filter(academicActive_).forEach((row) => {
    if (!studentByUid.has(String(row.student_uid))) return;
    if (recordUids.has(String(row.student_uid))) throw apiError_("ACADEMIC_SOURCE_CONFLICT", "학적 원장의 선택과목 명단이 중복되어 있습니다.");
    recordUids.add(String(row.student_uid));
    academicList_(row.choices_json).filter((choice) => inSemester(choice.semester)).forEach((choice) => {
      const student = studentByUid.get(String(row.student_uid));
      const original = String(choice.raw_subject || choice.name || "").trim();
      const normalized = academicNameKey_(original);
      const matches = catalog[String(student.grade)].filter((name) => normalized === academicNameKey_(name) || (normalized.startsWith(academicNameKey_(name)) && /^[a-f]\d+$/i.test(normalized.slice(academicNameKey_(name).length))));
      if (matches.length > 1 || (!matches.length && choice.review_needed)) { warnings.unreviewed_choices += 1; return; }
      const name = matches[0] || choice.name || choice.raw_subject;
      const sourceCode = choice.source_code || (/([A-F]\d+)$/i.exec(original) || [])[1] || "";
      const matchingSections = Array.from(sections.values()).filter((section) => {
        const meta = (tables.section_map || {})[section.section_id];
        return meta && String(meta.subject_id) === String(choice.subject_id) && String(meta.semester) === String(choice.semester) && String(meta.slot) === String(choice.slot) && String(meta.source_code || "") === String(sourceCode);
      });
      const sectionId = choice.section_id || (matchingSections.length === 1 ? matchingSections[0].section_id : "");
      if (name && !["-", "없음", "미선택"].includes(name)) addCourse(row.student_uid, name, choice.subject_id, sectionId);
    });
  });
  Object.values(catalog).forEach((names) => names.sort((a, b) => a.localeCompare(b, "ko")));
  return { academic_year: config.academic_year, semester: config.semester, classes: Array.from(classes.values()), students: students, courses: courses, catalog: catalog, warnings: warnings };
}

/** 기존 고사 ID를 재사용한다. 다른 학생은 같은 번호여도 새 ID를 받는다. */
function planAcademicMerge_(snapshot, previous, now) {
  const year = snapshot.academic_year;
  const classes = previous.classes.map((row) => ({ ...row })), students = previous.students.map((row) => ({ ...row }));
  const classIds = new Map(), sourceStudents = new Map();
  const adoptable = (row) => !row.source_academic_year || Number(row.source_academic_year) === year;
  snapshot.classes.forEach((row) => {
    let old = classes.find((item) => Number(item.grade) === row.grade && Number(item.class_num) === row.class_num && adoptable(item));
    if (!old) { old = { id: createId_("class") }; classes.push(old); }
    Object.assign(old, row, { active: true, source_academic_year: year, updated_at: now });
    classIds.set(row.grade + "|" + row.class_num, old.id);
  });
  const activeClassIds = new Set(classIds.values());
  classes.filter((row) => row.source_academic_year && !activeClassIds.has(row.id)).forEach((row) => { row.active = false; });
  const used = new Set();
  snapshot.students.forEach((row) => {
    const classId = classIds.get(row.class_key);
    let old = students.find((item) => item.source_student_uid === row.source_student_uid && Number(item.source_academic_year) === year);
    if (!old) {
      const candidates = students.filter((item) => !item.source_student_uid && String(item.class_id) === String(classId) && Number(item.number) === row.number && String(item.name).normalize("NFKC").trim() === row.name.normalize("NFKC"));
      if (candidates.length > 1) throw apiError_("ACADEMIC_SOURCE_CONFLICT", "기존 고사 명단에서 중복 학생을 확인하세요.");
      old = candidates[0];
    }
    if (!old) { old = { id: createId_("student") }; students.push(old); }
    Object.assign(old, { class_id: classId, number: row.number, name: row.name, active: true, source_student_uid: row.source_student_uid, source_academic_year: year, updated_at: now });
    used.add(old.id); sourceStudents.set(row.source_student_uid, old);
  });
  students.filter((row) => (row.source_academic_year || activeClassIds.has(row.class_id)) && !used.has(row.id)).forEach((row) => { row.active = false; row.updated_at = now; });
  const enrollments = previous.enrollments.filter((row) => !activeClassIds.has(row.class_id));
  snapshot.courses.forEach((row) => {
    const student = sourceStudents.get(row.source_student_uid);
    const old = previous.enrollments.find((item) => String(item.student_id) === String(student.id) && academicNameKey_(item.subject_name) === academicNameKey_(row.subject_name));
    const timetable = previous.timetable.filter((item) => Number(item.grade) === row.grade && academicNameKey_(item.subject_name) === academicNameKey_(row.subject_name));
    const names = Array.from(new Set(timetable.map((item) => item.subject_name)));
    if (names.length > 1) throw apiError_("ACADEMIC_SUBJECT_AMBIGUOUS", "시간표에서 같은 과목의 표기를 통일하세요.");
    enrollments.push({ ...row, id: old ? old.id : createId_("enrollment"), class_id: student.class_id, student_id: student.id,
      subject_name: names[0] || row.subject_name, room_name: old ? String(old.room_name || "") : "", source_academic_year: year, semester: snapshot.semester, updated_at: now });
  });
  return { classes: classes, students: students, enrollments: enrollments };
}

function syncAcademicData_(force, requestId, actorId) {
  const requestKey = force && requestId ? ["request", "academic", actorId, requestId].join(":") : "";
  const requestCache = CacheService.getScriptCache();
  if (requestKey && requestCache.get(requestKey)) return;
  const config = academicSourceConfig_();
  if (!config.spreadsheet_id) { if (force) throw apiError_("ACADEMIC_NOT_CONFIGURED", "학적 원장 연결이 설정되지 않았습니다."); return; }
  const status = academicLinkStatus_();
  if (!force && status.state === "error" && Date.now() - new Date(status.attempted_at || 0).getTime() < 60000) return;
  if (!force && status.academic_year === config.academic_year && status.semester === config.semester && Date.now() - new Date(status.synced_at || 0).getTime() < 300000) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(force ? 28000 : 1000)) { if (force) throw apiError_("BUSY", "다른 저장이 진행 중입니다. 다시 시도하세요."); return; }
  try {
    if (requestKey && requestCache.get(requestKey)) return;
    const latest = academicLinkStatus_();
    if (!force && latest.state === "connected" && latest.academic_year === config.academic_year && latest.semester === config.semester && Date.now() - new Date(latest.synced_at || 0).getTime() < 300000) return;
    const snapshot = readAcademicSnapshot_(config), now = nowIso_();
    const before = { classes: getRecordsFresh_(SHEETS.CLASSES), students: getRecordsFresh_(SHEETS.STUDENTS), enrollments: getRecordsFresh_(SHEETS.ENROLLMENTS), timetable: getRecordsFresh_(SHEETS.TIMETABLE), settings: getRecordsFresh_(SHEETS.SETTINGS) };
    const next = planAcademicMerge_(snapshot, before, now);
    const properties = PropertiesService.getScriptProperties();
    if (!properties.getProperty("ACADEMIC_INITIAL_BACKUP_SHEET_ID")) {
      const backup = getSpreadsheet_().copy("고사 통합 - 학적 연결 전 자동백업 " + Utilities.formatDate(new Date(), APP.TIMEZONE, "yyyyMMdd-HHmmss"));
      properties.setProperty("ACADEMIC_INITIAL_BACKUP_SHEET_ID", backup.getId());
    }
    // 데이터 검증과 계획을 모두 끝낸 뒤에만 쓰기. 중간 실패 시 전체를 원복한다.
    try {
      replaceRecords_(SHEETS.CLASSES, next.classes);
      replaceRecords_(SHEETS.STUDENTS, next.students);
      replaceRecords_(SHEETS.ENROLLMENTS, next.enrollments);
    const catalog = normalizeStoredSubjectCatalog_(getSettings_().subject_catalog_json);
    const priorCatalog = parseJson_(getSettings_().academic_catalog_json, {});
    Object.keys(snapshot.catalog).forEach((grade) => { catalog[grade] = Array.from(new Set([...(snapshot.catalog[grade] || []), ...(catalog[grade] || []).filter((name) => !(priorCatalog[grade] || []).includes(name))])).sort((a, b) => a.localeCompare(b, "ko")); });
    setSetting_("subject_catalog_json", JSON.stringify(catalog));
    setSetting_("academic_catalog_json", JSON.stringify(snapshot.catalog));
    setSetting_("academic_sync_status_json", JSON.stringify({ state: "connected", academic_year: snapshot.academic_year, semester: snapshot.semester, synced_at: now,
      counts: { classes: snapshot.classes.length, students: snapshot.students.length, enrollments: snapshot.courses.length, subjects: Object.values(snapshot.catalog).reduce((sum, names) => sum + names.length, 0) }, warnings: snapshot.warnings }));
    } catch (error) {
      replaceRecords_(SHEETS.CLASSES, before.classes); replaceRecords_(SHEETS.STUDENTS, before.students); replaceRecords_(SHEETS.ENROLLMENTS, before.enrollments); replaceRecords_(SHEETS.SETTINGS, before.settings); throw error;
    }
    safeAppendAudit_("system", "sync_academic_data", "records", "", snapshot.students.length + snapshot.courses.length, true, "OK", "academic_sync");
    if (requestKey) requestCache.put(requestKey, "1", APP.IDEMPOTENCY_TTL_SECONDS);
  } catch (error) {
    setSetting_("academic_sync_status_json", JSON.stringify(Object.assign({}, status, { state: "error", attempted_at: nowIso_(), error_code: error.code || "ACADEMIC_SYNC_FAILED" })));
    safeAppendAudit_("system", "sync_academic_data", "records", "", 0, false, error.code || "ACADEMIC_SYNC_FAILED", "academic_sync");
    if (force) throw error;
  } finally { lock.releaseLock(); }
}
