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
    academic_year: Number(notice.year) || currentYear, semester: Number(notice.semester) || (Number(today.slice(5, 7)) < 3 || Number(today.slice(5, 7)) >= 8 ? 2 : 1) };
}

function academicLinkStatus_() {
  const configured = Boolean(academicSourceConfig_().spreadsheet_id);
  const stored = parseJson_(PropertiesService.getScriptProperties().getProperty("ACADEMIC_SYNC_STATUS") || getSettings_().academic_sync_status_json, {});
  const target = academicSourceConfig_();
  const result = Object.assign({ configured: configured, state: configured ? "pending" : "not_configured" }, stored,
    { configured: configured, target_year: target.academic_year, target_semester: target.semester });
  result.period_mismatch = configured && Boolean(stored.synced_at) && (Number(stored.academic_year) !== target.academic_year || Number(stored.semester) !== target.semester);
  if (result.period_mismatch && result.state === "connected") result.state = "stale";
  if (configured && result.state === "connected" && Number(result.integration_version || 0) < 2) result.state = "pending";
  return result;
}

function readAcademicTable_(book, name, required) {
  const sheet = book.getSheetByName(name);
  if (!sheet) { if (required) throw apiError_("ACADEMIC_SOURCE_SCHEMA", "학적 원장의 필수 표를 찾지 못했습니다."); return []; }
  const values = sheet.getDataRange().getValues();
  const headers = (values.shift() || []).map(String);
  (required || []).forEach((key) => { if (headers.indexOf(key) < 0) throw apiError_("ACADEMIC_SOURCE_SCHEMA", "학적 원장의 열 구성을 확인하세요."); });
  return values.filter((row) => row.some((value) => value !== "")).map((row) => Object.fromEntries(headers.map((key, i) => [key, row[i]])));
}

function academicActive_(row, fallback) {
  const value = String(row.active == null ? "" : row.active).trim().toLowerCase();
  if (["false", "0", "n", "no", "비활성", "미사용"].includes(value)) return false;
  if (["true", "1", "y", "yes", "활성", "사용", "완료"].includes(value)) return true;
  return fallback !== false;
}
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
  const master = new Map();
  tables.students.forEach((row) => {
    const uid = String(row.student_uid || "").trim();
    if (!uid || master.has(uid)) throw apiError_("ACADEMIC_SOURCE_CONFLICT", "학적 원장의 학생 고유번호 중복·누락을 확인하세요.");
    master.set(uid, row);
  });
  const classes = new Map(), students = [], seenUid = new Set(), seenNumber = new Set();
  function addClass(grade, classNo) {
    if (![1, 2, 3].includes(Number(grade)) || !Number.isInteger(Number(classNo)) || Number(classNo) < 1) throw apiError_("ACADEMIC_SOURCE_INVALID", "학적 원장의 학급 정보를 확인하세요.");
    const key = Number(grade) + "|" + Number(classNo);
    if (!classes.has(key)) classes.set(key, { grade: Number(grade), class_num: Number(classNo) });
    return key;
  }
  tables.classes.filter((row) => academicActive_(row)).forEach((row) => addClass(row.grade, row.class_no));
  tables.assignments.filter((row) => academicActive_(row, false)).forEach((row) => {
    const uid = String(row.student_uid || ""), student = master.get(uid);
    if (!student) throw apiError_("ACADEMIC_SOURCE_INVALID", "학적 원장의 학생 연결을 확인하세요.");
    if (!academicActive_(student) || (student.enrollment_status && student.enrollment_status !== "ACTIVE")) return;
    const classKey = addClass(row.grade, row.class_no), number = Number(row.student_no), name = String(student.name || "").trim();
    if (!uid || !name || !Number.isInteger(number) || number < 1 || seenUid.has(uid) || seenNumber.has(classKey + "|" + number)) throw apiError_("ACADEMIC_SOURCE_CONFLICT", "학적 명단의 중복 학생·번호를 확인하세요. 기존 고사 명단은 유지됩니다.");
    seenUid.add(uid); seenNumber.add(classKey + "|" + number);
    students.push({ source_student_uid: uid, class_key: classKey, number: number, name: name, grade: Number(row.grade) });
  });
  if (!students.length) throw apiError_("ACADEMIC_EMPTY_SOURCE", "해당 학년도에 재학생 명단이 없습니다. 기존 고사 명단은 유지됩니다.");
  const catalog = { "1": [], "2": [], "3": [] }, sourceSubjects = new Map(), courses = [], warnings = { unreviewed_choices: 0, ambiguous_sections: 0 };
  const inSemester = (value) => !value || value === "YEAR" || String(value) === String(config.semester);
  function addSubject(grade, name) {
    name = String(name || "").trim();
    if (!catalog[String(grade)] || !name || name.length > 80) throw apiError_("ACADEMIC_SUBJECT_INVALID", "학적 원장의 과목명을 확인하세요.");
    if (!catalog[String(grade)].some((old) => academicNameKey_(old) === academicNameKey_(name))) catalog[String(grade)].push(name);
    return name;
  }
  tables.subjects.filter((row) => academicActive_(row)).forEach((row) => {
    const semester = String(row.code || "").split(":")[0];
    if (!row.subject_id || sourceSubjects.has(String(row.subject_id))) throw apiError_("ACADEMIC_SOURCE_CONFLICT", "학적 원장의 과목 고유번호 중복·누락을 확인하세요.");
    sourceSubjects.set(String(row.subject_id), row);
    if (!/^[12]$/.test(semester) || inSemester(semester)) addSubject(row.grade, row.name);
  });
  const plans = new Map(), choiceCatalog = new Map();
  tables.plans.forEach((plan) => {
    const grade = Number(plan.grade);
    if (plans.has(grade)) throw apiError_("ACADEMIC_SOURCE_CONFLICT", "학적 원장의 교육과정이 중복되어 있습니다.");
    const groups = academicList_(plan.groups_json).filter((group) => inSemester(group.semester));
    plans.set(grade, groups);
    choiceCatalog.set(grade, groups.flatMap((group) => academicList_(group.subjects).map((raw) => {
      const name = addSubject(grade, raw), semester = String(group.semester);
      const digest = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(academicNameKey_(name)), Utilities.Charset.UTF_8)).replace(/=+$/, "");
      return { name: name, semester: semester, group: String(group.key), subject_id: "course_" + config.academic_year + "_" + grade + "_" + semester + "_" + digest.slice(0, 16) };
    })));
  });
  const studentByUid = new Map(students.map((row) => [row.source_student_uid, row]));
  const sections = new Map(tables.sections.filter((row) => academicActive_(row)).map((row) => [String(row.section_id), row]));
  const canonical = new Map(), courseRecordUids = new Set();
  tables.courses.filter((row) => academicActive_(row)).forEach((row) => {
    const uid = String(row.student_uid), student = studentByUid.get(uid);
    if (!student) return;
    if (courseRecordUids.has(uid)) throw apiError_("ACADEMIC_SOURCE_CONFLICT", "학적 원장의 선택과목 명단이 중복되어 있습니다.");
    courseRecordUids.add(uid);
    const choices = academicList_(row.choices_json).filter((choice) => inSemester(choice.semester));
    const groups = plans.get(student.grade) || [];
    const candidates = choiceCatalog.has(student.grade) ? choiceCatalog.get(student.grade) : Array.from(sourceSubjects.values()).filter((subject) => Number(subject.grade) === student.grade).map((subject) => ({ name: subject.name, subject_id: subject.subject_id, semester: String(subject.code || "").split(":")[0] }));
    const interpreted = choices.map((choice) => {
      const raw = String(choice.raw_subject || choice.name || "").trim(), key = academicNameKey_(raw);
      if (!raw || ["-", "없음", "미선택"].includes(raw)) return null;
      const sameTerm = candidates.filter((item) => String(item.semester) === String(choice.semester));
      const exact = sameTerm.filter((item) => academicNameKey_(item.name) === key);
      const coded = sameTerm.filter((item) => key.startsWith(academicNameKey_(item.name)) && /^[a-f]\d+$/i.test(key.slice(academicNameKey_(item.name).length)));
      const matches = exact.length ? exact : coded;
      if (matches.length !== 1) throw apiError_("ACADEMIC_CHOICES_INCOMPLETE", "확인되지 않은 선택과목이 있어 반영하지 않았습니다. 학적 통합에서 과목 확인을 마친 뒤 다시 동기화하세요.");
      return Object.assign({}, matches[0], { slot: String(choice.slot || ""), source_code: exact.length ? "" : ((/[A-F]\d+$/i.exec(raw) || [])[0] || "") });
    }).filter(Boolean);
    groups.forEach((group) => {
      const selected = interpreted.filter((choice) => choice.group === String(group.key) && choice.semester === String(group.semester));
      if (new Set(selected.map((choice) => choice.subject_id)).size !== Number(group.required_count)) throw apiError_("ACADEMIC_CHOICES_INCOMPLETE", "선택과목의 필수 선택 개수가 맞지 않아 반영하지 않았습니다. 기존 고사 명단은 변경하지 않았습니다.");
    });
    if (student.grade === 3 && interpreted.some((choice) => choice.slot === "확통/심수" && choice.group !== "MATH")) throw apiError_("ACADEMIC_CHOICES_INCOMPLETE", "학적 통합의 확통·심수 선택 열을 확인하세요.");
    if (interpreted.length || groups.length) canonical.set(uid, interpreted);
  });
  const courseKeys = new Set();
  function addCourse(uid, name, subjectId, sectionId) {
    const student = studentByUid.get(String(uid)); if (!student) return;
    name = addSubject(student.grade, name);
    const key = uid + "|" + academicNameKey_(name);
    if (courseKeys.has(key)) return;
    courseKeys.add(key);
    const candidateSection = sections.get(String(sectionId || ""));
    const section = candidateSection && String(candidateSection.subject_id) === String(subjectId) ? candidateSection : null;
    courses.push({ source_student_uid: String(uid), grade: student.grade, subject_name: name,
      source_subject_id: String(subjectId || ""), source_section_id: section ? String(section.section_id) : "", section_name: section ? String(section.section_name || "") : "" });
  }
  tables.enrollments.filter((row) => academicActive_(row)).filter((row) => !row.status || ["ENROLLED", "ACTIVE", "CONFIRMED"].includes(row.status)).forEach((row) => {
    const subject = sourceSubjects.get(String(row.subject_id));
    if (!subject || !academicActive_(subject)) return;
    const semester = String(subject.code || "").split(":")[0];
    if (/^[12]$/.test(semester) && !inSemester(semester)) return;
    if (canonical.has(String(row.student_uid))) return;
    addCourse(row.student_uid, subject.name, row.subject_id, row.section_id);
  });
  canonical.forEach((choices, uid) => {
    choices.forEach((choice) => {
      const matchingSections = Array.from(sections.values()).filter((section) => {
        const meta = (tables.section_map || {})[section.section_id];
        return meta && String(section.subject_id) === String(choice.subject_id) && String(meta.subject_id) === String(choice.subject_id) && String(meta.semester) === String(choice.semester) && String(meta.slot) === String(choice.slot) && String(meta.source_code || "") === String(choice.source_code);
      });
      if (matchingSections.length > 1) warnings.ambiguous_sections += 1;
      const sectionId = matchingSections.length === 1 ? matchingSections[0].section_id : "";
      addCourse(uid, choice.name, choice.subject_id, sectionId);
    });
  });
  // Validate coverage from the active roster, not only from existing choice rows.
  // A missing canonical row must not turn an incomplete fallback into an empty replacement.
  students.forEach((student) => {
    const groups = plans.get(student.grade) || [];
    if (canonical.has(student.source_student_uid)) return;
    const selected = courses.filter((course) => course.source_student_uid === student.source_student_uid);
    const candidates = choiceCatalog.get(student.grade) || [];
    groups.forEach((group) => {
      const names = new Set(academicList_(group.subjects).map(academicNameKey_));
      const count = new Set(selected.filter((course) => names.has(academicNameKey_(course.subject_name))).map((course) => academicNameKey_(course.subject_name))).size;
      if (count !== Number(group.required_count)) throw apiError_("ACADEMIC_CHOICES_INCOMPLETE", "선택과목 자료가 없는 학생 또는 필수 선택 개수 오류가 있어 반영하지 않았습니다. 학적 통합에서 전체 학생의 선택과목을 확인하세요.");
    });
    if (groups.length && selected.some((course) => !candidates.some((candidate) => academicNameKey_(candidate.name) === academicNameKey_(course.subject_name)))) throw apiError_("ACADEMIC_CHOICES_INCOMPLETE", "수강 원장에 현재 교육과정과 맞지 않는 과목이 있어 반영하지 않았습니다.");
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
  classes.filter((row) => !activeClassIds.has(row.id)).forEach((row) => { row.active = false; });
  const used = new Set();
  snapshot.students.forEach((row) => {
    const classId = classIds.get(row.class_key);
    const linked = students.filter((item) => item.source_student_uid === row.source_student_uid && Number(item.source_academic_year) === year);
    if (linked.length > 1) throw apiError_("ACADEMIC_SOURCE_CONFLICT", "기존 고사 명단의 학적 고유번호 연결이 중복되어 있습니다.");
    let old = linked[0];
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
  const managedClasses = new Set(previous.classes.filter(adoptable).map((row) => String(row.id)));
  const managedStudents = new Set(previous.students.filter((row) => Number(row.source_academic_year) === year || managedClasses.has(String(row.class_id))).map((row) => String(row.id)));
  const enrollments = previous.enrollments.filter((row) => !(Number(row.source_academic_year) === year || managedClasses.has(String(row.class_id)) || managedStudents.has(String(row.student_id))));
  snapshot.courses.forEach((row) => {
    const student = sourceStudents.get(row.source_student_uid);
    const old = previous.enrollments.find((item) => String(item.student_id) === String(student.id) && academicNameKey_(item.subject_name) === academicNameKey_(row.subject_name));
    const timetable = previous.timetable.filter((item) => Number(item.grade) === row.grade && academicNameKey_(item.subject_name) === academicNameKey_(row.subject_name));
    const names = Array.from(new Set(timetable.map((item) => item.subject_name)));
    if (names.length > 1) throw apiError_("ACADEMIC_SUBJECT_AMBIGUOUS", "시간표에서 같은 과목의 표기를 통일하세요.");
    enrollments.push({ ...row, id: old ? old.id : createId_("enrollment"), class_id: student.class_id, student_id: student.id,
      subject_name: names[0] || row.subject_name, room_name: old ? String(old.room_name || "") : "", source_academic_year: year, semester: snapshot.semester, updated_at: now });
  });
  if (new Set(enrollments.map((row) => String(row.id))).size !== enrollments.length) throw apiError_("ACADEMIC_SOURCE_CONFLICT", "수강 연결 ID가 중복되어 반영하지 않았습니다.");
  return { classes: classes, students: students, enrollments: enrollments };
}

var academicWriteInProgress_ = false;

function assertAcademicReadable_() {
  if (academicWriteInProgress_) return;
  const journal = parseJson_(PropertiesService.getScriptProperties().getProperty("ACADEMIC_SYNC_JOURNAL"), null);
  if (journal && journal.stage !== "committed") throw apiError_("RECOVERY_REQUIRED", "학적 자료 복구가 필요해 조회·저장을 중지했습니다. 관리자가 학적 통합 연결에서 다시 동기화하세요.");
}

function assertAcademicCurrent_() {
  const status = academicLinkStatus_();
  if (status.configured && (!status.synced_at || status.period_mismatch || Number(status.integration_version || 0) < 2)) throw apiError_("ACADEMIC_SYNC_REQUIRED", "고사 학년도·학기에 맞는 학적 자료를 먼저 동기화하세요. 이전 기간 또는 검증 전 명단으로는 고사 작업을 진행할 수 없습니다.");
}

function snapshotAbsenceIdentity_(absences, students, classes) {
  const studentMap = new Map(students.map((row) => [String(row.id), row]));
  const classMap = new Map(classes.map((row) => [String(row.id), row]));
  return absences.map((row) => {
    const student = studentMap.get(String(row.student_id)) || {}, group = classMap.get(String(row.class_id)) || {};
    return Object.assign({}, row, { student_name_snapshot: row.student_name_snapshot || student.name || "",
      student_number_snapshot: row.student_number_snapshot || student.number || "", grade_snapshot: row.grade_snapshot || group.grade || "", class_num_snapshot: row.class_num_snapshot || group.class_num || "" });
  });
}

function academicProjection_(rows) {
  return JSON.stringify(rows.map((row) => Object.fromEntries(Object.keys(row).sort().filter((key) => key !== "updated_at" && row[key] !== "" && row[key] != null).map((key) => [key, String(row[key])]))).sort((a, b) => String(a.id || a.key).localeCompare(String(b.id || b.key))));
}

function restoreAcademicBackup_(journal) {
  const backup = SpreadsheetApp.openById(journal.backup_id), failed = [];
  [SHEETS.CLASSES, SHEETS.STUDENTS, SHEETS.ENROLLMENTS, SHEETS.ABSENCES, SHEETS.SETTINGS].forEach((name) => {
    try {
      const rows = readAcademicTable_(backup, name, HEADERS[name]).map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [key, normalizeCell_(value)])));
      replaceRecords_(name, rows);
    } catch (ignored) { failed.push(name); }
  });
  try { SpreadsheetApp.flush(); } catch (ignored) { failed.push("flush"); }
  try {
    const properties = PropertiesService.getScriptProperties();
    if (journal.previous_status != null) properties.setProperty("ACADEMIC_SYNC_STATUS", String(journal.previous_status));
    else properties.deleteProperty("ACADEMIC_SYNC_STATUS");
  } catch (ignored) { failed.push("status"); }
  if (failed.length) throw apiError_("RECOVERY_REQUIRED", "일부 학적 자료를 복구하지 못해 조회·저장을 중지했습니다. 다시 동기화하거나 관리자에게 문의하세요.");
}

function rememberAcademicCommit_(properties, journal) {
  // Keep bounded durable receipts; cache eviction or a lost response must not replay a commit.
  const receipts = parseJson_(properties.getProperty("ACADEMIC_SYNC_RECEIPTS"), []).filter((row) => Date.now() - Number(row.at) < APP.IDEMPOTENCY_TTL_SECONDS * 1000);
  if (journal.request_key) receipts.push({ key: journal.request_key, at: Date.now() });
  properties.setProperty("ACADEMIC_SYNC_RECEIPTS", JSON.stringify(receipts.slice(-40)));
  properties.deleteProperty("ACADEMIC_SYNC_JOURNAL");
}

function syncAcademicData_(force, requestId, actorId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(force ? 28000 : 1000)) { if (force) throw apiError_("BUSY", "다른 저장이 진행 중입니다. 다시 시도하세요."); return; }
  const properties = PropertiesService.getScriptProperties();
  const requestKey = force && requestId ? hashSecret_(String(actorId) + "|" + requestId, "academic-request-v2") : "";
  let status = {}, journal = null;
  academicWriteInProgress_ = true;
  try {
    const pending = parseJson_(properties.getProperty("ACADEMIC_SYNC_JOURNAL"), null);
    if (pending) {
      if (pending.stage !== "committed") restoreAcademicBackup_(pending);
      else if (pending.status) properties.setProperty("ACADEMIC_SYNC_STATUS", JSON.stringify(pending.status));
      if (pending.stage === "committed") {
        try { rememberAcademicCommit_(properties, pending); }
        catch (ignored) { if (!force || pending.request_key === requestKey) return pending.status; throw apiError_("BUSY", "이전 동기화 완료 상태를 정리하는 중입니다. 잠시 후 다시 시도하세요."); }
      }
      else properties.deleteProperty("ACADEMIC_SYNC_JOURNAL");
    }
    // A resumed administrator session may be the first request after deployment.
    // Recover first, then migrate under this same lock before reading/planning any replacement.
    ensureSchemaOnRequest_();
    if (requestKey && parseJson_(properties.getProperty("ACADEMIC_SYNC_RECEIPTS"), []).some((row) => row.key === requestKey && Date.now() - Number(row.at) < APP.IDEMPOTENCY_TTL_SECONDS * 1000)) return;
    const config = academicSourceConfig_();
    if (!config.spreadsheet_id) { if (force) throw apiError_("ACADEMIC_NOT_CONFIGURED", "학적 원장 연결이 설정되지 않았습니다."); return; }
    status = academicLinkStatus_();
    if (!force && status.state === "error" && Date.now() - new Date(status.attempted_at || 0).getTime() < 60000) return;
    if (!force && status.integration_version === 2 && status.state === "connected" && !status.period_mismatch && Date.now() - new Date(status.synced_at || 0).getTime() < 300000) return;
    const snapshot = readAcademicSnapshot_(config), now = nowIso_();
    const before = { classes: getRecordsFresh_(SHEETS.CLASSES), students: getRecordsFresh_(SHEETS.STUDENTS), enrollments: getRecordsFresh_(SHEETS.ENROLLMENTS), timetable: getRecordsFresh_(SHEETS.TIMETABLE), settings: getRecordsFresh_(SHEETS.SETTINGS), absences: getRecordsFresh_(SHEETS.ABSENCES) };
    const next = planAcademicMerge_(snapshot, before, now);
    next.absences = snapshotAbsenceIdentity_(before.absences, before.students, before.classes);
    const settings = Object.fromEntries(before.settings.map((row) => [String(row.key), row.value]));
    const catalog = normalizeStoredSubjectCatalog_(settings.subject_catalog_json), priorCatalog = parseJson_(settings.academic_catalog_json, {});
    Object.keys(snapshot.catalog).forEach((grade) => { catalog[grade] = Array.from(new Set([...(snapshot.catalog[grade] || []), ...(catalog[grade] || []).filter((name) => !(priorCatalog[grade] || []).includes(name))])).sort((a, b) => a.localeCompare(b, "ko")); });
    const connected = { state: "connected", integration_version: 2, academic_year: snapshot.academic_year, semester: snapshot.semester, synced_at: now, commit_id: createId_("academic"),
      counts: { classes: snapshot.classes.length, students: snapshot.students.length, enrollments: snapshot.courses.length, subjects: Object.values(snapshot.catalog).reduce((sum, names) => sum + names.length, 0) }, warnings: snapshot.warnings };
    const changed = ["classes", "students", "enrollments", "absences"].some((key) => academicProjection_(before[key]) !== academicProjection_(next[key])) || JSON.stringify(catalog) !== String(settings.subject_catalog_json || "") || JSON.stringify(snapshot.catalog) !== String(settings.academic_catalog_json || "");
    if (changed) {
      const backup = getSpreadsheet_().copy("고사 통합 - 학적 반영 전 자동백업 " + Utilities.formatDate(new Date(), APP.TIMEZONE, "yyyyMMdd-HHmmss"));
      const pendingJournal = { stage: "writing", backup_id: backup.getId(), started_at: now, request_key: requestKey, previous_status: properties.getProperty("ACADEMIC_SYNC_STATUS") };
      properties.setProperty("ACADEMIC_SYNC_JOURNAL", JSON.stringify(pendingJournal));
      journal = pendingJournal;
      if (!properties.getProperty("ACADEMIC_INITIAL_BACKUP_SHEET_ID")) properties.setProperty("ACADEMIC_INITIAL_BACKUP_SHEET_ID", backup.getId());
      const updates = { subject_catalog_json: JSON.stringify(catalog), academic_catalog_json: JSON.stringify(snapshot.catalog), academic_sync_status_json: JSON.stringify(connected) };
      const nextSettings = before.settings.filter((row) => !Object.prototype.hasOwnProperty.call(updates, row.key)).concat(Object.entries(updates).map(([key, value]) => ({ key: key, value: value, updated_at: now })));
      replaceRecords_(SHEETS.CLASSES, next.classes); replaceRecords_(SHEETS.STUDENTS, next.students); replaceRecords_(SHEETS.ENROLLMENTS, next.enrollments); replaceRecords_(SHEETS.ABSENCES, next.absences); replaceRecords_(SHEETS.SETTINGS, nextSettings);
      SpreadsheetApp.flush();
    }
    properties.setProperty("ACADEMIC_SYNC_STATUS", JSON.stringify(connected));
    journal = Object.assign({}, journal || {}, { stage: "committed", request_key: requestKey, status: connected });
    properties.setProperty("ACADEMIC_SYNC_JOURNAL", JSON.stringify(journal));
    try { rememberAcademicCommit_(properties, journal); } catch (ignored) {} // A committed marker remains durable if receipt cleanup fails.
    safeAppendAudit_("system", "sync_academic_data", "records", "", snapshot.students.length + snapshot.courses.length, true, "OK", actorId || "academic_sync");
    return connected;
  } catch (error) {
    if (journal && journal.stage !== "committed") {
      try { restoreAcademicBackup_(journal); properties.deleteProperty("ACADEMIC_SYNC_JOURNAL"); }
      catch (recoveryError) { error = recoveryError; }
    }
    try { properties.setProperty("ACADEMIC_SYNC_STATUS", JSON.stringify(Object.assign({}, status, { state: "error", attempted_at: nowIso_(), error_code: error.code || "ACADEMIC_SYNC_FAILED" }))); } catch (ignored) {}
    safeAppendAudit_("system", "sync_academic_data", "records", "", 0, false, error.code || "ACADEMIC_SYNC_FAILED", actorId || "academic_sync");
    if (force || error.code === "RECOVERY_REQUIRED") throw error;
  } finally { academicWriteInProgress_ = false; lock.releaseLock(); }
}
