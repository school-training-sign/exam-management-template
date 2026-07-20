const APP = Object.freeze({
  VERSION: "1.0.0",
  TIMEZONE: "Asia/Seoul",
  TEACHER_TTL_SECONDS: 300,
  ADMIN_TTL_SECONDS: 300,
  IDEMPOTENCY_TTL_SECONDS: 21600,
  PROPERTY_SHEET_ID: "SPREADSHEET_ID",
  PROPERTY_SCHOOL_SALT: "SCHOOL_CODE_SALT",
  PROPERTY_SCHOOL_HASH: "SCHOOL_CODE_HASH",
  PROPERTY_ADMIN_SALT: "ADMIN_PASSWORD_SALT",
  PROPERTY_ADMIN_HASH: "ADMIN_PASSWORD_HASH",
});

const SHEETS = Object.freeze({
  GUIDE: "사용설명서",
  SETTINGS: "설정",
  CLASSES: "학급",
  STUDENTS: "학생",
  EXAM_DATES: "고사일",
  TIMETABLE: "시간표",
  ENROLLMENTS: "선택과목",
  ABSENCES: "결시",
  COMPLETIONS: "입력완료",
  SEATING: "자리배치",
  AUDIT: "감사로그",
});

const HEADERS = Object.freeze({
  [SHEETS.SETTINGS]: ["key", "value", "updated_at"],
  [SHEETS.CLASSES]: ["id", "grade", "class_num", "active", "updated_at"],
  [SHEETS.STUDENTS]: ["id", "class_id", "number", "name", "active", "updated_at"],
  [SHEETS.EXAM_DATES]: ["id", "exam_date", "label", "active", "updated_at"],
  [SHEETS.TIMETABLE]: ["id", "exam_date", "grade", "period", "start_time", "end_time", "subject_name", "subject_type", "room_name", "class_ids", "updated_at"],
  [SHEETS.ENROLLMENTS]: ["id", "grade", "class_id", "student_id", "subject_name", "room_name", "updated_at"],
  [SHEETS.ABSENCES]: ["key", "id", "exam_date", "period", "class_id", "student_id", "reason", "reason_detail", "updated_at"],
  [SHEETS.COMPLETIONS]: ["key", "id", "exam_date", "period", "class_id", "completed_at", "updated_at"],
  [SHEETS.SEATING]: ["id", "mode", "exam_date", "period", "grade", "class_id", "subject_name", "room_name", "rows", "cols", "start_side", "disabled_seats_json", "examinee_ids_json", "assignment_json", "updated_at"],
  [SHEETS.AUDIT]: ["id", "occurred_at", "role", "action", "target_type", "target_id", "count", "success", "detail"],
});

const MUTATION_ACTIONS = new Set([
  "set_absence",
  "complete_input",
  "save_exam_dates",
  "save_class",
  "delete_class",
  "replace_students",
  "import_students",
  "save_timetable",
  "save_enrollments",
  "save_seat_chart",
  "save_seat_charts_batch",
  "delete_seat_chart",
  "cleanup",
]);

const ADMIN_ACTIONS = new Set([
  "get_admin_bootstrap",
  "get_hq_status",
  "get_period_report",
  "save_exam_dates",
  "save_class",
  "delete_class",
  "replace_students",
  "import_students",
  "save_timetable",
  "save_enrollments",
  "save_seat_chart",
  "save_seat_charts_batch",
  "delete_seat_chart",
  "cleanup",
]);
