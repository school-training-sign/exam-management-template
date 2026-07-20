import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const files = ["Config.gs", "Repository.gs", "Code.gs"];
const source = Object.fromEntries(
  files.map((name) => [name, fs.readFileSync(path.join(root, "apps-script", name), "utf8")]),
);
const all = Object.values(source).join("\n");

test("Apps Script 파일이 JavaScript 문법으로 파싱된다", () => {
  for (const [name, content] of Object.entries(source)) {
    assert.doesNotThrow(() => new Function(content), name);
  }
});

test("학교코드와 관리자 암호 원문이 배포 파일에 없다", () => {
  assert.equal(
    /(?:school[_-]?code|학교코드)\s*[:=]\s*["'][A-Z0-9_-]{6,}["']/i.test(all),
    false,
  );
  assert.equal(/ADMIN_PASSWORD\s*[:=]\s*["'][^"']+/.test(all), false);
});

test("요청 응답 형식과 권한 액션이 명시되어 있다", () => {
  assert.match(source["Code.gs"], /ok:\s*true,\s*data:/);
  assert.match(source["Code.gs"], /ok:\s*false,\s*data:\s*null,\s*error:/);
  for (const action of ["school_login", "admin_login", "presence_ping", "submit_absence_input", "get_hq_status", "save_seat_charts_batch", "cleanup"]) {
    assert.match(all, new RegExp(action));
  }
  assert.doesNotMatch(source["Config.gs"], /"set_absence"|"complete_input"/);
  assert.doesNotMatch(source["Code.gs"], /case "set_absence"|case "complete_input"/);
  assert.match(source["Config.gs"], /ADMIN_ACTIONS/);
});

test("모든 쓰기에 잠금과 중복 방지 키를 사용한다", () => {
  assert.match(source["Code.gs"], /MUTATION_ACTIONS\.has\(action\)/);
  assert.match(source["Code.gs"], /LockService\.getScriptLock\(\)/);
  assert.match(source["Code.gs"], /request_id/);
  assert.match(source["Code.gs"], /cachedAfterLock/);
});

test("소금값 해시와 5분 세션 만료 정책이 있다", () => {
  assert.match(source["Code.gs"], /DigestAlgorithm\.SHA_256/);
  assert.match(source["Config.gs"], /TEACHER_TTL_SECONDS:\s*300/);
  assert.match(source["Config.gs"], /ADMIN_TTL_SECONDS:\s*300/);
  assert.match(source["Code.gs"], /CacheService\.getScriptCache\(\)/);
});

test("계획한 11개 시트 이름과 감사로그 비식별 필드를 포함한다", () => {
  for (const name of ["사용설명서", "설정", "학급", "학생", "고사일", "시간표", "선택과목", "결시", "입력완료", "자리배치", "감사로그"]) {
    assert.ok(all.includes(`"${name}"`), name);
  }
  assert.match(source["Config.gs"], /\["id", "occurred_at", "role", "action", "target_type", "target_id", "count", "success", "detail"\]/);
});

test("배포용 원본 통합문서가 포함되어 있다", () => {
  const workbook = path.join(root, "학교 정기고사 관리 시스템 - 배포용 원본.xlsx");
  assert.ok(fs.existsSync(workbook));
  assert.ok(fs.statSync(workbook).size > 10_000);
});

test("기존 시트에 revision 열과 schema_version을 안전하게 추가한다", () => {
  assert.match(source["Config.gs"], /\[SHEETS\.COMPLETIONS\]:\s*\[[^\]]*"revision"/s);
  assert.match(source["Config.gs"], /SCHEMA_VERSION:\s*"2"/);
  assert.match(source["Repository.gs"], /requiredHeaders\.filter\(\(header\) => existing\.indexOf\(header\) === -1\)/);
  assert.match(source["Repository.gs"], /setSetting_\("schema_version", APP\.SCHEMA_VERSION\)/);
});

test("결시 일괄 제출은 revision 충돌을 검사하고 한 번에 상태를 교체한다", () => {
  assert.match(source["Code.gs"], /case "submit_absence_input": return submitAbsenceInput_\(request\)/);
  assert.match(source["Code.gs"], /expected_revision/);
  assert.match(source["Code.gs"], /REVISION_CONFLICT/);
  assert.match(source["Code.gs"], /replaceAbsenceStateAtomically_/);
  assert.match(source["Code.gs"], /submitted:\s*submitted/);
  assert.match(source["Code.gs"], /const preservedInactive = oldAbsences\.filter/);
  assert.match(source["Code.gs"], /absences:\s*incoming/);
  assert.match(source["Code.gs"], /new Set\(\["인정", "미인정", "질병", "기타"\]\)/);
  assert.match(source["Code.gs"], /ABSENCE_REASON_DETAIL_REQUIRED/);
});

test("학생 개별 저장·비활성화와 명단 ID 재사용 계약이 있다", () => {
  for (const action of ["save_student", "delete_student", "replace_students"]) {
    assert.match(source["Config.gs"], new RegExp(`"${action}"`));
  }
  assert.match(source["Code.gs"], /const existing = requestedById \|\| byNumber\.get\(number\)/);
  assert.match(source["Code.gs"], /active:\s*false/);
  assert.match(source["Code.gs"], /preview_only/);
});

test("고사일·학생·학급 삭제는 영향 건수를 미리 반환하고 참조 자료를 정리한다", () => {
  assert.match(source["Code.gs"], /case "delete_exam_date": return deleteExamDate_\(request\)/);
  assert.match(source["Code.gs"], /function examDateDeleteImpact_/);
  assert.match(source["Code.gs"], /function classDeleteImpact_/);
  assert.match(source["Code.gs"], /function studentDeleteImpact_/);
  assert.match(source["Code.gs"], /function timetableDeleteImpact_/);
  assert.match(source["Code.gs"], /case "delete_timetable": return deleteTimetable_\(request\)/);
  assert.match(source["Code.gs"], /EXAM_DATE_DELETE_ACTION_REQUIRED/);
  assert.match(source["Code.gs"], /TIMETABLE_DELETE_ACTION_REQUIRED/);
  assert.match(source["Code.gs"], /request\.preview_only/);
  assert.match(source["Code.gs"], /SHEETS\.SEATING/);
});

test("시간표 학급 참조와 선택과목 미매칭 행 오류를 검증한다", () => {
  assert.match(source["Code.gs"], /TIMETABLE_CLASS_REQUIRED/);
  assert.match(source["Code.gs"], /TIMETABLE_DATE_NOT_FOUND/);
  assert.match(source["Code.gs"], /TIMETABLE_CLASS_NOT_FOUND/);
  assert.match(source["Code.gs"], /TIMETABLE_CLASS_GRADE_MISMATCH/);
  for (const code of ["CLASS_NOT_FOUND", "STUDENT_NOT_FOUND", "STUDENT_NAME_MISMATCH", "SUBJECT_NOT_FOUND"]) {
    assert.match(source["Code.gs"], new RegExp(`code: "${code}"`));
  }
  assert.match(source["Code.gs"], /SUBJECT_CLASS_MISMATCH/);
  assert.match(source["Code.gs"], /ENROLLMENT_VALIDATION_FAILED/);
  assert.match(source["Code.gs"], /row_number:\s*rowNumber/);
});

test("고사본부와 기간 보고서가 제출 상태와 집계를 함께 반환한다", () => {
  assert.match(source["Code.gs"], /period_totals:\s*periodTotals/);
  assert.match(source["Code.gs"], /class_status:\s*classStatus/);
  assert.match(source["Code.gs"], /reason_summary:\s*aggregateRows_/);
  assert.match(source["Code.gs"], /class_summary:\s*aggregateRows_/);
  assert.match(source["Code.gs"], /date_summary:\s*aggregateRows_/);
});

test("설정 데이터 캐시, 관리자 세션 버전, 로그인 제한이 적용된다", () => {
  assert.match(source["Config.gs"], /CACHEABLE_SHEETS/);
  assert.match(source["Repository.gs"], /invalidateDataCaches_/);
  assert.match(source["Code.gs"], /PROPERTY_ADMIN_SESSION_VERSION/);
  assert.match(source["Code.gs"], /LOGIN_RATE_MAX_ATTEMPTS/);
  assert.match(source["Code.gs"], /LOGIN_RATE_GLOBAL_MAX_ATTEMPTS/);
  assert.match(source["Code.gs"], /LOGIN_RATE_LIMIT/);
  assert.match(source["Code.gs"], /function recordLoginFailure_[\s\S]*LockService\.getScriptLock\(\)/);
});

test("접속 상태 heartbeat는 인증 세션을 짧은 캐시에만 집계한다", () => {
  assert.match(source["Config.gs"], /PRESENCE_WINDOW_SECONDS:\s*90/);
  assert.match(source["Config.gs"], /PRESENCE_MAX_ENTRIES:\s*250/);
  assert.match(source["Code.gs"], /action === "presence_ping"[\s\S]*requireSession_\(request\.school_session,\s*"teacher"/);
  assert.match(source["Code.gs"], /request\.user_active === true/);
  assert.match(source["Code.gs"], /request\.action !== "presence_ping"/);
  assert.match(source["Code.gs"], /leavePresence_\(request\.school_session\)/);
  assert.match(source["Code.gs"], /function requireSession_\(token,\s*role,\s*touch\)/);

  const mutationBlock = source["Config.gs"].match(/const MUTATION_ACTIONS[\s\S]*?\]\);/)?.[0] || "";
  const adminBlock = source["Config.gs"].match(/const ADMIN_ACTIONS[\s\S]*?\]\);/)?.[0] || "";
  assert.doesNotMatch(mutationBlock, /presence_ping/);
  assert.doesNotMatch(adminBlock, /presence_ping/);

  const helpers = new Function(
    "Utilities",
    `${all}\nreturn { prunePresenceRegistry_, presenceResponse_ };`,
  )({ formatDate: () => "2030-01-01T00:00:00+09:00" });
  const now = Date.now();
  const active = helpers.prunePresenceRegistry_({
    mine: now,
    other: now - 10_000,
    expired: now - 91_000,
  }, now);
  assert.deepEqual(Object.keys(active), ["mine", "other"]);
  assert.equal(helpers.presenceResponse_(active, "mine", false).online_count, 1);
  assert.equal(helpers.presenceResponse_({ mine: now }, "mine", false).online_count, 0);

  const crowded = Object.fromEntries(
    Array.from({ length: 300 }, (_, index) => [`client-${index}`, now - index]),
  );
  assert.equal(Object.keys(helpers.prunePresenceRegistry_(crowded, now)).length, 250);
});

test("자리배치는 6행 5열 기본값과 자연키 upsert, 참조 검증을 사용한다", () => {
  assert.match(source["Repository.gs"], /rows:\s*Number\(chart\.rows\) \|\| 6/);
  assert.match(source["Repository.gs"], /cols:\s*Number\(chart\.cols\) \|\| 5/);
  assert.match(source["Code.gs"], /seatNaturalKey_/);
  assert.match(source["Code.gs"], /SEAT_STUDENT_NOT_FOUND/);
  assert.match(source["Code.gs"], /SEAT_CAPACITY_EXCEEDED/);
  assert.match(source["Code.gs"], /INVALID_ASSIGNMENT_SEAT/);
  assert.match(source["Code.gs"], /ASSIGNMENT_ON_DISABLED_SEAT/);
  assert.match(source["Code.gs"], /DUPLICATE_ASSIGNMENT_SEAT/);
});

test("집계와 학급 ID 정규화 헬퍼가 기대한 결과를 만든다", () => {
  const helpers = new Function(`${all}\nreturn { aggregateRows_, normalizeClassIds_, seatNaturalKey_ };`)();
  assert.deepEqual(
    helpers.normalizeClassIds_('["class_a","class_b","class_a"]'),
    ["class_a", "class_b"],
  );
  assert.deepEqual(
    helpers.normalizeClassIds_("class_a|class_b"),
    ["class_a", "class_b"],
  );
  assert.deepEqual(
    helpers.aggregateRows_([
      { reason: "질병" },
      { reason: "질병" },
      { reason: "인정" },
    ], ["reason"]),
    [
      { count: 1, reason: "인정" },
      { count: 2, reason: "질병" },
    ],
  );
  assert.equal(
    helpers.seatNaturalKey_({ mode: "own", exam_date: "2030-01-01", period: 1, grade: 3, class_id: "c1" }),
    "own|2030-01-01|1|3|c1||",
  );
});
