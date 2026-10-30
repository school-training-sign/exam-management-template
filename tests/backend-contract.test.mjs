import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const files = ["Config.gs", "Repository.gs", "Code.gs", "Academic.gs"];
const source = Object.fromEntries(
  files.map((name) => [name, fs.readFileSync(path.join(root, "apps-script", name), "utf8")]),
);
const all = Object.values(source).join("\n");

function readPublicTextFiles(directory, output = []) {
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    if ([".git", "node_modules"].includes(entry.name)) continue;
    const full = path.join(directory, entry.name);
    if (entry.isDirectory()) readPublicTextFiles(full, output);
    else if ([".gs", ".js", ".mjs", ".json", ".md", ".yaml", ".yml"].includes(path.extname(entry.name))) {
      output.push(fs.readFileSync(full, "utf8"));
    }
  }
  return output;
}

function createLoginRateHarness(startMs = 2_000_000_000_000) {
  const clock = { now: startMs };
  const cacheValues = new Map();
  const propertyValues = new Map();
  const cache = {
    get(key) {
      const item = cacheValues.get(key);
      if (!item) return null;
      if (item.expires_at <= clock.now) {
        cacheValues.delete(key);
        return null;
      }
      return item.value;
    },
    put(key, value, ttlSeconds) {
      cacheValues.set(key, { value: String(value), expires_at: clock.now + Number(ttlSeconds) * 1000 });
    },
    remove: (key) => cacheValues.delete(key),
    removeAll: (keys) => keys.forEach((key) => cacheValues.delete(key)),
  };
  const properties = {
    getProperties: () => Object.fromEntries(propertyValues),
    getProperty: (key) => propertyValues.get(key) || null,
    setProperty: (key, value) => propertyValues.set(key, String(value)),
    setProperties(values, deleteAllOthers) {
      if (deleteAllOthers) propertyValues.clear();
      Object.entries(values).forEach(([key, value]) => propertyValues.set(key, String(value)));
    },
    deleteProperty: (key) => propertyValues.delete(key),
  };
  const lock = { tryLock: () => true, releaseLock: () => {} };
  const utilities = {
    DigestAlgorithm: { SHA_256: "SHA_256" },
    Charset: { UTF_8: "UTF_8" },
    computeDigest(_algorithm, value) {
      return createHash("sha256").update(String(value), "utf8").digest();
    },
    base64EncodeWebSafe: (bytes) => Buffer.from(bytes).toString("base64url"),
  };
  const helpers = new Function(
    "CacheService",
    "PropertiesService",
    "LockService",
    "Utilities",
    "clock",
    `${all}
    nowMillis_ = () => clock.now;
    return {
      loginRateKeys_, loginRateStorageKey_, assertLoginAllowed_, recordLoginFailure_, clearLoginFailures_,
      readLoginRateState_, writeLoginRateState_
    };`,
  )(
    { getScriptCache: () => cache },
    { getScriptProperties: () => properties },
    { getScriptLock: () => lock },
    utilities,
    clock,
  );
  return { helpers, clock, cacheValues, propertyValues };
}

test("Apps Script 파일이 JavaScript 문법으로 파싱된다", () => {
  for (const [name, content] of Object.entries(source)) {
    assert.doesNotThrow(() => new Function(content), name);
  }
});

test("이전 학교코드 로그인 계약과 관리자 암호 원문이 배포 파일에 없다", () => {
  assert.equal(
    /(?:school[_-]?code|학교코드)\s*[:=]\s*["'][A-Z0-9_-]{6,}["']/i.test(all),
    false,
  );
  assert.equal(/ADMIN_PASSWORD\s*[:=]\s*["'][^"']+/.test(all), false);
  assert.doesNotMatch(all, /school_login|school_session|request\.school_code/);
});

test("공개 배포 자료에 운영 학교코드·외부 공유표 ID·Google Sheets 주소가 없다", () => {
  const publicText = readPublicTextFiles(root).join("\n");
  const retiredSchoolCode = ["HANYANG", "3520"].join("");
  const sourceSheetId = ["1ukMUIjGhRILLj", "XwpTmxm3NecSz", "-5cD3khRnD1Xk6Nx8"].join("");
  const sheetsPrefix = ["https://docs.google.com/", "spreadsheets/d/"].join("");
  assert.equal(publicText.includes(retiredSchoolCode), false);
  assert.equal(publicText.includes(sourceSheetId), false);
  assert.equal(publicText.includes(sheetsPrefix), false);
});

test("요청 응답 형식과 권한 액션이 명시되어 있다", () => {
  assert.match(source["Code.gs"], /ok:\s*true,\s*data:/);
  assert.match(source["Code.gs"], /ok:\s*false,\s*data:\s*null,\s*error:/);
  for (const action of ["user_login", "admin_entry_login", "admin_login", "presence_ping", "submit_absence_input", "get_hq_status", "save_seat_charts_batch", "cleanup"]) {
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

test("PIN 소금값·비공개 pepper 해시와 5분 세션 만료 정책이 있다", () => {
  assert.match(source["Code.gs"], /DigestAlgorithm\.SHA_256/);
  assert.match(source["Config.gs"], /USER_TTL_SECONDS:\s*300/);
  assert.match(source["Config.gs"], /ADMIN_TTL_SECONDS:\s*300/);
  assert.match(source["Config.gs"], /PROPERTY_ACCESS_USER_PEPPER:\s*"ACCESS_USER_PEPPER"/);
  assert.match(source["Code.gs"], /function hashAccessPin_[\s\S]*PROPERTY_ACCESS_USER_PEPPER/);
  assert.match(source["Code.gs"], /CacheService\.getScriptCache\(\)/);
});

test("접속 이름과 숫자 4~6자리 PIN 로그인은 단일 실패 코드와 세 가지 시도 제한을 사용한다", () => {
  const loginBlock = source["Code.gs"].match(/function userLogin_[\s\S]*?\n}\n\nfunction verifyAdminPassword_/)?.[0] || "";
  assert.match(loginBlock, /request\.login_name/);
  assert.match(loginBlock, /\^\\d\{4,6\}\$/);
  assert.match(loginBlock, /INVALID_LOGIN/);
  assert.doesNotMatch(loginBlock, /USER_NOT_FOUND|USER_INACTIVE|INVALID_PIN/);
  assert.match(source["Code.gs"], /:client:/);
  assert.match(source["Code.gs"], /:name:/);
  assert.match(source["Code.gs"], /:global/);
  assert.match(source["Code.gs"], /login-name-rate-v1/);
});

test("브라우저·이름 제한은 Cache가 사라져도 8회 실패 뒤 차단한다", () => {
  const runtime = createLoginRateHarness();
  const keys = runtime.helpers.loginRateKeys_("user", { client_id: "raw-client-id" }, "테스트 교사");
  for (let index = 0; index < 8; index += 1) {
    assert.doesNotThrow(() => runtime.helpers.assertLoginAllowed_(keys));
    runtime.helpers.recordLoginFailure_(keys);
  }
  runtime.cacheValues.clear();
  assert.throws(() => runtime.helpers.assertLoginAllowed_(keys), { code: "LOGIN_RATE_LIMIT" });

  const persisted = [...runtime.propertyValues.entries()]
    .filter(([key]) => key.startsWith("LOGIN_RATE_STATE_"))
    .map(([, value]) => value)
    .join("");
  assert.doesNotMatch(persisted, /raw-client-id|테스트 교사/);
});

test("서로 다른 브라우저와 이름도 전역 40회 실패 뒤 차단한다", () => {
  const runtime = createLoginRateHarness();
  for (let index = 0; index < 40; index += 1) {
    const keys = runtime.helpers.loginRateKeys_("user", { client_id: `client-${index}` }, `교사-${index}`);
    assert.doesNotThrow(() => runtime.helpers.assertLoginAllowed_(keys));
    runtime.helpers.recordLoginFailure_(keys);
  }
  const nextKeys = runtime.helpers.loginRateKeys_("user", { client_id: "new-client" }, "새 교사");
  assert.throws(() => runtime.helpers.assertLoginAllowed_(nextKeys), { code: "LOGIN_RATE_LIMIT" });
});

test("로그인 성공 시 브라우저·이름 카운터만 초기화하고 전역 카운터는 유지한다", () => {
  const runtime = createLoginRateHarness();
  const keys = runtime.helpers.loginRateKeys_("user", { client_id: "client-success" }, "성공 교사");
  for (let index = 0; index < 3; index += 1) runtime.helpers.recordLoginFailure_(keys);
  runtime.helpers.clearLoginFailures_(keys);

  const state = runtime.helpers.readLoginRateState_(runtime.clock.now).entries;
  const clientKey = runtime.helpers.loginRateStorageKey_(keys[0].key);
  const nameKey = runtime.helpers.loginRateStorageKey_(keys[1].key);
  const globalKey = runtime.helpers.loginRateStorageKey_(keys[2].key);
  assert.equal(state[clientKey], undefined);
  assert.equal(state[nameKey], undefined);
  assert.equal(state[globalKey].count, 3);
  assert.equal(runtime.cacheValues.has(keys[0].key), false);
  assert.equal(runtime.cacheValues.has(keys[1].key), false);
  assert.equal(runtime.cacheValues.has(keys[2].key), true);
});

test("로그인 제한 카운터는 마지막 실패 5분 뒤 만료된다", () => {
  const runtime = createLoginRateHarness();
  const keys = runtime.helpers.loginRateKeys_("user", { client_id: "client-expiry" }, "만료 교사");
  for (let index = 0; index < 8; index += 1) runtime.helpers.recordLoginFailure_(keys);
  assert.throws(() => runtime.helpers.assertLoginAllowed_(keys), { code: "LOGIN_RATE_LIMIT" });

  runtime.clock.now += 300_001;
  assert.doesNotThrow(() => runtime.helpers.assertLoginAllowed_(keys));
  const state = runtime.helpers.readLoginRateState_(runtime.clock.now).entries;
  keys.forEach((item) => assert.equal(state[runtime.helpers.loginRateStorageKey_(item.key)], undefined));
});

test("영속 로그인 제한 상태는 최근 해시 키 600개로 제한되고 shard 크기를 분산한다", () => {
  assert.match(source["Config.gs"], /LOGIN_RATE_MAX_ENTRIES:\s*600/);
  assert.match(source["Config.gs"], /LOGIN_RATE_STATE_SHARDS:\s*12/);
  assert.match(source["Config.gs"], /PROPERTY_LOGIN_RATE_STATE_PREFIX:\s*"LOGIN_RATE_STATE_"/);
  const runtime = createLoginRateHarness();
  const entries = {};
  for (let index = 0; index < 700; index += 1) {
    const key = createHash("sha256").update(`entry-${index}`).digest("base64url").slice(0, 32);
    entries[key] = { count: 1, expires_at: runtime.clock.now + 300_000, updated_at: runtime.clock.now + index };
  }
  runtime.helpers.writeLoginRateState_(entries, runtime.clock.now);
  assert.equal(Object.keys(runtime.helpers.readLoginRateState_(runtime.clock.now).entries).length, 600);
  const shardValues = [...runtime.propertyValues.entries()]
    .filter(([key]) => key.startsWith("LOGIN_RATE_STATE_"))
    .map(([, value]) => value);
  assert.equal(shardValues.length, 12);
  assert.ok(shardValues.every((value) => value.length < 9_000));
});

test("접속 이름 정규화와 PIN 검증 헬퍼가 기대한 결과를 만든다", () => {
  const helpers = new Function(`${all}\nreturn { canonicalLoginName_, loginLookupKey_, validatePin_, safeAccessUser_ };`)();
  assert.equal(helpers.canonicalLoginName_("  Ｔｅａｃｈｅｒ   김  "), "Teacher 김");
  assert.equal(helpers.loginLookupKey_("  ＴＥＡＣＨＥＲ 김  "), "teacher 김");
  assert.equal(helpers.validatePin_("1234"), "1234");
  assert.equal(helpers.validatePin_("12345"), "12345");
  assert.equal(helpers.validatePin_("123456"), "123456");
  assert.throws(() => helpers.validatePin_("123"), { code: "INVALID_PIN" });
  assert.throws(() => helpers.validatePin_("1234567"), { code: "INVALID_PIN" });
  assert.throws(() => helpers.validatePin_("12가4"), { code: "INVALID_PIN" });
  assert.deepEqual(
    helpers.safeAccessUser_({ id: "u1", login_name: "김민수", active: true, updated_at: "now", pin_salt: "secret", pin_hash: "secret" }),
    { id: "u1", login_name: "김민수", active: true, updated_at: "now" },
  );
});

test("접속 사용자 추가·수정·비활성화·PIN 재설정이 세션 버전을 올린다", () => {
  const users = [];
  let uuid = 0;
  const helpers = new Function(
    "Utilities",
    "users",
    `${all}
    getAccessUsers_ = () => users.map((user) => ({ ...user }));
    findAccessUserById_ = (id) => users.find((user) => user.id === String(id)) || null;
    writeAccessUser_ = (user) => {
      const index = users.findIndex((item) => item.id === user.id);
      if (index >= 0) users[index] = { ...user };
      else users.push({ ...user });
      return user;
    };
    hashAccessPin_ = (pin, salt) => salt + ":" + pin;
    adminBootstrap_ = () => ({ access_users: users.map(safeAccessUser_) });
    nowIso_ = () => "2030-01-01T00:00:00+09:00";
    randomToken_ = () => "salt";
    return { saveAccessUser_, setAccessUserActive_, resetAccessUserPin_ };`,
  )({ getUuid: () => `uuid-${++uuid}` }, users);

  helpers.saveAccessUser_({ login_name: " 김민수 ", pin: "123456" });
  assert.equal(users.length, 1);
  assert.equal(users[0].login_name, "김민수");
  assert.equal(users[0].session_version, 1);
  assert.throws(() => helpers.saveAccessUser_({ login_name: "  김민수  ", pin: "654321" }), { code: "DUPLICATE_LOGIN_NAME" });

  helpers.saveAccessUser_({ id: users[0].id, login_name: "김민수(1학년)" });
  assert.equal(users[0].session_version, 2);
  helpers.resetAccessUserPin_({ id: users[0].id, pin: "654321" });
  assert.equal(users[0].session_version, 3);
  helpers.setAccessUserActive_({ id: users[0].id, active: false });
  assert.equal(users[0].active, false);
  assert.equal(users[0].session_version, 4);
});

test("사용자 세션은 비활성화·버전 변경·관리자 암호 변경 직후 거부된다", () => {
  const cacheValues = new Map();
  const cache = {
    get: (key) => cacheValues.get(key) || null,
    put: (key, value) => cacheValues.set(key, value),
    remove: (key) => cacheValues.delete(key),
  };
  const properties = new Map([["ADMIN_SESSION_VERSION", "1"]]);
  const users = new Map();
  let uuid = 0;
  const helpers = new Function(
    "CacheService",
    "PropertiesService",
    "Utilities",
    "users",
    `${all}
    findAccessUserById_ = (id) => users.get(String(id)) || null;
    return { createUserSession_, requireSession_ };`,
  )(
    { getScriptCache: () => cache },
    { getScriptProperties: () => ({ getProperty: (key) => properties.get(key) || null }) },
    {
      getUuid: () => `uuid-${++uuid}`,
      formatDate: () => "2030-01-01T00:00:00+09:00",
    },
    users,
  );

  const user = { id: "u1", login_name: "김민수", active: true, session_version: 1 };
  users.set(user.id, user);
  const disabledToken = helpers.createUserSession_(user);
  user.active = false;
  assert.throws(() => helpers.requireSession_(disabledToken, "user"), { code: "SESSION_EXPIRED" });

  user.active = true;
  const changedToken = helpers.createUserSession_(user);
  user.session_version += 1;
  assert.throws(() => helpers.requireSession_(changedToken, "user"), { code: "SESSION_EXPIRED" });

  const adminToken = helpers.createUserSession_({ id: "system_admin", login_name: "관리자" }, true);
  properties.set("ADMIN_SESSION_VERSION", "2");
  assert.throws(() => helpers.requireSession_(adminToken, "user"), { code: "SESSION_EXPIRED" });
});

test("접속 사용자는 설정 시트 JSON에 저장하고 공개 설정에서 제외한다", () => {
  assert.match(source["Config.gs"], /ACCESS_USER_SETTING_PREFIX:\s*"access_user:"/);
  assert.match(source["Repository.gs"], /setSetting_\(accessUserSettingKey_\(stored\.id\), JSON\.stringify\(stored\)\)/);
  assert.match(source["Repository.gs"], /function getPublicSettings_[\s\S]*\["school_name", "app_name", "timezone", "version", "schema_version"\]/);
  assert.match(source["Repository.gs"], /settings:\s*getPublicSettings_\(\)/);
  assert.match(source["Repository.gs"], /result\.access_users = accessUsersForAdmin_\(\)/);
  assert.doesNotMatch(source["Repository.gs"].match(/function safeAccessUser_[\s\S]*?\n}/)?.[0] || "", /pin_salt|pin_hash|normalized_login_name/);
});

test("관리자 직접 진입은 사용자가 없어도 사용자·관리자 세션과 현재 사용자를 발급한다", () => {
  const block = source["Code.gs"].match(/function adminEntryLogin_[\s\S]*?\n}\n\nfunction adminLogin_/)?.[0] || "";
  assert.match(block, /createUserSession_\(currentUser, true\)/);
  assert.match(block, /createAdminSession_\(currentUser\.id\)/);
  assert.match(block, /user_session:/);
  assert.match(block, /admin_session:/);
  assert.match(block, /current_user:/);
  assert.match(block, /adminBootstrap_\(\)/);
  assert.doesNotMatch(block, /getAccessUsers_|findAccessUser/);
});

test("접속 사용자 관리 액션은 관리자 전용 잠금 작업이며 완전 삭제 액션은 없다", () => {
  const mutationBlock = source["Config.gs"].match(/const MUTATION_ACTIONS[\s\S]*?\]\);/)?.[0] || "";
  const adminBlock = source["Config.gs"].match(/const ADMIN_ACTIONS[\s\S]*?\]\);/)?.[0] || "";
  for (const action of ["save_access_user", "set_access_user_active", "reset_access_user_pin"]) {
    assert.match(mutationBlock, new RegExp(`"${action}"`));
    assert.match(adminBlock, new RegExp(`"${action}"`));
    assert.match(source["Code.gs"], new RegExp(`case "${action}"`));
  }
  assert.doesNotMatch(all, /delete_access_user/);
  assert.match(source["Code.gs"], /DUPLICATE_LOGIN_NAME/);
  assert.match(source["Code.gs"], /user\.session_version \+= 1/);
});

test("사용자 세션은 사용자 ID·버전을 검증하고 잠금 획득 뒤 다시 확인한다", () => {
  assert.match(source["Code.gs"], /role:\s*"user"[\s\S]*user_id:[\s\S]*session_version:/);
  assert.match(source["Code.gs"], /findAccessUserById_\(session\.user_id\)/);
  assert.match(source["Code.gs"], /!currentUser\.active/);
  assert.match(source["Code.gs"], /currentUser\.session_version[\s\S]*session\.session_version/);
  const mutationBlock = source["Code.gs"].match(/function withMutation_[\s\S]*?\n}\n\nfunction inferCount_/)?.[0] || "";
  assert.match(mutationBlock, /tryLock\(28000\)[\s\S]*requireSession_\(request\.user_session, "user", false\)/);
  assert.match(mutationBlock, /requireSession_\(request\.admin_session, "admin", false\)/);
});

test("세션 복원 bootstrap과 로그인 응답에 current_user가 포함된다", () => {
  assert.match(source["Code.gs"], /case "get_bootstrap": syncAcademicData_\(false\); return bootstrapForUser_\(false, userSession\)/);
  assert.match(source["Code.gs"], /case "get_admin_bootstrap": syncAcademicData_\(false\); return bootstrapForUser_\(true, userSession\)/);
  assert.match(source["Code.gs"], /bootstrapForCurrentUser_\(teacherBootstrap_\(\), currentUser\)/);
  assert.match(source["Code.gs"], /bootstrapForCurrentUser_\(adminBootstrap_\(\), currentUser\)/);
});

test("계획한 11개 시트 이름과 감사로그 비식별 필드를 포함한다", () => {
  for (const name of ["사용설명서", "설정", "학급", "학생", "고사일", "시간표", "선택과목", "결시", "입력완료", "자리배치", "감사로그"]) {
    assert.ok(all.includes(`"${name}"`), name);
  }
  assert.match(source["Config.gs"], /\["id", "occurred_at", "role", "actor_id", "action", "target_type", "target_id", "count", "success", "detail"\]/);
  assert.match(source["Repository.gs"], /headers\.map\(\(header\).*record\[header\]/);
});

test("배포용 원본 통합문서가 포함되어 있다", () => {
  const workbook = path.join(root, "학교 정기고사 관리 시스템 - 배포용 원본.xlsx");
  assert.ok(fs.existsSync(workbook));
  assert.ok(fs.statSync(workbook).size > 10_000);
});

test("기존 시트에 revision·seat_order 열과 schema_version을 안전하게 추가한다", () => {
  assert.match(source["Config.gs"], /\[SHEETS\.COMPLETIONS\]:\s*\[[^\]]*"revision"/s);
  assert.match(source["Config.gs"], /\[SHEETS\.SEATING\]:\s*\[[^\]]*"seat_order"/s);
  assert.match(source["Config.gs"], /SCHEMA_VERSION:\s*"7"/);
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
  for (const name of ["assertLoginAllowed_", "recordLoginFailure_", "clearLoginFailures_"]) {
    const block = source["Code.gs"].match(new RegExp(`function ${name}[\\s\\S]*?\\n}`))?.[0] || "";
    assert.match(block, /LockService\.getScriptLock\(\)/);
    assert.match(block, /tryLock\(5000\)/);
    assert.match(block, /finally[\s\S]*lock\.releaseLock\(\)/);
  }
});

test("접속 상태 heartbeat는 인증 세션을 짧은 캐시에만 집계한다", () => {
  assert.match(source["Config.gs"], /PRESENCE_WINDOW_SECONDS:\s*90/);
  assert.match(source["Config.gs"], /PRESENCE_MAX_ENTRIES:\s*250/);
  assert.match(source["Code.gs"], /action === "presence_ping"[\s\S]*requireSession_\(request\.user_session,\s*"user"/);
  assert.match(source["Code.gs"], /request\.user_active === true/);
  assert.match(source["Code.gs"], /request\.action !== "presence_ping"/);
  assert.match(source["Code.gs"], /leavePresence_\(request\.user_session,\s*"user"\)/);
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

test("자리배치는 두 순번 방식, 6행 5열 기본값, 자연키 upsert와 참조 검증을 사용한다", () => {
  assert.match(source["Repository.gs"], /rows:\s*Number\(chart\.rows\) \|\| 6/);
  assert.match(source["Repository.gs"], /cols:\s*Number\(chart\.cols\) \|\| 5/);
  assert.match(source["Repository.gs"], /seat_order:\s*chart\.seat_order === "column" \? "column" : "row"/);
  assert.match(source["Repository.gs"], /seat_order:\s*row\.seat_order === "column" \? "column" : "row"/);
  assert.match(source["Code.gs"], /seatNaturalKey_/);
  assert.match(source["Code.gs"], /INVALID_SEAT_ORDER/);
  assert.match(source["Code.gs"], /SEAT_STUDENT_NOT_FOUND/);
  assert.match(source["Code.gs"], /SEAT_CAPACITY_EXCEEDED/);
  assert.match(source["Code.gs"], /INVALID_ASSIGNMENT_SEAT/);
  assert.match(source["Code.gs"], /ASSIGNMENT_ON_DISABLED_SEAT/);
  assert.match(source["Code.gs"], /DUPLICATE_ASSIGNMENT_SEAT/);
  assert.match(source["Code.gs"], /INVALID_ASSIGNMENT_SEQUENCE/);
  assert.match(source["Code.gs"], /DUPLICATE_ASSIGNMENT_SEQUENCE/);
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
  assert.equal(
    helpers.seatNaturalKey_({ mode: "own", exam_date: "2030-01-01", period: 1, grade: 3, class_id: "c1", seat_order: "column" }),
    "own|2030-01-01|1|3|c1||",
  );
});

test("교과목·가정통신문 저장 액션은 관리자 전용 잠금 작업으로 라우팅된다", () => {
  const mutationBlock = source["Config.gs"].match(/const MUTATION_ACTIONS[\s\S]*?\]\);/)?.[0] || "";
  const adminBlock = source["Config.gs"].match(/const ADMIN_ACTIONS[\s\S]*?\]\);/)?.[0] || "";
  for (const action of ["save_subject_catalog", "save_exam_notice_settings"]) {
    assert.match(mutationBlock, new RegExp(`"${action}"`));
    assert.match(adminBlock, new RegExp(`"${action}"`));
    assert.match(source["Code.gs"], new RegExp(`case "${action}"`));
  }
  assert.match(source["Code.gs"], /setSetting_\("subject_catalog_json", JSON\.stringify\(catalog\)\)/);
  assert.match(source["Code.gs"], /setSetting_\("exam_notice_json", JSON\.stringify\(notice\)\)/);
  assert.match(source["Repository.gs"], /result\.subject_catalog = normalizeStoredSubjectCatalog_/);
  assert.match(source["Repository.gs"], /result\.exam_notice = normalizeStoredExamNotice_/);
  assert.match(source["Code.gs"], /Array\.isArray\(request\.subjects\)[\s\S]*request\.subjects\.length/);
});

test("교과목 카탈로그는 학년·개수·길이·trim 중복을 검증한다", () => {
  const helpers = new Function(
    `${all}\nreturn { normalizeSubjectCatalogInput_, normalizeStoredSubjectCatalog_, inferCount_ };`,
  )();
  assert.deepEqual(
    helpers.normalizeSubjectCatalogInput_(1, [" 국어 ", "영어"]),
    { grade: 1, subjects: ["국어", "영어"] },
  );
  assert.deepEqual(helpers.normalizeSubjectCatalogInput_(9, []), { grade: 9, subjects: [] });
  assert.deepEqual(
    helpers.normalizeStoredSubjectCatalog_('{"1":[" 국어 ","국어","영어"],"10":["제외"]}'),
    { "1": ["국어", "영어"] },
  );
  assert.equal(helpers.inferCount_({ subjects: ["국어", "영어"] }), 2);
  assert.throws(() => helpers.normalizeSubjectCatalogInput_(0, []), { code: "INVALID_GRADE" });
  assert.throws(() => helpers.normalizeSubjectCatalogInput_(1, "국어"), { code: "INVALID_SUBJECTS" });
  assert.throws(() => helpers.normalizeSubjectCatalogInput_(1, Array(201).fill("과목")), { code: "TOO_MANY_SUBJECTS" });
  assert.throws(() => helpers.normalizeSubjectCatalogInput_(1, ["국어", " 국어 "]), { code: "DUPLICATE_SUBJECT" });
  assert.throws(() => helpers.normalizeSubjectCatalogInput_(1, [" "]), { code: "INVALID_SUBJECT_NAME" });
  assert.throws(() => helpers.normalizeSubjectCatalogInput_(1, ["가".repeat(81)]), { code: "INVALID_SUBJECT_NAME" });
});

test("가정통신문 설정은 표준 형태와 모든 길이·날짜 제한을 적용한다", () => {
  const helpers = new Function(
    `${all}\nreturn { normalizeExamNoticeInput_, normalizeStoredExamNotice_, isValidIsoDate_ };`,
  )();
  const notice = helpers.normalizeExamNoticeInput_({
    year: "2026",
    semester: "2",
    exam_name: " 2학기 중간고사 ",
    greeting: " 안내드립니다. ",
    notes: [" 준비물을 확인하세요. "],
    issue_date: "2026-09-01",
    issuer: " 한양대학교사범대학부속고등학교장 ",
  });
  assert.deepEqual(notice, {
    year: 2026,
    semester: 2,
    exam_name: "2학기 중간고사",
    greeting: "안내드립니다.",
    notes: ["준비물을 확인하세요."],
    issue_date: "2026-09-01",
    issuer: "한양대학교사범대학부속고등학교장",
  });
  assert.equal(Object.hasOwn(notice, "academic_year"), false);
  assert.equal(helpers.isValidIsoDate_("2026-02-28"), true);
  assert.equal(helpers.isValidIsoDate_("2026-02-30"), false);
  assert.deepEqual(helpers.normalizeStoredExamNotice_("{}", "한양대학교사범대학부속고등학교"), {
    year: "",
    semester: "",
    exam_name: "",
    greeting: "",
    notes: [],
    issue_date: "",
    issuer: "한양대학교사범대학부속고등학교장",
  });

  const base = { year: 2026, semester: 1, exam_name: "중간고사", greeting: "", notes: [], issue_date: "", issuer: "학교장" };
  assert.throws(() => helpers.normalizeExamNoticeInput_({ ...base, year: 1899 }), { code: "INVALID_ACADEMIC_YEAR" });
  assert.throws(() => helpers.normalizeExamNoticeInput_({ ...base, semester: 3 }), { code: "INVALID_SEMESTER" });
  assert.throws(() => helpers.normalizeExamNoticeInput_({ ...base, exam_name: "" }), { code: "EXAM_NAME_REQUIRED" });
  assert.throws(() => helpers.normalizeExamNoticeInput_({ ...base, exam_name: "가".repeat(81) }), { code: "INVALID_EXAM_NAME" });
  assert.throws(() => helpers.normalizeExamNoticeInput_({ ...base, greeting: "가".repeat(1501) }), { code: "INVALID_GREETING" });
  assert.throws(() => helpers.normalizeExamNoticeInput_({ ...base, notes: Array(21).fill("안내") }), { code: "INVALID_NOTICE_NOTES" });
  assert.throws(() => helpers.normalizeExamNoticeInput_({ ...base, notes: [""] }), { code: "INVALID_NOTICE_NOTE" });
  assert.throws(() => helpers.normalizeExamNoticeInput_({ ...base, issue_date: "2026-02-30" }), { code: "INVALID_ISSUE_DATE" });
  assert.throws(() => helpers.normalizeExamNoticeInput_({ ...base, issuer: "" }), { code: "NOTICE_ISSUER_REQUIRED" });
  assert.throws(() => helpers.normalizeExamNoticeInput_({ ...base, issuer: "가".repeat(101) }), { code: "INVALID_NOTICE_ISSUER" });
});

test("새 설정은 시트 추가 없이 설정 JSON으로 저장된다", () => {
  assert.match(source["Config.gs"], /VERSION:\s*"3\.1\.1"/);
  assert.match(source["Config.gs"], /SCHEMA_VERSION:\s*"7"/);
  assert.match(source["Config.gs"], /\[SHEETS\.SETTINGS\]: \["key", "value", "updated_at"\]/);
  assert.equal((source["Config.gs"].match(/\[SHEETS\.[A-Z_]+\]: \[/g) || []).length, 10);
  assert.match(source["Repository.gs"], /setSetting_\("version", APP\.VERSION\)/);
  assert.match(source["Repository.gs"], /function setSetting_[\s\S]*replaceRecords_\(SHEETS\.SETTINGS, rows\)/);
});

test("초기 설정과 스키마 마이그레이션이 새 시스템 이름과 인증 구조를 적용한다", () => {
  assert.match(source["Config.gs"], /SYSTEM_NAME:\s*"한양대학교사범대학부속고등학교 통합 시스템"/);
  const initBlock = source["Code.gs"].match(/function initializeExamSystem[\s\S]*?\n}\n\nfunction changeAdminPassword/)?.[0] || "";
  assert.match(initBlock, /schoolNamePrompt/);
  assert.match(initBlock, /passwordPrompt/);
  assert.doesNotMatch(initBlock, /schoolCodePrompt|request\.school_code/);
  assert.match(initBlock, /setSetting_\("app_name", APP\.SYSTEM_NAME\)/);
  assert.match(source["Repository.gs"], /oldGuideRange\.getMergedRanges\(\)\.forEach\(\(range\) => range\.breakApart\(\)\)[\s\S]*oldGuideRange\.clearContent\(\)/);
  assert.doesNotMatch(source["Repository.gs"], /oldGuideRange\.breakApart\(\)/);
  assert.match(source["Repository.gs"], /PROPERTY_ACCESS_USER_PEPPER/);
  assert.match(source["Repository.gs"], /\["SCHOOL", "CODE", "SALT"\]\.join\("_"\)/);
});

test("관리자 암호는 초기 설정과 변경 화면에서 4자 이상을 요구한다", () => {
  const initBlock = source["Code.gs"].match(/function initializeExamSystem[\s\S]*?\n}\n\nfunction changeAdminPassword/)?.[0] || "";
  const changeBlock = source["Code.gs"].match(/function changeAdminPassword[\s\S]*?\n}\n\nfunction showDeploymentGuide/)?.[0] || "";
  assert.match(initBlock, /password\.length < 4/);
  assert.match(initBlock, /관리자 암호를 4자 이상/);
  assert.match(changeBlock, /password\.length < 4/);
  assert.match(changeBlock, /관리자 암호를 4자 이상/);
  assert.match(source["Repository.gs"], /4자 이상의 관리자 암호/);
  assert.match(fs.readFileSync(path.join(root, "README.md"), "utf8"), /4자 이상의 관리자 암호/);
  assert.doesNotMatch(source["Code.gs"], /password\.length < 8|8자 이상|10자 이상/);
});

test("기존 6자리 PIN은 호환하면서 로그인·등록·재설정이 숫자 4~6자리를 허용한다", () => {
  const loginBlock = source["Code.gs"].match(/function userLogin_[\s\S]*?\n}\n\nfunction verifyAdminPassword_/)?.[0] || "";
  const validateBlock = source["Code.gs"].match(/function validatePin_[\s\S]*?\n}/)?.[0] || "";
  assert.match(loginBlock, /\^\\d\{4,6\}\$/);
  assert.match(validateBlock, /\^\\d\{4,6\}\$/);
  assert.match(validateBlock, /숫자 4~6자리/);
  assert.match(source["Code.gs"], /function saveAccessUser_[\s\S]*validatePin_\(request\.pin\)/);
  assert.match(source["Code.gs"], /function resetAccessUserPin_[\s\S]*validatePin_\(request\.pin\)/);
  assert.match(source["Repository.gs"], /guideText\.includes\("PIN 6자리"\)/);
  assert.match(source["Repository.gs"], /숫자 4~6자리 개인 PIN/);
  assert.doesNotMatch(source["Code.gs"], /\^\\d\{6\}\$|개인 PIN은 숫자 6자리/);
});

test("원안 포장 스키마는 기존 11개 시트를 유지하면서 시간표 열만 안전하게 확장한다", () => {
  assert.match(source["Config.gs"], /VERSION:\s*"3\.1\.1"/);
  assert.match(source["Config.gs"], /SCHEMA_VERSION:\s*"7"/);
  const timetableHeaders = source["Config.gs"].match(/\[SHEETS\.TIMETABLE\]:\s*\[([^\]]+)\]/s)?.[1] || "";
  for (const header of [
    "representative_teacher", "answer_sheet_type", "packaging_date", "packaging_slot_id",
    "packaging_revision", "packaging_updated_at",
  ]) assert.match(timetableHeaders, new RegExp(`"${header}"`));
  assert.equal((source["Config.gs"].match(/\[SHEETS\.[A-Z_]+\]: \[/g) || []).length, 10);
  assert.match(source["Repository.gs"], /requiredHeaders\.filter\(\(header\) => existing\.indexOf\(header\) === -1\)/);
});

test("원안 포장 여섯 API가 일반·관리자·비밀 링크 권한에 맞게 라우팅된다", () => {
  for (const action of [
    "redeem_exam_packaging_invite", "get_exam_packaging", "save_exam_packaging_items",
    "save_exam_packaging_config", "create_exam_packaging_invite", "disable_exam_packaging_invite",
  ]) assert.match(source["Code.gs"], new RegExp(action));
  const mutationBlock = source["Config.gs"].match(/const MUTATION_ACTIONS[\s\S]*?\]\);/)?.[0] || "";
  const adminBlock = source["Config.gs"].match(/const ADMIN_ACTIONS[\s\S]*?\]\);/)?.[0] || "";
  assert.match(mutationBlock, /"save_exam_packaging_items"/);
  for (const action of ["save_exam_packaging_config", "create_exam_packaging_invite", "disable_exam_packaging_invite"]) {
    assert.match(mutationBlock, new RegExp(`"${action}"`));
    assert.match(adminBlock, new RegExp(`"${action}"`));
  }
  assert.doesNotMatch(adminBlock, /"save_exam_packaging_items"/);
  assert.match(source["Code.gs"], /if \(request\.packaging_session\)[\s\S]*get_exam_packaging[\s\S]*save_exam_packaging_items[\s\S]*throw apiError_\("FORBIDDEN"/);
});

test("시간표 일괄 저장은 기존 ID의 원안 포장 필드를 보존하고 요청 필드로 덮어쓰지 않는다", () => {
  const block = source["Code.gs"].match(/function saveTimetable_[\s\S]*?\n}\n\nfunction normalizeSubjectCatalogInput_/)?.[0] || "";
  assert.match(block, /const oldById = new Map/);
  assert.match(block, /const preserved = .*oldById/);
  for (const field of [
    "representative_teacher", "answer_sheet_type", "packaging_date", "packaging_slot_id",
    "packaging_revision", "packaging_updated_at",
  ]) assert.match(block, new RegExp(`${field}:`));
  assert.doesNotMatch(block, /item\.representative_teacher|item\.answer_sheet_type|item\.packaging_date|item\.packaging_slot_id/);
});

test("원안 포장 설정은 설정 JSON에 revision·마감·정원·날짜·행·담당자를 저장한다", () => {
  assert.match(source["Repository.gs"], /exam_packaging_config_json/);
  assert.match(source["Code.gs"], /setSetting_\("exam_packaging_config_json", JSON\.stringify\(next\)\)/);
  assert.match(source["Repository.gs"], /capacity:\s*3/);
  assert.match(source["Repository.gs"], /kind:\s*row\[1\]/);
  assert.match(source["Repository.gs"], /staff_assignments:/);
  assert.match(source["Repository.gs"], /result\.exam_packaging_config = getExamPackagingConfig_\(\)/);
});

test("원안 포장 설정 정규화는 날짜·시간대 중복과 정원 오류를 거부하고 쉬는 시간 담당자를 허용한다", () => {
  const helpers = new Function(`${all}\nreturn { normalizeExamPackagingConfigInput_, examPackagingConfigImpact_, apiError_ };`)();
  const base = {
    input_deadline: "2030-05-01T17:00:00+09:00",
    capacity: 3,
    packaging_dates: ["2030-05-02"],
    rows: [
      { id: "slot_a", kind: "slot", period_label: "1회", start_time: "09:00", end_time: "09:25", sort_order: 1 },
      { id: "break_a", kind: "break", period_label: "휴식", start_time: "09:25", end_time: "09:40", sort_order: 2 },
    ],
    staff_assignments: [
      { packaging_date: "2030-05-02", slot_id: "slot_a", staff_name: "가상담당" },
      { packaging_date: "2030-05-02", slot_id: "break_a", staff_name: "가상휴식담당" },
    ],
  };
  const normalized = helpers.normalizeExamPackagingConfigInput_(base);
  assert.equal(normalized.capacity, 3);
  assert.equal(normalized.rows[1].kind, "break");
  assert.equal(normalized.staff_assignments[1].slot_id, "break_a");
  assert.throws(() => helpers.normalizeExamPackagingConfigInput_({ ...base, capacity: 0 }), { code: "INVALID_PACKAGING_CAPACITY" });
  assert.throws(() => helpers.normalizeExamPackagingConfigInput_({ ...base, packaging_dates: ["2030-05-02", "2030-05-02"] }), { code: "DUPLICATE_PACKAGING_DATE" });
  assert.throws(() => helpers.normalizeExamPackagingConfigInput_({ ...base, rows: [base.rows[0], { ...base.rows[0], id: "slot_b" }] }), { code: "DUPLICATE_PACKAGING_TIME" });
  assert.throws(() => helpers.normalizeExamPackagingConfigInput_({ ...base, rows: [{ ...base.rows[0], end_time: "09:30" }] }), { code: "INVALID_PACKAGING_SLOT_DURATION" });
  assert.throws(() => helpers.normalizeExamPackagingConfigInput_({ ...base, rows: [base.rows[0], { ...base.rows[1], start_time: "09:20" }] }), { code: "OVERLAPPING_PACKAGING_TIME" });
  assert.deepEqual(
    helpers.examPackagingConfigImpact_(normalized, [
      { id: "ok", packaging_date: "2030-05-02", packaging_slot_id: "slot_a" },
      { id: "bad", packaging_date: "2030-05-03", packaging_slot_id: "slot_a" },
    ]),
    { affected_count: 1, assigned_items: 1, timetable_ids: ["bad"] },
  );
});

test("항목 저장은 설정·항목 revision과 쉬는 시간·정원 3건을 한 잠금 안에서 검사한다", () => {
  const block = source["Code.gs"].match(/function saveExamPackagingItems_[\s\S]*?\n}\n\nfunction cleanup_/)?.[0] || "";
  assert.match(block, /expected_config_revision/);
  assert.match(block, /PACKAGING_CONFIG_CONFLICT/);
  assert.match(block, /incoming\.expected_revision/);
  assert.match(block, /PACKAGING_REVISION_CONFLICT/);
  assert.match(block, /PACKAGING_BREAK_NOT_ASSIGNABLE/);
  assert.match(block, /count > config\.capacity[\s\S]*PACKAGING_SLOT_FULL/);
  const wrapper = source["Code.gs"].match(/function withPackagingMutation_[\s\S]*?\n}\n\nfunction inferCount_/)?.[0] || "";
  assert.match(wrapper, /LockService\.getScriptLock\(\)/);
  assert.match(wrapper, /tryLock\(28000\)/);
  assert.match(wrapper, /requirePackagingSession_\(request\.packaging_session, false\)/);
  assert.match(wrapper, /writeMutationCache_/);
});

test("비밀 링크는 256비트 원문을 해시로만 저장하고 30분 세션·버전 폐기를 적용한다", () => {
  assert.match(source["Config.gs"], /PACKAGING_SESSION_TTL_SECONDS:\s*1800/);
  assert.match(source["Code.gs"], /const token = randomSecret256_\(\)/);
  assert.match(source["Code.gs"], /function randomSecret256_[\s\S]*hashSecret_\(entropy/);
  assert.match(source["Code.gs"], /hashExamPackagingInviteToken_\(token, salt, pepper\)/);
  assert.match(source["Code.gs"], /PROPERTY_PACKAGING_INVITE_HASH/);
  assert.match(source["Code.gs"], /role:\s*"packaging"/);
  assert.match(source["Code.gs"], /Number\(session\.invite_version\) !== currentVersion/);
  assert.match(source["Code.gs"], /getScriptCache\(\)\.put\("session:" \+ token, JSON\.stringify\(session\), APP\.PACKAGING_SESSION_TTL_SECONDS\)/);
  assert.match(source["Code.gs"], /fragment:\s*"#packaging=" \+ token/);
  assert.match(source["Code.gs"], /request\.action === "create_exam_packaging_invite"[\s\S]*__committed:\s*true/);
  assert.doesNotMatch(source["Repository.gs"], /getPublicSettings_[\s\S]*PROPERTY_PACKAGING_INVITE_HASH/);
  assert.match(source["Code.gs"], /active:\s*Boolean\(enabled && !expired\)/);
  assert.match(source["Code.gs"], /action === "save_exam_packaging_items"/);
});

test("비밀 링크 컨텍스트는 학생·학급·접속 사용자 없이 최소 고사 자료와 완료 집계만 반환한다", () => {
  const block = source["Code.gs"].match(/function getExamPackaging_[\s\S]*?\n}\n\nfunction examPackagingInviteStatus_/)?.[0] || "";
  for (const field of ["settings", "exam_title", "exam_dates", "items", "config", "summary", "read_only"]) {
    assert.match(block, new RegExp(`${field}:`));
  }
  assert.doesNotMatch(block, /SHEETS\.STUDENTS|SHEETS\.CLASSES|getAccessUsers_|access_users/);
  const helpers = new Function(`${all}\nreturn { examPackagingSummary_, isPackagingDeadlinePassed_ };`)();
  assert.deepEqual(helpers.examPackagingSummary_([
    { packaging_date: "2030-05-02", packaging_slot_id: "slot_a", answer_sheet_type: "card" },
    { packaging_date: "", packaging_slot_id: "", answer_sheet_type: "" },
  ]), { total: 2, assigned: 1, unassigned: 1, answer_selected: 1, answer_unselected: 1 });
  assert.equal(helpers.isPackagingDeadlinePassed_({ input_deadline: "2030-05-01T00:00:00Z" }, Date.parse("2030-05-02T00:00:00Z")), true);
});

test("원안 포장 전용 접속 현황과 감사로그는 별도 범위와 비식별 세션 ID를 사용한다", () => {
  assert.match(source["Code.gs"], /presence:registry:packaging:v1/);
  assert.match(source["Code.gs"], /request\.packaging_active === true/);
  assert.match(source["Code.gs"], /request\.scope === "exam_packaging" \? "packaging" : "user"/);
  assert.match(source["Code.gs"], /request\.packaging_session \|\| request\.user_session/);
  assert.match(source["Code.gs"], /safeAppendAudit_\("packaging"[\s\S]*issued\.session\.session_id/);
  const wrapper = source["Code.gs"].match(/function withPackagingMutation_[\s\S]*?\n}\n\nfunction inferCount_/)?.[0] || "";
  assert.match(wrapper, /actorId = String\(session\.session_id/);
  assert.doesNotMatch(wrapper, /representative_teacher|staff_name|login_name/);
});
