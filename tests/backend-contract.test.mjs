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
  const schoolCode = ["HANYANG", "3520"].join("");
  assert.equal(all.includes(schoolCode), false);
  assert.equal(/ADMIN_PASSWORD\s*[:=]\s*["'][^"']+/.test(all), false);
});

test("요청 응답 형식과 권한 액션이 명시되어 있다", () => {
  assert.match(source["Code.gs"], /ok:\s*true,\s*data:/);
  assert.match(source["Code.gs"], /ok:\s*false,\s*data:\s*null,\s*error:/);
  for (const action of ["school_login", "admin_login", "set_absence", "complete_input", "get_hq_status", "save_seat_charts_batch", "cleanup"]) {
    assert.match(all, new RegExp(action));
  }
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
