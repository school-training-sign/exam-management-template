import fs from 'node:fs';
import test from 'node:test';
import assert from 'node:assert/strict';
const code=['Config','Repository','Code','Academic'].map(name=>fs.readFileSync(new URL(`../apps-script/${name}.gs`,import.meta.url),'utf8')).join('\n');
test('이전 검증 버전의 연결 완료도 새 검증 동기화 전에는 고사 입력에 사용하지 않는다',()=>{
 const status={configured:true,synced_at:'2026-10-30T00:00:00Z',integration_version:1,period_mismatch:false};
 const check=new Function('status',code+'\nacademicLinkStatus_=()=>status;return assertAcademicCurrent_;')(status);
 assert.throws(()=>check(),{code:'ACADEMIC_SYNC_REQUIRED'});status.integration_version=2;assert.doesNotThrow(()=>check());
});
