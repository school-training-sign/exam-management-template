import fs from 'node:fs';
import test from 'node:test';
import assert from 'node:assert/strict';
const source=['Config','Repository','Code','Academic'].map(name=>fs.readFileSync(new URL(`../apps-script/${name}.gs`,import.meta.url),'utf8')).join('\n');
function harness(failure,publicFailure=false){
 const called=[];
 const implementation=new Function('called','failure','publicFailure',source+`\n
 PropertiesService={getScriptProperties:()=>({getProperty:()=>null})};
 const injected=label=>{called.push(label);if(failure===label){const error=publicFailure?apiError_('INVALID_ADMIN_PASSWORD','관리자 암호가 올바르지 않습니다.'):new TypeError('PRIVATE_WORKBOOK_IDENTIFIER PRIVATE_STUDENT_NAME PRIVATE_CREDENTIAL');error.stack='TypeError: PRIVATE_WORKBOOK_IDENTIFIER\\n    at ensureSheets_ (Repository:42:3)\\n    at PRIVATE_STUDENT_NAME (Code:1:2)';throw error;}};
 verifyAdminPassword_=()=>injected('verify');ensureSchemaOnRequest_=()=>injected('schema');syncAcademicData_=()=>injected('sync');
 createUserSession_=()=>{injected('user_session');return 'synthetic-user-token';};createAdminSession_=()=>{injected('admin_session');return 'synthetic-admin-token';};safeAppendAudit_=()=>{};
 withDataRead_=fn=>fn();adminBootstrap_=()=>{injected('bootstrap');return {settings:{school_name:'합성 학교'}};};
 bootstrapForUser_=()=>{injected('bootstrap');return {settings:{school_name:'합성 학교'}};};
 return {entry:adminEntryLogin_,elevate:adminLogin_,phase:runAdminLoginPhase_,error:apiError_};
 `)(called,failure,publicFailure);
 return{...implementation,called};
}
test('실제 관리자 직접 진입은 기존 암호 확인·설정·동기화·세션·초기 조회 순서를 유지한다',()=>{
 const h=harness();const result=h.entry({password:'synthetic-credential'});assert.equal(result.user_session,'synthetic-user-token');assert.equal(result.admin_session,'synthetic-admin-token');assert.deepEqual(h.called,['verify','schema','sync','user_session','admin_session','bootstrap']);
});
test('잘못된 관리자 암호는 기존 공개 오류를 그대로 반환하고 후속 처리를 실행하지 않는다',()=>{
 const h=harness('verify',true);assert.throws(()=>h.entry({password:'synthetic-credential'}),{code:'INVALID_ADMIN_PASSWORD'});assert.deepEqual(h.called,['verify']);
});
test('예상 밖 로그인 단계 오류는 짧은 고정 진단번호만 공개하고 원문·스택·개인정보를 숨긴다',()=>{
 for(const [stage,code]of [['verify','A1'],['schema','A3'],['sync','A4'],['user_session','A5'],['admin_session','A6'],['bootstrap','A7']]){
  const h=harness(stage);let caught;try{h.entry({password:'synthetic-credential'});}catch(error){caught=error;}
  assert.equal(caught.code,'ADMIN_ENTRY_FAILED');assert.equal(caught.details.diagnostic,code+'-T-SI');
  const visible=JSON.stringify({code:caught.code,message:caught.message,details:caught.details});assert.ok(!/PRIVATE_|synthetic-credential|Repository:|Code:/.test(visible));
 }
});
test('복구·시도 제한 등 의도된 공개 오류의 객체·코드를 바꾸거나 재시도하지 않는다',()=>{
 const h=harness();for(const code of ['RECOVERY_REQUIRED','BUSY','LOGIN_RATE_LIMITED']){const original=h.error(code,'합성 공개 오류');let count=0;assert.throws(()=>h.phase('recovery',()=>{count++;throw original;}),error=>error===original);assert.equal(count,1);}
});
test('권한 승격 중 초기 조회 오류도 동일한 보호 진단을 반환한다',()=>{
 const h=harness('bootstrap');assert.throws(()=>h.elevate({password:'synthetic-credential'},{user_id:'synthetic-user'}),error=>error.code==='ADMIN_ENTRY_FAILED'&&error.details.diagnostic==='A7-T-SI');
});
test('알 수 없는 예외 종류·스택 함수명은 진단 값으로 유출되지 않는다',()=>{
 const h=harness();assert.throws(()=>h.phase('unknown',()=>{const error=Error('PRIVATE_DATA');error.name='PRIVATE_NAME';error.stack='PRIVATE_DATA\\n at PRIVATE_DATA (PRIVATE_PATH:1)';throw error;}),error=>error.details.diagnostic==='A0-E-X'&&!error.message.includes('PRIVATE'));
});
