import fs from 'node:fs';
import vm from 'node:vm';
import test from 'node:test';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';

// Actual Apps Script functions, deterministic synthetic services, separate execution lock owners.
// This is not an Apps Script load/quotas measurement and contains no school records.
const source=['Config','Repository','Code','Academic'].map(name=>fs.readFileSync(new URL(`../apps-script/${name}.gs`,import.meta.url),'utf8')).join('\n');
const clone=value=>JSON.parse(JSON.stringify(value));
function harness(){
 const state={properties:new Map(),cache:new Map(),tables:new Map(),backups:new Map(),owner:null,sequence:0,reads:0,writes:0,copies:0,hooks:{},audits:[]};
 let nextExecution=0;
 function sheet(tables,name){return {
  setFrozenRows(){},
  getLastColumn:()=>tables.get(name)?.[0]?.length||0,
  getLastRow:()=>{let rows=tables.get(name)||[];let end=rows.length;while(end&&rows[end-1].every(x=>x===''))end--;return end;},
  getDataRange:()=>({getValues:()=>clone(tables.get(name)||[])}),
  getRange(row,column,height,width){return {
   setBackground(){return this;},setFontColor(){return this;},setFontWeight(){return this;},
   getValues(){const values=Array.from({length:height},(_,y)=>Array.from({length:width},(_,x)=>tables.get(name)?.[row-1+y]?.[column-1+x]??''));state.hooks.afterRead?.(name);return values;},
   clearContent(){state.writes++;for(let y=row-1;y<row-1+height;y++)for(let x=column-1;x<column-1+width;x++)tables.get(name)[y][x]='';},
   setValues(values){state.hooks.beforeSet?.(name,values);state.writes++;for(let y=0;y<values.length;y++){if(!tables.get(name)[row-1+y])tables.get(name)[row-1+y]=[];for(let x=0;x<values[y].length;x++)tables.get(name)[row-1+y][column-1+x]=values[y][x];}state.hooks.afterSet?.(name);}
  };}
 };}
 const book=tables=>({getSheetByName:name=>sheet(tables,name),copy(){state.copies++;const id='backup-'+state.copies;state.backups.set(id,new Map([...tables].map(([k,v])=>[k,clone(v)])));return {getId:()=>id};}});
 const properties={getProperty:key=>state.properties.get(key)??null,setProperty(key,value){state.hooks.beforeProperty?.(key,value);state.properties.set(key,String(value));},deleteProperty:key=>state.properties.delete(key),getProperties:()=>Object.fromEntries(state.properties)};
 function execution(){
  const id=++nextExecution;
  const lock={hasLock:()=>state.owner===id,tryLock(){if(state.owner!==null&&state.owner!==id)return false;state.owner=id;return true;},releaseLock(){assert.equal(state.owner,id);state.owner=null;}};
  const context=vm.createContext({console,PropertiesService:{getScriptProperties:()=>properties},LockService:{getScriptLock:()=>lock},CacheService:{getScriptCache:()=>({get:key=>state.cache.get(key)??null,put(key,value){state.hooks.cachePut?.(key,value);state.cache.set(key,String(value));},remove:key=>state.cache.delete(key),removeAll:keys=>keys.forEach(key=>state.cache.delete(key))})},
   Utilities:{getUuid:()=>String(++state.sequence),formatDate:value=>value.toISOString().slice(0,10),DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},computeDigest:(_,value)=>[...createHash('sha256').update(value).digest()],base64EncodeWebSafe:bytes=>Buffer.from(bytes).toString('base64url')},
   SpreadsheetApp:{openById:id=>book(state.backups.get(id)),flush(){state.hooks.flush?.();}},ContentService:{MimeType:{JSON:'json'},createTextOutput:text=>({text,setMimeType(){return this;}})}});
  vm.runInContext(source,context);
  context.getSpreadsheet_=()=>book(state.tables);
  // Skip guide formatting/provisioning only; execute the real authenticated schema lock
  // and real ensureTableSheet_ header migration, including legacy-header rollout tests.
  context.ensureSheets_=()=>{state.schemaCalls=(state.schemaCalls||0)+1;const headers=vm.runInContext('HEADERS',context);for(const [name,columns]of Object.entries(headers))context.ensureTableSheet_(context.getSpreadsheet_(),name,columns);};
  context.academicSourceConfig_=()=>({spreadsheet_id:'synthetic-source-only',academic_year:2026,semester:2});
  context.readAcademicSnapshot_=()=>{state.reads++;return clone(state.snapshot);};
  context.safeAppendAudit_=(...args)=>state.audits.push(args[5]);
  context.appendAudit_=()=>state.audits.push(false);
  context.nowIso_=()=>new Date().toISOString();
  return context;
 }
 const writer=execution(),reader=execution(),{SHEETS:S,HEADERS:H}=vm.runInContext('({SHEETS,HEADERS})',writer);
 function seed(name,rows){state.tables.set(name,[clone(H[name]),...rows.map(row=>H[name].map(key=>row[key]??''))]);}
 for(const name of Object.keys(H))seed(name,[]);
 seed(S.CLASSES,[{id:'class-old',grade:2,class_num:1,active:true,source_academic_year:2026}]);
 seed(S.STUDENTS,[{id:'student-old',class_id:'class-old',number:1,name:'합성 학생',active:true,source_student_uid:'uid-1',source_academic_year:2026}]);
 seed(S.ENROLLMENTS,[{id:'enroll-old',grade:2,class_id:'class-old',student_id:'student-old',subject_name:'합성 과목',room_name:'시험실',source_academic_year:2026,semester:2}]);
 seed(S.ABSENCES,[{id:'absence-old',student_id:'student-old',class_id:'class-old',exam_date:'2026-05-01',period:1}]);
 seed(S.SETTINGS,[{key:'subject_catalog_json',value:'{"1":[],"2":["합성 과목"],"3":[]}'},{key:'academic_catalog_json',value:'{"1":[],"2":["합성 과목"],"3":[]}'}]);
 state.snapshot={academic_year:2026,semester:2,classes:[{grade:2,class_num:2}],students:[{source_student_uid:'uid-1',class_key:'2|2',number:9,name:'합성 학생',grade:2}],courses:[{source_student_uid:'uid-1',grade:2,subject_name:'합성 과목'}],catalog:{'1':[],'2':['합성 과목'],'3':[]},warnings:{unreviewed_choices:0,ambiguous_sections:0}};
 const read=name=>clone(writer.getRecordsFresh_(name));
 return {state,writer,reader,S,seed,read,execution,sync:id=>writer.syncAcademicData_(true,id,'synthetic-admin')};
}
test('학생·수강 ID와 호실을 보존하며 폐반 수강을 제거하고 과거 결시 신원을 고정한다',()=>{
 const h=harness();h.sync('first');const rows=h.read(h.S.ENROLLMENTS);
 assert.equal(rows.length,1);assert.equal(rows[0].id,'enroll-old');assert.equal(rows[0].class_id,h.read(h.S.STUDENTS)[0].class_id);assert.equal(rows[0].room_name,'시험실');
 const archived=h.writer.decorateAbsences_(h.read(h.S.ABSENCES))[0];assert.equal(archived.student_number,1);assert.equal(archived.class_num,1);
});
test('동기화 중 별도 실행의 명단 조회는 혼합 스냅샷 대신 BUSY를 반환한다',()=>{
 const h=harness();let attempted=0;h.state.hooks.afterSet=name=>{if(name===h.S.CLASSES){attempted++;assert.throws(()=>h.reader.withDataRead_(()=>h.reader.getRecords_(h.S.STUDENTS)),{code:'BUSY'});}};
 h.sync('writer');assert.equal(attempted,1);assert.equal(h.reader.getRecords_(h.S.STUDENTS)[0].number,9);
});
test('캐시 채우기 중 writer는 잠금을 얻지 못해 오래된 캐시 재삽입 경합이 없다',()=>{
 const h=harness();let checked=false;h.state.hooks.afterRead=name=>{if(name!==h.S.STUDENTS||checked)return;checked=true;assert.throws(()=>h.reader.syncAcademicData_(true,'other','synthetic-admin'),{code:'BUSY'});};
 assert.equal(h.writer.getRecords_(h.S.STUDENTS)[0].number,1);h.state.hooks={};h.sync('later');assert.equal(h.writer.getRecords_(h.S.STUDENTS)[0].number,9);
});
test('쓰기·첫 원복 이중 실패에도 나머지 표를 원복하고 영속 복구 표식으로 차단한다',()=>{
 const h=harness();let phase=0;h.state.hooks.beforeSet=name=>{if(phase===0&&name===h.S.ENROLLMENTS){phase=1;throw Error('FORWARD_IO');}if(phase===1&&name===h.S.CLASSES){phase=2;throw Error('RESTORE_IO');}};
 assert.throws(()=>h.sync('double-fault'),{code:'RECOVERY_REQUIRED'});assert.equal(h.read(h.S.STUDENTS)[0].number,1);assert.equal(h.read(h.S.ENROLLMENTS)[0].id,'enroll-old');
 assert.ok(h.state.properties.get('ACADEMIC_SYNC_JOURNAL'));assert.throws(()=>h.reader.getRecords_(h.S.STUDENTS),{code:'RECOVERY_REQUIRED'});
 h.state.hooks={};h.sync('recover');assert.equal(h.read(h.S.STUDENTS)[0].number,9);assert.equal(h.state.properties.has('ACADEMIC_SYNC_JOURNAL'),false);
});
test('커밋 뒤 receipt 저장 실패도 성공을 실패로 바꾸지 않고 같은 요청을 재실행하지 않는다',()=>{
 const h=harness();h.state.hooks.beforeProperty=key=>{if(key==='ACADEMIC_SYNC_RECEIPTS')throw Error('RECEIPT_IO');};
 h.sync('committed');assert.equal(h.writer.academicLinkStatus_().state,'connected');assert.equal(h.state.reads,1);assert.equal(JSON.parse(h.state.properties.get('ACADEMIC_SYNC_JOURNAL')).stage,'committed');
 h.sync('committed');assert.equal(h.state.reads,1);assert.deepEqual(h.state.audits,[true]);
});
test('Cache 초기화·새 실행·응답 유실 후 같은 요청은 영속 receipt로 한 번만 반영한다',()=>{
 const h=harness();h.sync('same');h.state.cache.clear();const restarted=h.execution();restarted.syncAcademicData_(true,'same','synthetic-admin');assert.equal(h.state.reads,1);assert.equal(h.state.copies,1);
});
test('변경 없는 강제 갱신은 시트 재작성·백업을 만들지 않는다',()=>{
 const h=harness();h.sync('first');const writes=h.state.writes,copies=h.state.copies;h.sync('new-request');assert.equal(h.state.reads,2);assert.equal(h.state.writes,writes);assert.equal(h.state.copies,copies);
});
test('변경된 학기와 반영된 학기가 다르면 고사 작업을 차단한다',()=>{
 const h=harness();h.sync('first');h.writer.academicSourceConfig_=()=>({spreadsheet_id:'synthetic-source-only',academic_year:2026,semester:1});
 assert.equal(h.writer.academicLinkStatus_().state,'stale');assert.throws(()=>h.writer.assertAcademicCurrent_(),{code:'ACADEMIC_SYNC_REQUIRED'});
});
test('공개 GET은 초기화·동기화·시트 읽기를 실행하지 않는다',()=>{
 const h=harness();h.writer.ensureSchemaOnRequest_=()=>{throw Error('MUST_NOT_RUN');};h.writer.syncAcademicData_=()=>{throw Error('MUST_NOT_RUN');};h.writer.academicLinkStatus_=()=>{throw Error('MUST_NOT_READ');};
 const response=JSON.parse(h.writer.doGet().text);assert.equal(response.ok,true);assert.equal(response.data.version,'3.1.1');assert.equal(h.state.writes,0);
});
test('null·배열·문자열·잘못된 JSON은 고정 JSON 오류이고 내부 상세는 공개하지 않는다',()=>{
 const h=harness();for(const body of ['null','[]','"text"','{bad']){const result=JSON.parse(h.writer.doPost({postData:{contents:body}}).text);assert.equal(result.ok,false);}
 h.writer.dispatch_=()=>{const error=Error('PRIVATE_IDENTIFIER');error.details={sensitive:'PRIVATE_IDENTIFIER'};throw error;};
 const result=h.writer.doPost({postData:{contents:'{"action":"synthetic"}'}}).text;assert.ok(!result.includes('PRIVATE_IDENTIFIER'));assert.equal(JSON.parse(result).error.details,null);assert.equal(h.state.writes,0);assert.equal(h.state.audits.length,1);
});
test('중복 학생 고유번호는 계획 전에 거부한다',()=>{
 const h=harness();const base={students:[{student_uid:'dup'},{student_uid:'dup'}]};assert.throws(()=>h.writer.buildAcademicSnapshot_(base,{academic_year:2026,semester:2}),{code:'ACADEMIC_SOURCE_CONFLICT'});
});
test('합성 25개 별도 실행은 잠금 경합 뒤 동일 ID 재시도에서 한 번만 커밋한다',()=>{
 const h=harness(),peers=Array.from({length:25},()=>h.execution());let attempted=false;
 h.state.hooks.afterSet=name=>{if(name!==h.S.CLASSES||attempted)return;attempted=true;for(const peer of peers)assert.throws(()=>peer.syncAcademicData_(true,'shared','synthetic-admin'),{code:'BUSY'});};
 h.sync('shared');h.state.hooks={};for(const peer of peers)peer.syncAcademicData_(true,'shared','synthetic-admin');
 assert.equal(h.state.reads,1);assert.equal(h.state.copies,1);assert.deepEqual(h.state.audits,[true]);
});
test('실행 중단 표식을 다음 실행에서 복구한 뒤 원장 오류가 나도 완전한 이전 자료를 유지한다',()=>{
 const h=harness();h.sync('first');const backup=h.writer.getSpreadsheet_().copy();h.state.properties.set('ACADEMIC_SYNC_JOURNAL',JSON.stringify({stage:'writing',backup_id:backup.getId()}));
 h.seed(h.S.STUDENTS,[{id:'student-old',number:123,name:'합성 중단값'}]);const resumed=h.execution();resumed.readAcademicSnapshot_=()=>{throw resumed.apiError_('ACADEMIC_SOURCE_INVALID','합성 원장 오류');};
 assert.throws(()=>resumed.syncAcademicData_(true,'resume','synthetic-admin'),{code:'ACADEMIC_SOURCE_INVALID'});assert.equal(h.read(h.S.STUDENTS)[0].number,9);assert.equal(h.state.properties.has('ACADEMIC_SYNC_JOURNAL'),false);
});
test('기존 관리자 세션의 첫 요청도 schema6 결시 열을 확장한 뒤 번호 이동을 반영한다',()=>{
 const h=harness(),old=h.state.tables.get(h.S.ABSENCES).map(row=>row.slice(0,9));h.state.tables.set(h.S.ABSENCES,old);
 h.state.properties.set('ADMIN_SESSION_VERSION','1');
 h.state.cache.set('session:existing-user',JSON.stringify({role:'user',user_id:'system_admin',system_admin:true,session_version:1}));
 h.state.cache.set('session:existing-admin',JSON.stringify({role:'admin',actor_id:'system_admin',session_version:1}));
 h.writer.dispatch_({action:'sync_academic_data',user_session:'existing-user',admin_session:'existing-admin',request_id:'first-after-rollout'});
 assert.equal(h.state.schemaCalls,1);assert.ok(h.state.tables.get(h.S.ABSENCES)[0].includes('student_number_snapshot'));
 assert.equal(h.read(h.S.ABSENCES)[0].student_number_snapshot,1);assert.equal(h.writer.decorateAbsences_(h.read(h.S.ABSENCES))[0].student_number,1);assert.equal(h.read(h.S.STUDENTS)[0].number,9);
});
test('중첩 스키마 확인은 동기화 잠금을 풀지 않는다',()=>{
 const h=harness();let checked=false;h.state.hooks.beforeSet=name=>{if(name===h.S.CLASSES){checked=true;assert.throws(()=>h.reader.getRecords_(h.S.STUDENTS),{code:'BUSY'});}};
 h.sync('nested-schema');assert.equal(checked,true);assert.equal(h.state.owner,null);
});
test('새 status 게시 직후 중단되면 재시작은 표와 이전 status·commit을 함께 복구한다',()=>{
 const h=harness();h.sync('old-commit');const previous=h.state.properties.get('ACADEMIC_SYNC_STATUS'),backup=h.writer.getSpreadsheet_().copy();
 const journal={stage:'writing',backup_id:backup.getId(),previous_status:previous,request_key:'synthetic-interrupted'};
 h.state.properties.set('ACADEMIC_SYNC_JOURNAL',JSON.stringify(journal));
 h.seed(h.S.STUDENTS,[{...h.read(h.S.STUDENTS)[0],number:10}]);
 h.state.properties.set('ACADEMIC_SYNC_STATUS',JSON.stringify({...JSON.parse(previous),commit_id:'interrupted-new-commit',counts:{students:99}}));
 const resumed=h.execution();resumed.syncAcademicData_(false);
 assert.equal(h.read(h.S.STUDENTS)[0].number,9);assert.equal(h.state.properties.get('ACADEMIC_SYNC_STATUS'),previous);assert.equal(h.state.properties.has('ACADEMIC_SYNC_JOURNAL'),false);assert.equal(h.state.reads,1);assert.doesNotThrow(()=>resumed.assertAcademicCurrent_());
});
test('복구 status 원복 실패는 표가 돌아왔더라도 조회 차단 표식을 유지한다',()=>{
 const h=harness();h.sync('old');const previous=h.state.properties.get('ACADEMIC_SYNC_STATUS'),backup=h.writer.getSpreadsheet_().copy();
 h.state.properties.set('ACADEMIC_SYNC_JOURNAL',JSON.stringify({stage:'writing',backup_id:backup.getId(),previous_status:previous}));
 h.state.hooks.beforeProperty=key=>{if(key==='ACADEMIC_SYNC_STATUS')throw Error('RESTORE_STATUS_FAILED');};
 assert.throws(()=>h.execution().syncAcademicData_(true,'recover','synthetic-admin'),{code:'RECOVERY_REQUIRED'});assert.ok(h.state.properties.has('ACADEMIC_SYNC_JOURNAL'));assert.throws(()=>h.reader.getRecords_(h.S.STUDENTS),{code:'RECOVERY_REQUIRED'});
});
