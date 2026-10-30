import fs from 'node:fs';
import assert from 'node:assert/strict';
import test from 'node:test';
const code=['Config','Repository','Code','Academic'].map(name=>fs.readFileSync(new URL(`../apps-script/${name}.gs`,import.meta.url),'utf8')).join('\n');
function harness(options={}){
 const calls=[],formatCalls=[];
 const values=options.values||[['key','value'],['school_name','합성 학교']];
 const nativeError=options.error||Object.assign(Error(options.message||'Service Spreadsheets failed while accessing document PRIVATE_ID'),{name:'Exception',stack:'Exception: PRIVATE_ID PRIVATE_PERSON\n at getRecordsFresh_ (Repository:137:2)'});
 const fail=step=>{if(options.fail===step)throw nativeError;};
 const sheet={
  getLastRow(){assert.fail('Reader must not call getLastRow');},
  getLastColumn(){assert.fail('Reader must not call getLastColumn');},
  getRange(){assert.fail('Reader must not request a size-based range');},
  getDataRange(){calls.push('getDataRange');fail('D');return{getValues(){calls.push('getValues');fail('V');return values;}};},
 };
 const utilities={formatDate(value,timezone,format){formatCalls.push({value,timezone,format});fail('N');return new Intl.DateTimeFormat('sv-SE',{timeZone:timezone,year:'numeric',month:'2-digit',day:'2-digit'}).format(value);}};
 const helpers=new Function('sheet','options','fail','Utilities',code+`\ngetSpreadsheet_=()=>({getSheetByName:()=>{fail('S');return options.missing?null:sheet;}});return {read:getRecordsFresh_,phase:runAdminLoginPhase_,classify:classifySheetReadError_};`)(sheet,options,fail,utilities);
 return{...helpers,calls,formatCalls,nativeError};
}
test('실패한 행·열 수 조회 없이 데이터 범위를 한 번 직접 읽고 값을 바꾸지 않는다',()=>{
 const original=[['key','value'],['school_name','합성 학교']];const h=harness({values:original});assert.deepEqual(h.read('설정'),[{key:'school_name',value:'합성 학교'}]);assert.deepEqual(h.calls,['getDataRange','getValues']);assert.deepEqual(original,[['key','value'],['school_name','합성 학교']]);
});
test('기존 추가 열·빈 헤더·헤더 공백·숫자 0·논리 값·공백 행 처리는 유지한다',()=>{
 const values=[[' key ','value','updated_at','extra',''],['setting',false,0,'kept','ignored'],['','','','','']];const h=harness({values});assert.deepEqual(h.read('설정'),[{key:'setting',value:false,updated_at:0,extra:'kept'}]);assert.deepEqual(h.calls,['getDataRange','getValues']);
});
test('헤더뿐인 시트·실제 빈 셀·빈 배열은 데이터 범위를 읽고 빈 목록을 반환한다',()=>{
 for(const values of [[['key','value']],[['']],[],[[],[]]]){const h=harness({values});assert.deepEqual(h.read('설정'),[]);assert.deepEqual(h.calls,['getDataRange','getValues']);}
});

test('날짜 셀은 기존 한국 시간대 정규화를 유지하고 원본 셀을 변경하지 않는다',()=>{
 const date=new Date('2026-10-29T15:30:00.000Z');const h=harness({values:[['date','enabled'],[date,true]]});
 assert.deepEqual(h.read('고사일'),[{date:'2026-10-30',enabled:true}]);assert.equal(h.formatCalls.length,1);assert.equal(h.formatCalls[0].value,date);assert.equal(h.formatCalls[0].timezone,'Asia/Seoul');assert.equal(h.formatCalls[0].format,'yyyy-MM-dd');assert.equal(date.toISOString(),'2026-10-29T15:30:00.000Z');
});
test('Google 범위·권한·서비스·날짜 예외는 고정 범주로만 표시한다',()=>{
 for(const [message,category]of [['Those columns are out of bounds PRIVATE_ID','BOUND'],['You do not have permission to read PRIVATE_ID','PERM'],['Authorization is required PRIVATE_ID','AUTH'],['Service invoked too many times PRIVATE_PERSON','QUOTA'],['Invalid date PRIVATE_PERSON','DATE'],['Range not found PRIVATE_ID','RANGE'],['The parameters do not match signature PRIVATE_ID','TYPE'],['Service Spreadsheets failed while accessing PRIVATE_ID','SERVICE'],['PRIVATE_PERSON PRIVATE_ID','OTHER']]){
  const h=harness({fail:'V',message});let caught;try{h.phase('schema',()=>h.read('설정'));}catch(error){caught=error;}
  assert.equal(caught.code,'ADMIN_ENTRY_FAILED');assert.equal(caught.details.diagnostic,'A3-G-RR-CF-V-'+category);assert.ok(!/PRIVATE_|Repository:/.test(JSON.stringify({message:caught.message,details:caught.details})));
 }
});
test('조회·범위 획득·값 읽기·정규화 실패는 고정 단계만 표시하며 시트 이름은 숨긴다',()=>{
 for(const [fail,stage]of [['S','S'],['D','V'],['V','V'],['N','N']]){
  const h=harness({fail,message:'PRIVATE_ERROR',values:[['date'],[new Date('2026-10-30T00:00:00.000Z')]]});
  assert.throws(()=>h.phase('schema',()=>h.read('PRIVATE_SHEET')),error=>error.details.diagnostic.endsWith('XX-'+stage+'-OTHER')&&!error.message.includes('PRIVATE'));
 }
});
test('없어진 시트는 빈 명단으로 위장하지 않고 기존 오류로 실패한다',()=>{
 const h=harness({missing:true});assert.throws(()=>h.read('설정'),/시트를 찾지 못했습니다/);assert.deepEqual(h.calls,[]);
});
test('읽기 실패는 원본 예외를 그대로 던지며 공개 오류를 진단으로 바꾸지 않는다',()=>{
 const h=harness({fail:'V'});assert.throws(()=>h.read('설정'),error=>error===h.nativeError&&error.sheet_read_hint==='CF-V-SERVICE');
 const publicError=Object.assign(Error('공개 검증 오류'),{public_error:true,code:'VALIDATION'});const p=harness({fail:'V',error:publicError});assert.throws(()=>p.phase('schema',()=>p.read('설정')),error=>error===publicError);assert.equal(publicError.sheet_read_hint,undefined);
});

test('변경 불가능한 예외도 원본 정체성을 보존하고 예외 내용은 공개하지 않는다',()=>{
 const error=Object.freeze(Object.assign(Error('PRIVATE_ID PRIVATE_PERSON'),{name:'Exception',stack:'Exception: PRIVATE_ID\n at getRecordsFresh_ (Repository:137:2)'}));const h=harness({fail:'V',error});
 assert.throws(()=>h.read('설정'),caught=>caught===error);assert.throws(()=>h.phase('schema',()=>h.read('설정')),caught=>caught.code==='ADMIN_ENTRY_FAILED'&&caught.details.diagnostic==='A3-G-RR'&&!/PRIVATE/.test(caught.message));
});
