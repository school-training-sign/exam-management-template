import fs from 'node:fs';
import assert from 'node:assert/strict';
import test from 'node:test';
const code=['Config','Repository','Code','Academic'].map(name=>fs.readFileSync(new URL(`../apps-script/${name}.gs`,import.meta.url),'utf8')).join('\n');
const clone=value=>JSON.parse(JSON.stringify(value));
function harness(options={}){
 const calls=[];const values=options.values||[['key','value'],['school_name','합성 학교']];
 const nativeError=()=>{const error=Error(options.message||'Service Spreadsheets failed while accessing document PRIVATE_ID');error.name='Exception';error.stack='Exception: PRIVATE_ID PRIVATE_PERSON\n at getRecordsFresh_ (Repository:137:2)';return error;};
 const fail=step=>{if(options.fail===step)throw nativeError();};
 const sheet={getLastRow(){fail('L');return values.length;},getLastColumn(){fail('C');return values[0]?.length||0;},getRange(row,column,height,width){calls.push([row,column,height,width]);fail('V');if(width>(values[0]?.length||0))throw nativeError();return{getValues(){return values;}};}};
 const helpers=new Function('sheet','options',code+`\ngetSpreadsheet_=()=>({getSheetByName:()=>{if(options.fail==='S'){const error=new Error(options.message);error.name='Exception';throw error;}return options.missing?null:sheet;}});normalizeCell_=v=>v;return {read:getRecordsFresh_,phase:runAdminLoginPhase_,classify:classifySheetReadError_};`)(sheet,options);
 return{...helpers,calls};
}
test('두 열뿐인 기존 설정은 실제 두 열만 읽고 범위를 넓히거나 값을 바꾸지 않는다',()=>{
 const original=[['key','value'],['school_name','합성 학교']];const h=harness({values:original});assert.deepEqual(h.read('설정'),[{key:'school_name',value:'합성 학교'}]);assert.deepEqual(h.calls,[[1,1,2,2]]);assert.deepEqual(original,[['key','value'],['school_name','합성 학교']]);
});
test('스키마보다 넓은 기존 열과 숫자·논리 값·공백 행 처리는 유지한다',()=>{
 const values=[['key','value','updated_at','extra'],['setting',true,2026,'kept'],['','','','']];const h=harness({values});assert.deepEqual(h.read('설정'),[{key:'setting',value:true,updated_at:2026,extra:'kept'}]);assert.equal(h.calls[0][3],4);
});
test('헤더뿐이거나 빈 시트는 0 높이·너비 범위를 요청하지 않는다',()=>{
 for(const values of [[['key','value']],[]]){const h=harness({values});assert.deepEqual(h.read('설정'),[]);assert.equal(h.calls.length,0);}
});
test('Google 범위·권한·서비스·날짜 예외는 고정 범주로만 표시한다',()=>{
 for(const [message,category]of [['Those columns are out of bounds PRIVATE_ID','BOUND'],['You do not have permission to read PRIVATE_ID','PERM'],['Authorization is required PRIVATE_ID','AUTH'],['Service invoked too many times PRIVATE_PERSON','QUOTA'],['Invalid date PRIVATE_PERSON','DATE'],['Range not found PRIVATE_ID','RANGE'],['The parameters do not match signature PRIVATE_ID','TYPE'],['Service Spreadsheets failed while accessing PRIVATE_ID','SERVICE'],['PRIVATE_PERSON PRIVATE_ID','OTHER']]){
  const h=harness({fail:'V',message});let caught;try{h.phase('schema',()=>h.read('설정'));}catch(error){caught=error;}
  assert.equal(caught.code,'ADMIN_ENTRY_FAILED');assert.equal(caught.details.diagnostic,'A3-G-RR-CF-V-'+category);assert.ok(!/PRIVATE_|Repository:/.test(JSON.stringify({message:caught.message,details:caught.details})));
 }
});
test('시트 읽기의 실패 단계와 고정 표 코드만 식별하고 임의 시트 이름은 공개하지 않는다',()=>{
 for(const stage of ['S','L','C','V']){const h=harness({fail:stage,message:'PRIVATE_ERROR'});assert.throws(()=>h.phase('schema',()=>h.read('PRIVATE_SHEET')),error=>error.details.diagnostic.endsWith('XX-'+stage+'-OTHER')&&!error.message.includes('PRIVATE'));}
});
test('없어진 시트는 빈 명단으로 위장하지 않고 기존 오류로 실패한다',()=>{
 const h=harness({missing:true});assert.throws(()=>h.read('설정'),/시트를 찾지 못했습니다/);
});
test('원본 예외와 기존 공개 오류의 정체성은 변경하지 않는다',()=>{
 const h=harness();let original;try{harness({fail:'V'}).read('설정');}catch(error){original=error;}assert.equal(original.name,'Exception');assert.match(original.message,/PRIVATE_ID/);assert.equal(h.classify({message:''}),'OTHER');
});
