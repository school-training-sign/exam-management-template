import fs from 'node:fs';
import test from 'node:test';
import assert from 'node:assert/strict';
const files=Object.fromEntries(['Config','Repository','Code','Academic'].map(name=>[name,fs.readFileSync(new URL(`../apps-script/${name}.gs`,import.meta.url),'utf8')]));
const nativeMessage='셀을 병합하거나 병합을 취소하려면 해당 범위의 모든 셀을 선택해야 합니다.';
const clone=value=>JSON.parse(JSON.stringify(value));
function harness(options={}){
 const merges=options.merges||['A1:F1','A3:F3','B6:F6','B36:F36'];
 const state={merges:[...merges],unmerged:[],partialUnmerges:0,guideWrites:0,pending:null,guideRows:options.current?[['한양대학교사범대학부속고등학교 통합 시스템','3.1.1']]:[['Legacy guide','old instructions']],tableInit:0};
 const settings=[['key','value','updated_at'],['schema_version','7','synthetic'],['version','3.1.1','synthetic'],['app_name','한양대학교사범대학부속고등학교 통합 시스템','synthetic']];
 const originalSettings=clone(settings),properties=new Map([['ACCESS_USER_PEPPER','synthetic-existing-pepper'],['EXAM_PACKAGING_INVITE_PEPPER','synthetic-existing-pepper'],['EXAM_PACKAGING_INVITE_VERSION','1']]);
 const deferMergeError=()=>{state.pending=Object.assign(new Error(nativeMessage),{name:'Exception',stack:'Exception: '+nativeMessage+'\n at getRecordsFresh_ (Repository:134:26)'});};
 const flush=()=>{if(state.pending)throw state.pending;};
 const fullMerge=a1=>({getA1Notation:()=>a1,breakApart(){state.unmerged.push(a1);state.merges=state.merges.filter(x=>x!==a1);return this;}});
 function dataRange(){return{
  getDisplayValues(){flush();return clone(state.guideRows);},
  getMergedRanges(){flush();return state.merges.map(fullMerge);},
  breakApart(){state.partialUnmerges++;if(state.merges.some(x=>x.includes('F')||x==='A30:A40'))deferMergeError();else state.merges=[];return this;},
  clearContent(){state.guideWrites++;state.guideRows=[];return this;},
 };}
 const guide={getLastRow:()=>state.guideRows.length?36:0,getDataRange:dataRange,
  getRange(a1){return{setValues(values){assert.equal(a1,'A1:B8');state.guideWrites++;state.guideRows=clone(values);return this;},merge(){assert.equal(a1,'A1:B1');if(state.merges.includes('A1:F1'))deferMergeError();state.merges.push(a1);return this;},setBackground(){return this;},setFontColor(){return this;},setFontWeight(){return this;},setWrap(){return this;},setVerticalAlignment(){return this;}};},setColumnWidth(){return this;}};
 let missing=Boolean(options.missing);if(missing){state.guideRows=[];state.merges=[];}
 const book={getSheetByName(name){if(name==='사용설명서')return missing?null:guide;if(name==='설정')return{getDataRange(){flush();return{getValues(){flush();return clone(settings);}};}};throw Error('Unexpected sheet '+name);},insertSheet(name,index){assert.equal(name,'사용설명서');assert.equal(index,0);missing=false;return guide;}};
 let repository=files.Repository;
 if(options.legacy){const marker='oldGuideRange.getMergedRanges().forEach((range) => range.breakApart());';assert.ok(repository.includes(marker));repository=repository.replace(marker,'oldGuideRange.breakApart();');}
 const source=[files.Config,repository,files.Code,files.Academic].join('\n');
 const app=new Function('book','state','PropertiesService','CacheService','LockService',source+`\ngetSpreadsheet_=()=>book;ensureTableSheet_=()=>{state.tableInit++;};assertAcademicReadable_=()=>{};return{initialize:ensureSheets_,phase:runAdminLoginPhase_};`)(book,state,{getScriptProperties:()=>({getProperty:key=>properties.get(key)||null,setProperty:(key,value)=>properties.set(key,value),deleteProperty:key=>properties.delete(key)})},{getScriptCache:()=>({get:()=>null,put(){}})},{getScriptLock:()=>({hasLock:()=>true})});
 return{...app,state,settings,originalSettings,flush};
}

test('기존 부분 범위 병합 해제는 뒤의 설정 읽기에서 실제와 같은 Google 오류·진단을 재현한다',()=>{
 const h=harness({legacy:true});assert.throws(()=>h.phase('schema',h.initialize),error=>error.code==='ADMIN_ENTRY_FAILED'&&error.details.diagnostic==='A3-G-RR-CF-V-OTHER');assert.equal(h.state.partialUnmerges,1);assert.equal(h.state.pending.message,nativeMessage);
});

test('데이터는 A:B, 병합은 F열까지인 사용설명서를 전체 병합 단위로 해제한다',()=>{
 const h=harness();h.initialize();h.flush();assert.deepEqual(h.state.unmerged,['A1:F1','A3:F3','B6:F6','B36:F36']);assert.equal(h.state.partialUnmerges,0);assert.deepEqual(h.state.merges,['A1:B1']);assert.equal(h.state.guideRows.length,8);assert.deepEqual(h.settings,h.originalSettings);assert.equal(h.state.tableInit,10);
});

test('행·열 경계를 넘는 여러 병합도 하드코딩 없이 완전한 범위로 처리한다',()=>{
 const h=harness({merges:['A1:F1','A30:A40','B3:F4']});h.initialize();h.flush();assert.deepEqual(h.state.unmerged,['A1:F1','A30:A40','B3:F4']);assert.equal(h.state.partialUnmerges,0);assert.deepEqual(h.settings,h.originalSettings);
});

test('기존 병합이 없어도 안내문을 만들고 학급·학생·설정 데이터는 다시 쓰지 않는다',()=>{
 const h=harness({merges:[]});h.initialize();h.flush();assert.deepEqual(h.state.unmerged,[]);assert.equal(h.state.guideWrites,2);assert.deepEqual(h.settings,h.originalSettings);
});

test('수정 후 반복 초기화는 사용설명서를 다시 지우거나 병합 해제하지 않는다',()=>{
 const h=harness();h.initialize();const writes=h.state.guideWrites,unmerged=[...h.state.unmerged];h.initialize();h.flush();assert.equal(h.state.guideWrites,writes);assert.deepEqual(h.state.unmerged,unmerged);assert.deepEqual(h.settings,h.originalSettings);
});

test('현재 버전 안내문은 그대로 보존하고 없는 안내문은 새로 만들 수 있다',()=>{
 const current=harness({current:true});current.initialize();assert.equal(current.state.guideWrites,0);assert.deepEqual(current.state.unmerged,[]);
 const missing=harness({missing:true});missing.initialize();missing.flush();assert.equal(missing.state.guideRows.length,8);assert.equal(missing.state.partialUnmerges,0);assert.deepEqual(missing.settings,missing.originalSettings);
});
