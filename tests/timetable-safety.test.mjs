import fs from 'node:fs';
import assert from 'node:assert/strict';
import test from 'node:test';
const source=['Config','Repository','Code','Academic'].map(name=>fs.readFileSync(new URL(`../apps-script/${name}.gs`,import.meta.url),'utf8')).join('\n');
function fixture(){const data={'학급':[{id:'c1',grade:2,class_num:1,active:true}],'고사일':[{exam_date:'2026-10-30',active:true}],'시간표':[],'선택과목':[{source_academic_year:2026,grade:2,subject_name:'합성 과목'}]};const save=new Function('data',source+`\ngetRecords_=name=>data[name]||[];replaceRecords_=(name,rows)=>{data[name]=rows;};adminBootstrap_=()=>({timetable:data['시간표']});createId_=()=>'synthetic-id';nowIso_=()=>'';return saveTimetable_;`)(data);return{data,save};}
const row=()=>({id:'new',exam_date:'2026-10-30',grade:2,period:2,subject_name:'합성과목',subject_type:'elective',class_ids:['c1'],start_time:'09:00',end_time:'09:50'});
test('1·2·7교시와 유일한 학적 과목명 별칭이 저장 직후 연결된다',()=>{for(const period of [1,2,7]){const h=fixture();const result=h.save({timetable:[{...row(),period}]});assert.equal(result.timetable[0].period,period);assert.equal(result.timetable[0].subject_name,'합성 과목');}});
test('역전·같은 시각·한쪽 누락·잘못된 시각은 쓰기 전에 거부한다',()=>{for(const [start_time,end_time]of[['11:00','09:00'],['09:00','09:00'],['','09:00'],['09:00',''],['25:00','26:00']]){const h=fixture();assert.throws(()=>h.save({timetable:[{...row(),start_time,end_time}]}),{code:'INVALID_TIMETABLE_TIME'});assert.equal(h.data['시간표'].length,0);}});
test('1~2월 기본 연결은 이전 학년도 2학기다',()=>{const config=new Function('PropertiesService','Utilities',source+`\ngetSettings_=()=>({});return academicSourceConfig_;`)({getScriptProperties:()=>({getProperty:()=> 'synthetic-only'})},{formatDate:()=> '2027-01-05'})();assert.equal(config.academic_year,2026);assert.equal(config.semester,2);});
