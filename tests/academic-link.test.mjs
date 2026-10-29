import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
const code=['Config.gs','Repository.gs','Code.gs','Academic.gs'].map(name=>fs.readFileSync(new URL('../apps-script/'+name,import.meta.url),'utf8')).join('\n');
let nextId=0;
const helpers=new Function('Utilities', code+'\ncreateId_=()=>"new-"+Utilities.getUuid(); return {buildAcademicSnapshot_,planAcademicMerge_};')({getUuid:()=>++nextId});
const config={academic_year:2026,semester:2};
function fixture(){return {students:[{student_uid:'synthetic-a',name:'가상학생',active:true,enrollment_status:'ACTIVE',phone:'must-not-copy'}],assignments:[{student_uid:'synthetic-a',grade:2,class_no:1,student_no:1,active:true}],classes:[],subjects:[{subject_id:'subject-a',grade:2,name:'가상 과목',code:'2:CHOICE',active:true}],sections:[{section_id:'section-a',subject_id:'subject-a',section_name:'가상분반',active:true}],enrollments:[],courses:[{student_uid:'synthetic-a',active:true,choices_json:JSON.stringify([{name:'가상 과목',raw_subject:'가상 과목A1',semester:'2',subject_id:'subject-a'},{name:'다른 학기',semester:'1'},{name:'연간 과목',semester:'YEAR'},{name:'확인 필요',semester:'2',review_needed:true}])}],plans:[]};}
const previous=()=>({classes:[{id:'class-original',grade:2,class_num:1,active:true}],students:[{id:'student-original',class_id:'class-original',number:1,name:'가상학생',active:true}],enrollments:[{id:'enroll-original',student_id:'student-original',class_id:'class-original',subject_name:'가상과목',room_name:'시험실'}],timetable:[{id:'schedule',grade:2,subject_name:'가상과목'}]});
test('학적 연결은 재학생과 현재 학기·연간 과목만 가져오며 부가 개인정보를 제외한다',()=>{
 const source=fixture();source.students.push({student_uid:'inactive',name:'가상비활성',active:false});source.assignments.push({student_uid:'inactive',grade:2,class_no:1,student_no:2,active:true});
 const data=helpers.buildAcademicSnapshot_(source,config);
 assert.equal(data.students.length,1);assert.equal(data.courses.length,2);assert.equal(data.warnings.unreviewed_choices,1);assert.equal(data.courses[0].subject_name,'가상 과목');assert.ok(!JSON.stringify(data).includes('must-not-copy'));
});
test('중복 학생·번호와 잘못된 선택과목 JSON은 저장 계획 전에 차단한다',()=>{
 const source=fixture();source.assignments.push({...source.assignments[0]});assert.throws(()=>helpers.buildAcademicSnapshot_(source,config),{code:'ACADEMIC_SOURCE_CONFLICT'});
 const bad=fixture();bad.courses[0].choices_json='broken';assert.throws(()=>helpers.buildAcademicSnapshot_(bad,config),{code:'ACADEMIC_SOURCE_INVALID'});
});
test('초기 연결과 반복 동기화는 학생·학급·수강 ID와 시험 호실을 유지한다',()=>{
 const data=helpers.buildAcademicSnapshot_(fixture(),config),before=previous(),first=helpers.planAcademicMerge_(data,before,'first');
 assert.equal(first.classes[0].id,'class-original');assert.equal(first.students[0].id,'student-original');assert.equal(first.enrollments[0].id,'enroll-original');assert.equal(first.enrollments[0].room_name,'시험실');assert.equal(first.enrollments[0].subject_name,'가상과목');
 const again=helpers.planAcademicMerge_(data,{...first,timetable:before.timetable},'second');assert.deepEqual(again.students.map(x=>x.id),first.students.map(x=>x.id));assert.deepEqual(again.enrollments.map(x=>x.id),first.enrollments.map(x=>x.id));
});
test('같은 번호의 다른 학생은 새 ID를 받고 예전 학생 행은 비활성 상태로 보존된다',()=>{
 const before=previous();before.students[0].name='다른 가상학생';const next=helpers.planAcademicMerge_(helpers.buildAcademicSnapshot_(fixture(),config),before,'now');assert.equal(next.students[0].active,false);assert.equal(next.students[0].name,'다른 가상학생');assert.notEqual(next.students[1].id,'student-original');
});
test('반 이동 뒤에도 고유 ID로 연결하고 이전 학년도 학생·수강 이력은 남긴다',()=>{
 const data=helpers.buildAcademicSnapshot_(fixture(),config),before=previous();before.students[0].source_student_uid='synthetic-a';before.students[0].source_academic_year=2026;data.classes[0].class_num=2;data.students[0].class_key='2|2';
 before.classes.push({id:'old-year-class',grade:1,class_num:1,source_academic_year:2025,active:true});before.students.push({id:'old-year-student',class_id:'old-year-class',source_academic_year:2025,active:true});before.enrollments.push({id:'old-year-enrollment',class_id:'old-year-class',student_id:'old-year-student'});
 const next=helpers.planAcademicMerge_(data,before,'now');assert.equal(next.students.find(x=>x.id==='student-original').class_id,next.classes.find(x=>x.class_num===2).id);assert.equal(next.students.find(x=>x.id==='old-year-student').active,false);assert.ok(next.enrollments.find(x=>x.id==='old-year-enrollment'));
});
