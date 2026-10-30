import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {createHash} from 'node:crypto';
const code=['Config.gs','Repository.gs','Code.gs','Academic.gs'].map(name=>fs.readFileSync(new URL('../apps-script/'+name,import.meta.url),'utf8')).join('\n');
let nextId=0;
const helpers=new Function('Utilities', code+'\ncreateId_=()=>"new-"+Utilities.getUuid(); return {buildAcademicSnapshot_,planAcademicMerge_,academicActive_,snapshotAbsenceIdentity_};')({getUuid:()=>++nextId, DigestAlgorithm:{SHA_256:'sha256'},Charset:{UTF_8:'utf8'},computeDigest:(_,text)=>[...createHash('sha256').update(text).digest()],base64EncodeWebSafe:bytes=>Buffer.from(bytes).toString('base64url')});
const config={academic_year:2026,semester:2};
function fixture(){return {students:[{student_uid:'synthetic-a',name:'가상학생',active:true,enrollment_status:'ACTIVE',phone:'must-not-copy'}],assignments:[{student_uid:'synthetic-a',grade:2,class_no:1,student_no:1,active:true}],classes:[],subjects:[{subject_id:'subject-a',grade:2,name:'가상 과목',code:'2:CHOICE',active:true},{subject_id:'year-a',grade:2,name:'연간 과목',code:'YEAR:CHOICE',active:true}],sections:[{section_id:'section-a',subject_id:'subject-a',section_name:'가상분반',active:true}],enrollments:[],courses:[{student_uid:'synthetic-a',active:true,choices_json:JSON.stringify([{name:'가상 과목',raw_subject:'가상 과목A1',semester:'2',subject_id:'subject-a'},{name:'다른 학기',semester:'1'},{name:'연간 과목',semester:'YEAR'}])}],plans:[]};}
const previous=()=>({classes:[{id:'class-original',grade:2,class_num:1,active:true}],students:[{id:'student-original',class_id:'class-original',number:1,name:'가상학생',active:true}],enrollments:[{id:'enroll-original',student_id:'student-original',class_id:'class-original',subject_name:'가상과목',room_name:'시험실'}],timetable:[{id:'schedule',grade:2,subject_name:'가상과목'}]});
test('학적 연결은 재학생과 현재 학기·연간 과목만 가져오며 부가 개인정보를 제외한다',()=>{
 const source=fixture();source.students.push({student_uid:'inactive',name:'가상비활성',active:false});source.assignments.push({student_uid:'inactive',grade:2,class_no:1,student_no:2,active:true});
 const data=helpers.buildAcademicSnapshot_(source,config);
 assert.equal(data.students.length,1);assert.equal(data.courses.length,2);assert.equal(data.warnings.unreviewed_choices,0);assert.equal(data.courses[0].subject_name,'가상 과목');assert.ok(!JSON.stringify(data).includes('must-not-copy'));
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
 assert.equal(next.enrollments.filter(x=>x.id==='enroll-original').length,1);assert.equal(next.enrollments.find(x=>x.id==='enroll-original').class_id,next.students.find(x=>x.id==='student-original').class_id);
});
test('미확인 선택과목과 필수 개수 부족은 기존 수강을 지우는 계획을 만들지 않는다',()=>{
 const source=fixture();source.courses[0].choices_json=JSON.stringify([{raw_subject:'존재하지 않는 과목',semester:'2'}]);
 assert.throws(()=>helpers.buildAcademicSnapshot_(source,config),{code:'ACADEMIC_CHOICES_INCOMPLETE'});
 const partial=fixture();partial.courses[0].choices_json=JSON.stringify([{raw_subject:'가상 과목',semester:'2'}]);partial.plans=[{grade:2,groups_json:JSON.stringify([{key:'G',semester:'2',required_count:2,subjects:['가상 과목','다른 과목']}])}];
 assert.throws(()=>helpers.buildAcademicSnapshot_(partial,config),{code:'ACADEMIC_CHOICES_INCOMPLETE'});
});
test('선택 명단 행 누락·빈 fallback·부분 fallback을 모든 재학생 기준으로 차단한다',()=>{
 const source=fixture();source.courses=[];source.plans=[{grade:2,groups_json:JSON.stringify([{key:'G',semester:'2',required_count:2,subjects:['가상 과목','다른 과목']}])}];
 source.subjects.push({subject_id:'other-subject',grade:2,name:'다른 과목',code:'2:CHOICE',active:true});
 assert.throws(()=>helpers.buildAcademicSnapshot_(source,config),{code:'ACADEMIC_CHOICES_INCOMPLETE'});
 source.enrollments=[{student_uid:'synthetic-a',subject_id:'subject-a',active:true}];assert.throws(()=>helpers.buildAcademicSnapshot_(source,config),{code:'ACADEMIC_CHOICES_INCOMPLETE'});
 source.enrollments.push({student_uid:'synthetic-a',subject_id:'other-subject',active:true});assert.equal(helpers.buildAcademicSnapshot_(source,config).courses.length,2);
 source.students.push({student_uid:'synthetic-b',name:'가상 다른학생',active:true,enrollment_status:'ACTIVE'});source.assignments.push({student_uid:'synthetic-b',grade:2,class_no:1,student_no:2,active:true});
 assert.throws(()=>helpers.buildAcademicSnapshot_(source,config),{code:'ACADEMIC_CHOICES_INCOMPLETE'});
});
test('현재 교육과정으로 과목 ID를 재해석하고 정확한 이름을 접미사 일치보다 우선한다',()=>{
 const source=fixture();source.courses[0].choices_json=JSON.stringify([{raw_subject:'가상 과목A1',name:'옛 이름',subject_id:'stale',semester:'2',slot:'선택1',review_needed:true}]);
 source.plans=[{grade:2,groups_json:JSON.stringify([{key:'G',semester:'2',required_count:1,subjects:['가상 과목','가상 과목A1']}])}];
 const first=helpers.buildAcademicSnapshot_(source,config);assert.equal(first.courses[0].subject_name,'가상 과목A1');assert.match(first.courses[0].source_subject_id,/^course_2026_2_2_/);
 source.plans[0].groups_json=JSON.stringify([{key:'G',semester:'2',required_count:1,subjects:['다른 과목']}]);assert.throws(()=>helpers.buildAcademicSnapshot_(source,config),{code:'ACADEMIC_CHOICES_INCOMPLETE'});
});
test('원장 비활성 별칭과 빈 배정 active는 제외한다',()=>{
 for(const active of [' false ','N','no','비활성','미사용',0,false])assert.equal(helpers.academicActive_({active}),false);
 assert.equal(helpers.academicActive_({},false),false);
});
test('반 이동 전 결시 이름·번호·학급을 스냅샷으로 고정하고 재실행해도 덮지 않는다',()=>{
 const before=previous(),rows=[{student_id:'student-original',class_id:'class-original'}];
 const snapshot=helpers.snapshotAbsenceIdentity_(rows,before.students,before.classes);assert.equal(snapshot[0].student_number_snapshot,1);assert.equal(snapshot[0].class_num_snapshot,1);
 before.students[0].number=20;before.students[0].name='변경';assert.deepEqual(helpers.snapshotAbsenceIdentity_(snapshot,before.students,before.classes),snapshot);
});
test('분반 연결이 두 개면 임의 선택하지 않고 경고 건수를 반환한다',()=>{
 const source=fixture();const choice=JSON.parse(source.courses[0].choices_json);choice[0].slot='선택1';source.courses[0].choices_json=JSON.stringify(choice);
 source.sections.push({...source.sections[0],section_id:'section-b'});source.section_map=Object.fromEntries(['section-a','section-b'].map(id=>[id,{subject_id:'subject-a',semester:'2',slot:'선택1',source_code:'A1'}]));
 const data=helpers.buildAcademicSnapshot_(source,config);assert.equal(data.courses[0].source_section_id,'');assert.equal(data.warnings.ambiguous_sections,1);
});
test('선택과목 명단을 수강 원장보다 우선하고 학적 분반 매핑을 함께 연결한다',()=>{
 const source=fixture();source.subjects.push({subject_id:'stale-subject',grade:2,name:'가상 이전 과목',code:'2:CHOICE',active:true});source.enrollments.push({student_uid:'synthetic-a',subject_id:'stale-subject',status:'ACTIVE',active:true});
 source.section_map={'section-a':{subject_id:'subject-a',semester:'2',slot:'선택1',source_code:'A1'}};
 const choices=JSON.parse(source.courses[0].choices_json);choices[0].slot='선택1';source.courses[0].choices_json=JSON.stringify(choices);
 const snapshot=helpers.buildAcademicSnapshot_(source,config);assert.ok(!snapshot.courses.some(row=>row.subject_name==='가상 이전 과목'));assert.equal(snapshot.courses[0].source_section_id,'section-a');assert.equal(snapshot.courses[0].section_name,'가상분반');assert.equal(snapshot.courses[0].room_name,undefined);
});
