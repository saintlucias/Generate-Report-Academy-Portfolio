const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
function load(overrides = {}, functionOverrides = {}) {
  const context = vm.createContext({console, ...overrides});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../Code.gs'), 'utf8') + `
    globalThis.api = {REPORT, demoReport_, matchingStudents_, selectStudent_, parseWrongAnswers_,
      buildReportFromSource_, loadSourceData_, readScore_, niceChartMax_,
      pageTransform_, buildReportModel_, fitText_, modelToRequests_, renderReport_,
      extractDriveId_, readCellColorStyle_, colorHex_, styledText_, joinStyledText_,
      fittedColorRuns_, characterColors_, nativeFilledShape_, setupReportAutomation, generateReportByName,
      reportHeader_, syncReportHeaderSettings_, createReportForDialog, reportDialogHtml_,
      extractDriveReference_, readImageBlob_, readConfig_, checkReportFileAccess};`, context);
  Object.assign(context, functionOverrides);
  return context.api;
}
const api = load();
const config = {academy:'국어 날개달기', chartTitle:'모의고사 출제영역', title:'11월 모의고사 국어영역 분석표', chartMax:null};
const BLACK='#000000', BLUE='#0000FF', RED='#FF0000';
const style = (text,color=BLACK,runs=[]) => ({text,color,runs});
const colorObject = hex => ({getColorType:()=> 'RGB',asRgbColor:()=>({asHexString:()=>hex})});
function richMock(text,runs) {
  return {getText:()=>text,getRuns:()=>runs.map(r=>({getStartIndex:()=>r.start,getEndIndex:()=>r.end,
    getTextStyle:()=>({getForegroundColorObject:()=>r.color?colorObject(r.color):null})}))};
}
function row(number,values) {return {row:number,values,display:values.map(String),styles:values.map(v=>style(String(v)))};}
function fixture() {
  const selected=row(7,['028001','국','데모학생','학교',52,1,'1, 2, 4, 8',4,16,12,4,16,99,98,97,'조건 확인 후 다시 풀어 보세요.']);
  selected.styles[6]=style(selected.display[6],RED);
  selected.styles[15]=style(selected.display[15],BLACK,[{start:0,end:5,color:BLUE}]);
  const examNote='시험지 공통 설명\n조건 확인과 근거 찾기';
  return {requestedName:'데모학생',students:[row(3,['28002','국','다른학생','','','','',16,12,12,12,12,'','','','다른 학생의 전달사항']),selected],
    examNote,examNoteStyle:style(examNote,BLACK,[{start:examNote.indexOf('조건'),end:examNote.indexOf('조건')+5,color:BLUE}]),
    averages:[12.65,17.76,10.59,7.25,13.5],questions:Array.from({length:25},(_,i)=>row(i+2,[i+1,'B열','C열','영역'+(i+1),'갈래'+(i+1),'62.2%',i%4+1])).reverse()};
}
function report(source=fixture(),options=config) {return api.buildReportFromSource_(source,source.students[1],options);}

test('S7 이름으로 학생 행을 찾아 A 수험번호의 앞자리 0과 P 전달사항을 보존한다',()=>{
  const source=fixture(), found=api.selectStudent_(api.matchingStudents_(source.students,' 데모학생 '),'');
  const r=api.buildReportFromSource_(source,found,config);
  assert.equal(found.row,7);assert.equal(r.examNumber,'028001');assert.equal(r.message,source.students[1].display[15]);
});
test('빈 이름과 존재하지 않는 이름을 중단한다',()=>{
  const source=fixture();assert.throws(()=>api.matchingStudents_(source.students,''),/S7/);
  assert.throws(()=>api.matchingStudents_(source.students,'없는학생'),/찾지/);
});
test('동명이인은 수험번호를 지정해야 한 명 선택된다',()=>{
  const s=fixture();s.students.push(row(12,['028002','국','데모학생','','','','',1,2,3,4,5,'','','','']));
  const matches=api.matchingStudents_(s.students,'데모학생');assert.equal(matches.length,2);
  assert.throws(()=>api.selectStudent_(matches,''),/한 명/);assert.equal(api.selectStudent_(matches,'028002').row,12);
});
test('A열 수험번호 누락과 중복을 거부한다',()=>{
  const a=fixture();a.students[1].display[0]='';assert.throws(()=>report(a),/수험번호가 비어/);
  const b=fixture();b.students[0].display[0]='028001';assert.throws(()=>report(b),/수험번호가 중복/);
});
test('G열의 1,2,4,8만 X로 표시하고 나머지는 O로 표시한다',()=>{
  const r=report();assert.equal(r.answers.length,25);
  assert.deepEqual(Array.from(r.answers,(answer,i)=>answer==='X'?i+1:null).filter(Boolean),[1,2,4,8]);
  assert.equal(r.answerStyles[0].color,RED);
});
test('빈 오답 목록·없음·대시는 모두 정답, 숫자와 쉼표·줄바꿈 구분자를 지원한다',()=>{
  for(const text of ['', '없음','-','—'])assert.equal(api.parseWrongAnswers_(text,null).numbers.length,0);
  assert.deepEqual(Array.from(api.parseWrongAnswers_(8,null).numbers),[8]);
  assert.deepEqual(Array.from(api.parseWrongAnswers_('(1번，2번; 4\n8 8)',null).numbers),[1,2,4,8]);
});
test('오답 목록의 범위 밖 숫자·음수·소수·범위 기호·잘못된 문자열을 거부한다',()=>{
  for(const text of ['0,1','26','-1','1.5','1~3','1-3','abc','1, 2 ...'])assert.throws(()=>api.parseWrongAnswers_(text,null),/문제 번호/);
});
test('오답 목록의 문항별 빨강·파랑·검정을 해당 X에 연결한다',()=>{
  const text='1, 12, 25',p=api.parseWrongAnswers_(text,style(text,BLACK,[{start:0,end:1,color:RED},{start:3,end:5,color:BLUE}]));
  assert.equal(p.colors[1],RED);assert.equal(p.colors[12],BLUE);assert.equal(p.colors[25],BLACK);
});
test('시험지데이터는 A/D/E/G를 사용하고 과정은 D 영역 다음 E 갈래 순서다',()=>{
  const r=report();assert.equal(r.questions[0].number,1);assert.equal(r.questions[24].number,25);
  assert.equal(r.questions[0].area,'영역1');assert.equal(r.questions[0].branch,'갈래1');
  const m=api.buildReportModel_(r,config,false);assert.equal(m.find(m=>m.label==='1번 과정').text,'영역1\n갈래1');
  assert.ok(!m.find(m=>m.label==='1번 과정').text.includes('62.2%'));
});
test('숫자 난이도 2와 4가 별 2개와 4개로 표시된다',()=>{
  const m=api.buildReportModel_(report(),config,false);
  assert.equal(m.find(m=>m.label==='2번 난이도').text,'★★');assert.equal(m.find(m=>m.label==='4번 난이도').text,'★★★★');
});
test('난이도는 1~4 정수만 허용하고 빈 D/E 태그를 거부한다',()=>{
  for(const value of ['',0,5,2.5]){const s=fixture();s.questions[0].values[6]=value;assert.throws(()=>report(s),/난이도/);}
  for(const col of [3,4]){const s=fixture();s.questions[0].display[col]='';assert.throws(()=>report(s),/태그/);}
});
test('문항번호 중복과 25문항 누락을 거부한다',()=>{
  const a=fixture();a.questions.pop();assert.throws(()=>report(a),/1~25/);
  const b=fixture();b.questions[0].values[0]=1;assert.throws(()=>report(b),/문항번호/);
});
test('학생 점수 H:I:J:K:L과 평균의 다섯 영역 순서를 유지한다',()=>{
  const r=report();assert.deepEqual(Array.from(r.scores),[4,16,12,4,16]);
  assert.deepEqual(Array.from(r.averages),[12.65,17.76,10.59,7.25,13.5]);assert.equal(r.chartMax,20);
  assert.deepEqual(Array.from(api.REPORT.AXES),['운문 문학','산문 문학','비문학','문법','초등']);
});
test('0점은 보존하고 점수 누락·수식 오류·음수를 거부한다',()=>{
  assert.equal(api.readScore_(0,'운문'),0);
  for(const v of ['', 'abc','#DIV/0!',-1,true])assert.throws(()=>api.readScore_(v,'운문'),/점수/);
  const s=fixture();s.students[1].values[7]='';assert.throws(()=>report(s),/운문 문학/);
  const b=fixture();b.averages[3]='';assert.throws(()=>report(b),/평균 문법/);
});
test('그래프는 평균과 학생의 공통 눈금을 사용하고 최댓값 초과를 거부한다',()=>{
  assert.throws(()=>report(fixture(),{...config,chartMax:10}),/최댓값/);
  assert.equal(report(fixture(),{...config,chartMax:40}).chartMax,40);
  assert.equal(api.niceChartMax_([0,0,0]),1);assert.equal(api.niceChartMax_([21,13.5]),50);
});
test('실제 읽기는 S7·A3:P·A2:G26·H27·M6/M8/M10/M12/M14를 사용한다',()=>{
  const calls=[];
  function sheet(name){return{getLastRow:()=>name==='수험생데이터'?14:27,getParent:()=>({getSpreadsheetTheme:()=>null}),getRange(...args){
    calls.push([name,...args]);if(args[0]==='S7')return{getDisplayValue:()=> '데모학생'};
    if(args[0]==='M6:M14')return{getValues:()=>[[11],[999],[12],[999],[13],[999],[14],[999],[15]]};
    const note=name==='시험지데이터'&&args[0]===27&&args[1]===8;
    const values=Array.from({length:args[2]},(_,i)=>Array.from({length:args[3]},(_,j)=>note?'시험지 H27 설명':j===0?String(i+1):j===2&&name==='수험생데이터'?'데모학생':''));
    return{getValues:()=>values,getDisplayValues:()=>values,getRichTextValues:()=>values.map(r=>r.map(()=>null)),getFontColorObjects:()=>values.map(r=>r.map(()=>colorObject(BLACK)))};
  } };}
  const source=api.loadSourceData_({getSheetByName:name=>sheet(name)});assert.equal(source.requestedName,'데모학생');
  assert.deepEqual(Array.from(source.averages),[11,12,13,14,15]);
  assert.ok(calls.some(c=>JSON.stringify(c)===JSON.stringify(['수험생데이터',3,1,12,16])));
  assert.ok(calls.some(c=>JSON.stringify(c)===JSON.stringify(['시험지데이터',2,1,25,7])));
  assert.ok(calls.some(c=>JSON.stringify(c)===JSON.stringify(['시험지데이터',27,8,1,1])));
  assert.ok(!calls.some(c=>c[1]==='S6'));
  assert.equal(source.questions.length,25);assert.equal(source.examNote,'시험지 H27 설명');
});
test('상단에는 수험 번호와 이름, 같은 학생 행의 D열 학교명을 표시한다',()=>{
  const m=api.buildReportModel_(report(),config,false);assert.equal(m.find(m=>m.label==='수험 번호').text,'028001');
  assert.equal(m.find(m=>m.label==='학생이름').text,'데모학생');assert.equal(m.find(m=>m.label==='학교명').text,'학교');
  assert.ok(m.find(m=>m.label==='학교명라벨'));assert.ok(!m.find(m=>m.label==='반 명라벨'));assert.ok(!m.find(m=>m.label==='응시일'));
  assert.equal(m.find(m=>m.label==='담당 T 전달사항').text,fixture().students[1].display[15]);
});
test('25문항 10+10+5 배치와 텍스트가 양식 안에 들어간다',()=>{
  const model=api.buildReportModel_(report(),config,false);
  for(const suffix of ['번 번호','번 과정','번 난이도','번 정오'])assert.equal(model.filter(m=>m.label?.endsWith(suffix)).length,25);
  assert.ok(model.find(m=>m.label==='25번 번호').x>1422);
  for(const m of model.filter(m=>m.kind==='text')){assert.ok(m.x>=0&&m.y>=0&&m.x+m.w<=3216&&m.y+m.h<=1960,m.label);assert.ok(api.fitText_(m,0.2066).text);}
});
test('평균선 뒤·불투명도 0.25 빨간 면·학생선 앞 순서와 파란 테두리를 만든다',()=>{
  const m=api.buildReportModel_(report(),config,false);
  const indices=f=>m.map((v,i)=>f(v)?i:-1).filter(i=>i>=0);
  const blue=indices(v=>v.kind==='line'&&v.color===api.REPORT.BLUE),fill=indices(v=>v.kind==='triangle'),red=indices(v=>v.kind==='line'&&v.width===5.5);
  assert.equal(fill.length,5);assert.equal(red.length,5);assert.ok(Math.max(...blue)<Math.min(...fill));assert.ok(Math.max(...fill)<Math.min(...red));
  assert.ok(fill.every(i=>m[i].alpha===0.25&&m[i].color===api.REPORT.RED));assert.equal(m.find(m=>m.kind==='rect').color,api.REPORT.BLUE);
});
test('모든 학생 점수가 0일 때 면·길이 0 선·NaN 요청이 없다',()=>{
  const s=fixture();s.students[1].values.splice(7,5,0,0,0,0,0);
  const m=api.buildReportModel_(report(s),config,false);assert.equal(m.filter(m=>m.kind==='triangle').length,0);
  assert.ok(!JSON.stringify(api.modelToRequests_(m,'slide_1',api.pageTransform_(720,405),'test_zero')).includes('null'));
});
test('반투명 native 삼각형 좌표가 세 꼭짓점과 정확히 일치한다',()=>{
  const m=api.buildReportModel_(report(),config,false),t=api.pageTransform_(720,405);
  for(const item of m.filter(m=>m.kind==='triangle')){const p=api.nativeFilledShape_(item,'slide_1',t),tr=p.elementProperties.transform;
    [[0.5,0],[0,1],[1,1]].forEach(([x,y],i)=>{assert.ok(Math.abs(tr.scaleX*x+tr.shearX*y+tr.translateX-(t.x+item.points[i].x*t.scale))<1e-8);assert.ok(Math.abs(tr.shearY*x+tr.scaleY*y+tr.translateY-(t.y+item.points[i].y*t.scale))<1e-8);});
    assert.equal(p.shapeProperties.shapeBackgroundFill.solidFill.alpha,0.25);assert.equal(p.shapeProperties.outline.propertyState,'NOT_RENDERED');
  }
});
test('API 도형 ID·선 끝점이 유효하다',()=>{
  const model=api.buildReportModel_(report(),config,false),t=api.pageTransform_(720,405),requests=api.modelToRequests_(model,'slide_1',t,'test_prefix'),ids=new Set();
  for(const r of requests){assert.equal(Object.keys(r).length,1);const c=r.createShape||r.createLine;if(c){assert.match(c.objectId,/^[a-zA-Z0-9_][a-zA-Z0-9_:-]{4,49}$/);assert.ok(!ids.has(c.objectId));ids.add(c.objectId);}}
  model.forEach((m,i)=>{if(m.kind!=='line')return;const tr=requests.find(r=>r.createLine?.objectId==='test_prefix_'+i).createLine.elementProperties.transform;
    assert.ok(Math.abs(tr.translateX+tr.scaleX+tr.shearX-(t.x+m.x2*t.scale))<1e-8);assert.ok(Math.abs(tr.translateY+tr.shearY+tr.scaleY-(t.y+m.y2*t.scale))<1e-8);});
});
test('가로·세로 페이지 모두 원본 비율을 보존하고 중앙에 배치한다',()=>{
  for(const[w,h]of[[720,405],[900,550],[550,900]]){const t=api.pageTransform_(w,h);assert.ok(t.x>=-1e-8&&t.y>=-1e-8);assert.ok(Math.abs(t.x*2+3216*t.scale-w)<1e-8);assert.ok(Math.abs(t.y*2+1960*t.scale-h)<1e-8);}
});
test('지나치게 긴 전달사항은 자르지 않고 오류로 알린다',()=>{
  const r=report();r.message='아주 긴 전달사항 '.repeat(300);const m=api.buildReportModel_(r,config,false).find(m=>m.label==='담당 T 전달사항');
  assert.throws(()=>api.fitText_(m,0.2),/내용이 칸보다/);
});
test('셀 안의 빨강·파랑·검정 구간과 앞뒤 공백 제거 위치를 읽는다',()=>{
  const raw='  빨강 파랑 검정  ',s=api.readCellColorStyle_(raw,richMock(raw,[{start:0,end:5,color:RED},{start:5,end:8,color:BLUE},{start:8,end:raw.length,color:BLACK}]),colorObject(BLACK),null);
  assert.equal(s.text,'빨강 파랑 검정');assert.deepEqual(JSON.parse(JSON.stringify(s.runs)),[{start:0,end:3,color:RED},{start:3,end:6,color:BLUE},{start:6,end:8,color:BLACK}]);
});
test('테마 색을 원본 시트 RGB 값으로 변환한다',()=>{
  const c={getColorType:()=> 'THEME',asThemeColor:()=>({getThemeColorType:()=> 'ACCENT1'})};assert.equal(api.colorHex_(c,{getConcreteColor:()=>colorObject('#aabbcc')}),'#AABBCC');
});
test('줄바꿈·반복 단어·이모지가 있어도 색이 밀리지 않는다',()=>{
  const text='같은 말\r\n같은 말 😀 끝',second=text.indexOf('같은',1),emoji=text.indexOf('😀');
  const item={text,label:'검사',color:BLACK,colorRuns:[{start:second,end:second+4,color:BLUE},{start:emoji,end:emoji+2,color:RED}]},fitted='같은\n말\n같은\n말 😀\n끝';
  const c=api.characterColors_(fitted,BLACK,api.fittedColorRuns_(item,fitted));assert.equal(c[0],BLACK);assert.equal(c[fitted.indexOf('같은',1)],BLUE);assert.equal(c[fitted.indexOf('😀')],RED);assert.equal(c[fitted.indexOf('😀')+1],RED);assert.equal(c[fitted.lastIndexOf('끝')],BLACK);
});
test('D/E 태그 색과 P 전달사항 색이 native 스타일 요청에 들어간다',()=>{
  const s=fixture();s.questions[24].styles[3]=style('영역1',BLUE);s.questions[24].styles[4]=style('갈래1',RED);
  const model=api.buildReportModel_(report(s),config,false),q=model.find(m=>m.label==='1번 과정'),f=api.fitText_(q,0.2),c=api.characterColors_(f.text,q.color,api.fittedColorRuns_(q,f.text));
  assert.equal(c[0],BLUE);assert.equal(c[f.text.indexOf('갈래')],RED);
  const requests=api.modelToRequests_([model.find(m=>m.label==='담당 T 전달사항')],'slide_1',api.pageTransform_(720,405),'color_test');
  assert.ok(requests.some(r=>r.updateTextStyle?.textRange.type==='FIXED_RANGE'));
});
test('난이도 원본 셀이 빨강·파랑이어도 모든 별을 검정으로 고정한다',()=>{
  const s=fixture();s.questions.forEach(q=>q.styles[6]=style(q.display[6],RED,[{start:0,end:1,color:BLUE}]));
  const stars=api.buildReportModel_(report(s),config,false).filter(m=>m.label?.endsWith('번 난이도'));
  assert.equal(stars.length,25);
  for(const star of stars){
    assert.equal(star.color,BLACK);assert.ok(!star.colorRuns);
    const requests=api.modelToRequests_([star],'slide_1',api.pageTransform_(720,405),'black_star');
    const colors=requests.filter(r=>r.updateTextStyle).map(r=>r.updateTextStyle.style.foregroundColor.opaqueColor.rgbColor);
    assert.ok(colors.every(c=>c.red===0&&c.green===0&&c.blue===0));
  }
});
test('H27 설명을 오른쪽 아래 회색 병합칸 안에 원본 부분 색과 함께 입력한다',()=>{
  const r=report(),m=api.buildReportModel_(r,config,false).find(m=>m.label==='시험지데이터 H27');
  assert.equal(m.text,fixture().examNote);
  assert.ok(m.x>1422&&m.y>1138&&m.x+m.w<2130&&m.y+m.h<1886);
  assert.equal(m.align,'START');assert.equal(m.vertical,'TOP');
  const f=api.fitText_(m,0.2066),colors=api.characterColors_(f.text,m.color,api.fittedColorRuns_(m,f.text));
  assert.equal(colors[f.text.indexOf('조건')],BLUE);
});
test('H27은 학생 공통 값이며 비어 있으면 회색 칸을 비워 둔다',()=>{
  const s=fixture(),a=api.buildReportFromSource_(s,s.students[0],config),b=report(s);
  assert.equal(a.examNote,b.examNote);
  s.examNote='';assert.ok(!api.buildReportModel_(report(s),config,false).find(m=>m.label==='시험지데이터 H27'));
});
test('칸을 넘는 H27 설명은 셀 주소를 포함한 오류로 알린다',()=>{
  const s=fixture();s.examNote='매우 긴 시험지 분석 문구 '.repeat(300);
  const m=api.buildReportModel_(report(s),config,false).find(m=>m.label==='시험지데이터 H27');
  assert.throws(()=>api.fitText_(m,0.2066),/시험지데이터 H27.*내용이 칸보다/);
});
test('Drive 링크·ID를 읽고 로컬 경로는 거부한다',()=>{
  const id='abcdefghij_1234567890';assert.equal(api.extractDriveId_('https://drive.google.com/file/d/'+id+'/view',true),id);assert.equal(api.extractDriveId_('',false),'');assert.throws(()=>api.extractDriveId_('C:/template.png',true),/Drive/);
});

test('공유 이미지의 전체 링크에서 resourcekey를 보존하며 기존 파일 ID도 지원한다',()=>{
  const id='abcdefghij_1234567890';
  for(const link of [
    'https://drive.google.com/file/d/'+id+'/view?usp=sharing&resourcekey=0%2Dabc_DEF#preview',
    'https://drive.google.com/open?id='+id+'&resourcekey=0-abc_DEF'
  ]){
    const ref=api.extractDriveReference_(link,true,'템플릿이미지');
    assert.equal(ref.id,id);assert.equal(ref.resourceKey,'0-abc_DEF');
  }
  assert.equal(api.extractDriveReference_(id,true,'템플릿이미지').resourceKey,'');
  assert.equal(api.extractDriveReference_('',false,'로고이미지').id,'');
});

test('잘못된 링크와 손상된 resourcekey는 입력 위치와 함께 중단한다',()=>{
  assert.throws(()=>api.extractDriveReference_('C:/image.png',true,'템플릿이미지'),/B2.*Drive/);
  assert.throws(()=>api.extractDriveReference_('',true,'템플릿이미지'),/B2/);
  for(const key of ['', '%XX', 'abc%0Adef', 'abc%2Fdef']){
    assert.throws(()=>api.extractDriveReference_('https://drive.google.com/file/d/abcdefghij_1234567890/view?resourcekey='+key,false,'로고이미지'),/B4/);
  }
});

test('보고서설정에서 이미지 접근 키와 기존 슬라이드 ID를 함께 읽는다',()=>{
  const template='https://drive.google.com/file/d/template_file_123456/view?resourcekey=0-template';
  const logo='https://drive.google.com/file/d/logo_file_123456789/view?resourcekey=0-logo';
  const data=[['템플릿이미지',template],['로고이미지',logo],['결과슬라이드','https://docs.google.com/presentation/d/slides_file_12345678/edit']];
  const local=load({Utilities:{formatDate:()=> '11'}},{
    syncReportHeaderSettings_(){},readSourceRows_:()=>data.map((values,i)=>row(i+2,values))
  });
  const settings=local.readConfig_({getSheetByName:()=>({getLastRow:()=>8})});
  assert.equal(settings.templateId,'template_file_123456');assert.equal(settings.templateResourceKey,'0-template');
  assert.equal(settings.logoId,'logo_file_123456789');assert.equal(settings.logoResourceKey,'0-logo');
  assert.equal(settings.deckId,'slides_file_12345678');
});

test('PNG를 읽을 때 resourcekey가 있으면 키 포함 API를 쓰고 없으면 기존 API를 쓴다',()=>{
  const calls=[],bytes=Array.from(fs.readFileSync(path.join(__dirname,'../report-template.png')));
  const blob={getContentType:()=> 'image/png',getBytes:()=>bytes},file={getBlob:()=>blob};
  const local=load({DriveApp:{
    getFileById:id=>{calls.push(['id',id]);return file;},
    getFileByIdAndResourceKey:(id,key)=>{calls.push(['key',id,key]);return file;}
  }});
  assert.equal(local.readImageBlob_('imageid','템플릿이미지','0-access'),blob);
  assert.equal(local.readImageBlob_('imageid','템플릿이미지',''),blob);
  assert.deepEqual(calls,[['key','imageid','0-access'],['id','imageid']]);
});

test('이미지 접근 오류는 실패한 셀·파일 ID·원래 오류를 남긴다',()=>{
  const local=load({DriveApp:{getFileById(){throw new Error('입력한 ID에 해당하는 항목이 없습니다.');}}});
  for(const[label,cell]of[['템플릿이미지','B2'],['로고이미지','B4']]){
    assert.throws(()=>local.readImageBlob_('missing_file_id',label,''),error=>{
      assert.ok(error.message.includes(cell));assert.ok(error.message.includes('missing_file_id'));
      assert.match(error.message,/해당하는 항목/);assert.match(error.message,/현재 버튼을 누른 계정/);return true;
    });
  }
});

function accessHarness(values,overrides={}){
  const alerts=[];
  const menuUi={alert(...args){alerts.push(args);},ButtonSet:{OK:'OK'}};
  const sheet={getLastRow:()=>values.length+1,getRange:()=>({getDisplayValues:()=>values})};
  const local=load({SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:()=>sheet}),getUi:()=>menuUi},...overrides});
  return{local,alerts};
}

test('파일 접근 확인은 양식 실패 후에도 슬라이드를 확인하고 보고서를 생성하지 않는다',()=>{
  const image='https://drive.google.com/file/d/template_file_123456/view?resourcekey=0-access';
  const calls=[];
  const h=accessHarness([['템플릿이미지',image],['로고이미지',''],['결과슬라이드','slides_file_12345678']],{
    DriveApp:{
      getFileByIdAndResourceKey(id,key){calls.push([id,key]);throw new Error('이미지 접근 거부');},
      getFileById:id=>({getMimeType:()=> 'application/vnd.google-apps.presentation'})
    },
    SlidesApp:{openById:id=>({getName:()=> '국어 분석 보고서'}),create(){throw new Error('새 파일 생성 금지');}}
  });
  const result=h.local.checkReportFileAccess();
  assert.deepEqual(Array.from(result,r=>r.status),['오류','생략','확인']);
  assert.match(result[0].message,/B2/);assert.match(result[2].message,/B5/);
  assert.deepEqual(calls,[['template_file_123456','0-access']]);
  assert.match(h.alerts[0][1],/보고서나 새 파일을 만들지 않습니다/);
  assert.match(h.alerts[0][1],/실제 오류가 난 사용자 계정/);
});

test('접근 확인에서 빈 B5는 새 파일을 만들지 않으며 잘못된 B5 형식은 구분한다',()=>{
  const blob={getContentType:()=> 'image/png',getBytes:()=>Array.from(fs.readFileSync(path.join(__dirname,'../report-template.png')))};
  const h=accessHarness([['템플릿이미지','template_file_123456'],['결과슬라이드','']],{
    DriveApp:{getFileById:()=>({getBlob:()=>blob})}
  });
  assert.equal(h.local.checkReportFileAccess()[2].status,'생략');
  const wrong=accessHarness([['템플릿이미지','template_file_123456'],['결과슬라이드','slides_file_12345678']],{
    DriveApp:{getFileById:()=>({getBlob:()=>blob,getMimeType:()=> 'application/vnd.google-apps.folder'})}
  });
  const result=wrong.local.checkReportFileAccess()[2];
  assert.equal(result.status,'오류');assert.match(result.message,/B5/);assert.match(result.message,/Google Slides 파일이 아닙니다/);
});

test('양식 형식과 크기 오류도 B2를 표시한다',()=>{
  for(const blob of [
    {getContentType:()=> 'application/pdf'},
    {getContentType:()=> 'image/png',getBytes:()=>new Array(24).fill(0)}
  ]){
    const local=load({DriveApp:{getFileById:()=>({getBlob:()=>blob})}});
    assert.throws(()=>local.readImageBlob_('wrong_image','템플릿이미지',''),/B2/);
  }
});

function remoteMock({failBatch=false,failProperty=false}={}){
  const stored=new Map([['REPORT_SLIDE_deckid_028001','oldslide']]),calls=[];
  const slide={getObjectId:()=> 'newslide',getBackground:()=>({setSolidFill(){}}),insertImage:()=>({setTitle(){return this;}}),getNotesPage:()=>({getSpeakerNotesShape:()=>({getText:()=>({setText(){}})})})};
  const deck={getPageWidth:()=>720,getPageHeight:()=>405,getSlides:()=>['manualslide','oldslide'].map(id=>({getObjectId:()=>id})),appendSlide:()=>slide,saveAndClose(){}};
  const local=load({PropertiesService:{getScriptProperties:()=>({getProperty:k=>stored.get(k),setProperty(k,v){if(failProperty)throw new Error('property failure');stored.set(k,v);},deleteProperty:k=>stored.delete(k)})},SlidesApp:{openById:()=>deck,PredefinedLayout:{BLANK:'BLANK'}},Utilities:{getUuid:()=> 'abcdefghijklmnop1234567890'},Slides:{Presentations:{batchUpdate(body){calls.push(body.requests);if(failBatch&&calls.length===1)throw new Error('API failure');}}}});
  return{api:local,stored,calls};
}
const ss={getUrl:()=> 'https://docs.google.com/spreadsheets/d/sourceid'};
test('같은 수험번호 갱신은 순서를 유지하고 수동 슬라이드를 보존한다',()=>{
  const m=remoteMock(),url=m.api.renderReport_(ss,'deckid',report(),config,{},null);assert.match(url,/slide=id.newslide$/);assert.equal(m.calls.length,1);assert.equal(m.calls[0].filter(r=>r.deleteObject).length,1);assert.equal(m.calls[0].find(r=>r.deleteObject).deleteObject.objectId,'oldslide');assert.equal(m.stored.get('REPORT_SLIDE_deckid_028001'),'newslide');
});
test('API 실패 시 기존 매핑을 유지하고 임시 페이지만 정리한다',()=>{
  const m=remoteMock({failBatch:true});assert.throws(()=>m.api.renderReport_(ss,'deckid',report(),config,{},null),/API failure/);assert.equal(m.calls[1][0].deleteObject.objectId,'newslide');assert.equal(m.stored.get('REPORT_SLIDE_deckid_028001'),'oldslide');
});
test('갱신 성공 뒤 속성 저장 오류가 나도 성공한 페이지는 삭제하지 않는다',()=>{
  const m=remoteMock({failProperty:true});assert.throws(()=>m.api.renderReport_(ss,'deckid',report(),config,{},null),/property failure/);assert.equal(m.calls.length,1);
});
test('초기 설정은 원본 셀을 쓰지 않고 T6 버튼에 공개 함수를 연결한다',()=>{
  const assigned=[],button={getAltTextTitle:()=> 'ACADEMY_REPORT_GENERATE_BUTTON_V2'};
  for(const name of ['setAnchorCell','setWidth','setHeight','setAltTextTitle','setAltTextDescription'])button[name]=()=>button;
  button.assignScript=name=>{assigned.push(name);return button;};
  const original={getImages:()=>[button],getRange(address){assert.equal(address,'T6');return{};}},settings={getLastRow:()=>8,getRange:()=>({getValues:()=>[]})},log={getLastRow:()=>1,getRange:()=>({getDisplayValues:()=>[Array.from(api.REPORT.LOG_HEADERS)]})};
  const menu=new Proxy({},{get:()=>()=>menu}),spreadsheet={getSheetByName:name=>name==='수험생데이터'?original:name==='시험지데이터'?{}:name==='보고서설정'?settings:log,toast(){}};
  const local=load({SpreadsheetApp:{getActiveSpreadsheet:()=>spreadsheet,getUi:()=>({createMenu:()=>menu})},ScriptApp:{getProjectTriggers:()=>[]},Utilities:{formatDate:()=> '11'}});local.setupReportAutomation();assert.deepEqual(assigned,['generateReportByName']);
});

function buttonHarness(source,functions={}){
  const state={renders:[],logs:[],alerts:[],dialogs:[],reads:0,released:0,lockAllowed:true};
  const ui={alert(...args){state.alerts.push(args);},showModalDialog(html,title){state.dialogs.push({html:html.text,title});}};
  const spreadsheet={...ss,getSheetByName:name=>({getRange:address=>({getDisplayValue(){assert.equal(address,'S7');return source.requestedName;}})})};
  const local=load({SpreadsheetApp:{getActiveSpreadsheet:()=>spreadsheet,getUi:()=>ui,flush(){}},
    Slides:{},LockService:{getDocumentLock:()=>({tryLock:()=>state.lockAllowed,releaseLock(){state.released++;}})},
    HtmlService:{createHtmlOutput:text=>({text,setWidth(){return this;},setHeight(){return this;}})}},
    {loadSourceData_:()=>{state.reads++;return source;},readConfig_:()=>config,
      readImageBlob_:()=>({}),ensureOutputDeck_:()=> 'deckid',
      renderReport_:(spreadsheet,deckid,r)=>{state.renders.push(r);return 'https://docs.google.com/presentation/d/deckid/edit#slide=id.report';},
      appendReportLog_:(spreadsheet,student,status,url,error)=>state.logs.push({student,status,url,error}),...functions});
  return{local,state};
}
test('버튼은 로딩 창을 먼저 열고, 비동기 요청에서 보고서 생성 결과를 반환한다',()=>{
  const h=buttonHarness(fixture());h.local.generateReportByName();
  assert.equal(h.state.renders.length,0);assert.equal(h.state.logs.length,0);assert.equal(h.state.dialogs.length,1);
  assert.match(h.state.dialogs[0].html,/class="spinner"/);assert.match(h.state.dialogs[0].html,/createReportForDialog/);
  const result=h.local.createReportForDialog({requestedName:'데모학생',examNumber:''});
  assert.equal(h.state.alerts.length,0);assert.equal(h.state.renders.length,1);assert.equal(result.state,'complete');
  assert.equal(h.state.renders[0].sourceRow,7);assert.equal(h.state.renders[0].examNumber,'028001');
  assert.equal(result.schoolName,'학교');assert.match(result.url,/presentation\/d\//);assert.equal(h.state.logs[0].status,'완료');
  assert.equal(h.state.released,1);
});
test('동명이인은 선택 결과를 반환하고 다음 요청에서 고른 수험번호 행을 생성한다',()=>{
  const s=fixture();s.students.push(row(12,['028002','국','데모학생','','','','',1,2,3,4,5,'','','','선택한 학생의 메시지']));
  const h=buttonHarness(s);h.local.generateReportByName();
  const first=h.local.createReportForDialog({requestedName:'데모학생',examNumber:''});
  assert.equal(first.state,'choose');assert.equal(first.candidates.length,2);assert.equal(h.state.renders.length,0);assert.equal(h.state.logs.length,0);
  const result=h.local.createReportForDialog({requestedName:'데모학생',examNumber:'028002'});
  assert.equal(result.state,'complete');assert.equal(h.state.alerts.length,0);assert.equal(h.state.reads,2);
  assert.equal(h.state.renders[0].sourceRow,12);assert.equal(h.state.renders[0].message,'선택한 학생의 메시지');
});
test('로딩 창을 연 뒤 S7 이름이 바뀌면 다른 학생의 보고서를 만들지 않는다',()=>{
  const s=fixture(),h=buttonHarness(s);h.local.generateReportByName();s.requestedName='다른학생';
  const result=h.local.createReportForDialog({requestedName:'데모학생'});
  assert.equal(result.state,'error');assert.match(result.message,/S7 이름이 바뀌/);assert.equal(h.state.renders.length,0);
});
test('생성 실패는 오류 결과와 기록으로 반환하고 잠금을 해제한다',()=>{
  const h=buttonHarness(fixture(),{renderReport_(){throw new Error('Drive 파일을 읽지 못했습니다.');}});
  const result=h.local.createReportForDialog({requestedName:'데모학생'});
  assert.equal(result.state,'error');assert.match(result.message,/Drive/);assert.equal(h.state.logs[0].status,'오류');assert.equal(h.state.released,1);assert.equal(h.state.alerts.length,0);
});

test('비동기 생성은 양식·로고의 접근 키를 전달하고 실패한 파일의 단계를 기록한다',()=>{
  const files={...config,templateId:'template_file_123456',templateResourceKey:'0-template',
    logoId:'logo_file_123456789',logoResourceKey:'0-logo'};
  for(const[label,cell]of[['템플릿이미지','B2'],['로고이미지','B4']]){
    const calls=[];
    const h=buttonHarness(fixture(),{readConfig_:()=>files,readImageBlob_(id,name,key){
      calls.push([id,name,key]);if(name===label)throw new Error('파일을 찾을 수 없습니다.');return{};
    }});
    const result=h.local.createReportForDialog({requestedName:'데모학생'});
    assert.equal(result.state,'error');assert.ok(result.message.includes(cell));
    assert.equal(h.state.logs[0].error,result.message);assert.equal(h.state.renders.length,0);
    assert.deepEqual(calls[0],['template_file_123456','템플릿이미지','0-template']);
    if(label==='로고이미지')assert.deepEqual(calls[1],['logo_file_123456789','로고이미지','0-logo']);
  }
});

test('슬라이드 작성 실패는 B5로 알리고 이미지 오류로 표시하지 않는다',()=>{
  const h=buttonHarness(fixture(),{renderReport_(){throw new Error('입력한 ID에 해당하는 항목이 없습니다.');}});
  const result=h.local.createReportForDialog({requestedName:'데모학생'});
  assert.equal(result.state,'error');assert.match(result.message,/단계: 보고서설정!B5 결과 슬라이드 작성/);
  assert.match(result.message,/입력한 ID/);assert.equal(h.state.logs[0].error,result.message);
});
test('처리 중 잠금은 오류로 안내하고 획득하지 않은 잠금을 해제하지 않는다',()=>{
  const h=buttonHarness(fixture());h.state.lockAllowed=false;
  const result=h.local.createReportForDialog({requestedName:'데모학생'});
  assert.equal(result.state,'error');assert.match(result.message,/처리 중/);assert.equal(h.state.released,0);assert.equal(h.state.renders.length,0);
});
test('HTML 창의 초기 이름은 스크립트 태그로 실행되지 않도록 이스케이프한다',()=>{
  const name='</script><img src=x onerror=alert(1)>';
  const html=api.reportDialogHtml_({requestedName:name});
  assert.ok(html.includes('\\u003c/script>'));assert.equal((html.match(/<\/script>/g)||[]).length,1);
  assert.equal(JSON.parse(html.match(/var initial = ([^;]+);/)[1]).requestedName,name);
});
test('학교명 D열의 부분 글자색도 상단 학교명에 반영된다',()=>{
  const s=fixture();s.students[1].styles[3]=style('학교',BLUE);
  const item=api.buildReportModel_(report(s),config,false).find(m=>m.label==='학교명');assert.equal(item.color,BLUE);
});
test('제목은 한국 시간의 생성 월을 사용하고 월말·연말 경계도 처리한다',()=>{
  const local=load({Utilities:{formatDate(date,timeZone,format){
    assert.equal(timeZone,'Asia/Seoul');assert.equal(format,'M');
    return new Intl.DateTimeFormat('en-US',{timeZone,month:'numeric'}).format(date);
  }}});
  for(const[instant,month]of[
    ['2026-10-31T14:59:59Z',10],['2026-10-31T15:00:00Z',11],
    ['2026-12-31T14:59:59Z',12],['2026-12-31T15:00:00Z',1]
  ]){
    const header=local.reportHeader_(new Date(instant));
    assert.equal(header.academy,'국어 날개달기');
    assert.equal(header.title,month+'월 모의고사 국어영역 분석표');
  }
});
test('학원명·당월 제목을 갱신해도 양식·슬라이드 링크와 다른 설정은 유지한다',()=>{
  const rows=[['템플릿이미지','image-file-id'],['학원명','국어학원'],['로고이미지','logo-file-id'],
    ['결과슬라이드','presentation-file-id'],['그래프최댓값',20],['그래프제목','모의고사 출제영역'],['보고서제목','국어 모의고사 분석표']];
  const writes=[];
  const sheet={getLastRow:()=>8,getRange(r,c,n,columns){
    if(r===2&&c===1)return{getValues:()=>rows};
    assert.equal(c,2);return{setValue(value){rows[r-2][1]=value;writes.push(r);}};
  }};
  const header={academy:'국어 날개달기',title:'11월 모의고사 국어영역 분석표'};
  api.syncReportHeaderSettings_(sheet,header);
  assert.deepEqual(writes,[3,8]);assert.equal(rows[0][1],'image-file-id');assert.equal(rows[3][1],'presentation-file-id');
  assert.equal(rows[4][1],20);assert.equal(rows[5][1],'모의고사 출제영역');
  api.syncReportHeaderSettings_(sheet,header);assert.equal(writes.length,2);
});
