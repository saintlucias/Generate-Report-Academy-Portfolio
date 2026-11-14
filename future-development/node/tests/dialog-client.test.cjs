// 브라우저를 열지 않는 클라이언트 단위 검증. 렌더링·글꼴·레이아웃 검증은 포함하지 않습니다.
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const server=vm.createContext({});
vm.runInContext(fs.readFileSync(path.join(__dirname,'../../../apps-script/Code.gs'),'utf8')+'\nglobalThis.render=reportDialogHtml_;',server);
const html=server.render({requestedName:'데모학생'});
const script=html.match(/<script>([\s\S]*?)<\/script>/)[1];

class Element {
  constructor(tag,id=''){this.tagName=tag;this.id=id;this.children=[];this.events={};this.hidden=false;this.dataset={};this._text='';this._value='';}
  set textContent(value){this._text=String(value);this.children=[];this._value='';}
  get textContent(){return this._text;}
  get options(){return this.children;}
  get value(){return this._value;}
  set value(value){this._value=String(value);}
  appendChild(child){this.children.push(child);if(!this._value)this._value=child.value;}
  addEventListener(type,handler){this.events[type]=handler;}
  click(){if(this.events.click)this.events.click();}
  focus(){this.focused=true;}
}
function client(){
  const nodes=new Map();
  for(const match of html.matchAll(/<([a-z]+)\b[^>]*id="([^"]+)"[^>]*>/g))nodes.set(match[2],new Element(match[1],match[2]));
  const views=Array.from(html.matchAll(/<section\b([^>]*data-view="([^"]+)"[^>]*)>/g),m=>{
    const element=new Element('section');element.dataset.view=m[2];element.hidden=/\bhidden\b/.test(m[1]);return element;
  });
  const closes=Array.from(html.matchAll(/<button\b[^>]*class="[^"]*close-button[^"]*"[^>]*>/g),()=>new Element('button'));
  const requests=[];let closed=0;
  const document={getElementById:id=>{assert.ok(nodes.has(id),'Unknown element: '+id);return nodes.get(id);},
    createElement:tag=>new Element(tag),querySelectorAll:selector=>{
      if(selector==='[data-view]')return views;if(selector==='.close-button')return closes;
      throw new Error('Unexpected selector: '+selector);
    }};
  const runner=(success,failure)=>({withSuccessHandler:fn=>runner(fn,failure),withFailureHandler:fn=>runner(success,fn),
    createReportForDialog:payload=>requests.push({payload,success,failure})});
  const google={script:{run:runner(),host:{close(){closed++;}}}};
  const context=vm.createContext({document,google,window:{google}});vm.runInContext(script,context);
  return{nodes,views,closes,requests,context,closed:()=>closed,view:()=>views.find(v=>!v.hidden).dataset.view};
}
const result={state:'complete',name:'데모학생',schoolName:'예시초등학교',examNumber:'28009',title:'11월 모의고사 국어영역 분석표',url:'https://docs.google.com/presentation/d/example-id/edit#slide=id.report'};

test('클라이언트는 로딩 상태에서 한 번 요청하고 처리 중 중복 호출을 막는다',()=>{
  const c=client();assert.equal(c.view(),'loading');assert.equal(c.requests.length,1);
  assert.equal(c.requests[0].payload.requestedName,'데모학생');
  c.context.requestReport('');c.nodes.get('retry').click();assert.equal(c.requests.length,1);
});
test('완료 응답에서 학생·학교·수험번호·링크를 표시하고 닫기를 연결한다',()=>{
  const c=client();c.requests[0].success(result);assert.equal(c.view(),'complete');
  assert.equal(c.nodes.get('done-name').textContent,'데모학생 학생');assert.equal(c.nodes.get('done-school').textContent,'학교 예시초등학교');
  assert.equal(c.nodes.get('done-number').textContent,'수험번호 28009');assert.equal(c.nodes.get('open-report').href,result.url);
  c.closes[0].click();assert.equal(c.closed(),1);
});
test('동명이인 선택을 표시하고 고른 수험번호로 다시 요청한다',()=>{
  const c=client();c.requests[0].success({state:'choose',candidates:[{examNumber:'28001',school:'예시초등학교'},{examNumber:'28009',school:'가상초등학교'}]});
  assert.equal(c.view(),'choose');assert.equal(c.nodes.get('student-choice').options.length,2);
  c.nodes.get('student-choice').value='28009';c.nodes.get('choose-student').click();
  assert.equal(c.view(),'loading');assert.equal(c.requests[1].payload.examNumber,'28009');
  c.requests[1].success(result);assert.equal(c.view(),'complete');
});
test('업무 오류와 통신 실패를 표시하고 사용자가 다시 시도할 수 있다',()=>{
  const c=client();c.requests[0].success({state:'error',message:'H27 내용을 확인하세요.'});
  assert.equal(c.view(),'error');assert.equal(c.nodes.get('error-message').textContent,'H27 내용을 확인하세요.');
  c.nodes.get('retry').click();assert.equal(c.requests.length,2);assert.equal(c.view(),'loading');
  c.requests[1].failure({message:'연결 실패'});assert.equal(c.view(),'error');assert.equal(c.nodes.get('error-message').textContent,'연결 실패');
});
test('잘못된 완료 링크는 열기 버튼에 연결하지 않고 오류로 안내한다',()=>{
  const c=client();c.requests[0].success({...result,url:'javascript:alert(1)'});
  assert.equal(c.view(),'error');assert.equal(c.nodes.get('open-report').href,undefined);
});
