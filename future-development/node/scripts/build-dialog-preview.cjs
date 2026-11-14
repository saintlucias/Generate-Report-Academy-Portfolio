const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const context=vm.createContext({});
vm.runInContext(fs.readFileSync(path.resolve(__dirname,'../../../apps-script/Code.gs'),'utf8')+'\nglobalThis.render=reportDialogHtml_;',context);
let html=context.render({requestedName:'데모학생'});
const previewStyle=`<style>
body{background:#edf1f7;padding:28px 18px 40px}.preview-bar{max-width:650px;margin:0 auto 20px;text-align:center;color:#52647d;font:13px/1.8 'Malgun Gothic',sans-serif}.preview-bar strong{font-size:18px;display:block;color:#2d425f}.preview-bar nav{display:flex;gap:8px;justify-content:center;flex-wrap:wrap;margin:13px 0}.preview-bar button{min-height:0;padding:7px 15px;background:#fff;color:#4d6282;border:1px solid #d1dbe9;border-radius:20px;font-size:12px}.preview-bar button:hover{background:#e7eef9}.preview-bar output{min-height:22px;display:block;color:#7b899f}.preview-frame{width:500px;max-width:100%;margin:auto;background:white;border-radius:14px;box-shadow:0 12px 50px #223c6520;overflow:hidden;border:1px solid #e1e6ee}.preview-chrome{padding:18px 24px 12px;font-size:17px;color:#3e4e65;display:flex;justify-content:space-between}.preview-chrome span{font-weight:300;color:#7e8b9e}.shell{min-height:570px}.preview-bar p{font-size:13px}
</style>`;
const bridge=`<script>
var previewModeValue='auto';var previewTicket=0;
function previewNote(message){var node=document.getElementById('preview-note');if(node)node.textContent=message;}
function makePreviewRunner(success,failure){return{
 withSuccessHandler:function(fn){return makePreviewRunner(fn,failure);},
 withFailureHandler:function(fn){return makePreviewRunner(success,fn);},
 createReportForDialog:function(request){
  var ticket=++previewTicket;var mode=previewModeValue;
  if(mode==='loading')return;
  setTimeout(function(){
   if(ticket!==previewTicket)return;
   if(mode==='error'){success({state:'error',message:'시험지데이터 H27 내용이 칸보다 깁니다. 해당 입력값을 줄여 주세요.'});return;}
   if(mode==='choose'&&!request.examNumber){success({state:'choose',requestedName:'데모학생',candidates:[{examNumber:'28001',school:'예시초등학교'},{examNumber:'28009',school:'가상초등학교'}]});return;}
   var number=request.examNumber||'28009';
   success({state:'complete',name:'데모학생',examNumber:number,schoolName:number==='28001'?'예시초등학교':'가상초등학교',title:'11월 모의고사 국어영역 분석표',url:'https://docs.google.com/presentation/d/DEMO_PREVIEW_ONLY/edit'});
  },mode==='auto'?2400:350);
 }
};}
window.google={script:{run:makePreviewRunner(),host:{close:function(){previewNote('닫기 버튼 동작을 확인했습니다. 실제 시트에서는 창이 닫힙니다.');}}}};
function previewMode(mode){previewModeValue=mode;previewNote('');busy=false;requestReport('');}
document.addEventListener('click',function(event){if(event.target.closest('#open-report')){event.preventDefault();previewNote('미리보기용 버튼입니다. 실제 보고서는 구글 시트에서 생성하세요.');}});
</script>`;
const toolbar=`<div class="preview-bar"><strong>보고서 알림 창 · 디자인·동작 미리보기</strong><p>예시 데이터를 사용합니다. 실제 보고서를 생성하지 않습니다.</p><nav><button type="button" onclick="previewMode('loading')">로딩</button><button type="button" onclick="previewMode('complete')">완료</button><button type="button" onclick="previewMode('error')">오류</button><button type="button" onclick="previewMode('choose')">동명이인</button></nav><output id="preview-note" aria-live="polite"></output></div><div class="preview-frame"><div class="preview-chrome">보고서 생성<span>×</span></div>`;
html=html.replace('</head>',previewStyle+bridge+'</head>').replace('<body>','<body>'+toolbar).replace('</main>','</main></div>');
fs.writeFileSync(path.resolve(__dirname,'../../../apps-script/dialog-preview.html'),html,'utf8');
console.log('dialog-preview.html 생성: 실제 클라이언트 UI + 예시 응답');
