/** 개발용: Apps Script와 같은 좌표 모델에서 독립적인 HTML 배치 미리보기를 만듭니다. */
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = __dirname;
const context = vm.createContext({console});
vm.runInContext(fs.readFileSync(path.join(root, 'Code.gs'), 'utf8') + `
  globalThis.api = {demoReport_, buildReportModel_, fitText_, fittedColorRuns_};`, context);
const config = {academy: '국어 날개달기', chartTitle: '모의고사 출제영역', title: '11월 모의고사 국어영역 분석표', chartMax: 20};
const model = context.api.buildReportModel_(context.api.demoReport_(), config, false);
const escape = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[c]));
const parts = [];
for (const item of model) {
  if (item.kind === 'triangle') {
    parts.push(`<polygon points="${item.points.map(p => p.x + ',' + p.y).join(' ')}" fill="${item.color}" fill-opacity="${item.alpha}"/>`);
  } else if (item.kind === 'rect') {
    parts.push(`<rect x="${item.x}" y="${item.y}" width="${item.w}" height="${item.h}" fill="none" stroke="${item.color}" stroke-width="${item.width}"/>`);
  } else if (item.kind === 'line') {
    parts.push(`<line x1="${item.x1}" y1="${item.y1}" x2="${item.x2}" y2="${item.y2}" stroke="${item.color}" stroke-width="${item.width}"/>`);
  } else {
    const fit = context.api.fitText_(item, 720 / 3216);
    const lines = fit.text.split('\n');
    const lineHeight = fit.size * 1.2;
    const x = item.align === 'START' ? item.x + 2 : item.align === 'END' ? item.x + item.w - 2 : item.x + item.w / 2;
    const startY = item.vertical === 'TOP' ? item.y + fit.size : item.y + (item.h - lines.length * lineHeight) / 2 + fit.size;
    const colors = Array(fit.text.length).fill(item.color);
    context.api.fittedColorRuns_(item, fit.text).forEach(run => {
      for (let i = run.start; i < run.end; i++) colors[i] = run.color;
    });
    let offset = 0;
    parts.push(`<text xml:space="preserve" fill="${item.color}" font-size="${fit.size}" font-weight="${item.bold ? 700 : 400}" text-anchor="${item.align === 'START' ? 'start' : item.align === 'END' ? 'end' : 'middle'}">` +
      lines.map((line, i) => {
        const segments = [];
        for (let j = 0; j < line.length;) {
          const color = colors[offset + j];
          let end = j + 1;
          while (end < line.length && colors[offset + end] === color) end++;
          segments.push(`<tspan fill="${color}">${escape(line.slice(j, end))}</tspan>`);
          j = end;
        }
        offset += line.length + 1;
        return `<tspan x="${x}" y="${startY + i * lineHeight}">${segments.join('')}</tspan>`;
      }).join('') + '</text>');
  }
}
const data = fs.readFileSync(path.join(root, 'report-template.png')).toString('base64');
const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="3216" height="1960" viewBox="0 0 3216 1960" style="font-family:'Noto Sans KR','Malgun Gothic',Arial,sans-serif">
<image href="data:image/png;base64,${data}" width="3216" height="1960"/>${parts.join('\n')}</svg>`;
fs.writeFileSync(path.join(root, 'preview.html'), `<!doctype html><html lang="ko"><meta charset="utf-8"><title>국어 보고서 배치 미리보기</title>
<style>body{margin:0;background:#e9edf3;color:#263449;font:15px 'Malgun Gothic',sans-serif}header{padding:20px 28px}h1{font-size:22px;margin:0 0 8px}p{margin:0}main{margin:0 24px 24px;background:white;box-shadow:0 2px 14px #0002}svg{display:block;width:100%;height:auto}@media print{header{display:none}main{margin:0;box-shadow:none}}</style>
<header><h1>국어 분석 보고서 · S7 이름 검색 예시</h1><p>데모학생 학생을 가정한 합성 데이터입니다. 파란 평균선 뒤·빨간 학생선 앞, 빨간 면 불투명도 0.25. 실제 Google Slides 렌더링과 글꼴·줄바꿈은 다를 수 있습니다.</p></header><main>${svg}</main></html>`, 'utf8');
console.log('preview.html 생성: ' + model.filter(i => i.kind === 'text').length + '개 텍스트, ' + model.filter(i => i.kind === 'line').length + '개 선');
