/**
 * 국어 보고서 자동화 — 수험생데이터!S7 이름 입력 → 버튼 클릭.
 * Google Sheets에 연결된 Apps Script / V8 / Google Slides 고급 서비스.
 * 설치: setupReportAutomation 실행. 원본 데이터 셀은 변경하지 않습니다.
 */
const REPORT = Object.freeze({
  WIDTH: 3216, HEIGHT: 1960, QUESTION_COUNT: 25,
  STUDENTS: '수험생데이터', QUESTIONS: '시험지데이터',
  SETTINGS: '보고서설정', LOG: '보고서생성기록', NAME_CELL: 'S7', BUTTON_CELL: 'T6',
  STUDENT_FIRST_ROW: 3, QUESTION_FIRST_ROW: 2,
  AVERAGE_CELLS: ['M6', 'M8', 'M10', 'M12', 'M14'],
  AXES: ['운문 문학', '산문 문학', '비문학', '문법', '초등'],
  BLUE: '#4472C4', RED: '#E53935', FILL_ALPHA: 0.25,
  ACADEMY_NAME: '국어 날개달기', TIME_ZONE: 'Asia/Seoul',
  ROW_Y: [380, 526, 678, 836, 988, 1138, 1298, 1450, 1596, 1740, 1886],
  COL_X: [[63, 199, 462, 640, 743], [743, 874, 1136, 1317, 1422], [1422, 1556, 1822, 2012, 2130]],
  LOG_HEADERS: ['생성시각', '수험번호', '학생명', '원본행', '상태', '슬라이드URL', '오류내용'],
  BUTTON_TAG: 'ACADEMY_REPORT_GENERATE_BUTTON_V2'
});

function onOpen() {
  SpreadsheetApp.getUi().createMenu('보고서 자동화')
    .addItem('초기 설정·생성 버튼 준비', 'setupReportAutomation')
    .addItem('S7 학생 보고서 생성', 'generateReportByName')
    .addItem('파일 접근 권한 확인', 'checkReportFileAccess')
    .addItem('생성 버튼 다시 배치', 'installReportButton')
    .addItem('기존 기록의 오류 글자색 적용', 'formatReportLogErrors')
    .addToUi();
}

/** 전용 설정·기록 탭과 T6 위의 버튼만 준비합니다. A:P 및 S7의 값·색은 보존합니다. */
function setupReportAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  requiredSheet_(ss, REPORT.STUDENTS);
  requiredSheet_(ss, REPORT.QUESTIONS);
  const heading = reportHeader_();
  let settings = ss.getSheetByName(REPORT.SETTINGS);
  if (!settings) settings = ss.insertSheet(REPORT.SETTINGS);
  if (!settings.getLastRow()) {
    settings.getRange(1, 1, 8, 3).setValues([
      ['설정항목', '입력값', '설명'],
      ['템플릿이미지', '', 'report-template.png의 Drive 파일 링크 또는 ID (필수)'],
      ['학원명', heading.academy, '자동 입력: 국어 날개달기'],
      ['로고이미지', '', '선택: Drive PNG/JPG/GIF 링크 또는 ID'],
      ['결과슬라이드', '', '비우면 최초 생성 때 새 Google Slides를 만들고 링크 기록'],
      ['그래프최댓값', '', '비우면 평균·학생 점수에 맞춰 자동 설정. 고정하려면 양수 입력'],
      ['그래프제목', '모의고사 출제영역', '오른쪽 위 오각형 그래프 제목'],
      ['보고서제목', heading.title, '자동 입력: 한국 시간 기준 생성 월 + 월 모의고사 국어영역 분석표']
    ]);
    formatOwnedSheet_(settings, 3);
    settings.setColumnWidth(1, 160).setColumnWidth(2, 430).setColumnWidth(3, 630);
  }
  syncReportHeaderSettings_(settings, heading);
  colorLogErrorStatuses_(ensureLogSheet_(ss));
  // 이전 체크박스 방식의 이 프로젝트 트리거만 정리합니다.
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'handleReportRequest')
    .forEach(t => ScriptApp.deleteTrigger(t));
  installReportButton();
  onOpen();
  ss.toast('보고서설정 B2에 양식 이미지 링크를 입력하세요. 수험생데이터 S7 입력 후 생성 버튼을 누릅니다.', '설정 완료', 10);
}

function requiredSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('필요한 탭을 찾지 못했습니다: ' + name);
  return sheet;
}

function formatOwnedSheet_(sheet, columns) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, columns).setBackground('#334766').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setColumnWidths(1, columns, 150);
  sheet.getRange(1, 1, sheet.getMaxRows(), columns).setWrap(true).setVerticalAlignment('middle');
}

function ensureLogSheet_(ss) {
  let sheet = ss.getSheetByName(REPORT.LOG);
  if (!sheet) sheet = ss.insertSheet(REPORT.LOG);
  if (!sheet.getLastRow()) {
    sheet.getRange(1, 1, 1, REPORT.LOG_HEADERS.length).setValues([REPORT.LOG_HEADERS]);
    formatOwnedSheet_(sheet, REPORT.LOG_HEADERS.length);
    sheet.setColumnWidth(6, 470).setColumnWidth(7, 560);
  }
  const headers = sheet.getRange(1, 1, 1, REPORT.LOG_HEADERS.length).getDisplayValues()[0];
  if (headers.some((value, i) => value !== REPORT.LOG_HEADERS[i])) throw new Error(REPORT.LOG + ' 탭의 헤더가 다릅니다.');
  return sheet;
}

/** 기존 기록 중 E열 상태가 '오류'인 셀에만 빨간 글자색을 적용합니다. */
function colorLogErrorStatuses_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const values = sheet.getRange(2, 5, lastRow - 1, 1).getDisplayValues();
  const errors = values.map((row, i) => clean_(row[0]) === '오류' ? 'E' + (i + 2) : null)
    .filter(address => address !== null);
  if (errors.length) sheet.getRangeList(errors).setFontColor('#D32F2F');
  return errors.length;
}

/** 코드 업데이트 후 한 번 실행하면 기존 오류 기록에도 적용됩니다. */
function formatReportLogErrors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = colorLogErrorStatuses_(ensureLogSheet_(ss));
  ss.toast('기존 오류 상태 ' + count + '건에 빨간 글자색을 적용했습니다.', '보고서생성기록', 5);
}

/** T6에 셀 위 이미지 버튼을 만듭니다. 버튼 함수명에는 괄호를 붙이지 않습니다. */
function installReportButton() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = requiredSheet_(ss, REPORT.STUDENTS);
  const anchor = sheet.getRange(REPORT.BUTTON_CELL);
  let button = sheet.getImages().find(image => image.getAltTextTitle() === REPORT.BUTTON_TAG);
  if (!button) {
    const blob = Utilities.newBlob(Utilities.base64Decode(REPORT_BUTTON_PNG), 'image/png', '보고서생성.png');
    button = sheet.insertImage(blob, anchor.getColumn(), anchor.getRow());
  }
  button.setAnchorCell(anchor).setWidth(150).setHeight(44)
    .setAltTextTitle(REPORT.BUTTON_TAG).setAltTextDescription('S7에 입력한 학생의 보고서 생성')
    .assignScript('generateReportByName');
}

/** 사용자가 누를 버튼의 공개 함수입니다. */
function generateReportByName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  const sheet = ss.getSheetByName(REPORT.STUDENTS);
  const requestedName = sheet ? clean_(sheet.getRange(REPORT.NAME_CELL).getDisplayValue()) : '';
  const html = HtmlService.createHtmlOutput(reportDialogHtml_({requestedName: requestedName}))
    .setWidth(500).setHeight(590);
  SpreadsheetApp.getUi().showModalDialog(html, '보고서 생성');
}

/** HTML 창이 비동기로 호출합니다. 반환값에는 문자열·숫자·배열만 사용합니다. */
function createReportForDialog(request) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let selected = null;
  let lock = null;
  let locked = false;
  let phase = '학생 데이터 읽기';
  try {
    if (!request || typeof request.requestedName !== 'string') {
      throw new Error('수험생데이터 탭의 보고서 생성 버튼으로 실행하세요.');
    }
    SpreadsheetApp.flush();
    const source = loadSourceData_(ss);
    if (clean_(request.requestedName) !== source.requestedName) {
      throw new Error('S7 이름이 바뀌었습니다. 창을 닫고 보고서 생성 버튼을 다시 누르세요.');
    }
    const candidates = matchingStudents_(source.students, source.requestedName);
    const chosenNumber = clean_(request.examNumber);
    if (candidates.length > 1 && !chosenNumber) {
      return {state: 'choose', requestedName: source.requestedName,
        candidates: candidates.map(student => ({examNumber: clean_(student.display[0]), school: clean_(student.display[3])}))};
    }
    selected = selectStudent_(candidates, chosenNumber);
    lock = LockService.getDocumentLock();
    locked = lock.tryLock(1000);
    if (!locked) throw new Error('다른 보고서를 처리 중입니다. 잠시 후 다시 시도하세요.');
    if (typeof Slides === 'undefined') throw new Error('Apps Script 서비스 + 에서 Google Slides API를 추가하세요.');
    phase = '보고서설정 읽기';
    const config = readConfig_(ss);
    phase = '보고서 입력값 확인';
    const report = buildReportFromSource_(source, selected, config);
    phase = '보고서설정!B2 양식 이미지 읽기';
    const template = readImageBlob_(config.templateId, '템플릿이미지', config.templateResourceKey);
    phase = '보고서설정!B4 로고 이미지 읽기';
    const logo = config.logoId ? readImageBlob_(config.logoId, '로고이미지', config.logoResourceKey) : null;
    phase = '보고서설정!B5 결과 슬라이드 준비';
    const deckId = ensureOutputDeck_(ss, config);
    phase = '보고서설정!B5 결과 슬라이드 작성';
    const url = renderReport_(ss, deckId, report, config, template, logo);
    phase = '보고서생성기록 저장';
    appendReportLog_(ss, selected, '완료', url, '');
    return {state: 'complete', name: report.name, schoolName: report.schoolName, examNumber: report.examNumber, title: report.title, url: url};
  } catch (error) {
    const message = ('단계: ' + phase + '\n' + String(error.message || error)).slice(0, 1500);
    try { appendReportLog_(ss, selected, '오류', '', message); }
    catch (logError) { console.warn(logError.message); }
    return {state: 'error', message: message};
  } finally {
    if (locked) lock.releaseLock();
  }
}

function reportDialogHtml_(initial) {
  const json = JSON.stringify(initial || {}).replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
  return REPORT_DIALOG_HTML.replace('__REPORT_INITIAL_DATA__', json);
}

function escapeHtml_(text) {
  return String(text).replace(/[&<>"']/g, char => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[char]));
}

function appendReportLog_(ss, student, status, url, error) {
  const log = ensureLogSheet_(ss);
  log.appendRow([new Date(), student ? student.display[0] : '', student ? student.display[2] : '',
    student ? student.row : '', status, url, error.slice(0, 1500)]);
  const row = log.getLastRow();
  log.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  log.getRange(row, 5).setFontColor(status === '오류' ? '#D32F2F' : '#000000');
}

/** 값, 표시값(수험번호의 앞자리 0 포함), 부분 글자색을 같은 범위에서 읽습니다. */
function readSourceRows_(sheet, firstRow, firstColumn, count, columns) {
  if (count <= 0) return [];
  const range = sheet.getRange(firstRow, firstColumn, count, columns);
  const values = range.getValues();
  const display = range.getDisplayValues();
  const rich = range.getRichTextValues();
  const fontColors = range.getFontColorObjects();
  const theme = sheet.getParent().getSpreadsheetTheme();
  return values.map((row, i) => ({row: firstRow + i, values: row, display: display[i],
    styles: row.map((value, j) => readCellColorStyle_(value, rich[i][j], fontColors[i][j], theme))}));
}

function loadSourceData_(ss) {
  const students = requiredSheet_(ss, REPORT.STUDENTS);
  const questions = requiredSheet_(ss, REPORT.QUESTIONS);
  const examNote = readSourceRows_(questions, 27, 8, 1, 1)[0]; // 시험지데이터!H27
  return {
    requestedName: clean_(students.getRange(REPORT.NAME_CELL).getDisplayValue()),
    students: readSourceRows_(students, 3, 1, students.getLastRow() - 2, 16)
      .filter(row => clean_(row.display[2])),
    // 문항은 A2:G26만 읽어 H27 설명이 26번째 문항으로 섞이지 않게 합니다.
    questions: readSourceRows_(questions, 2, 1, REPORT.QUESTION_COUNT, 7)
      .filter(row => clean_(row.display[0])),
    examNote: clean_(examNote.display[0]), examNoteStyle: examNote.styles[0],
    averages: students.getRange('M6:M14').getValues().filter((_, i) => i % 2 === 0).map(row => row[0])
  };
}

function matchingStudents_(students, requestedName) {
  const name = clean_(requestedName);
  if (!name) throw new Error('수험생데이터!S7에 학생 이름을 입력하세요.');
  const matches = students.filter(student => clean_(student.display[2]) === name);
  if (!matches.length) throw new Error('C3:C에서 학생을 찾지 못했습니다: ' + name);
  return matches;
}

function selectStudent_(candidates, examNumber) {
  const found = examNumber ? candidates.filter(student => clean_(student.display[0]) === examNumber) : candidates;
  if (found.length !== 1) throw new Error('수험번호로 학생을 한 명 선택할 수 없습니다. A열의 수험번호 중복 또는 입력값을 확인하세요.');
  return found[0];
}

/** 실행 호스트·시트의 시간대와 관계없이 한국 시간의 생성 월을 사용합니다. */
function reportHeader_(generatedAt) {
  const month = Utilities.formatDate(generatedAt || new Date(), REPORT.TIME_ZONE, 'M');
  return {academy: REPORT.ACADEMY_NAME, title: month + '월 모의고사 국어영역 분석표'};
}

/** 이전 설정에 남은 학원명·제목만 갱신하고 이미지·결과 파일 링크는 보존합니다. */
function syncReportHeaderSettings_(sheet, heading) {
  if (sheet.getLastRow() < 2) return;
  const generated = {'학원명': heading.academy, '보고서제목': heading.title};
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  rows.forEach((row, i) => {
    if (Object.prototype.hasOwnProperty.call(generated, row[0]) && clean_(row[1]) !== generated[row[0]]) {
      sheet.getRange(i + 2, 2).setValue(generated[row[0]]);
    }
  });
}

function readConfig_(ss) {
  const sheet = ss.getSheetByName(REPORT.SETTINGS);
  if (!sheet) throw new Error('setupReportAutomation을 먼저 실행하세요.');
  const heading = reportHeader_();
  syncReportHeaderSettings_(sheet, heading);
  const map = {}, styles = {};
  readSourceRows_(sheet, 2, 1, sheet.getLastRow() - 1, 3).forEach(row => {
    map[row.values[0]] = row.values[1]; styles[row.values[0]] = row.styles[1];
  });
  const max = clean_(map['그래프최댓값']) ? Number(map['그래프최댓값']) : null;
  if (max !== null && (!Number.isFinite(max) || max <= 0)) throw new Error('그래프최댓값은 비워 두거나 0보다 큰 숫자를 입력하세요.');
  const template = extractDriveReference_(map['템플릿이미지'], true, '템플릿이미지');
  const logo = extractDriveReference_(map['로고이미지'], false, '로고이미지');
  const deck = extractDriveReference_(map['결과슬라이드'], false, '결과슬라이드');
  return {
    templateId: template.id, templateResourceKey: template.resourceKey, academy: heading.academy,
    logoId: logo.id, logoResourceKey: logo.resourceKey, deckId: deck.id,
    chartMax: max, chartTitle: clean_(map['그래프제목']) || '모의고사 출제영역',
    title: heading.title,
    styles: {academy: styles['학원명'], chartTitle: styles['그래프제목'], title: styles['보고서제목']}
  };
}

function readScore_(value, label) {
  if (clean_(value) === '' || typeof value === 'boolean') throw new Error(label + ' 점수가 비어 있거나 숫자가 아닙니다.');
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) throw new Error(label + ' 점수에 0 이상의 숫자를 입력하세요.');
  return number;
}

/** 빈 오답 목록은 모두 정답. 목록의 숫자별 글자색을 해당 X에 적용합니다. */
function parseWrongAnswers_(value, style) {
  const text = clean_(value);
  if (!text || ['없음', '-', '—'].includes(text)) return {numbers: [], colors: {}};
  const residue = text.replace(/\d+/g, '').replace(/번/g, '');
  if (!/^[\s,，、;；()\[\]{}]*$/.test(residue)) throw new Error('틀린 문제 번호는 1, 2, 4처럼 입력하세요. 범위·소수·기호는 사용할 수 없습니다.');
  const matches = Array.from(text.matchAll(/\d+/g));
  if (!matches.length) throw new Error('틀린 문제 번호에 숫자가 없습니다.');
  const palette = characterColors_(text, style ? style.color : '#000000', style && style.text === text ? style.runs : []);
  const colors = {};
  matches.forEach(match => {
    const number = Number(match[0]);
    if (!Number.isInteger(number) || number < 1 || number > 25) throw new Error('틀린 문제 번호는 1~25입니다: ' + match[0]);
    if (!(number in colors)) colors[number] = palette[match.index];
  });
  return {numbers: Object.keys(colors).map(Number).sort((a, b) => a - b), colors: colors};
}

function niceChartMax_(values) {
  const largest = Math.max(...values);
  if (!largest) return 1;
  const power = Math.pow(10, Math.floor(Math.log10(largest)));
  return [1, 2, 5, 10].map(factor => factor * power).find(value => value >= largest);
}

function buildReportFromSource_(source, student, config) {
  const examNumber = clean_(student.display[0]);
  const name = clean_(student.display[2]);
  if (!examNumber) throw new Error('선택한 학생의 A' + student.row + ' 수험번호가 비어 있습니다.');
  if (source.students.filter(row => clean_(row.display[0]) === examNumber).length !== 1) throw new Error('A열 수험번호가 중복됩니다: ' + examNumber);
  if (source.questions.length !== 25) throw new Error('시험지데이터 A2:A에 문항 1~25를 정확히 한 번씩 입력하세요.');
  const byNumber = {};
  source.questions.forEach(row => {
    const number = Number(row.values[0]);
    const difficulty = Number(row.values[6]);
    if (!Number.isInteger(number) || number < 1 || number > 25 || byNumber[number]) throw new Error('문항번호는 중복 없이 정수 1~25입니다.');
    if (!Number.isInteger(difficulty) || difficulty < 1 || difficulty > 4) throw new Error(number + '번 난이도는 정수 1~4입니다.');
    const area = clean_(row.display[3]), branch = clean_(row.display[4]);
    if (!area || !branch) throw new Error(number + '번 D열 영역 태그와 E열 갈래 태그를 모두 입력하세요.');
    byNumber[number] = {number: number, area: area, branch: branch, difficulty: difficulty,
      styles: {number: row.styles[0], area: row.styles[3], branch: row.styles[4], difficulty: row.styles[6]}};
  });
  const wrong = parseWrongAnswers_(student.values[6], student.styles[6]);
  const scores = student.values.slice(7, 12).map((value, i) => readScore_(value, REPORT.AXES[i]));
  const averages = source.averages.map((value, i) => readScore_(value, '평균 ' + REPORT.AXES[i]));
  const chartMax = config.chartMax === null || config.chartMax === undefined ? niceChartMax_(scores.concat(averages)) : config.chartMax;
  if (scores.concat(averages).some(value => value > chartMax)) throw new Error('그래프최댓값보다 큰 점수가 있습니다. 설정을 높이거나 비워 자동 설정하세요.');
  return {
    id: examNumber, examNumber: examNumber, name: name, sourceRow: student.row, title: config.title,
    schoolName: clean_(student.display[3]), message: clean_(student.display[15]), examNote: clean_(source.examNote),
    questions: Array.from({length: 25}, (_, i) => byNumber[i + 1]),
    wrongNumbers: wrong.numbers,
    answers: Array.from({length: 25}, (_, i) => wrong.numbers.includes(i + 1) ? 'X' : 'O'),
    answerStyles: Array.from({length: 25}, (_, i) => ({text: wrong.numbers.includes(i + 1) ? 'X' : 'O',
      color: wrong.colors[i + 1] || '#000000', runs: []})),
    scores: scores, averages: averages, chartMax: chartMax,
    styles: {examNumber: student.styles[0], name: student.styles[2], schoolName: student.styles[3], message: student.styles[15], examNote: source.examNoteStyle || null}
  };
}

function buildReportModel_(report, config, hasLogo) {
  const items = [];
  inputHeader_(items, report, config, hasLogo);
  inputQuestions_(items, report);
  inputExamNote_(items, report);
  inputRadarChart_(items, report, config);
  inputTeacherMessage_(items, report);
  return items;
}

function inputHeader_(items, report, config, hasLogo) {
  text_(items, '학원명', styledText_(config.academy, (config.styles || {}).academy), hasLogo ? 253 : 85, 82, hasLogo ? 285 : 450, 142, 48, {bold: true});
  text_(items, '보고서제목', styledText_(report.title, (config.styles || {}).title), 588, 80, 1925, 145, 68, {minSize: 42, bold: true});
  ['수험 번호', '학교명', '학생이름'].forEach((label, i) => {
    text_(items, label + '라벨', label, 2543, [66, 123, 183][i], 285, [55, 58, 59][i], 34, {bold: true});
    const value = [report.examNumber, report.schoolName, report.name][i];
    text_(items, label, styledText_(value, [report.styles.examNumber, report.styles.schoolName, report.styles.name][i]),
      2847, [66, 123, 183][i], 293, [55, 58, 59][i], 34);
  });
}

function inputQuestions_(items, report) {
  report.questions.forEach((q, index) => {
    const block = Math.floor(index / 10), row = index % 10;
    const xs = REPORT.COL_X[block], y = REPORT.ROW_Y[row] + 3;
    const h = REPORT.ROW_Y[row + 1] - y - 3;
    const course = joinStyledText_([styledText_(q.area, q.styles.area), styledText_(q.branch, q.styles.branch)], '\n');
    text_(items, q.number + '번 번호', styledText_(q.number + '번', q.styles.number), xs[0] + 3, y, xs[1] - xs[0] - 6, h, 29);
    text_(items, q.number + '번 과정', course, xs[1] + 3, y, xs[2] - xs[1] - 6, h, 29, {minSize: 22, keepLines: true});
    text_(items, q.number + '번 난이도', '★'.repeat(q.difficulty), xs[2] + 3, y, xs[3] - xs[2] - 6, h, 31,
      {minSize: 24, bold: true, color: '#000000'});
    text_(items, q.number + '번 정오', styledText_(report.answers[index], report.answerStyles[index]), xs[3] + 3, y, xs[4] - xs[3] - 6, h, 34, {bold: report.answers[index] === 'X'});
  });
}

function inputExamNote_(items, report) {
  text_(items, '시험지데이터 H27', styledText_(report.examNote, report.styles.examNote), 1440, 1158, 670, 706, 32,
    {minSize: 24, align: 'START', vertical: 'TOP'});
}

function inputTeacherMessage_(items, report) {
  text_(items, '담당 T 전달사항 제목', '담당 T 전달사항', 2193, 1220, 937, 80, 43, {bold: true, color: REPORT.BLUE, align: 'START'});
  text_(items, '담당 T 전달사항', styledText_(report.message, report.styles.message), 2193, 1315, 937, 548, 34,
    {minSize: 27, align: 'START', vertical: 'TOP'});
}

function inputRadarChart_(items, report, config) {
  items.push({kind: 'rect', x: 2175, y: 283, w: 975, h: 849, color: REPORT.BLUE, width: 5});
  text_(items, '그래프제목', styledText_(config.chartTitle, (config.styles || {}).chartTitle), 2205, 303, 915, 70, 43, {color: '#666666'});
  const cx = 2662, cy = 785, radius = 248;
  const point = (axis, value) => ({x: cx + Math.sin(axis * Math.PI * 2 / 5) * radius * value / report.chartMax,
    y: cy - Math.cos(axis * Math.PI * 2 / 5) * radius * value / report.chartMax});
  for (let level = 1; level <= 4; level++) {
    for (let axis = 0; axis < 5; axis++) {
      const a = point(axis, report.chartMax * level / 4), b = point((axis + 1) % 5, report.chartMax * level / 4);
      line_(items, a.x, a.y, b.x, b.y, '#D9DDE3', 1.7);
    }
    // Slides의 내부 여백 때문에 두 자리 숫자가 자동 줄바꿈되지 않도록 폭을 확보합니다.
    text_(items, '그래프눈금' + level, String(Math.round(report.chartMax * level / 4 * 100) / 100),
      cx - 152, cy - radius * level / 4 - 29, 140, 58, 19,
      {minSize: 19, keepLines: true, align: 'END', color: '#737373'});
  }
  for (let axis = 0; axis < 5; axis++) {
    const p = point(axis, report.chartMax);
    line_(items, cx, cy, p.x, p.y, '#E2E5EA', 1.5);
  }
  const labels = [[2530,472,264,55],[2930,680,205,60],[2780,1010,235,60],[2320,1010,220,60],[2195,680,190,60]];
  labels.forEach((v, i) => text_(items, '영역 ' + REPORT.AXES[i], REPORT.AXES[i], v[0], v[1], v[2], v[3], 33, {color: '#747474'}));
  line_(items, 2375, 410, 2435, 410, REPORT.BLUE, 5);
  text_(items, '평균 범례', '평균점수', 2440, 387, 230, 46, 28, {align: 'START'});
  line_(items, 2700, 410, 2760, 410, REPORT.RED, 5);
  text_(items, '학생 범례', report.name + ' 점수', 2765, 387, 310, 46, 28, {minSize: 22, align: 'START'});
  // z 순서: 평균선 → 학생 반투명 면 → 학생선. 두 그래프는 같은 좌표축에 겹칩니다.
  const averagePoints = report.averages.map((value, axis) => point(axis, value));
  const studentPoints = report.scores.map((value, axis) => point(axis, value));
  for (let axis = 0; axis < 5; axis++) {
    const a = averagePoints[axis], b = averagePoints[(axis + 1) % 5];
    line_(items, a.x, a.y, b.x, b.y, REPORT.BLUE, 5);
  }
  for (let axis = 0; axis < 5; axis++) {
    const a = {x: cx, y: cy}, b = studentPoints[axis], c = studentPoints[(axis + 1) % 5];
    if (Math.abs((b.x - a.x) * (c.y - a.y) - (c.x - a.x) * (b.y - a.y)) > 0.001) {
      items.push({kind: 'triangle', points: [a,b,c], color: REPORT.RED, alpha: REPORT.FILL_ALPHA});
    }
  }
  for (let axis = 0; axis < 5; axis++) {
    const a = studentPoints[axis], b = studentPoints[(axis + 1) % 5];
    line_(items, a.x, a.y, b.x, b.y, REPORT.RED, 5.5);
  }
}

/** 예시 미리보기 전용 합성 데이터. 사용자의 데이터 탭에 쓰지 않습니다. */
function demoReport_() {
  const message = '시의 화자와 상황을 파악하는 연습이 필요합니다.\n\n산문 문학에서는 인물의 말과 행동을 근거로 정답을 찾는 연습을 진행합니다.\n\n문법 오답은 개념을 정리한 뒤 유사 문항을 다시 풀어 보세요.\n\n이 문장은 색상·배치 확인용 예시입니다.';
  const source = {
    requestedName: '데모학생', averages: [12.65,17.76,10.59,7.2,13.5],
    examNote: '시험지 분석 및 학습 방향\n\n여러 영역의 개념을 함께 확인하는 문항이 포함되어 있습니다.\n\n문제를 풀 때 조건을 끝까지 확인하고, 글에서 근거를 찾는 연습을 진행하세요.\n\n이 영역에는 시험지데이터 H27의 내용을 입력합니다.\n\n배치 확인용 예시 문장입니다.',
    students: [{row: 8, values: ['28006','국','데모학생','예시초등학교','','', '1, 2, 4, 8', 16, 12, 16, 8, 16, '', '', '', message],
      display: ['28006','국','데모학생','예시초등학교','','', '1, 2, 4, 8', '16','12','16','8','16','','','',message], styles: Array(16).fill(null)}],
    questions: Array.from({length: 25}, (_, i) => {
      const area = REPORT.AXES[i % 5], branch = ['화자와 상황','인물과 사건','내용 이해','문장과 품사','어휘와 표현'][i % 5];
      return {row: i + 2, values: [i + 1,'','',area,branch,'62.2%',i % 4 + 1],
        display: [String(i + 1),'','',area,branch,'62.2%',String(i % 4 + 1)], styles: Array(7).fill(null)};
    })
  };
  source.students[0].styles[6] = {text: '1, 2, 4, 8', color: '#D53939', runs: []};
  source.students[0].styles[15] = exampleColorStyle_(message, [['화자와 상황', '#D53939'], ['인물의 말과 행동', '#4472C4']]);
  source.questions.forEach(q => {
    q.styles[3] = {text: q.display[3], color: REPORT.BLUE, runs: []};
    q.styles[6] = {text: q.display[6], color: '#FF0000', runs: []}; // 별은 원본 색과 관계없이 검정
  });
  source.examNoteStyle = exampleColorStyle_(source.examNote, [['조건을 끝까지 확인', '#D53939'], ['근거를 찾는 연습', '#4472C4']]);
  return buildReportFromSource_(source, source.students[0], {title: '11월 모의고사 국어영역 분석표', chartMax: 20});
}


function extractDriveId_(value, required) {
  const text = clean_(value);
  if (!text && !required) return '';
  const match = text.match(/\/d\/([a-zA-Z0-9_-]+)/) || text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  const id = match ? match[1] : text;
  if (!/^[a-zA-Z0-9_-]{15,}$/.test(id)) throw new Error('설정에 올바른 Drive 파일 링크 또는 ID를 입력하세요.');
  return id;
}

/** 공유 링크의 resourcekey를 보존합니다. 기존 파일 ID 설정도 지원합니다. */
function extractDriveReference_(value, required, label) {
  try {
    const id = extractDriveId_(value, required);
    const match = clean_(value).match(/[?&]resourcekey=([^&#]*)/i);
    const resourceKey = match ? decodeURIComponent(match[1]) : '';
    if (match && !/^[a-zA-Z0-9_-]+$/.test(resourceKey)) {
      throw new Error('공유 링크의 resourcekey가 올바르지 않습니다. Drive의 공유에서 전체 링크를 다시 복사하세요.');
    }
    return {id: id, resourceKey: resourceKey};
  } catch (error) {
    throw new Error(reportFileLabel_(label) + ': ' + error.message);
  }
}

function reportFileLabel_(label) {
  const cells = {'템플릿이미지': 'B2', '로고이미지': 'B4', '결과슬라이드': 'B5'};
  return '보고서설정!' + (cells[label] || '?') + ' (' + label + ')';
}

function openDriveFile_(id, resourceKey) {
  return resourceKey ? DriveApp.getFileByIdAndResourceKey(id, resourceKey) : DriveApp.getFileById(id);
}

function reportFileAccessError_(label, id, error) {
  const role = label === '결과슬라이드' ? '편집자' : '뷰어';
  return new Error(reportFileLabel_(label) + ' 파일 접근에 실패했습니다.\n파일 ID: ' + id +
    '\n현재 버튼을 누른 계정에 이 파일의 ' + role + ' 권한이 있는지 확인하세요. ' +
    '해당 계정의 이메일을 파일 공유에 직접 추가하면 링크 접근 문제를 확인할 수 있습니다.' +
    (label === '결과슬라이드' ? '\nB5에는 폴더·PPTX가 아닌 Google Slides 파일 링크를 입력하세요.' :
      '\n공유 링크에 resourcekey가 있으면 B2/B4에 ID만 넣지 말고 공유 → 링크 복사로 가져온 전체 링크를 입력하세요.') +
    '\n원래 오류: ' + String(error.message || error).slice(0, 350));
}

/** 보고서를 생성하지 않고, 실행한 사용자의 계정으로 각 파일을 확인합니다. */
function checkReportFileAccess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = requiredSheet_(ss, REPORT.SETTINGS);
  const map = {};
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues()
      .forEach(row => { map[row[0]] = row[1]; });
  }
  const checks = ['템플릿이미지', '로고이미지', '결과슬라이드'].map(label => {
    const labelText = reportFileLabel_(label);
    try {
      const ref = extractDriveReference_(map[label], label === '템플릿이미지', label);
      if (!ref.id) return {label: label, status: '생략', message: labelText + ': ' +
        (label === '결과슬라이드' ? '비어 있음. 생성 버튼을 누르면 실행 계정의 새 슬라이드가 만들어집니다.' : '사용하지 않음')};
      if (label !== '결과슬라이드') {
        readImageBlob_(ref.id, label, ref.resourceKey);
        return {label: label, status: '확인', message: labelText + ': 이미지 읽기 확인'};
      }
      try {
        const file = openDriveFile_(ref.id, ref.resourceKey);
        if (file.getMimeType() !== 'application/vnd.google-apps.presentation') {
          throw new Error('Google Slides 파일이 아닙니다. 폴더·이미지·PPTX 링크인지 확인하세요.');
        }
        const deck = SlidesApp.openById(ref.id);
        return {label: label, status: '확인', message: labelText + ': 슬라이드 열기 확인 · ' + deck.getName()};
      } catch (error) {
        throw reportFileAccessError_(label, ref.id, error);
      }
    } catch (error) {
      return {label: label, status: '오류', message: String(error.message || error)};
    }
  });
  const ui = SpreadsheetApp.getUi();
  ui.alert('실행 계정의 파일 접근 확인',
    checks.map(check => '[' + check.status + '] ' + check.message).join('\n\n') +
    '\n\n이 검사는 파일 읽기·열기만 확인하며 보고서나 새 파일을 만들지 않습니다. ' +
    '결과 슬라이드의 편집자 권한은 공유 설정에서 별도로 확인하세요. ' +
    '실제 오류가 난 사용자 계정으로 실행해야 합니다.', ui.ButtonSet.OK);
  return checks;
}

function readImageBlob_(id, label, resourceKey) {
  let blob;
  try { blob = openDriveFile_(id, resourceKey).getBlob(); }
  catch (error) { throw reportFileAccessError_(label, id, error); }
  if (!['image/png', 'image/jpeg', 'image/gif'].includes(blob.getContentType())) {
    throw new Error(reportFileLabel_(label) + '는 PNG/JPG/GIF 파일이어야 합니다.');
  }
  // 이 템플릿 전용 좌표이므로 비율이 같은 다른 양식도 임의로 받지 않습니다.
  if (label === '템플릿이미지') {
    const bytes = blob.getBytes();
    const u = i => bytes[i] & 255;
    const uint32 = i => u(i) * 16777216 + u(i + 1) * 65536 + u(i + 2) * 256 + u(i + 3);
    if (blob.getContentType() !== 'image/png' || u(0) !== 137 || u(1) !== 80 ||
        uint32(16) !== REPORT.WIDTH || uint32(20) !== REPORT.HEIGHT) {
      throw new Error(reportFileLabel_(label) + ': 동봉된 report-template.png (3216 × 1960)를 사용하세요. 다른 템플릿은 좌표 조정이 필요합니다.');
    }
  }
  return blob;
}

function ensureOutputDeck_(ss, config) {
  if (config.deckId) return config.deckId;
  const deck = SlidesApp.create('국어 분석 보고서');
  const id = deck.getId();
  const initial = deck.getSlides()[0];
  PropertiesService.getScriptProperties().setProperty('REPORT_INITIAL_' + id, initial.getObjectId());
  deck.saveAndClose();
  const settings = ss.getSheetByName(REPORT.SETTINGS);
  const labels = settings.getRange(2, 1, settings.getLastRow() - 1, 1).getValues();
  const index = labels.findIndex(r => r[0] === '결과슬라이드');
  if (index < 0) throw new Error('설정 탭에 결과슬라이드 행이 없습니다.');
  settings.getRange(index + 2, 2).setValue('https://docs.google.com/presentation/d/' + id + '/edit');
  config.deckId = id;
  return id;
}

function clean_(value) { return value === null || value === undefined ? '' : String(value).trim(); }

function colorHex_(color, theme) {
  if (!color) return null;
  const type = String(color.getColorType());
  if (type === 'RGB') return color.asRgbColor().asHexString().toUpperCase();
  if (type === 'THEME' && theme) {
    return theme.getConcreteColor(color.asThemeColor().getThemeColorType()).asRgbColor().asHexString().toUpperCase();
  }
  return null;
}

function readCellColorStyle_(value, rich, fontColor, theme) {
  const raw = rich ? rich.getText() : clean_(value);
  const text = raw.trim();
  const offset = raw.length - raw.trimStart().length;
  const color = colorHex_(fontColor, theme) || '#000000';
  const runs = rich ? rich.getRuns().map(run => ({
    start: Math.max(0, run.getStartIndex() - offset),
    end: Math.min(text.length, run.getEndIndex() - offset),
    color: colorHex_(run.getTextStyle().getForegroundColorObject(), theme) || color
  })).filter(run => run.end > run.start) : [];
  return {text: text, color: color, runs: runs};
}

function styledText_(text, style) {
  if (!style) return text;
  return {text: String(text), color: style.color,
    runs: style.text === String(text) ? style.runs : []};
}

function characterColors_(text, baseColor, runs) {
  const colors = Array(text.length).fill(baseColor);
  (runs || []).forEach(run => {
    for (let i = Math.max(0, run.start); i < Math.min(text.length, run.end); i++) colors[i] = run.color;
  });
  return colors;
}

function compressColorRuns_(colors) {
  const runs = [];
  colors.forEach((color, i) => {
    const previous = runs[runs.length - 1];
    if (previous && previous.color === color) previous.end = i + 1;
    else runs.push({start: i, end: i + 1, color: color});
  });
  return runs;
}

function joinStyledText_(parts, separator) {
  let text = '';
  let colors = [];
  parts.forEach((part, index) => {
    if (index) { text += separator; colors = colors.concat(Array(separator.length).fill('#000000')); }
    const value = typeof part === 'object' ? part : {text: String(part), color: '#000000', runs: []};
    text += value.text;
    colors = colors.concat(characterColors_(value.text, value.color, value.runs));
  });
  return {text: text, color: '#000000', runs: compressColorRuns_(colors)};
}

function fittedColorRuns_(item, fittedText) {
  if (!item.colorRuns || !item.colorRuns.length) return [];
  const sourceColors = characterColors_(item.text, item.color, item.colorRuns);
  const colors = [];
  let sourceIndex = 0;
  let previousColor = item.color;
  for (const char of Array.from(fittedText)) {
    let color = previousColor;
    if (!/\s/.test(char)) {
      while (sourceIndex < item.text.length && /\s/.test(item.text[sourceIndex])) sourceIndex++;
      if (item.text.slice(sourceIndex, sourceIndex + char.length) !== char) {
        throw new Error(item.label + ' 색상 구간을 연결하지 못했습니다.');
      }
      color = sourceColors[sourceIndex];
      sourceIndex += char.length;
      previousColor = color;
    }
    for (let i = 0; i < char.length; i++) colors.push(color);
  }
  return compressColorRuns_(colors).filter(run => run.color !== item.color);
}

function pageTransform_(width, height) {
  const scale = Math.min(width / REPORT.WIDTH, height / REPORT.HEIGHT);
  return {scale: scale, x: (width - REPORT.WIDTH * scale) / 2, y: (height - REPORT.HEIGHT * scale) / 2};
}

function text_(items, label, text, x, y, w, h, size, options) {
  const source = text && typeof text === 'object' ? text : null;
  const value = source ? source.text : (text === null || text === undefined ? '' : String(text));
  if (!clean_(value)) return;
  items.push(Object.assign({kind: 'text', label: label, text: value, x: x, y: y, w: w, h: h,
    size: size, minSize: Math.max(22, size - 8), align: 'CENTER', vertical: 'MIDDLE',
    color: '#111111', bold: false}, options || {}, source ? {color: source.color, colorRuns: source.runs} : {}));
}

function line_(items, x1, y1, x2, y2, color, width) {
  if (Math.hypot(x2 - x1, y2 - y1) < 0.001) return; // 원점의 0값 연속 구간
  items.push({kind: 'line', x1: x1, y1: y1, x2: x2, y2: y2, color: color, width: width});
}

function fitText_(item, scale) {
  const availableW = item.w - 5;
  const availableH = item.h - 4;
  // 과정은 입력한 단원별 줄을 우선 보존합니다. 그래도 길면 어절 기준으로 줄바꿈합니다.
  if (item.keepLines) {
    const lines = item.text.split(/\r?\n/);
    for (let size = item.size; size >= item.minSize; size--) {
      if (lines.every(line => textUnits_(line) * size <= availableW) &&
          lines.length * size * 1.22 <= availableH) return {text: lines.join('\n'), size: size};
    }
  }
  for (let size = item.size; size >= item.minSize; size--) {
    const lines = wrapText_(item.text, availableW / size);
    if (availableW > 0 && lines.every(line => textUnits_(line) * size <= availableW + 0.01) &&
        lines.length * size * 1.22 <= availableH) return {text: lines.join('\n'), size: size};
  }
  throw new Error(item.label + ' 내용이 칸보다 깁니다. 해당 입력값을 줄여 주세요.');
}

function textUnits_(text) {
  return Array.from(text).reduce((sum, char) => sum +
    (/\s/.test(char) ? 0.32 : /[.,:;!|'\-]/.test(char) ? 0.35 : /[\x00-\x7F]/.test(char) ? 0.58 : 1.03), 0);
}

function wrapText_(text, maxUnits) {
  const lines = [];
  String(text).split(/\r?\n/).forEach(paragraph => {
    let current = '';
    (paragraph.match(/\S+|\s+/g) || []).forEach(token => {
      if (!current && /^\s+$/.test(token)) return;
      if (textUnits_(current + token) <= maxUnits) { current += token; return; }
      if (current.trim()) { lines.push(current.trimEnd()); current = ''; }
      if (/^\s+$/.test(token)) return;
      Array.from(token).forEach(char => {
        if (current && textUnits_(current + char) > maxUnits) { lines.push(current); current = ''; }
        current += char;
      });
    });
    lines.push(current.trimEnd());
  });
  return lines;
}

function rgb_(hex) {
  return {red: parseInt(hex.slice(1, 3), 16) / 255, green: parseInt(hex.slice(3, 5), 16) / 255,
    blue: parseInt(hex.slice(5, 7), 16) / 255};
}

function modelToRequests_(items, slideId, transform, prefix) {
  const requests = [];
  const s = transform.scale;
  items.forEach((item, index) => {
    const objectId = prefix + '_' + index;
    if (item.kind === 'text') {
      const fit = fitText_(item, s);
      requests.push({createShape: {objectId: objectId, shapeType: 'TEXT_BOX', elementProperties: {
        pageObjectId: slideId,
        size: {width: {magnitude: item.w * s + 7.2, unit: 'PT'}, height: {magnitude: item.h * s + 7.2, unit: 'PT'}},
        transform: {scaleX: 1, scaleY: 1, translateX: transform.x + item.x * s - 3.6, translateY: transform.y + item.y * s - 3.6, unit: 'PT'}
      }}}, {insertText: {objectId: objectId, text: fit.text, insertionIndex: 0}},
      {updateShapeProperties: {objectId: objectId, shapeProperties: {
        contentAlignment: item.vertical, shapeBackgroundFill: {propertyState: 'NOT_RENDERED'},
        outline: {propertyState: 'NOT_RENDERED'}
      }, fields: 'contentAlignment,shapeBackgroundFill,outline'}},
      {updateTextStyle: {objectId: objectId, textRange: {type: 'ALL'}, style: {
        fontFamily: 'Noto Sans KR', fontSize: {magnitude: fit.size * s, unit: 'PT'}, bold: item.bold,
        foregroundColor: {opaqueColor: {rgbColor: rgb_(item.color)}}
      }, fields: 'fontFamily,fontSize,bold,foregroundColor'}},
      {updateParagraphStyle: {objectId: objectId, textRange: {type: 'ALL'}, style: {
        alignment: item.align, lineSpacing: 100, spaceAbove: {magnitude: 0, unit: 'PT'},
        spaceBelow: {magnitude: 0, unit: 'PT'}
      }, fields: 'alignment,lineSpacing,spaceAbove,spaceBelow'}});
      fittedColorRuns_(item, fit.text).forEach(run => requests.push({updateTextStyle: {objectId: objectId,
        textRange: {type: 'FIXED_RANGE', startIndex: run.start, endIndex: run.end},
        style: {foregroundColor: {opaqueColor: {rgbColor: rgb_(run.color)}}}, fields: 'foregroundColor'}}));
    } else if (item.kind === 'triangle' || item.kind === 'rect') {
      const props = nativeFilledShape_(item, slideId, transform);
      requests.push({createShape: {objectId: objectId, shapeType: item.kind === 'triangle' ? 'TRIANGLE' : 'RECTANGLE', elementProperties: props.elementProperties}},
        {updateShapeProperties: {objectId: objectId, shapeProperties: props.shapeProperties, fields: 'shapeBackgroundFill,outline'}});
    } else {
      const dx = item.x2 - item.x1, dy = item.y2 - item.y1;
      requests.push({createLine: {objectId: objectId, category: 'STRAIGHT', elementProperties: {
        pageObjectId: slideId,
        size: {width: {magnitude: 1, unit: 'PT'}, height: {magnitude: 1, unit: 'PT'}},
        transform: {scaleX: (dx + dy) * s / 2, scaleY: (dy + dx) * s / 2,
          shearX: (dx - dy) * s / 2, shearY: (dy - dx) * s / 2,
          translateX: transform.x + item.x1 * s, translateY: transform.y + item.y1 * s, unit: 'PT'}
      }}}, {updateLineProperties: {objectId: objectId, lineProperties: {
        lineFill: {solidFill: {color: {rgbColor: rgb_(item.color)}, alpha: 1}},
        weight: {magnitude: item.width * s, unit: 'PT'}, startArrow: 'NONE', endArrow: 'NONE'
      }, fields: 'lineFill,weight,startArrow,endArrow'}});
    }
  });
  return requests;
}

function renderReport_(ss, deckId, report, config, template, logo) {
  const properties = PropertiesService.getScriptProperties();
  const mappingKey = 'REPORT_SLIDE_' + deckId + '_' + report.id;
  const oldId = properties.getProperty(mappingKey);
  const deck = SlidesApp.openById(deckId);
  const transform = pageTransform_(deck.getPageWidth(), deck.getPageHeight());
  const items = buildReportModel_(report, config, !!logo);
  // 원격 변경 전에 모든 본문 길이 검사.
  items.filter(item => item.kind === 'text').forEach(item => fitText_(item, transform.scale));
  const knownIds = deck.getSlides().map(slide => slide.getObjectId());
  const slide = deck.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  const slideId = slide.getObjectId();
  let committed = false;
  try {
    slide.getBackground().setSolidFill('#FFFFFF');
    slide.insertImage(template, transform.x, transform.y, REPORT.WIDTH * transform.scale, REPORT.HEIGHT * transform.scale)
      .setTitle('보고서 원본 템플릿');
    if (logo) {
      const image = slide.insertImage(logo);
      const factor = Math.min(155 * transform.scale / image.getWidth(), 130 * transform.scale / image.getHeight());
      const w = image.getWidth() * factor, h = image.getHeight() * factor;
      image.setWidth(w).setHeight(h).setLeft(transform.x + 92 * transform.scale + (155 * transform.scale - w) / 2)
        .setTop(transform.y + 87 * transform.scale + (130 * transform.scale - h) / 2);
    }
    slide.getNotesPage().getSpeakerNotesShape().getText().setText(
      'ACADEMY_REPORT_V1\n보고서ID: ' + report.id + '\n수험번호: ' + report.examNumber + '\n원본시트: ' + ss.getUrl());
    deck.saveAndClose(); // SlidesApp 변경을 고급 API 호출 전에 확정.
    const prefix = 'r_' + Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    const requests = modelToRequests_(items, slideId, transform, prefix);
    if (oldId && knownIds.includes(oldId)) {
      requests.push({updateSlidesPosition: {slideObjectIds: [slideId], insertionIndex: knownIds.indexOf(oldId)}});
      requests.push({deleteObject: {objectId: oldId}});
    }
    const initialKey = 'REPORT_INITIAL_' + deckId;
    const initialId = properties.getProperty(initialKey);
    if (initialId && knownIds.includes(initialId)) requests.push({deleteObject: {objectId: initialId}});
    Slides.Presentations.batchUpdate({requests: requests}, deckId);
    committed = true;
    properties.setProperty(mappingKey, slideId);
    properties.deleteProperty(initialKey);
    return 'https://docs.google.com/presentation/d/' + deckId + '/edit#slide=id.' + slideId;
  } catch (error) {
    // 실패한 임시 페이지만 제거합니다. 기존 보고서는 성공한 batchUpdate에서만 교체됩니다.
    if (!committed) {
      try { Slides.Presentations.batchUpdate({requests: [{deleteObject: {objectId: slideId}}]}, deckId); }
      catch (cleanupError) { console.warn('임시 슬라이드 정리 실패: ' + cleanupError.message); }
    }
    throw error;
  }
}

function exampleColorStyle_(text, highlights) {
  const runs = [];
  highlights.forEach(([phrase, color]) => {
    let index = text.indexOf(phrase);
    while (index >= 0) {
      runs.push({start: index, end: index + phrase.length, color: color});
      index = text.indexOf(phrase, index + phrase.length);
    }
  });
  return {text: text, color: '#000000', runs: runs};
}

function nativeFilledShape_(item, slideId, transform) {
  const s = transform.scale;
  if (item.kind === 'rect') return {
    elementProperties: {pageObjectId: slideId, size: {width: {magnitude: item.w * s, unit: 'PT'}, height: {magnitude: item.h * s, unit: 'PT'}},
      transform: {scaleX: 1, scaleY: 1, translateX: transform.x + item.x * s, translateY: transform.y + item.y * s, unit: 'PT'}},
    shapeProperties: {shapeBackgroundFill: {propertyState: 'NOT_RENDERED'}, outline: {propertyState: 'RENDERED',
      outlineFill: {solidFill: {color: {rgbColor: rgb_(item.color)}, alpha: 1}}, weight: {magnitude: item.width * s, unit: 'PT'}}}
  };
  // 기본 TRIANGLE의 꼭짓점 (0.5,0), (0,1), (1,1)을 지정한 세 점으로 변환합니다.
  const [a,b,c] = item.points;
  const scaleX = c.x - b.x, shearY = c.y - b.y;
  const tx = a.x - scaleX / 2, ty = a.y - shearY / 2;
  return {
    elementProperties: {pageObjectId: slideId, size: {width: {magnitude: 1, unit: 'PT'}, height: {magnitude: 1, unit: 'PT'}},
      transform: {scaleX: scaleX * s, shearY: shearY * s, shearX: (b.x - tx) * s, scaleY: (b.y - ty) * s,
        translateX: transform.x + tx * s, translateY: transform.y + ty * s, unit: 'PT'}},
    shapeProperties: {shapeBackgroundFill: {solidFill: {color: {rgbColor: rgb_(item.color)}, alpha: item.alpha}},
      outline: {propertyState: 'NOT_RENDERED'}}
  };
}

const REPORT_BUTTON_PNG = "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";

const REPORT_DIALOG_HTML = "<!doctype html>\n<html lang=\"ko\">\n<head>\n  <meta charset=\"utf-8\">\n  <meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n  <base target=\"_top\">\n  <style>\n    :root{font-family:'Noto Sans KR','Malgun Gothic',Arial,sans-serif;color:#22334e;background:#fff;font-synthesis:none}\n    *{box-sizing:border-box}body{margin:0}button,a,select{-webkit-tap-highlight-color:transparent}button,select{font:inherit}[hidden]{display:none!important}\n    .shell{width:100%;min-height:560px;padding:20px 24px 18px;display:flex;flex-direction:column}\n    .brand{display:flex;gap:9px;align-items:center;font-size:14px;font-weight:750;color:#294b82;letter-spacing:-.3px}\n    .brand-mark{width:28px;height:28px;display:grid;place-items:center;border-radius:9px;background:#edf3fe;color:#4472c4}.brand-mark svg{width:18px;height:18px}\n    .view{flex:1;display:flex;align-items:center;flex-direction:column;text-align:center;padding-top:18px;animation:appear .2s ease-out}\n    .visual{height:94px;width:94px;display:grid;place-items:center;position:relative;margin-bottom:16px}.spinner{position:absolute;inset:0;border:4px solid #eaf0fa;border-top-color:#4472c4;border-right-color:#8ba9df;border-radius:50%;animation:spin 1.05s linear infinite}\n    .document-icon{width:64px;height:64px;border-radius:21px;background:#f0f5ff;display:grid;place-items:center;color:#4472c4}.document-icon svg{width:28px;height:28px}\n    .status-icon{width:80px;height:80px;border-radius:50%;display:grid;place-items:center;animation:pop .28s ease-out}.status-icon svg{height:36px;width:36px}.success-icon{background:#e9f7ef;color:#199565}.error-icon{background:#fff0f0;color:#d65361}.choice-icon{background:#edf3ff;color:#4472c4}\n    h1{font-size:23px;font-weight:750;line-height:1.4;letter-spacing:-1px;margin:0 0 10px}p{font-size:14px;line-height:1.65;color:#6b788f;margin:0}.name-chip{display:inline-block;color:#3f62a0;background:#eef3fc;border-radius:20px;font-size:12px;font-weight:650;padding:5px 12px;margin-bottom:10px}\n    .loading-copy{min-height:48px}.activity{height:4px;width:190px;background:#edf2fa;border-radius:9px;overflow:hidden;margin:27px 0 13px}.activity span{display:block;width:55%;height:100%;background:linear-gradient(90deg,#a8c1ed,#4472c4);border-radius:9px;animation:travel 1.8s ease-in-out infinite}\n    .subtle{font-size:12px;color:#8a96aa}.details{width:100%;background:#f6f8fc;border:1px solid #e9edf4;border-radius:14px;padding:12px 16px;text-align:left;margin:16px 0}.student{font-size:16px;font-weight:700;color:#2c405f;margin-bottom:6px}.student span{font-size:12px;color:#7d8ba2;font-weight:400;margin-left:10px}.report-title{font-size:13px;color:#728099;line-height:1.7;word-break:keep-all}\n    .actions{display:flex;gap:10px;width:100%;margin-top:auto;padding-top:8px}button,.primary{display:flex;align-items:center;justify-content:center;gap:8px;border:0;border-radius:11px;padding:12px 16px;font-size:14px;font-weight:650;text-decoration:none;cursor:pointer;min-height:44px;transition:background .15s,transform .15s}button:active,.primary:active{transform:translateY(1px)}.primary{background:#365d9b;color:white;flex:1}.primary:hover{background:#294b82}.primary svg{height:16px;width:16px}.secondary{background:#f0f3f8;color:#62718a;min-width:84px}.secondary:hover{background:#e6ebf3}.outline{background:white;border:1px solid #dbe3ef;color:#476487;flex:1}button:disabled{opacity:.55;cursor:wait}\n    .report-school{font-size:12px;color:#7d8ba2;margin-bottom:6px}.error-details{font-size:13px;line-height:1.75;max-height:90px;overflow:auto;text-align:left;background:#fff7f7;border:1px solid #f8e4e7;border-radius:12px;width:100%;padding:12px 14px;color:#9f5160;white-space:pre-wrap;overflow-wrap:anywhere;margin:16px 0}\n    .field{width:100%;text-align:left;margin:22px 0}.field label{font-size:12px;color:#7c8ba2;display:block;margin-bottom:8px}.field select{width:100%;background:#f7f9fc;color:#2c405f;border:1px solid #dbe3ef;border-radius:10px;padding:13px 12px;min-height:47px}\n    .footer{color:#98a2b2;text-align:center;font-size:11px;padding-top:14px;line-height:1.6}.view .actions{flex:none}.spacer{flex:1;min-height:12px}\n    :focus-visible{outline:3px solid #a5c5fa;outline-offset:3px}\n    @keyframes spin{to{transform:rotate(360deg)}}@keyframes travel{0%{transform:translateX(-110%)}100%{transform:translateX(285%)}}@keyframes appear{from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:translateY(0)}}@keyframes pop{from{transform:scale(.88);opacity:0}to{transform:scale(1);opacity:1}}\n    @media(prefers-reduced-motion:reduce){*,*:before,*:after{animation:none!important;transition:none!important}}\n  </style>\n</head>\n<body>\n<main class=\"shell\">\n  <header class=\"brand\"><span class=\"brand-mark\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"1.7\"><path d=\"M3 5.5c4-1 6 .2 9 2.5 3-2.3 5-3.5 9-2.5V19c-4-1-6 .2-9 2-3-1.8-5-3-9-2z\"/><path d=\"M12 8v13\"/></svg></span>국어 날개달기</header>\n  <section class=\"view\" data-view=\"loading\" aria-live=\"polite\">\n    <div class=\"visual\" aria-hidden=\"true\"><span class=\"spinner\"></span><span class=\"document-icon\"><svg viewBox=\"0 0 32 32\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"1.8\"><path d=\"M8 3h11l6 6v20H8z\"/><path d=\"M19 3v7h6M12 15h9M12 20h9M12 25h5\"/></svg></span></div>\n    <div id=\"loading-name\" class=\"name-chip\"></div>\n    <h1>보고서를 만들고 있어요</h1>\n    <p class=\"loading-copy\">입력된 내용을 확인하고<br>학생의 분석표를 작성하고 있습니다.</p>\n    <div class=\"activity\" aria-hidden=\"true\"><span></span></div>\n    <p class=\"subtle\">완성되면 자동으로 안내해 드릴게요.</p>\n  </section>\n  <section class=\"view\" data-view=\"complete\" hidden aria-live=\"polite\">\n    <div class=\"visual\" aria-hidden=\"true\"><span class=\"status-icon success-icon\"><svg viewBox=\"0 0 48 48\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"3.2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"m12 25 8 8 17-19\"/></svg></span></div>\n    <h1>보고서가 완성됐어요</h1><p>아래 버튼으로 분석표를 확인하세요.</p>\n    <div class=\"details\"><div class=\"student\"><b id=\"done-name\"></b><span id=\"done-number\"></span></div><div id=\"done-school\" class=\"report-school\"></div><div id=\"done-title\" class=\"report-title\"></div></div>\n    <div class=\"spacer\"></div>\n    <div class=\"actions\"><a id=\"open-report\" class=\"primary\" target=\"_blank\" rel=\"noopener noreferrer\">보고서 열기<svg viewBox=\"0 0 20 20\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"1.7\"><path d=\"M5 15 15 5M5 5h10v10\"/></svg></a><button class=\"secondary close-button\" type=\"button\">닫기</button></div>\n  </section>\n  <section class=\"view\" data-view=\"error\" hidden aria-live=\"assertive\">\n    <div class=\"visual\" aria-hidden=\"true\"><span class=\"status-icon error-icon\"><svg viewBox=\"0 0 48 48\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"3\" stroke-linecap=\"round\"><path d=\"M24 12v15M24 35h.01\"/></svg></span></div>\n    <h1>내용을 확인해 주세요</h1><p>보고서를 만드는 중 문제가 발생했어요.</p>\n    <div id=\"error-message\" class=\"error-details\" role=\"alert\"></div>\n    <div class=\"spacer\"></div><div class=\"actions\"><button id=\"retry\" class=\"outline\" type=\"button\">다시 시도</button><button class=\"primary close-button\" type=\"button\">닫기</button></div>\n  </section>\n  <section class=\"view\" data-view=\"choose\" hidden aria-live=\"polite\">\n    <div class=\"visual\" aria-hidden=\"true\"><span class=\"status-icon choice-icon\"><svg viewBox=\"0 0 48 48\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.5\" stroke-linecap=\"round\"><circle cx=\"20\" cy=\"17\" r=\"7\"/><path d=\"M7 38v-3c0-8 26-8 26 0v3M34 13a6 6 0 0 1 0 12M39 32c4 1 5 3 5 6\"/></svg></span></div>\n    <h1>같은 이름의 학생이 있어요</h1><p>생성할 학생의 수험번호를 선택해 주세요.</p>\n    <div class=\"field\"><label for=\"student-choice\">수험번호 · 학교</label><select id=\"student-choice\"></select></div>\n    <div class=\"spacer\"></div><div class=\"actions\"><button id=\"choose-student\" class=\"primary\" type=\"button\">선택한 학생으로 생성</button><button class=\"secondary close-button\" type=\"button\">닫기</button></div>\n  </section>\n  <footer class=\"footer\">생성 결과는 ‘보고서생성기록’에서도 확인할 수 있습니다.</footer>\n</main>\n<script>\n  var initial = __REPORT_INITIAL_DATA__;\n  var busy = false;\n  var chosenNumber = '';\n  function showView(name) {\n    document.querySelectorAll('[data-view]').forEach(function(view){view.hidden=view.dataset.view!==name;});\n  }\n  function showError(message) {\n    busy=false;\n    document.getElementById('error-message').textContent=String(message||'잠시 후 다시 시도해 주세요.');\n    showView('error');\n  }\n  function handleResult(result) {\n    busy=false;\n    if(!result||typeof result.state!=='string'){showError('결과를 확인하지 못했습니다. 생성 기록을 확인한 뒤 다시 시도해 주세요.');return;}\n    if(result.state==='choose'){\n      var select=document.getElementById('student-choice');select.textContent='';\n      (result.candidates||[]).forEach(function(student){\n        var option=document.createElement('option');option.value=student.examNumber;\n        option.textContent=student.examNumber+(student.school?' · '+student.school:'');select.appendChild(option);\n      });\n      if(!select.options.length){showError('선택할 학생을 찾지 못했습니다. S7의 이름을 확인해 주세요.');return;}\n      showView('choose');select.focus();return;\n    }\n    if(result.state==='complete'){\n      if(!/^https:\\/\\/docs\\.google\\.com\\/presentation\\/d\\//.test(String(result.url||''))){showError('보고서 링크를 확인하지 못했습니다. 생성 기록을 확인해 주세요.');return;}\n      document.getElementById('done-name').textContent=result.name+' 학생';\n      document.getElementById('done-number').textContent='수험번호 '+result.examNumber;\n      document.getElementById('done-school').textContent=result.schoolName?'학교 '+result.schoolName:'';\n      document.getElementById('done-school').hidden=!result.schoolName;\n      document.getElementById('done-title').textContent=result.title;\n      document.getElementById('open-report').href=result.url;\n      showView('complete');document.getElementById('open-report').focus();return;\n    }\n    showError(result.message);\n  }\n  function requestReport(number) {\n    if(busy)return;busy=true;chosenNumber=number||'';\n    document.getElementById('loading-name').textContent=initial.requestedName?initial.requestedName+' 학생':'보고서 준비';\n    showView('loading');\n    if(!window.google||!google.script||!google.script.run){showError('구글 시트의 보고서 생성 버튼으로 실행해 주세요.');return;}\n    google.script.run.withSuccessHandler(handleResult).withFailureHandler(function(error){\n      showError(error&&error.message?error.message:'연결을 확인하지 못했습니다. 생성 기록을 확인한 뒤 다시 시도해 주세요.');\n    }).createReportForDialog({requestedName:initial.requestedName||'',examNumber:chosenNumber});\n  }\n  document.getElementById('retry').addEventListener('click',function(){requestReport(chosenNumber);});\n  document.getElementById('choose-student').addEventListener('click',function(){requestReport(document.getElementById('student-choice').value);});\n  document.querySelectorAll('.close-button').forEach(function(button){button.addEventListener('click',function(){if(window.google&&google.script&&google.script.host)google.script.host.close();});});\n  requestReport('');\n</script>\n</body>\n</html>\n";
