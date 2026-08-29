var MA_CFG = {
  VERSION: 'MA-V1.0.0',
  SHEETS: {
    CONTENT: '콘텐츠DB',
    CALENDAR: '콘텐츠캘린더',
    BLOG: '블로그초안',
    THREADS: 'Threads초안',
    PUBLISH: '발행관리',
    SETTINGS: '설정',
    LOG: '자동화로그'
  },
  HEADERS: {
    CONTENT: [
      'Content ID','등록일','Pillar','콘텐츠유형','주제','핵심키워드','보조키워드','대상독자',
      '검색의도','콘텐츠목적','CTA','참고자료','사실확인메모','콘텐츠상태','우선순위',
      '목표발행일','블로그ID','Threads 생성수','관리자메모','최종수정일'
    ],
    BLOG: [
      'Blog ID','Content ID','생성일','제목1','제목2','제목3','최종제목','핵심키워드',
      '도입부','본문','CTA','해시태그','이미지요청','사실검수','문체검수','SEO검수',
      '관리자수정본','최종본문','검수상태','승인일','발행상태','발행일','발행URL'
    ],
    THREADS: [
      'Thread ID','Content ID','Blog ID','생성일','Thread 유형','Hook','본문','CTA',
      '최종본문','검수상태','승인일','발행상태','발행예정일','발행일','발행URL','관리자메모'
    ],
    PUBLISH: [
      'Publish ID','Content ID','콘텐츠ID','플랫폼','최종제목','최종본문 존재','검수상태',
      '승인여부','발행예정일','발행예정시간','발행상태','발행일','발행URL',
      '오류여부','오류메시지','재시도횟수'
    ],
    LOG: ['실행시각','Run ID','기능','Content ID','대상ID','실행결과','수준','메시지','처리시간(ms)','실행주체']
  }
};

function ma_getSettings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MA_CFG.SHEETS.SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(3, sheet.getLastColumn())).getValues();
  var out = {};
  values.forEach(function(r) {
    var key = String(r[0] || '').trim();
    if (key) out[key] = r[1];
  });
  return out;
}

function ma_headerMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error('헤더가 없습니다: ' + sheet.getName());
  var row = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var map = {};
  row.forEach(function(v, i) {
    var key = String(v || '').trim();
    if (key) map[key] = i + 1;
  });
  return map;
}

function ma_valueByHeader_(row, headerMap, header) {
  var col = headerMap[header];
  return col ? row[col - 1] : '';
}

function ma_colLetter_(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function ma_escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

