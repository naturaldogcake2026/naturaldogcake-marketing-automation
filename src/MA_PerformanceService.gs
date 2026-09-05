var MA_V15_PERFORMANCE_SHEET = '성과관리';
var MA_V15_PERFORMANCE_HEADERS = [
  'Metric ID',
  '기록일시',
  'Content ID',
  '콘텐츠ID',
  '플랫폼',
  '발행일',
  '발행URL',
  '측정기준일',
  '발행후경과일',
  '조회수',
  '좋아요·공감',
  '댓글·답글',
  '공유·리포스트',
  '저장',
  '링크·프로필클릭',
  '문의·전환',
  '데이터출처',
  '비고',
  '등록방식',
  '최종수정일'
];

function ma_installV15PerformanceSheet() {
  var started = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MA_V15_PERFORMANCE_SHEET);
  var created = false;

  if (!sheet) {
    sheet = ss.insertSheet(MA_V15_PERFORMANCE_SHEET);
    created = true;
  }

  if (sheet.getMaxColumns() < MA_V15_PERFORMANCE_HEADERS.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      MA_V15_PERFORMANCE_HEADERS.length - sheet.getMaxColumns()
    );
  }

  var headerRange = sheet.getRange(1, 1, 1, MA_V15_PERFORMANCE_HEADERS.length);
  var current = headerRange.getValues()[0];
  var hasAnyHeader = current.some(function(v) { return String(v || '').trim() !== ''; });

  if (hasAnyHeader) {
    for (var i = 0; i < MA_V15_PERFORMANCE_HEADERS.length; i++) {
      var actual = String(current[i] || '').trim();
      var expected = MA_V15_PERFORMANCE_HEADERS[i];
      if (actual && actual !== expected) {
        throw new Error(
          '성과관리 시트 헤더가 예상과 다릅니다. ' +
          ma_colLetterV15_(i + 1) + '1: [' + actual + '] / 예상 [' + expected + ']'
        );
      }
    }
  }

  headerRange.setValues([MA_V15_PERFORMANCE_HEADERS]);
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange('A:T').setVerticalAlignment('middle');
  sheet.getRange('B:B').setNumberFormat('yyyy. m. d hh:mm:ss');
  sheet.getRange('F:F').setNumberFormat('yyyy. m. d');
  sheet.getRange('H:H').setNumberFormat('yyyy. m. d');
  sheet.getRange('T:T').setNumberFormat('yyyy. m. d hh:mm:ss');
  sheet.getRange('I:P').setNumberFormat('0');

  var widths = [150,150,145,155,85,105,260,110,105,85,95,95,110,80,115,95,120,250,95,150];
  widths.forEach(function(width, idx) {
    sheet.setColumnWidth(idx + 1, width);
  });

  // 직접 입력 영역: 측정기준일 및 성과수치/출처/비고
  sheet.getRange(2, 8, Math.max(sheet.getMaxRows() - 1, 1), 11)
    .setBackground('#fff2cc');

  // 플랫폼/데이터출처 유효성은 자유 입력을 막지 않고 대표 값만 안내한다.
  var platformRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Blog', 'Threads'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 5, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(platformRule);

  ma_log_(
    'PERFORMANCE_V15_INSTALL',
    '',
    '',
    'SUCCESS',
    'INFO',
    'V1.5 성과관리 시트 적용: ' + (created ? '신규 생성' : '기존 시트 보존'),
    Date.now() - started,
    'menu'
  );

  ss.toast('V1.5 성과관리 시트가 준비되었습니다.', '성과관리', 5);
  return 'V1.5 성과관리 시트 적용 완료';
}

function ma_addSelectedPerformanceSnapshot() {
  var started = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();

  if (activeSheet.getName() !== MA_CFG.SHEETS.PUBLISH) {
    throw new Error('발행관리 시트에서 발행완료 행을 선택한 뒤 실행해 주세요.');
  }

  var row = activeSheet.getActiveRange().getRow();
  if (row < 2) throw new Error('헤더가 아닌 발행관리 데이터 행을 선택해 주세요.');

  var values = activeSheet.getRange(row, 1, 1, 16).getValues()[0];
  var publishId = ma_text_(values[0]);
  var contentId = ma_text_(values[1]);
  var assetId = ma_text_(values[2]);
  var platform = ma_text_(values[3]);
  var publishStatus = ma_text_(values[10]);
  var publishDate = values[11];
  var publishUrl = ma_text_(values[12]);

  if (publishStatus !== '발행완료') {
    throw new Error('성과 기록은 발행상태가 발행완료인 콘텐츠만 가능합니다.');
  }
  if (!publishUrl) {
    throw new Error('발행URL이 없습니다. 먼저 발행URL을 기록해 주세요.');
  }

  var sheet = ss.getSheetByName(MA_V15_PERFORMANCE_SHEET);
  if (!sheet) {
    ma_installV15PerformanceSheet();
    sheet = ss.getSheetByName(MA_V15_PERFORMANCE_SHEET);
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var published = ma_dateOnlyV15_(publishDate);
  var elapsedDays = published ? Math.max(0, Math.floor((today.getTime() - published.getTime()) / 86400000)) : '';

  if (ma_hasPerformanceSnapshotForDateV15_(sheet, assetId, today)) {
    throw new Error('이 콘텐츠는 오늘 날짜의 성과 기록행이 이미 있습니다. 기존 행을 입력하거나 다음 측정일에 새 행을 추가해 주세요.');
  }

  var metricId = ma_nextMetricIdV15_(sheet, today);
  var now = new Date();
  var nextRow = Math.max(sheet.getLastRow() + 1, 2);
  var rowValues = [
    metricId,
    now,
    contentId,
    assetId,
    platform,
    published || '',
    publishUrl,
    today,
    elapsedDays,
    '', '', '', '', '', '', '',
    '수동확인',
    '',
    'menu',
    now
  ];

  sheet.getRange(nextRow, 1, 1, rowValues.length).setValues([rowValues]);
  sheet.getRange(nextRow, 8, 1, 11).setBackground('#fff2cc');

  ma_log_(
    'PERFORMANCE_SNAPSHOT_ADD',
    contentId,
    assetId,
    'SUCCESS',
    'INFO',
    '성과 기록행 추가: ' + metricId + ', source=' + publishId,
    Date.now() - started,
    'menu'
  );

  ss.setActiveSheet(sheet);
  sheet.setActiveRange(sheet.getRange(nextRow, 10));
  ss.toast('성과 기록행을 추가했습니다. 노란색 성과 수치를 입력해 주세요.', '성과관리', 5);
  return metricId;
}

function ma_hasPerformanceSnapshotForDateV15_(sheet, assetId, targetDate) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var values = sheet.getRange(2, 4, lastRow - 1, 5).getValues(); // D:H
  var targetKey = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  return values.some(function(row) {
    if (ma_text_(row[0]) !== assetId) return false;
    var d = ma_dateOnlyV15_(row[4]);
    if (!d) return false;
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') === targetKey;
  });
}

function ma_nextMetricIdV15_(sheet, date) {
  var year = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy');
  var prefix = 'MET-' + year + '-';
  var maxNo = 0;
  var lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(function(row) {
      var id = ma_text_(row[0]);
      if (id.indexOf(prefix) !== 0) return;
      var n = Number(id.substring(prefix.length));
      if (isFinite(n) && n > maxNo) maxNo = n;
    });
  }

  return prefix + Utilities.formatString('%06d', maxNo + 1);
}

function ma_dateOnlyV15_(value) {
  if (!value) return null;
  var d = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function ma_colLetterV15_(col) {
  var s = '';
  while (col > 0) {
    var m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

// V1.5.1: 발행 후 D+1 / D+3 / D+7 성과 측정대상 관리
var MA_V151_MEASUREMENT_DAYS = [1, 3, 7];

function ma_showTodayPerformanceTargets() {
  var targets = ma_getTodayPerformanceTargetsV151_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!targets.length) {
    ss.toast('오늘(D+1/D+3/D+7) 측정할 발행완료 콘텐츠가 없습니다.', '성과관리', 6);
    return '오늘 성과 측정대상 0건';
  }

  var lines = targets.map(function(t) {
    return t.assetId + ' (' + t.platform + ', D+' + t.elapsedDays + ')' +
      (t.hasSnapshot ? ' - 오늘 기록행 있음' : ' - 기록 필요');
  });

  SpreadsheetApp.getUi().alert(
    '오늘 성과 측정대상 ' + targets.length + '건',
    lines.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return '오늘 성과 측정대상 ' + targets.length + '건';
}

function ma_createTodayPerformanceSnapshots() {
  var started = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var perf = ss.getSheetByName(MA_V15_PERFORMANCE_SHEET);
  if (!perf) {
    ma_installV15PerformanceSheet();
    perf = ss.getSheetByName(MA_V15_PERFORMANCE_SHEET);
  }

  var targets = ma_getTodayPerformanceTargetsV151_();
  var pending = targets.filter(function(t) { return !t.hasSnapshot; });

  if (!pending.length) {
    ss.toast(
      targets.length ? '오늘 측정대상은 모두 기록행이 이미 있습니다.' : '오늘 측정할 콘텐츠가 없습니다.',
      '성과관리',
      6
    );
    return '신규 성과 기록행 0건';
  }

  var now = new Date();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var rows = [];

  pending.forEach(function(t) {
    rows.push([
      ma_nextMetricIdV15_(perf, today),
      now,
      t.contentId,
      t.assetId,
      t.platform,
      t.publishDate,
      t.publishUrl,
      today,
      t.elapsedDays,
      '', '', '', '', '', '', '',
      '수동확인',
      'V1.5.1 D+' + t.elapsedDays + ' 정기측정',
      'menu-auto-target',
      now
    ]);
    // 다음 ID 계산 시 현재 rows가 아직 시트에 없으므로 충돌 방지를 위해 임시 ID를 후처리한다.
  });

  // 배치 안에서 ID가 중복되지 않도록 시작 번호를 다시 계산한다.
  var year = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy');
  var prefix = 'MET-' + year + '-';
  var maxNo = 0;
  var lastRow = perf.getLastRow();
  if (lastRow >= 2) {
    perf.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function(r) {
      var id = ma_text_(r[0]);
      if (id.indexOf(prefix) !== 0) return;
      var n = Number(id.substring(prefix.length));
      if (isFinite(n) && n > maxNo) maxNo = n;
    });
  }
  rows.forEach(function(r, idx) {
    r[0] = prefix + Utilities.formatString('%06d', maxNo + idx + 1);
  });

  var startRow = Math.max(perf.getLastRow() + 1, 2);
  perf.getRange(startRow, 1, rows.length, MA_V15_PERFORMANCE_HEADERS.length).setValues(rows);
  perf.getRange(startRow, 8, rows.length, 11).setBackground('#fff2cc');

  ma_log_(
    'PERFORMANCE_TARGETS_V151_CREATE',
    '',
    '',
    'SUCCESS',
    'INFO',
    'V1.5.1 오늘 성과 기록행 생성: 대상=' + targets.length + '건, 신규=' + rows.length + '건',
    Date.now() - started,
    'menu'
  );

  ss.setActiveSheet(perf);
  perf.setActiveRange(perf.getRange(startRow, 10));
  ss.toast('오늘 측정할 성과 기록행 ' + rows.length + '건을 만들었습니다.', '성과관리', 6);
  return '신규 성과 기록행 ' + rows.length + '건';
}

function ma_getTodayPerformanceTargetsV151_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var publish = ss.getSheetByName(MA_CFG.SHEETS.PUBLISH);
  if (!publish) throw new Error('발행관리 시트를 찾을 수 없습니다.');

  var perf = ss.getSheetByName(MA_V15_PERFORMANCE_SHEET);
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var lastRow = publish.getLastRow();
  if (lastRow < 2) return [];

  var values = publish.getRange(2, 1, lastRow - 1, 16).getValues();
  var targets = [];

  values.forEach(function(row) {
    var status = ma_text_(row[10]);
    var publishDate = ma_dateOnlyV15_(row[11]);
    var publishUrl = ma_text_(row[12]);
    if (status !== '발행완료' || !publishDate || !publishUrl) return;

    var elapsed = Math.floor((today.getTime() - publishDate.getTime()) / 86400000);
    if (MA_V151_MEASUREMENT_DAYS.indexOf(elapsed) < 0) return;

    var assetId = ma_text_(row[2]);
    targets.push({
      publishId: ma_text_(row[0]),
      contentId: ma_text_(row[1]),
      assetId: assetId,
      platform: ma_text_(row[3]),
      publishDate: publishDate,
      publishUrl: publishUrl,
      elapsedDays: elapsed,
      hasSnapshot: perf ? ma_hasPerformanceSnapshotForDateV15_(perf, assetId, today) : false
    });
  });

  return targets;
}
