function ma_runHealthCheck() {
  var start = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var errors = [];
  var warnings = [];

  Object.keys(MA_CFG.SHEETS).forEach(function(k) {
    var name = MA_CFG.SHEETS[k];
    if (!ss.getSheetByName(name)) errors.push('필수 시트 없음: ' + name);
  });

  ma_checkHeaders_(ss, MA_CFG.SHEETS.CONTENT, MA_CFG.HEADERS.CONTENT, errors);
  ma_checkHeaders_(ss, MA_CFG.SHEETS.BLOG, MA_CFG.HEADERS.BLOG, errors);
  ma_checkHeaders_(ss, MA_CFG.SHEETS.THREADS, MA_CFG.HEADERS.THREADS, errors);
  ma_checkHeaders_(ss, MA_CFG.SHEETS.PUBLISH, MA_CFG.HEADERS.PUBLISH, errors);
  ma_checkHeaders_(ss, MA_CFG.SHEETS.LOG, MA_CFG.HEADERS.LOG, errors);

  var settings = ma_getSettings_();
  if (String(settings.SYSTEM_VERSION || '') !== MA_CFG.VERSION) {
    warnings.push('설정 SYSTEM_VERSION이 코드 버전과 다름: ' +
      (settings.SYSTEM_VERSION || '(없음)'));
  }

  var level = errors.length ? 'ERROR' : (warnings.length ? 'WARN' : 'INFO');
  var result = errors.length ? 'FAIL' : 'SUCCESS';
  var msg = 'health check: errors=' + errors.length +
    ', warnings=' + warnings.length;

  if (errors.length) msg += ' | ' + errors.join(' / ');
  if (warnings.length) msg += ' | ' + warnings.join(' / ');

  ma_log_('HEALTH_CHECK', '', '', result, level, msg,
    Date.now() - start, 'menu');
  ss.toast(msg, 'Marketing Automation', 8);

  return {ok: errors.length === 0, errors: errors, warnings: warnings};
}

function ma_checkHeaders_(ss, sheetName, expected, errors) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  var actual = sheet.getRange(1, 1, 1, expected.length).getValues()[0].map(String);

  expected.forEach(function(h, i) {
    if (actual[i] !== h) {
      errors.push(sheetName + ' 헤더 불일치 ' + ma_colLetter_(i + 1) +
        '1: expected=' + h + ', actual=' + actual[i]);
    }
  });
}
