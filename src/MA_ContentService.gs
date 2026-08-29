function ma_applySelectedContentDefaults() {
  var start = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  if (sheet.getName() !== MA_CFG.SHEETS.CONTENT) {
    throw new Error('콘텐츠DB 시트에서 실행해 주세요.');
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) throw new Error('헤더 행이 아닌 콘텐츠 행을 선택해 주세요.');

  var contentId = ma_applyContentDefaultsForRow_(sheet, row);
  ma_log_('CONTENT_DEFAULTS', contentId, contentId, 'SUCCESS', 'INFO',
    '선택 행 ID/기본값 적용 완료', Date.now() - start, 'menu');
  ss.toast('적용 완료: ' + (contentId || 'ID 미생성'), 'Marketing Automation', 5);
}

function ma_applyContentDefaultsForRow_(sheet, row) {
  var headers = ma_headerMap_(sheet);
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var topic = ma_valueByHeader_(data, headers, '주제');
  var keyword = ma_valueByHeader_(data, headers, '핵심키워드');
  var pillar = ma_valueByHeader_(data, headers, 'Pillar');
  var hasMeaningfulInput = String(topic || keyword || pillar || '').trim() !== '';
  if (!hasMeaningfulInput) return '';

  var settings = ma_getSettings_();
  var idCell = sheet.getRange(row, headers['Content ID']);
  var contentId = String(idCell.getValue() || '').trim();

  if (!contentId) {
    contentId = ma_nextId_(
      sheet, headers['Content ID'],
      settings.ID_CONTENT_PREFIX || 'CNT', 4
    );
    idCell.setValue(contentId);
  }

  if (!sheet.getRange(row, headers['등록일']).getValue()) {
    sheet.getRange(row, headers['등록일']).setValue(new Date());
  }
  if (!sheet.getRange(row, headers['콘텐츠상태']).getValue()) {
    sheet.getRange(row, headers['콘텐츠상태'])
      .setValue(settings.DEFAULT_CONTENT_STATUS || '기획중');
  }
  if (!sheet.getRange(row, headers['우선순위']).getValue()) {
    sheet.getRange(row, headers['우선순위'])
      .setValue(settings.DEFAULT_PRIORITY || '보통');
  }
  if (!sheet.getRange(row, headers['Threads 생성수']).getValue()) {
    sheet.getRange(row, headers['Threads 생성수'])
      .setValue(Number(settings.THREAD_COUNT || 3));
  }

  sheet.getRange(row, headers['최종수정일']).setValue(new Date());
  return contentId;
}

function ma_nextId_(sheet, idCol, prefix, digits) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    var year = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || 'Asia/Seoul',
      'yyyy'
    );
    var re = new RegExp('^' + ma_escapeRegex_(prefix) + '-' + year + '-(\\d+)$');
    var max = 0;

    if (sheet.getLastRow() >= 2) {
      var vals = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getDisplayValues();
      vals.forEach(function(r) {
        var m = String(r[0] || '').match(re);
        if (m) max = Math.max(max, Number(m[1]) || 0);
      });
    }

    return prefix + '-' + year + '-' + String(max + 1).padStart(digits, '0');
  } finally {
    lock.releaseLock();
  }
}
