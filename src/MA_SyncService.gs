function ma_recalcBlogReviewState_(sheet, row) {
  var h = ma_headerMap_(sheet);
  var vals = [
    String(sheet.getRange(row, h['사실검수']).getValue() || ''),
    String(sheet.getRange(row, h['문체검수']).getValue() || ''),
    String(sheet.getRange(row, h['SEO검수']).getValue() || '')
  ];
  var next = '미검수';

  if (vals.indexOf('수정필요') >= 0) {
    next = '수정필요';
  } else if (vals.every(function(v){ return v === '검수완료'; })) {
    next = '검수완료';
  } else if (vals.some(function(v){ return v === '검수중' || v === '검수완료'; })) {
    next = '검수중';
  }

  var cell = sheet.getRange(row, h['검수상태']);
  if (String(cell.getValue() || '') !== next) cell.setValue(next);
}

function ma_recalcPublishState_(sheet, row) {
  var h = ma_headerMap_(sheet);
  var review = String(sheet.getRange(row, h['검수상태']).getValue() || '');
  var approval = String(sheet.getRange(row, h['승인여부']).getValue() || '');
  var finalExists = sheet.getRange(row, h['최종본문 존재']).getValue() === true;
  var stateCell = sheet.getRange(row, h['발행상태']);
  var current = String(stateCell.getValue() || '');

  if (current === '발행완료' || current === '발행중') return;

  if (approval === '승인' && review === '검수완료' && finalExists) {
    stateCell.setValue('발행대기');
    sheet.getRange(row, h['오류여부']).setValue(false);
    sheet.getRange(row, h['오류메시지']).clearContent();
  } else if (approval === '승인' && !finalExists) {
    stateCell.setValue('미준비');
    sheet.getRange(row, h['오류여부']).setValue(true);
    sheet.getRange(row, h['오류메시지']).setValue('최종본문이 없습니다.');
  } else if (approval === '보류' || approval === '반려') {
    stateCell.setValue('미준비');
  }
}

function ma_syncStates() {
  var start = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var blog = ss.getSheetByName(MA_CFG.SHEETS.BLOG);
  var publish = ss.getSheetByName(MA_CFG.SHEETS.PUBLISH);

  if (blog && blog.getLastRow() >= 2) {
    for (var r = 2; r <= blog.getLastRow(); r++) ma_recalcBlogReviewState_(blog, r);
  }
  if (publish && publish.getLastRow() >= 2) {
    for (var p = 2; p <= publish.getLastRow(); p++) ma_recalcPublishState_(publish, p);
  }

  ma_log_('STATE_SYNC', '', '', 'SUCCESS', 'INFO', '상태 동기화 완료',
    Date.now() - start, 'menu');
  ss.toast('상태 동기화 완료', 'Marketing Automation', 5);
}

function ma_syncAll() {
  var start = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var content = ss.getSheetByName(MA_CFG.SHEETS.CONTENT);

  if (content && content.getLastRow() >= 2) {
    for (var r = 2; r <= content.getLastRow(); r++) {
      ma_applyContentDefaultsForRow_(content, r);
    }
  }

  ma_syncStates();
  ma_syncPublishQueue();

  ma_log_('SYNC_ALL', '', '', 'SUCCESS', 'INFO', '전체 동기화 완료',
    Date.now() - start, 'menu');
  ss.toast('전체 동기화 완료', 'Marketing Automation', 5);
}
