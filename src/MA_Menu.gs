function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('자연담은멍케이크 콘텐츠')
    .addItem('선택 행 ID/기본값 적용', 'ma_applySelectedContentDefaults')
    .addItem('선택 콘텐츠 블로그 초안 생성', 'ma_createSelectedBlogDraft')
    .addItem('선택 콘텐츠 AI 블로그 작성', 'ma_generateSelectedBlogDraft')
    .addSeparator()
    .addItem('발행대기 생성', 'ma_syncPublishQueue')
    .addItem('상태 동기화', 'ma_syncStates')
    .addSeparator()
    .addItem('상태 점검', 'ma_runHealthCheck')
    .addItem('전체 동기화', 'ma_syncAll')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  if (row < 2) return;

  try {
    if (sheet.getName() === MA_CFG.SHEETS.CONTENT) {
      ma_applyContentDefaultsForRow_(sheet, row);
    } else if (sheet.getName() === MA_CFG.SHEETS.BLOG) {
      ma_recalcBlogReviewState_(sheet, row);
    } else if (sheet.getName() === MA_CFG.SHEETS.PUBLISH) {
      ma_recalcPublishState_(sheet, row);
    }
  } catch (err) {
    ma_log_(
      'ON_EDIT',
      '',
      '',
      'ERROR',
      'ERROR',
      String(err && err.message ? err.message : err),
      0,
      'onEdit'
    );
  }
}
