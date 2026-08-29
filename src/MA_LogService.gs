function ma_log_(feature, contentId, targetId, result, level, message, elapsedMs, actor) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MA_CFG.SHEETS.LOG);
    if (!sheet) return;

    var runId = 'RUN-' + Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || 'Asia/Seoul',
      'yyyyMMdd-HHmmss-SSS'
    );

    sheet.appendRow([
      new Date(), runId, feature || '', contentId || '', targetId || '',
      result || '', level || 'INFO', message || '', Number(elapsedMs || 0),
      actor || ''
    ]);
  } catch (ignore) {}
}
