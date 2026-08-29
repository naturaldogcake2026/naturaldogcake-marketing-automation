function ma_syncPublishQueue() {
  var start = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var publish = ss.getSheetByName(MA_CFG.SHEETS.PUBLISH);
  var created = 0;

  created += ma_syncBlogToPublish_(ss, publish);
  created += ma_syncThreadsToPublish_(ss, publish);

  ma_log_('PUBLISH_QUEUE_SYNC', '', '', 'SUCCESS', 'INFO',
    '발행관리 신규 등록: ' + created + '건', Date.now() - start, 'menu');
  ss.toast('발행관리 신규 등록 ' + created + '건', 'Marketing Automation', 5);
  return created;
}

function ma_syncBlogToPublish_(ss, publishSheet) {
  var blog = ss.getSheetByName(MA_CFG.SHEETS.BLOG);
  if (!blog || blog.getLastRow() < 2) return 0;

  var h = ma_headerMap_(blog);
  var ph = ma_headerMap_(publishSheet);
  var rows = blog.getRange(2, 1, blog.getLastRow() - 1, blog.getLastColumn()).getValues();
  var existing = ma_existingPublishTargets_(publishSheet, ph);
  var count = 0;

  rows.forEach(function(r) {
    var blogId = String(r[h['Blog ID'] - 1] || '').trim();
    var contentId = String(r[h['Content ID'] - 1] || '').trim();
    var review = String(r[h['검수상태'] - 1] || '').trim();
    var finalBody = String(r[h['최종본문'] - 1] || '').trim();
    var finalTitle = String(r[h['최종제목'] - 1] || r[h['제목1'] - 1] || '').trim();

    if (!blogId || review !== '검수완료' || existing[blogId]) return;

    var pubId = ma_nextId_(publishSheet, ph['Publish ID'], 'PUB', 6);
    publishSheet.appendRow([
      pubId, contentId, blogId, 'Blog', finalTitle, !!finalBody, review, '대기',
      '', '', '미준비', '', '', false, '', 0
    ]);
    existing[blogId] = true;
    count++;
  });
  return count;
}

function ma_syncThreadsToPublish_(ss, publishSheet) {
  var th = ss.getSheetByName(MA_CFG.SHEETS.THREADS);
  if (!th || th.getLastRow() < 2) return 0;

  var h = ma_headerMap_(th);
  var ph = ma_headerMap_(publishSheet);
  var rows = th.getRange(2, 1, th.getLastRow() - 1, th.getLastColumn()).getValues();
  var existing = ma_existingPublishTargets_(publishSheet, ph);
  var count = 0;

  rows.forEach(function(r) {
    var threadId = String(r[h['Thread ID'] - 1] || '').trim();
    var contentId = String(r[h['Content ID'] - 1] || '').trim();
    var review = String(r[h['검수상태'] - 1] || '').trim();
    var finalBody = String(r[h['최종본문'] - 1] || '').trim();
    var scheduleDate = r[h['발행예정일'] - 1] || '';

    if (!threadId || review !== '검수완료' || existing[threadId]) return;

    var pubId = ma_nextId_(publishSheet, ph['Publish ID'], 'PUB', 6);
    publishSheet.appendRow([
      pubId, contentId, threadId, 'Threads', '', !!finalBody, review, '대기',
      scheduleDate, '', '미준비', '', '', false, '', 0
    ]);
    existing[threadId] = true;
    count++;
  });
  return count;
}

function ma_existingPublishTargets_(sheet, h) {
  var map = {};
  if (sheet.getLastRow() < 2) return map;
  var vals = sheet.getRange(2, h['콘텐츠ID'], sheet.getLastRow() - 1, 1).getValues();
  vals.forEach(function(r) {
    var v = String(r[0] || '').trim();
    if (v) map[v] = true;
  });
  return map;
}
