/**
 * Marketing Automation V1
 * Publish queue service
 *
 * Important:
 * Google Sheets Tables / BOOLEAN columns can leave FALSE values in visually empty rows.
 * Do not use sheet.getLastRow() to determine the logical end of the publish queue.
 * Publish ID (column A) is the source of truth for occupied publish rows.
 */

function ma_syncPublishQueue() {
  var start = Date.now();
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var publishSheet = ss.getSheetByName(MA_CFG.SHEETS.PUBLISH);
    var blogSheet = ss.getSheetByName(MA_CFG.SHEETS.BLOG);
    var threadsSheet = ss.getSheetByName(MA_CFG.SHEETS.THREADS);

    if (!publishSheet) throw new Error('발행관리 시트를 찾을 수 없습니다.');
    if (!blogSheet) throw new Error('블로그초안 시트를 찾을 수 없습니다.');
    if (!threadsSheet) throw new Error('Threads초안 시트를 찾을 수 없습니다.');

    var existingAssetIds = ma_pubExistingAssetIds_(publishSheet);
    var created = 0;

    created += ma_pubSyncBlogs_(
      publishSheet,
      blogSheet,
      existingAssetIds
    );

    created += ma_pubSyncThreads_(
      publishSheet,
      threadsSheet,
      existingAssetIds
    );

    ma_log_(
      'PUBLISH_QUEUE_SYNC',
      '',
      '',
      'SUCCESS',
      'INFO',
      '발행관리 신규 등록: ' + created + '건',
      Date.now() - start,
      'menu'
    );

    ss.toast(
      created > 0
        ? '발행관리 신규 등록: ' + created + '건'
        : '새로 등록할 발행 항목이 없습니다.',
      'Marketing Automation',
      7
    );

    return created;

  } catch (err) {
    try {
      ma_log_(
        'PUBLISH_QUEUE_SYNC',
        '',
        '',
        'FAIL',
        'ERROR',
        String(err && err.message ? err.message : err),
        Date.now() - start,
        'menu'
      );
    } catch (ignore) {}

    throw err;
  } finally {
    lock.releaseLock();
  }
}


function ma_pubSyncBlogs_(publishSheet, blogSheet, existingAssetIds) {
  if (blogSheet.getLastRow() < 2) return 0;

  var h = ma_headerMap_(blogSheet);
  var values = blogSheet
    .getRange(2, 1, blogSheet.getLastRow() - 1, blogSheet.getLastColumn())
    .getDisplayValues();

  var created = 0;

  values.forEach(function(r) {
    var blogId = ma_pubText_(r[h['Blog ID'] - 1]);
    var contentId = ma_pubText_(r[h['Content ID'] - 1]);
    var finalTitle = ma_pubText_(r[h['최종제목'] - 1]);
    var finalBody = ma_pubText_(r[h['최종본문'] - 1]);
    var reviewStatus = ma_pubText_(r[h['검수상태'] - 1]);

    if (!blogId || !contentId) return;
    if (reviewStatus !== '검수완료') return;
    if (!finalBody) return;
    if (existingAssetIds[blogId]) return;

    var publishId = ma_pubNextId_(publishSheet);

    ma_pubWriteRow_(publishSheet, [
      publishId,          // A Publish ID
      contentId,          // B Content ID
      blogId,             // C 콘텐츠ID
      'Blog',             // D 플랫폼
      finalTitle,         // E 최종제목
      true,               // F 최종본문 존재
      '검수완료',         // G 검수상태
      '대기',             // H 승인여부
      '',                 // I 발행예정일
      '',                 // J 발행예정시간
      '미준비',           // K 발행상태
      '',                 // L 발행일
      '',                 // M 발행URL
      false,              // N 오류여부
      '',                 // O 오류메시지
      0                   // P 재시도횟수
    ]);

    existingAssetIds[blogId] = true;
    created++;
  });

  return created;
}


function ma_pubSyncThreads_(publishSheet, threadsSheet, existingAssetIds) {
  if (threadsSheet.getLastRow() < 2) return 0;

  var h = ma_headerMap_(threadsSheet);
  var values = threadsSheet
    .getRange(2, 1, threadsSheet.getLastRow() - 1, threadsSheet.getLastColumn())
    .getDisplayValues();

  var created = 0;

  values.forEach(function(r) {
    var threadId = ma_pubText_(r[h['Thread ID'] - 1]);
    var contentId = ma_pubText_(r[h['Content ID'] - 1]);
    var finalBody = ma_pubText_(r[h['최종본문'] - 1]);
    var reviewStatus = ma_pubText_(r[h['검수상태'] - 1]);

    if (!threadId || !contentId) return;
    if (reviewStatus !== '검수완료') return;
    if (!finalBody) return;
    if (existingAssetIds[threadId]) return;

    var publishId = ma_pubNextId_(publishSheet);

    ma_pubWriteRow_(publishSheet, [
      publishId,
      contentId,
      threadId,
      'Threads',
      '',
      true,
      '검수완료',
      '대기',
      '',
      '',
      '미준비',
      '',
      '',
      false,
      '',
      0
    ]);

    existingAssetIds[threadId] = true;
    created++;
  });

  return created;
}


/**
 * Writes immediately after the last row that has a real Publish ID in column A.
 * This avoids FALSE/default values in other columns affecting queue position.
 */
function ma_pubWriteRow_(sheet, rowValues) {
  var targetRow = ma_pubLogicalLastRow_(sheet) + 1;

  if (targetRow < 2) targetRow = 2;

  if (targetRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), Math.max(100, targetRow - sheet.getMaxRows()));
  }

  sheet
    .getRange(targetRow, 1, 1, rowValues.length)
    .setValues([rowValues]);

  return targetRow;
}


/**
 * Logical last row = last non-empty Publish ID in A.
 * Header row is 1. If there are no publish records, returns 1.
 */
function ma_pubLogicalLastRow_(sheet) {
  var maxRows = sheet.getMaxRows();
  if (maxRows < 2) return 1;

  var ids = sheet
    .getRange(2, 1, maxRows - 1, 1)
    .getDisplayValues();

  for (var i = ids.length - 1; i >= 0; i--) {
    if (ma_pubText_(ids[i][0])) {
      return i + 2;
    }
  }

  return 1;
}


/**
 * Existing source assets already registered in publish queue.
 * Uses column C (콘텐츠ID), but only rows with a real Publish ID are considered.
 */
function ma_pubExistingAssetIds_(sheet) {
  var out = {};
  var logicalLast = ma_pubLogicalLastRow_(sheet);
  if (logicalLast < 2) return out;

  var h = ma_headerMap_(sheet);
  var rows = sheet
    .getRange(2, 1, logicalLast - 1, sheet.getLastColumn())
    .getDisplayValues();

  rows.forEach(function(r) {
    var publishId = ma_pubText_(r[h['Publish ID'] - 1]);
    var assetId = ma_pubText_(r[h['콘텐츠ID'] - 1]);
    if (publishId && assetId) out[assetId] = true;
  });

  return out;
}


/**
 * Next PUB-yyyy-###### ID.
 * Scans real Publish IDs only, regardless of physical row position.
 */
function ma_pubNextId_(sheet) {
  var year = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'Asia/Seoul',
    'yyyy'
  );

  var max = 0;
  var maxRows = sheet.getMaxRows();

  if (maxRows >= 2) {
    var ids = sheet
      .getRange(2, 1, maxRows - 1, 1)
      .getDisplayValues();

    ids.forEach(function(r) {
      var id = ma_pubText_(r[0]);
      var m = id.match(/^PUB-(\d{4})-(\d{6})$/);
      if (m && m[1] === year) {
        max = Math.max(max, Number(m[2]));
      }
    });
  }

  return 'PUB-' + year + '-' + String(max + 1).padStart(6, '0');
}


function ma_pubText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}
