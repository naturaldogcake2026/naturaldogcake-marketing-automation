/**
 * Marketing Automation V1.1
 * Manual publish completion workflow
 */

function ma_markSelectedPublishComplete() {
  var start = Date.now();
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    if (sheet.getName() !== MA_CFG.SHEETS.PUBLISH) {
      throw new Error('발행관리 시트에서 발행 항목을 선택해 주세요.');
    }

    var range = sheet.getActiveRange();
    if (!range) {
      throw new Error('발행 항목 행을 선택해 주세요.');
    }

    var row = range.getRow();
    if (row < 2) {
      throw new Error('발행 항목 행을 선택해 주세요.');
    }

    var h = ma_headerMap_(sheet);

    var publishId = ma_pubText_(sheet.getRange(row, h['Publish ID']).getDisplayValue());
    var contentId = ma_pubText_(sheet.getRange(row, h['Content ID']).getDisplayValue());
    var assetId = ma_pubText_(sheet.getRange(row, h['콘텐츠ID']).getDisplayValue());
    var platform = ma_pubText_(sheet.getRange(row, h['플랫폼']).getDisplayValue());
    var reviewStatus = ma_pubText_(sheet.getRange(row, h['검수상태']).getDisplayValue());
    var approval = ma_pubText_(sheet.getRange(row, h['승인여부']).getDisplayValue());
    var publishStatus = ma_pubText_(sheet.getRange(row, h['발행상태']).getDisplayValue());
    var publishUrl = ma_pubText_(sheet.getRange(row, h['발행URL']).getDisplayValue());

    if (!publishId || !contentId || !assetId) {
      throw new Error('선택한 행의 발행 데이터가 올바르지 않습니다.');
    }

    if (reviewStatus !== '검수완료') {
      throw new Error('검수완료된 콘텐츠만 발행완료 처리할 수 있습니다.');
    }

    if (approval !== '승인') {
      throw new Error('승인된 콘텐츠만 발행완료 처리할 수 있습니다.');
    }

    if (publishStatus === '발행완료') {
      var existingPublishDate = ma_pubText_(sheet.getRange(row, h['발행일']).getDisplayValue());
      if (existingPublishDate) {
        throw new Error('이미 발행완료 처리된 콘텐츠입니다.');
      }

      if (!publishUrl) {
        throw new Error('발행완료 상태지만 발행URL이 없습니다. 발행URL을 먼저 입력해 주세요.');
      }

      var plannedDateValue = sheet.getRange(row, h['발행예정일']).getValue();
      var recoveredDate = ma_manualPublishSafeDate_(plannedDateValue);
      if (!recoveredDate) {
        throw new Error('발행완료 상태지만 발행일이 비어 있고 발행예정일도 없습니다. 실제 발행일을 확인해 발행일 열에 직접 입력해 주세요.');
      }

      if (platform === 'Blog') {
        ma_manualPublishSyncBlog_(ss, assetId, recoveredDate, publishUrl);
      } else if (platform === 'Threads') {
        ma_manualPublishSyncThread_(ss, assetId, recoveredDate, publishUrl);
      } else {
        throw new Error('지원하지 않는 플랫폼입니다: ' + platform);
      }

      sheet.getRange(row, h['발행일']).setValue(recoveredDate);
      sheet.getRange(row, h['오류여부']).setValue(false);
      sheet.getRange(row, h['오류메시지']).clearContent();
      ma_manualPublishSyncContent_(ss, contentId);

      ma_log_(
        'MANUAL_PUBLISH_REPAIR_V152',
        contentId,
        assetId,
        'SUCCESS',
        'INFO',
        platform + ' 발행완료 누락정보 복구: ' + publishId,
        Date.now() - start,
        'menu'
      );

      ss.toast('누락된 발행일과 원본 시트 연결정보를 복구했습니다.', 'Marketing Automation', 7);
      return publishId;
    }

    if (publishStatus !== '발행대기') {
      throw new Error('발행상태가 발행대기인 콘텐츠만 처리할 수 있습니다.');
    }

    if (!publishUrl) {
      throw new Error('발행URL을 먼저 입력해 주세요.');
    }

    var now = new Date();

    if (platform === 'Blog') {
      ma_manualPublishSyncBlog_(ss, assetId, now, publishUrl);
    } else if (platform === 'Threads') {
      ma_manualPublishSyncThread_(ss, assetId, now, publishUrl);
    } else {
      throw new Error('지원하지 않는 플랫폼입니다: ' + platform);
    }

    sheet.getRange(row, h['발행상태']).setValue('발행완료');
    sheet.getRange(row, h['발행일']).setValue(now);
    sheet.getRange(row, h['오류여부']).setValue(false);
    sheet.getRange(row, h['오류메시지']).clearContent();

    ma_manualPublishSyncContent_(ss, contentId);

    ma_log_(
      'MANUAL_PUBLISH_COMPLETE',
      contentId,
      assetId,
      'SUCCESS',
      'INFO',
      platform + ' 수동 발행완료 처리: ' + publishId,
      Date.now() - start,
      'menu'
    );

    ss.toast(
      platform + ' 발행완료 처리되었습니다.',
      'Marketing Automation',
      7
    );

    return publishId;

  } catch (err) {
    try {
      ma_log_(
        'MANUAL_PUBLISH_COMPLETE',
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

function ma_manualPublishSyncBlog_(ss, blogId, publishDate, publishUrl) {
  var sheet = ss.getSheetByName(MA_CFG.SHEETS.BLOG);
  if (!sheet) throw new Error('블로그초안 시트를 찾을 수 없습니다.');

  var h = ma_headerMap_(sheet);
  var row = ma_manualPublishFindRow_(sheet, h['Blog ID'], blogId);
  if (!row) throw new Error('블로그 원본을 찾을 수 없습니다: ' + blogId);

  sheet.getRange(row, h['발행상태']).setValue('발행완료');
  sheet.getRange(row, h['발행일']).setValue(publishDate);
  sheet.getRange(row, h['발행URL']).setValue(publishUrl);
}

function ma_manualPublishSyncThread_(ss, threadId, publishDate, publishUrl) {
  var sheet = ss.getSheetByName(MA_CFG.SHEETS.THREADS);
  if (!sheet) throw new Error('Threads초안 시트를 찾을 수 없습니다.');

  var h = ma_headerMap_(sheet);
  var row = ma_manualPublishFindRow_(sheet, h['Thread ID'], threadId);
  if (!row) throw new Error('Threads 원본을 찾을 수 없습니다: ' + threadId);

  sheet.getRange(row, h['발행상태']).setValue('발행완료');
  sheet.getRange(row, h['발행일']).setValue(publishDate);
  sheet.getRange(row, h['발행URL']).setValue(publishUrl);
}

function ma_manualPublishSyncContent_(ss, contentId) {
  var contentSheet = ss.getSheetByName(MA_CFG.SHEETS.CONTENT);
  var publishSheet = ss.getSheetByName(MA_CFG.SHEETS.PUBLISH);

  if (!contentSheet || !publishSheet) {
    throw new Error('콘텐츠DB 또는 발행관리 시트를 찾을 수 없습니다.');
  }

  var ph = ma_headerMap_(publishSheet);
  var logicalLast = ma_pubLogicalLastRow_(publishSheet);
  if (logicalLast < 2) return;

  var rows = publishSheet
    .getRange(2, 1, logicalLast - 1, publishSheet.getLastColumn())
    .getDisplayValues();

  var related = rows.filter(function(r) {
    return ma_pubText_(r[ph['Content ID'] - 1]) === contentId;
  });
  if (!related.length) return;

  var hasPending = related.some(function(r) {
    return ma_pubText_(r[ph['발행상태'] - 1]) !== '발행완료';
  });

  var ch = ma_headerMap_(contentSheet);
  var contentRow = ma_manualPublishFindRow_(contentSheet, ch['Content ID'], contentId);
  if (!contentRow) return;

  contentSheet
    .getRange(contentRow, ch['콘텐츠상태'])
    .setValue(hasPending ? '발행중' : '발행완료');

  if (ch['최종수정일']) {
    contentSheet.getRange(contentRow, ch['최종수정일']).setValue(new Date());
  }
}

function ma_manualPublishFindRow_(sheet, idColumn, idValue) {
  if (!idColumn || idColumn < 1) {
    throw new Error('ID 열을 찾을 수 없습니다.');
  }

  var maxRows = sheet.getMaxRows();
  if (maxRows < 2) return 0;

  var values = sheet
    .getRange(2, idColumn, maxRows - 1, 1)
    .getDisplayValues();

  for (var i = 0; i < values.length; i++) {
    if (ma_pubText_(values[i][0]) === idValue) {
      return i + 2;
    }
  }

  return 0;
}


/**
 * V1.5.2 one-time / maintenance repair.
 * Repairs only rows that are already 발행완료, have a URL, have no 발행일,
 * and still have a usable 발행예정일. It never guesses a date when no schedule exists.
 */
function ma_repairCompletedPublishMetadataV152() {
  var start = Date.now();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(MA_CFG.SHEETS.PUBLISH);
    if (!sheet) throw new Error('발행관리 시트를 찾을 수 없습니다.');

    var h = ma_headerMap_(sheet);
    var last = ma_pubLogicalLastRow_(sheet);
    if (last < 2) {
      ss.toast('복구할 발행 데이터가 없습니다.', 'Marketing Automation', 5);
      return 0;
    }

    var repaired = 0;
    var skippedNoSchedule = 0;
    var contentIds = {};

    for (var row = 2; row <= last; row++) {
      var status = ma_pubText_(sheet.getRange(row, h['발행상태']).getDisplayValue());
      var publishDateText = ma_pubText_(sheet.getRange(row, h['발행일']).getDisplayValue());
      var publishUrl = ma_pubText_(sheet.getRange(row, h['발행URL']).getDisplayValue());
      if (status !== '발행완료' || publishDateText || !publishUrl) continue;

      var plannedDate = ma_manualPublishSafeDate_(sheet.getRange(row, h['발행예정일']).getValue());
      if (!plannedDate) {
        skippedNoSchedule++;
        continue;
      }

      var contentId = ma_pubText_(sheet.getRange(row, h['Content ID']).getDisplayValue());
      var assetId = ma_pubText_(sheet.getRange(row, h['콘텐츠ID']).getDisplayValue());
      var platform = ma_pubText_(sheet.getRange(row, h['플랫폼']).getDisplayValue());

      if (platform === 'Blog') {
        ma_manualPublishSyncBlog_(ss, assetId, plannedDate, publishUrl);
      } else if (platform === 'Threads') {
        ma_manualPublishSyncThread_(ss, assetId, plannedDate, publishUrl);
      } else {
        continue;
      }

      sheet.getRange(row, h['발행일']).setValue(plannedDate);
      sheet.getRange(row, h['오류여부']).setValue(false);
      sheet.getRange(row, h['오류메시지']).clearContent();
      contentIds[contentId] = true;
      repaired++;
    }

    Object.keys(contentIds).forEach(function(contentId) {
      ma_manualPublishSyncContent_(ss, contentId);
    });

    ma_log_(
      'PUBLISH_METADATA_REPAIR_V152',
      '',
      '',
      'SUCCESS',
      'INFO',
      'V1.5.2 발행완료 누락정보 복구: 복구=' + repaired + '건, 일정없음 건너뜀=' + skippedNoSchedule + '건',
      Date.now() - start,
      'menu'
    );

    ss.toast(
      '복구 ' + repaired + '건' + (skippedNoSchedule ? ', 일정이 없어 건너뜀 ' + skippedNoSchedule + '건' : ''),
      'V1.5.2 발행정보 복구',
      8
    );
    return repaired;
  } finally {
    lock.releaseLock();
  }
}

function ma_manualPublishSafeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (!value) return null;
  var d = new Date(value);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
