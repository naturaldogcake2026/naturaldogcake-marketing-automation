/**
 * Marketing Automation V1.3.2 - Publish Schedule Service
 * ERP와 무관한 마케팅 전용 수동 발행 스케줄 보조 기능.
 *
 * 콘텐츠캘린더 스키마:
 * A 발행예정일 / B 요일 / C Content ID / D 플랫폼 / E 주제 / F Pillar
 * G 콘텐츠상태 / H 검수상태 / I 발행상태 / J 발행시간 / K 발행URL / L 메모
 *
 * L열(메모)에 asset ID를 저장해 일정 행을 안정적으로 식별한다.
 * K열(발행URL)은 절대 asset ID로 덮어쓰지 않는다.
 */

function ma_showTodayPublishItems() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('발행관리');
  if (!sheet) throw new Error('발행관리 시트를 찾을 수 없습니다.');

  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var lastRow = ma_schedLastRow_(sheet, 1);
  var items = [];

  if (lastRow >= 2) {
    var rows = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    rows.forEach(function(r, i) {
      if (!ma_schedText_(r[0]) || !r[8] || ma_schedText_(r[10]) === '발행완료') return;
      if (ma_schedDate_(r[8], tz) !== today) return;
      items.push({
        row: i + 2,
        platform: ma_schedText_(r[3]),
        assetId: ma_schedText_(r[2]),
        time: ma_schedTime_(r[9], tz),
        status: ma_schedText_(r[10])
      });
    });
  }

  var ui = SpreadsheetApp.getUi();
  if (!items.length) {
    ui.alert('오늘 발행할 콘텐츠', '오늘 발행 예정인 미완료 콘텐츠가 없습니다.', ui.ButtonSet.OK);
    return [];
  }

  items.sort(function(a, b) {
    return (a.time || '99:99').localeCompare(b.time || '99:99');
  });

  var msg = items.map(function(x) {
    return (x.time || '시간 미지정') + ' · ' + x.platform + ' · ' +
      x.assetId + ' · [' + x.status + '] · 발행관리 ' + x.row + '행';
  }).join('\n\n');

  ui.alert(
    '오늘 발행할 콘텐츠 (' + items.length + '건)',
    msg + '\n\n게시 후 발행관리 M열에 URL을 입력하고 "선택 행 수동 발행완료"를 실행하세요.',
    ui.ButtonSet.OK
  );

  return items;
}

function ma_syncSelectedPublishSchedule() {
  var startedAt = new Date().getTime();
  var contentId = '';
  var assetId = '';
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    if (!sheet || sheet.getName() !== '발행관리') {
      throw new Error('발행관리 시트에서 발행 항목을 선택해 주세요.');
    }

    var rowNo = sheet.getActiveRange().getRow();
    if (rowNo < 2) throw new Error('발행 항목 행을 선택해 주세요.');

    var r = sheet.getRange(rowNo, 1, 1, 16).getValues()[0];
    contentId = ma_schedText_(r[1]);
    assetId = ma_schedText_(r[2]);
    var platform = ma_schedText_(r[3]);
    var scheduledDate = r[8];
    var scheduledTime = r[9];
    var status = ma_schedText_(r[10]);

    if (!contentId || !assetId) {
      throw new Error('Content ID 또는 콘텐츠ID가 없습니다.');
    }

    if (!scheduledDate) {
      throw new Error('발행예정일을 먼저 입력해 주세요.');
    }

    if (status === '발행완료') {
      throw new Error('이미 발행완료된 항목의 일정은 변경하지 않습니다.');
    }

    if (platform === 'Threads') {
      var thr = ss.getSheetByName('Threads초안');
      if (!thr) throw new Error('Threads초안 시트를 찾을 수 없습니다.');

      var threadRow = ma_schedFind_(thr, assetId);
      if (!threadRow) {
        throw new Error('Threads초안에서 Thread ID를 찾을 수 없습니다: ' + assetId);
      }

      thr.getRange(threadRow, 13).setValue(scheduledDate);
    }

    ma_schedCalendar_(
      ss,
      contentId,
      assetId,
      platform,
      scheduledDate,
      scheduledTime,
      status
    );

    ma_log_(
      'PUBLISH_SCHEDULE_SYNC',
      contentId,
      assetId,
      'SUCCESS',
      'INFO',
      '발행 일정 동기화 완료',
      new Date().getTime() - startedAt,
      'menu'
    );

    ss.toast('발행 일정이 동기화되었습니다.', '발행 일정', 4);
    return '발행 일정 동기화 완료: ' + assetId;

  } catch (err) {
    ma_log_(
      'PUBLISH_SCHEDULE_SYNC',
      contentId,
      assetId,
      'ERROR',
      'ERROR',
      String(err && err.message ? err.message : err),
      new Date().getTime() - startedAt,
      'menu'
    );
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function ma_schedCalendar_(ss, contentId, assetId, platform, date, time, status, publishUrl) {
  var sh = ss.getSheetByName('콘텐츠캘린더');
  if (!sh) return;

  // L열(메모)까지 보장한다.
  if (sh.getMaxColumns() < 12) {
    sh.insertColumnsAfter(sh.getMaxColumns(), 12 - sh.getMaxColumns());
  }
  if (ma_schedText_(sh.getRange(1, 12).getValue()) !== '메모') {
    sh.getRange(1, 12).setValue('메모');
  }

  var last = ma_schedLastRow_(sh, 3);
  var target = 0;
  var legacyAssetInUrl = false;

  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 12).getValues();
    for (var i = 0; i < vals.length; i++) {
      var sameContent = ma_schedText_(vals[i][2]) === contentId;
      var samePlatform = ma_schedText_(vals[i][3]) === platform;
      var memoAsset = ma_schedText_(vals[i][11]);
      var urlValue = ma_schedText_(vals[i][10]);

      if (sameContent && samePlatform && memoAsset === assetId) {
        target = i + 2;
        break;
      }

      // V1.2.1 이전 데이터 호환: K열에 asset ID가 잘못 들어간 행.
      if (sameContent && samePlatform && !memoAsset && urlValue === assetId) {
        target = i + 2;
        legacyAssetInUrl = true;
        break;
      }
    }
  }

  if (!target) target = Math.max(2, last + 1);

  var row = sh.getRange(target, 1, 1, 12).getValues()[0];
  row[0] = date;
  row[1] = ma_schedKoreanWeekday_(date);
  row[2] = contentId;
  row[3] = platform;
  row[8] = status;
  row[9] = time || '';

  // K = 발행URL. 발행관리 URL이 있으면 동기화하고,
  // 과거 버그로 K에 asset ID가 있던 경우에는 비운다.
  if (publishUrl) {
    row[10] = publishUrl;
  } else if (legacyAssetInUrl && ma_schedText_(row[10]) === assetId) {
    row[10] = '';
  }

  // L = 메모. 일정 행 식별용 asset ID 저장.
  row[11] = assetId;

  sh.getRange(target, 1, 1, 12).setValues([row]);
}

function ma_schedKoreanWeekday_(value) {
  var d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return '';
  return ['일', '월', '화', '수', '목', '금', '토'][d.getDay()];
}

function ma_schedFind_(sheet, id) {
  var vals = sheet.getRange(
    2,
    1,
    Math.max(1, sheet.getMaxRows() - 1),
    1
  ).getDisplayValues();

  for (var i = 0; i < vals.length; i++) {
    if (ma_schedText_(vals[i][0]) === id) return i + 2;
  }
  return 0;
}

function ma_schedLastRow_(sheet, col) {
  var vals = sheet.getRange(
    2,
    col,
    Math.max(1, sheet.getMaxRows() - 1),
    1
  ).getDisplayValues();

  for (var i = vals.length - 1; i >= 0; i--) {
    if (ma_schedText_(vals[i][0])) return i + 2;
  }
  return 1;
}

function ma_schedDate_(v, tz) {
  var d = v instanceof Date ? v : new Date(v);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function ma_schedTime_(v, tz) {
  if (!v) return '';
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, tz, 'HH:mm');
  }
  return ma_schedText_(v);
}

function ma_schedText_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

/**
 * Marketing Automation V1.3.0
 * Threads daily auto-scheduling
 *
 * Rules:
 * - Threads only. Blog rows are never auto-scheduled here.
 * - Only review-complete rows with final body in publish queue are eligible.
 * - Existing scheduled dates are never changed.
 * - New rows are assigned one per day at 19:00.
 * - Start date = day after the latest existing Threads scheduled date.
 *   If that date would be today/past, start tomorrow.
 * - Because Thread finalization is already a human approval step,
 *   newly auto-scheduled Threads are set to 승인 / 발행대기.
 * - Threads초안 M(발행예정일) and 콘텐츠캘린더 are synchronized.
 */
function ma_autoSchedulePendingThreads_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var publishSheet = ss.getSheetByName('발행관리');
  var threadsSheet = ss.getSheetByName('Threads초안');

  if (!publishSheet) throw new Error('발행관리 시트를 찾을 수 없습니다.');
  if (!threadsSheet) throw new Error('Threads초안 시트를 찾을 수 없습니다.');

  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
  var logicalLast = ma_pubLogicalLastRow_(publishSheet);
  if (logicalLast < 2) return 0;

  var rows = publishSheet.getRange(2, 1, logicalLast - 1, 16).getValues();

  var latestDate = null;
  var pending = [];

  rows.forEach(function(r, i) {
    var publishId = ma_schedText_(r[0]);
    var contentId = ma_schedText_(r[1]);
    var assetId = ma_schedText_(r[2]);
    var platform = ma_schedText_(r[3]);
    var hasFinalBody = r[5] === true || ma_schedText_(r[5]).toUpperCase() === 'TRUE';
    var reviewStatus = ma_schedText_(r[6]);
    var scheduledDate = r[8];
    var publishStatus = ma_schedText_(r[10]);

    if (!publishId || platform !== 'Threads') return;

    if (scheduledDate) {
      var normalized = ma_schedNormalizeDate_(scheduledDate);
      if (normalized && (!latestDate || normalized.getTime() > latestDate.getTime())) {
        latestDate = normalized;
      }
      return;
    }

    if (!assetId || !contentId) return;
    if (!hasFinalBody || reviewStatus !== '검수완료') return;
    if (publishStatus === '발행완료') return;

    pending.push({
      rowNo: i + 2,
      contentId: contentId,
      assetId: assetId
    });
  });

  if (!pending.length) return 0;

  var today = ma_schedToday_(tz);
  var cursor = latestDate ? ma_schedAddDays_(latestDate, 1) : ma_schedAddDays_(today, 1);

  if (cursor.getTime() <= today.getTime()) {
    cursor = ma_schedAddDays_(today, 1);
  }

  var scheduledCount = 0;

  pending.forEach(function(item) {
    var dateValue = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate());
    var timeValue = '19:00';

    // H 승인여부 / I 발행예정일 / J 발행예정시간 / K 발행상태
    publishSheet.getRange(item.rowNo, 8, 1, 4).setValues([[
      '승인',
      dateValue,
      timeValue,
      '발행대기'
    ]]);

    var threadRow = ma_schedFind_(threadsSheet, item.assetId);
    if (!threadRow) {
      throw new Error('Threads초안에서 Thread ID를 찾을 수 없습니다: ' + item.assetId);
    }

    // Threads초안 M = 발행예정일
    threadsSheet.getRange(threadRow, 13).setValue(dateValue);

    ma_schedCalendar_(
      ss,
      item.contentId,
      item.assetId,
      'Threads',
      dateValue,
      timeValue,
      '발행대기'
    );

    scheduledCount++;
    cursor = ma_schedAddDays_(cursor, 1);
  });

  ma_log_(
    'THREADS_AUTO_SCHEDULE',
    '',
    '',
    'SUCCESS',
    'INFO',
    'Threads 자동 일정 배정: ' + scheduledCount + '건, 매일 19:00',
    0,
    'system'
  );

  return scheduledCount;
}


/**
 * Marketing Automation V1.3.3
 * 발행관리 = 일정 source-of-truth.
 *
 * 핵심 원칙
 * 1) 발행완료 Threads는 절대 재배치하지 않는다.
 * 2) 발행완료 Threads의 기존 발행예정일은 해당 날짜를 점유한 것으로 본다.
 * 3) 미완료 Threads는 오늘 이후 빈 날짜에 하루 1건, 19:00로 순차 재배치한다.
 * 4) 콘텐츠캘린더의 자동생성 행(L열이 BLOG-/THR- asset ID)은 매번 제거 후
 *    발행관리 기준으로 재작성한다. 수동/Seed 행은 보존한다.
 * 5) Threads초안 M열은 발행관리 I열과 동기화한다.
 */
function ma_reflowFuturePendingThreadsV133_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var publishSheet = ss.getSheetByName('발행관리');
  var threadsSheet = ss.getSheetByName('Threads초안');
  if (!publishSheet) throw new Error('발행관리 시트를 찾을 수 없습니다.');
  if (!threadsSheet) throw new Error('Threads초안 시트를 찾을 수 없습니다.');

  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
  var today = ma_schedToday_(tz);
  var logicalLast = ma_pubLogicalLastRow_(publishSheet);
  if (logicalLast < 2) return {pending:0, changed:0};

  var rows = publishSheet.getRange(2, 1, logicalLast - 1, 16).getValues();
  var occupied = {};
  var pending = [];

  // 먼저 발행완료 Threads의 예약일을 점유 처리한다.
  rows.forEach(function(r) {
    var publishId = ma_schedText_(r[0]);
    var platform = ma_schedText_(r[3]);
    var status = ma_schedText_(r[10]);
    if (!publishId || platform !== 'Threads' || status !== '발행완료') return;
    var scheduled = ma_schedNormalizeDate_(r[8]);
    if (scheduled) occupied[ma_schedDateKey_(scheduled, tz)] = true;
  });

  // 미완료 Threads를 현재 예정일 순서대로 모은다.
  rows.forEach(function(r, i) {
    var publishId = ma_schedText_(r[0]);
    var contentId = ma_schedText_(r[1]);
    var assetId = ma_schedText_(r[2]);
    var platform = ma_schedText_(r[3]);
    var status = ma_schedText_(r[10]);
    if (!publishId || platform !== 'Threads' || status === '발행완료') return;
    if (!contentId || !assetId) return;

    var reviewStatus = ma_schedText_(r[6]);
    var hasFinal = r[5] === true || ma_schedText_(r[5]).toUpperCase() === 'TRUE';
    if (!hasFinal || reviewStatus !== '검수완료') return;

    var scheduled = ma_schedNormalizeDate_(r[8]);
    // 과거 미완료 일정도 운영상 오류이므로 오늘 이후로 다시 배치한다.
    pending.push({
      rowNo: i + 2,
      contentId: contentId,
      assetId: assetId,
      scheduledDate: scheduled,
      originalIndex: i
    });
  });

  pending.sort(function(a, b) {
    var at = a.scheduledDate ? a.scheduledDate.getTime() : Number.MAX_SAFE_INTEGER;
    var bt = b.scheduledDate ? b.scheduledDate.getTime() : Number.MAX_SAFE_INTEGER;
    if (at !== bt) return at - bt;
    return a.originalIndex - b.originalIndex;
  });

  var cursor = ma_schedAddDays_(today, 1);
  var changed = 0;

  pending.forEach(function(item) {
    while (occupied[ma_schedDateKey_(cursor, tz)]) {
      cursor = ma_schedAddDays_(cursor, 1);
    }

    var newDate = new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate());
    var oldKey = item.scheduledDate ? ma_schedDateKey_(item.scheduledDate, tz) : '';
    var newKey = ma_schedDateKey_(newDate, tz);

    publishSheet.getRange(item.rowNo, 8, 1, 4).setValues([[
      '승인', newDate, '19:00', '발행대기'
    ]]);

    var threadRow = ma_schedFind_(threadsSheet, item.assetId);
    if (!threadRow) throw new Error('Threads초안에서 Thread ID를 찾을 수 없습니다: ' + item.assetId);
    threadsSheet.getRange(threadRow, 13).setValue(newDate);

    if (oldKey !== newKey) changed++;
    occupied[newKey] = true;
    cursor = ma_schedAddDays_(cursor, 1);
  });

  ma_log_(
    'THREADS_SCHEDULE_REFLOW_V133', '', '', 'SUCCESS', 'INFO',
    'V1.3.3 미완료 Threads 재배치: 대상=' + pending.length + '건, 날짜변경=' + changed + '건',
    0, 'system'
  );

  return {pending:pending.length, changed:changed};
}

/**
 * 콘텐츠캘린더 자동생성 행을 발행관리 기준으로 완전 재작성한다.
 * L열 메모가 BLOG-* 또는 THR-* 인 행만 자동행으로 간주한다.
 * Seed/수동 행은 보존한다.
 */
function ma_rebuildPublishCalendarV133_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var publishSheet = ss.getSheetByName('발행관리');
  var calendar = ss.getSheetByName('콘텐츠캘린더');
  if (!publishSheet) throw new Error('발행관리 시트를 찾을 수 없습니다.');
  if (!calendar) throw new Error('콘텐츠캘린더 시트를 찾을 수 없습니다.');

  if (calendar.getMaxColumns() < 12) {
    calendar.insertColumnsAfter(calendar.getMaxColumns(), 12 - calendar.getMaxColumns());
  }
  if (ma_schedText_(calendar.getRange(1, 12).getValue()) !== '메모') {
    calendar.getRange(1, 12).setValue('메모');
  }

  // 자동행을 아래에서부터 제거해 중복/레거시를 확실히 청소한다.
  var lastCal = ma_schedLastRow_(calendar, 3);
  var removed = 0;
  if (lastCal >= 2) {
    var calRows = calendar.getRange(2, 1, lastCal - 1, 12).getDisplayValues();
    for (var i = calRows.length - 1; i >= 0; i--) {
      var memo = ma_schedText_(calRows[i][11]);
      if (/^(BLOG|THR)-/.test(memo)) {
        calendar.deleteRow(i + 2);
        removed++;
      }
    }
  }

  var logicalLast = ma_pubLogicalLastRow_(publishSheet);
  if (logicalLast < 2) return {removed:removed, written:0};

  var rows = publishSheet.getRange(2, 1, logicalLast - 1, 16).getValues();
  var output = [];
  rows.forEach(function(r) {
    var publishId = ma_schedText_(r[0]);
    var contentId = ma_schedText_(r[1]);
    var assetId = ma_schedText_(r[2]);
    var platform = ma_schedText_(r[3]);
    var scheduledDate = r[8];
    if (!publishId || !contentId || !assetId || !scheduledDate) return;
    if (platform !== 'Threads' && platform !== 'Blog') return;

    output.push([
      scheduledDate,
      ma_schedKoreanWeekday_(scheduledDate),
      contentId,
      platform,
      '', '', '', '',
      ma_schedText_(r[10]),
      r[9] || '',
      ma_schedText_(r[12]),
      assetId
    ]);
  });

  if (output.length) {
    var startRow = ma_schedLastRow_(calendar, 3) + 1;
    if (startRow < 2) startRow = 2;
    if (startRow + output.length - 1 > calendar.getMaxRows()) {
      calendar.insertRowsAfter(calendar.getMaxRows(), startRow + output.length - 1 - calendar.getMaxRows());
    }
    calendar.getRange(startRow, 1, output.length, 12).setValues(output);
  }

  ma_log_(
    'PUBLISH_CALENDAR_REBUILD_V133', '', '', 'SUCCESS', 'INFO',
    'V1.3.3 캘린더 재작성: 기존 자동행 제거=' + removed + '건, 신규 작성=' + output.length + '건',
    0, 'system'
  );

  return {removed:removed, written:output.length};
}

/** 발행관리 -> Threads초안 M 동기화. */
function ma_syncThreadsScheduleFromPublishV133_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var publishSheet = ss.getSheetByName('발행관리');
  var threadsSheet = ss.getSheetByName('Threads초안');
  if (!publishSheet || !threadsSheet) return 0;

  var logicalLast = ma_pubLogicalLastRow_(publishSheet);
  if (logicalLast < 2) return 0;
  var rows = publishSheet.getRange(2, 1, logicalLast - 1, 16).getValues();
  var count = 0;
  rows.forEach(function(r) {
    if (ma_schedText_(r[3]) !== 'Threads') return;
    var assetId = ma_schedText_(r[2]);
    if (!assetId) return;
    var tr = ma_schedFind_(threadsSheet, assetId);
    if (!tr) return;
    threadsSheet.getRange(tr, 13).setValue(r[8] || '');
    count++;
  });
  return count;
}

/** 메뉴용 V1.3.3 정리 함수. */
function ma_reconcileAllPublishSchedules() {
  var startedAt = Date.now();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) throw new Error('다른 마케팅 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');

  try {
    var reflow = ma_reflowFuturePendingThreadsV133_();
    var threadSync = ma_syncThreadsScheduleFromPublishV133_();
    var calendar = ma_rebuildPublishCalendarV133_();

    ma_log_(
      'PUBLISH_SCHEDULE_RECONCILE_V133', '', '', 'SUCCESS', 'INFO',
      'V1.3.3 완료: Threads 대상=' + reflow.pending + ', 날짜변경=' + reflow.changed +
      ', Threads초안 동기화=' + threadSync + ', 캘린더 제거=' + calendar.removed +
      ', 캘린더 작성=' + calendar.written,
      Date.now() - startedAt, 'menu'
    );

    SpreadsheetApp.getActiveSpreadsheet().toast(
      'V1.3.3 완료 · Threads ' + reflow.pending + '건 정렬 / 캘린더 ' + calendar.written + '건 재작성',
      '발행 일정 정리', 8
    );
    return 'V1.3.3 정리 완료';
  } catch (err) {
    try {
      ma_log_('PUBLISH_SCHEDULE_RECONCILE_V133', '', '', 'FAIL', 'ERROR',
        String(err && err.message ? err.message : err), Date.now() - startedAt, 'menu');
    } catch (ignore) {}
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function ma_schedDateKey_(date, tz) {
  if (!date || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, tz || 'Asia/Seoul', 'yyyy-MM-dd');
}

function ma_schedNormalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  var s = ma_schedText_(value);
  var m = s.match(/^(\d{4})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})\.?$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  var d = new Date(value);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function ma_schedToday_(tz) {
  var parts = Utilities.formatDate(new Date(), tz, 'yyyy,MM,dd').split(',');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function ma_schedAddDays_(date, days) {
  var d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() + days);
  return d;
}
