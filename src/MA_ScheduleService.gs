/**
 * Marketing Automation V1.2.3 - Publish Schedule Service
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

function ma_schedCalendar_(ss, contentId, assetId, platform, date, time, status) {
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

  // K = 발행URL. 과거 버그로 K에 asset ID가 있던 경우에만 비운다.
  if (legacyAssetInUrl && ma_schedText_(row[10]) === assetId) {
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
