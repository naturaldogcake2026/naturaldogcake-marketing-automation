/**
 * Marketing Automation V1
 * Threads review/finalization flow
 *
 * Workflow:
 * AI source = F Hook + G 본문 + H CTA
 * I 최종본문 = human review working copy / final copy
 * J 검수상태
 * K 승인일
 *
 * Explicit finalization is the approval gate.
 */

function ma_prepareSelectedThreadReview() {
  var start = Date.now();
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    if (sheet.getName() !== MA_CFG.SHEETS.THREADS) {
      throw new Error('Threads초안 시트에서 실행해 주세요.');
    }

    var row = sheet.getActiveRange().getRow();
    if (row < 2) {
      throw new Error('헤더가 아닌 Thread 행을 선택해 주세요.');
    }

    var h = ma_headerMap_(sheet);
    var v = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

    var threadId = ma_threadReviewText_(v[h['Thread ID'] - 1]);
    var contentId = ma_threadReviewText_(v[h['Content ID'] - 1]);
    var blogId = ma_threadReviewText_(v[h['Blog ID'] - 1]);
    var hook = ma_threadReviewText_(v[h['Hook'] - 1]);
    var body = ma_threadReviewText_(v[h['본문'] - 1]);
    var cta = ma_threadReviewText_(v[h['CTA'] - 1]);
    var finalText = ma_threadReviewText_(v[h['최종본문'] - 1]);
    var reviewStatus = ma_threadReviewText_(v[h['검수상태'] - 1]);

    if (!threadId) throw new Error('Thread ID가 없습니다.');
    if (!hook || !body || !cta) {
      throw new Error('AI 원본 Hook/본문/CTA가 완전하지 않습니다.');
    }
    if (reviewStatus === '검수완료') {
      throw new Error('이미 검수완료된 Thread입니다: ' + threadId);
    }

    var composed = ma_thrComposeFinal_(hook, body, cta);

    // Compatibility with rows generated before this review flow:
    // old generator pre-filled I with exactly the AI-composed text.
    if (!finalText) {
      sheet.getRange(row, h['최종본문']).setValue(composed);
    } else if (finalText !== composed && reviewStatus === '미검수') {
      throw new Error(
        '최종본문(I열)이 AI 원본과 다릅니다. 이미 사람이 수정한 내용일 수 있어 자동으로 덮어쓰지 않았습니다.'
      );
    }

    sheet.getRange(row, h['검수상태']).setValue('검수중');
    sheet.getRange(row, h['발행상태']).setValue('미준비');

    ma_log_(
      'THREAD_REVIEW_PREP',
      contentId,
      threadId,
      'SUCCESS',
      'INFO',
      'Thread 검수본 준비 완료: ' + threadId,
      Date.now() - start,
      'menu'
    );

    ss.toast(
      'I열 최종본문을 확인·수정한 뒤 "선택 Thread 최종본 확정"을 실행해 주세요.',
      'Marketing Automation',
      8
    );

    return threadId;

  } catch (err) {
    try {
      ma_log_(
        'THREAD_REVIEW_PREP',
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


function ma_finalizeSelectedThread() {
  var start = Date.now();
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    if (sheet.getName() !== MA_CFG.SHEETS.THREADS) {
      throw new Error('Threads초안 시트에서 실행해 주세요.');
    }

    var row = sheet.getActiveRange().getRow();
    if (row < 2) {
      throw new Error('헤더가 아닌 Thread 행을 선택해 주세요.');
    }

    var h = ma_headerMap_(sheet);
    var v = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

    var threadId = ma_threadReviewText_(v[h['Thread ID'] - 1]);
    var contentId = ma_threadReviewText_(v[h['Content ID'] - 1]);
    var finalText = ma_threadReviewText_(v[h['최종본문'] - 1]);
    var reviewStatus = ma_threadReviewText_(v[h['검수상태'] - 1]);

    if (!threadId) throw new Error('Thread ID가 없습니다.');
    if (!finalText) throw new Error('최종본문(I열)이 없습니다.');
    if (reviewStatus === '검수완료') {
      throw new Error('이미 검수완료된 Thread입니다: ' + threadId);
    }
    if (reviewStatus !== '검수중') {
      throw new Error('먼저 "선택 Thread 검수본 준비"를 실행해 주세요.');
    }

    sheet.getRange(row, h['검수상태']).setValue('검수완료');
    sheet.getRange(row, h['승인일']).setValue(new Date());
    sheet.getRange(row, h['발행상태']).setValue('미준비');

    ma_log_(
      'THREAD_FINALIZE',
      contentId,
      threadId,
      'SUCCESS',
      'INFO',
      'Thread 최종본 확정 완료: ' + threadId,
      Date.now() - start,
      'menu'
    );

    ss.toast(
      'Thread 최종본 확정 완료. 발행대기 생성 대상이 되었습니다.',
      'Marketing Automation',
      7
    );

    return threadId;

  } catch (err) {
    try {
      ma_log_(
        'THREAD_FINALIZE',
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


function ma_threadReviewText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}
