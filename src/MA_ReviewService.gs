function ma_prepareSelectedBlogReview() {
  var start = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if (sheet.getName() !== MA_CFG.SHEETS.BLOG) throw new Error('블로그초안 시트에서 실행해 주세요.');

  var row = sheet.getActiveRange().getRow();
  if (row < 2) throw new Error('헤더가 아닌 블로그 초안 행을 선택해 주세요.');

  var h = ma_headerMap_(sheet);
  var v = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var blogId = ma_reviewText_(v[h['Blog ID'] - 1]);
  var contentId = ma_reviewText_(v[h['Content ID'] - 1]);
  var title1 = ma_reviewText_(v[h['제목1'] - 1]);
  var finalTitle = ma_reviewText_(v[h['최종제목'] - 1]);
  var intro = ma_reviewText_(v[h['도입부'] - 1]);
  var body = ma_reviewText_(v[h['본문'] - 1]);
  var adminDraft = ma_reviewText_(v[h['관리자수정본'] - 1]);
  var finalBody = ma_reviewText_(v[h['최종본문'] - 1]);

  if (!blogId) throw new Error('Blog ID가 없습니다.');
  if (finalBody) throw new Error('이미 최종본문이 확정된 블로그입니다: ' + blogId);

  if (adminDraft) {
    ma_log_('BLOG_REVIEW_PREP', contentId, blogId, 'SKIP', 'INFO',
      '관리자수정본이 이미 있어 덮어쓰지 않음', Date.now() - start, 'menu');
    ss.toast('관리자수정본이 이미 있습니다. 기존 내용을 덮어쓰지 않았습니다.', 'Marketing Automation', 6);
    return;
  }

  if (!intro && !body) throw new Error('AI 초안이 없습니다. 먼저 AI 블로그 작성을 실행해 주세요.');

  var reviewBody = [intro, body].filter(function(x) {
    return !!String(x || '').trim();
  }).join('\n\n');

  if (!finalTitle && title1) sheet.getRange(row, h['최종제목']).setValue(title1);
  sheet.getRange(row, h['관리자수정본']).setValue(reviewBody);

  sheet.getRange(row, h['사실검수']).setValue('검수중');
  sheet.getRange(row, h['문체검수']).setValue('검수중');
  sheet.getRange(row, h['SEO검수']).setValue('검수중');
  sheet.getRange(row, h['검수상태']).setValue('검수중');

  ma_reviewUpdateContentStatus_(ss, contentId, '검수중');

  ma_log_('BLOG_REVIEW_PREP', contentId, blogId, 'SUCCESS', 'INFO',
    'AI 초안을 관리자수정본으로 복사 완료', Date.now() - start, 'menu');

  ss.toast('관리자수정본(Q열)에 복사했습니다. 내용을 확인·수정한 뒤 최종본 확정을 실행해 주세요.',
    'Marketing Automation', 8);
}

function ma_finalizeSelectedBlog() {
  var start = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  if (sheet.getName() !== MA_CFG.SHEETS.BLOG) throw new Error('블로그초안 시트에서 실행해 주세요.');

  var row = sheet.getActiveRange().getRow();
  if (row < 2) throw new Error('헤더가 아닌 블로그 초안 행을 선택해 주세요.');

  var h = ma_headerMap_(sheet);
  var v = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var blogId = ma_reviewText_(v[h['Blog ID'] - 1]);
  var contentId = ma_reviewText_(v[h['Content ID'] - 1]);
  var finalTitle = ma_reviewText_(v[h['최종제목'] - 1]);
  var adminDraft = ma_reviewText_(v[h['관리자수정본'] - 1]);
  var existingFinal = ma_reviewText_(v[h['최종본문'] - 1]);

  if (!blogId) throw new Error('Blog ID가 없습니다.');
  if (!finalTitle) throw new Error('최종제목(G열)을 먼저 확인해 주세요.');
  if (!adminDraft) throw new Error('관리자수정본(Q열)이 비어 있습니다. 먼저 검수본 준비를 실행해 주세요.');
  if (existingFinal) throw new Error('이미 최종본문이 확정되어 있습니다.');

  sheet.getRange(row, h['최종본문']).setValue(adminDraft);
  sheet.getRange(row, h['사실검수']).setValue('검수완료');
  sheet.getRange(row, h['문체검수']).setValue('검수완료');
  sheet.getRange(row, h['SEO검수']).setValue('검수완료');
  sheet.getRange(row, h['검수상태']).setValue('검수완료');
  sheet.getRange(row, h['승인일']).setValue(new Date());
  sheet.getRange(row, h['발행상태']).setValue('미준비');

  ma_reviewUpdateContentStatus_(ss, contentId, '검수완료');

  ma_log_('BLOG_FINALIZE', contentId, blogId, 'SUCCESS', 'INFO',
    '관리자수정본을 최종본문으로 확정 완료', Date.now() - start, 'menu');

  ss.toast('최종본문 확정 완료. 다음 단계는 발행대기 생성입니다.', 'Marketing Automation', 7);
}

function ma_reviewUpdateContentStatus_(ss, contentId, status) {
  if (!contentId) return;
  var sheet = ss.getSheetByName(MA_CFG.SHEETS.CONTENT);
  if (!sheet || sheet.getLastRow() < 2) return;

  var h = ma_headerMap_(sheet);
  var ids = sheet.getRange(2, h['Content ID'], sheet.getLastRow() - 1, 1).getDisplayValues();

  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === String(contentId).trim()) {
      var r = i + 2;
      sheet.getRange(r, h['콘텐츠상태']).setValue(status);
      sheet.getRange(r, h['최종수정일']).setValue(new Date());
      return;
    }
  }
}

function ma_reviewText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}
