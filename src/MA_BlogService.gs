function ma_createSelectedBlogDraft() {
  var start = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  if (sheet.getName() !== MA_CFG.SHEETS.CONTENT) {
    throw new Error('콘텐츠DB 시트에서 실행해 주세요.');
  }

  var row = sheet.getActiveRange().getRow();
  if (row < 2) {
    throw new Error('헤더가 아닌 콘텐츠 행을 선택해 주세요.');
  }

  try {
    var result = ma_createBlogDraftForContentRow_(sheet, row);

    ma_log_(
      'BLOG_SCAFFOLD',
      result.contentId || '',
      result.blogId || '',
      result.created ? 'SUCCESS' : 'SKIP',
      'INFO',
      result.message || '',
      Date.now() - start,
      'menu'
    );

    ss.toast(result.message || '블로그 초안 처리가 완료되었습니다.', 'Marketing Automation', 6);
    return result;
  } catch (err) {
    var message = String(err && err.message ? err.message : err);
    try {
      ma_log_('BLOG_SCAFFOLD','','','FAIL','ERROR',message,Date.now()-start,'menu');
    } catch (ignore) {}
    throw err;
  }
}

function ma_createBlogDraftForContentRow_(contentSheet, row) {
  var ss = contentSheet.getParent();
  var blogSheet = ss.getSheetByName(MA_CFG.SHEETS.BLOG);
  if (!blogSheet) throw new Error('블로그초안 시트를 찾을 수 없습니다.');

  var contentId = ma_applyContentDefaultsForRow_(contentSheet, row);
  var ch = ma_headerMap_(contentSheet);
  var bh = ma_headerMap_(blogSheet);

  var contentValues = contentSheet
    .getRange(row, 1, 1, contentSheet.getLastColumn())
    .getValues()[0];

  var linkedBlogId = String(
    ma_valueByHeader_(contentValues, ch, '블로그ID') || ''
  ).trim();

  if (linkedBlogId && ma_blogIdExists_(blogSheet, bh, linkedBlogId)) {
    return {
      created: false,
      contentId: contentId,
      blogId: linkedBlogId,
      message: '이미 연결된 블로그 초안이 있습니다: ' + linkedBlogId
    };
  }

  var existing = ma_findBlogByContentId_(blogSheet, bh, contentId);
  if (existing) {
    contentSheet.getRange(row, ch['블로그ID']).setValue(existing.blogId);
    contentSheet.getRange(row, ch['콘텐츠상태']).setValue('초안생성');
    contentSheet.getRange(row, ch['최종수정일']).setValue(new Date());
    return {
      created: false,
      contentId: contentId,
      blogId: existing.blogId,
      message: '기존 블로그 초안 연결을 복구했습니다: ' + existing.blogId
    };
  }

  var settings = ma_getSettings_();
  var blogPrefix = String(settings.ID_BLOG_PREFIX || 'BLOG').trim();
  var blogId = ma_nextId_(blogSheet, bh['Blog ID'], blogPrefix, 4);

  var keyword = String(
    ma_valueByHeader_(contentValues, ch, '핵심키워드') || ''
  ).trim();

  var cta = String(
    ma_valueByHeader_(contentValues, ch, 'CTA') || ''
  ).trim();

  var targetRow = ma_firstEmptyRowById_(blogSheet, bh['Blog ID']);

  blogSheet.getRange(targetRow, 1, 1, 23).setValues([[
    blogId, contentId, new Date(),
    '', '', '', '',
    keyword,
    '', '',
    cta,
    '', '',
    '미검수','미검수','미검수',
    '', '',
    '미검수',
    '',
    '미준비',
    '', ''
  ]]);

  contentSheet.getRange(row, ch['블로그ID']).setValue(blogId);
  contentSheet.getRange(row, ch['콘텐츠상태']).setValue('초안생성');
  contentSheet.getRange(row, ch['최종수정일']).setValue(new Date());

  return {
    created: true,
    contentId: contentId,
    blogId: blogId,
    message: '블로그 초안 행 생성 완료: ' + blogId
  };
}

function ma_blogIdExists_(sheet, h, blogId) {
  if (!blogId || sheet.getLastRow() < 2) return false;
  var col = h['Blog ID'];
  var values = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === String(blogId).trim()) return true;
  }
  return false;
}

function ma_findBlogByContentId_(sheet, h, contentId) {
  if (!contentId || sheet.getLastRow() < 2) return null;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][h['Content ID'] - 1] || '').trim() === String(contentId).trim()) {
      return {
        row: i + 2,
        blogId: String(values[i][h['Blog ID'] - 1] || '').trim()
      };
    }
  }
  return null;
}

function ma_firstEmptyRowById_(sheet, idCol) {
  var lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastRow < 2) return 2;
  var values = sheet.getRange(2, idCol, lastRow - 1, 1).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (!String(values[i][0] || '').trim()) return i + 2;
  }
  return lastRow + 1;
}
