/**
 * Blog scaffold service
 * AI 본문 생성 전 단계:
 * - Blog ID 생성
 * - 블로그초안 행 생성
 * - 콘텐츠DB Blog ID 연결
 * - 콘텐츠상태 = 초안생성
 *
 * 실제 AI 본문이 생성되기 전에는 검수대기로 변경하지 않는다.
 */

function ma_createSelectedBlogDraft() {
  var start = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contentSheet = ss.getActiveSheet();

  if (contentSheet.getName() !== MA_CFG.SHEETS.CONTENT) {
    throw new Error('콘텐츠DB 시트에서 실행해 주세요.');
  }

  var row = contentSheet.getActiveRange().getRow();
  if (row < 2) {
    throw new Error('헤더 행이 아닌 콘텐츠 행을 선택해 주세요.');
  }

  var result = ma_createBlogDraftForContentRow_(contentSheet, row);

  ma_log_(
    'BLOG_SCAFFOLD',
    result.contentId,
    result.blogId,
    result.created ? 'SUCCESS' : 'SKIP',
    'INFO',
    result.message,
    Date.now() - start,
    'menu'
  );

  ss.toast(result.message, 'Marketing Automation', 6);
  return result;
}

function ma_createBlogDraftForContentRow_(contentSheet, row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var blogSheet = ss.getSheetByName(MA_CFG.SHEETS.BLOG);

  if (!blogSheet) {
    throw new Error('블로그초안 시트를 찾을 수 없습니다.');
  }

  // 먼저 Content ID 및 기본값을 보장한다.
  var contentId = ma_applyContentDefaultsForRow_(contentSheet, row);
  if (!contentId) {
    throw new Error('주제, 핵심키워드 또는 Pillar 중 하나 이상을 입력해 주세요.');
  }

  var ch = ma_headerMap_(contentSheet);
  var bh = ma_headerMap_(blogSheet);
  var contentRow = contentSheet
    .getRange(row, 1, 1, contentSheet.getLastColumn())
    .getValues()[0];

  var existingLinkedBlogId = String(
    contentSheet.getRange(row, ch['블로그ID']).getValue() || ''
  ).trim();

  // 1) 이미 콘텐츠DB에 Blog ID가 연결되어 있고 실제 초안 행도 존재하면 중복 생성하지 않는다.
  if (existingLinkedBlogId && ma_blogIdExists_(blogSheet, bh, existingLinkedBlogId)) {
    return {
      created: false,
      contentId: contentId,
      blogId: existingLinkedBlogId,
      message: '이미 연결된 블로그 초안이 있습니다: ' + existingLinkedBlogId
    };
  }

  // 2) 과거 부분 실행으로 콘텐츠DB 연결만 빠진 경우, Content ID 기준으로 기존 Blog 행을 복구 연결한다.
  var existingByContent = ma_findBlogByContentId_(blogSheet, bh, contentId);
  if (existingByContent) {
    contentSheet.getRange(row, ch['블로그ID']).setValue(existingByContent);
    contentSheet.getRange(row, ch['콘텐츠상태']).setValue('초안생성');
    contentSheet.getRange(row, ch['최종수정일']).setValue(new Date());

    return {
      created: false,
      contentId: contentId,
      blogId: existingByContent,
      message: '기존 블로그 초안을 다시 연결했습니다: ' + existingByContent
    };
  }

  var settings = ma_getSettings_();
  var blogPrefix = settings.ID_BLOG_PREFIX || 'BLOG';
  var blogId = ma_nextId_(blogSheet, bh['Blog ID'], blogPrefix, 4);

  var keyword = String(ma_valueByHeader_(contentRow, ch, '핵심키워드') || '').trim();
  var cta = String(ma_valueByHeader_(contentRow, ch, 'CTA') || '').trim();

  // 블로그초안 23열 순서에 맞춘 scaffold.
  // 제목/도입부/본문 등 AI 산출물은 의도적으로 비워 둔다.
  var values = [
    blogId,          // A Blog ID
    contentId,       // B Content ID
    new Date(),      // C 생성일
    '',              // D 제목1
    '',              // E 제목2
    '',              // F 제목3
    '',              // G 최종제목
    keyword,         // H 핵심키워드
    '',              // I 도입부
    '',              // J 본문
    cta,             // K CTA
    '',              // L 해시태그
    '',              // M 이미지요청
    '미검수',         // N 사실검수
    '미검수',         // O 문체검수
    '미검수',         // P SEO검수
    '',              // Q 관리자수정본
    '',              // R 최종본문
    '미검수',         // S 검수상태
    '',              // T 승인일
    '미준비',         // U 발행상태
    '',              // V 발행일
    ''               // W 발행URL
  ];

  var targetRow = ma_firstEmptyRowById_(blogSheet, bh['Blog ID']);
  blogSheet.getRange(targetRow, 1, 1, values.length).setValues([values]);

  // 블로그초안 생성이 성공한 뒤에만 부모 콘텐츠를 연결/상태 변경한다.
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

  var values = sheet
    .getRange(2, h['Blog ID'], sheet.getLastRow() - 1, 1)
    .getDisplayValues();

  return values.some(function(r) {
    return String(r[0] || '').trim() === blogId;
  });
}

function ma_findBlogByContentId_(sheet, h, contentId) {
  if (!contentId || sheet.getLastRow() < 2) return '';

  var numRows = sheet.getLastRow() - 1;
  var ids = sheet
    .getRange(2, h['Content ID'], numRows, 1)
    .getDisplayValues();
  var blogIds = sheet
    .getRange(2, h['Blog ID'], numRows, 1)
    .getDisplayValues();

  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === contentId) {
      return String(blogIds[i][0] || '').trim();
    }
  }
  return '';
}

function ma_firstEmptyRowById_(sheet, idCol) {
  var lastRow = Math.max(sheet.getLastRow(), 1);

  if (lastRow < 2) return 2;

  var values = sheet
    .getRange(2, idCol, Math.max(lastRow - 1, 1), 1)
    .getDisplayValues();

  for (var i = 0; i < values.length; i++) {
    if (!String(values[i][0] || '').trim()) {
      return i + 2;
    }
  }

  return lastRow + 1;
}
