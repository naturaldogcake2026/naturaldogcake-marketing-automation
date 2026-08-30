/**
 * Marketing Automation V1
 * AI review / rewrite for Naver blog administrator draft.
 *
 * Safety:
 * - Never modifies original AI draft columns I/J.
 * - Never modifies final body R.
 * - Refuses to overwrite Q when Q has already been manually edited.
 * - Only runs by explicit menu action.
 */

function ma_aiReviewSelectedBlog() {
  var start = Date.now();
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    if (sheet.getName() !== MA_CFG.SHEETS.BLOG) {
      throw new Error('블로그초안 시트에서 실행해 주세요.');
    }

    var row = sheet.getActiveRange().getRow();
    if (row < 2) {
      throw new Error('헤더가 아닌 블로그 초안 행을 선택해 주세요.');
    }

    var h = ma_headerMap_(sheet);
    var v = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    var blogId = ma_aiReviewText_(v[h['Blog ID'] - 1]);
    var contentId = ma_aiReviewText_(v[h['Content ID'] - 1]);
    var title1 = ma_aiReviewText_(v[h['제목1'] - 1]);
    var title2 = ma_aiReviewText_(v[h['제목2'] - 1]);
    var title3 = ma_aiReviewText_(v[h['제목3'] - 1]);
    var currentFinalTitle = ma_aiReviewText_(v[h['최종제목'] - 1]);
    var keyword = ma_aiReviewText_(v[h['핵심키워드'] - 1]);
    var intro = ma_aiReviewText_(v[h['도입부'] - 1]);
    var body = ma_aiReviewText_(v[h['본문'] - 1]);
    var cta = ma_aiReviewText_(v[h['CTA'] - 1]);
    var hashtags = ma_aiReviewText_(v[h['해시태그'] - 1]);
    var adminDraft = ma_aiReviewText_(v[h['관리자수정본'] - 1]);
    var finalBody = ma_aiReviewText_(v[h['최종본문'] - 1]);

    if (!blogId) throw new Error('Blog ID가 없습니다.');
    if (finalBody) throw new Error('이미 최종본문이 확정된 블로그입니다: ' + blogId);
    if (!adminDraft) {
      throw new Error('관리자수정본(Q열)이 없습니다. 먼저 "선택 블로그 검수본 준비"를 실행해 주세요.');
    }

    // Human-edit protection:
    // Q must still equal the automatically prepared I+J draft.
    var originalPreparedDraft = [intro, body]
      .filter(function(x) { return !!String(x || '').trim(); })
      .join('\n\n')
      .trim();

    if (adminDraft.trim() !== originalPreparedDraft) {
      throw new Error(
        '관리자수정본(Q열)이 이미 수정된 것으로 보입니다. 사람의 수정 내용을 보호하기 위해 AI 검수로 덮어쓰지 않았습니다.'
      );
    }

    var content = ma_aiReviewLoadContent_(ss, contentId);
    var settings = ma_aiReviewLoadSettings_(ss);

    var result = ma_aiReviewCallOpenAI_({
      blogId: blogId,
      contentId: contentId,
      title1: title1,
      title2: title2,
      title3: title3,
      currentFinalTitle: currentFinalTitle,
      keyword: keyword,
      adminDraft: adminDraft,
      cta: cta,
      hashtags: hashtags,
      content: content,
      settings: settings
    });

    if (!result || !result.finalTitle || !result.revisedBody) {
      throw new Error('AI 검수 결과가 올바르지 않습니다.');
    }

    sheet.getRange(row, h['최종제목']).setValue(String(result.finalTitle).trim());
    sheet.getRange(row, h['관리자수정본']).setValue(String(result.revisedBody).trim());

    // AI review is done, but human review is still required.
    sheet.getRange(row, h['사실검수']).setValue('검수중');
    sheet.getRange(row, h['문체검수']).setValue('검수중');
    sheet.getRange(row, h['SEO검수']).setValue('검수중');
    sheet.getRange(row, h['검수상태']).setValue('검수중');

    ma_log_(
      'BLOG_AI_REVIEW',
      contentId,
      blogId,
      'SUCCESS',
      'INFO',
      'AI 1차 검수·교정 완료. 사람 최종 확인 필요',
      Date.now() - start,
      'menu'
    );

    ss.toast(
      'AI 검수·교정 완료. G열 제목과 Q열 본문을 직접 확인한 뒤 최종본 확정을 실행해 주세요.',
      'Marketing Automation',
      8
    );

    return result;

  } catch (err) {
    try {
      ma_log_(
        'BLOG_AI_REVIEW',
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


function ma_aiReviewLoadContent_(ss, contentId) {
  var sheet = ss.getSheetByName(MA_CFG.SHEETS.CONTENT);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var h = ma_headerMap_(sheet);
  var ids = sheet
    .getRange(2, h['Content ID'], sheet.getLastRow() - 1, 1)
    .getDisplayValues();

  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === String(contentId || '').trim()) {
      var row = i + 2;
      var v = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

      function get(name) {
        return h[name] ? String(v[h[name] - 1] || '').trim() : '';
      }

      return {
        pillar: get('Pillar'),
        contentType: get('콘텐츠유형'),
        topic: get('주제'),
        primaryKeyword: get('핵심키워드'),
        secondaryKeywords: get('보조키워드'),
        audience: get('대상독자'),
        searchIntent: get('검색의도'),
        contentGoal: get('콘텐츠목적'),
        cta: get('CTA'),
        references: get('참고자료'),
        factNotes: get('사실확인메모')
      };
    }
  }

  return {};
}


function ma_aiReviewLoadSettings_(ss) {
  var sheet = ss.getSheetByName(MA_CFG.SHEETS.SETTINGS);
  var out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
  values.forEach(function(r) {
    var key = String(r[0] || '').trim();
    if (key) out[key] = String(r[1] || '').trim();
  });
  return out;
}


function ma_aiReviewCallOpenAI_(ctx) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = String(props.getProperty('OPENAI_API_KEY') || '').trim();

  if (!apiKey) {
    throw new Error('Script Properties에 OPENAI_API_KEY가 없습니다.');
  }

  var s = ctx.settings || {};
  var model = s.AI_MODEL || 'gpt-5.6-terra';
  var factPolicy = s.AI_FACT_POLICY ||
    '제공된 사실만 사용하고 실제 사례가 없으면 만들지 않는다.';
  var brandName = s.BRAND_NAME || '자연담은멍케이크';
  var positioning = s.BLOG_POSITIONING || '경산·대구 반려견 케이크 전문 공방';
  var tone = s.BLOG_TONE || '친근하지만 전문적인 공방 운영자';

  var instructions = [
    '너는 네이버 블로그 편집자다.',
    '목표는 검색 노출을 위해 억지로 키워드를 반복하는 글이 아니라 사람이 자연스럽게 읽는 지역 기반 전문 공방 글을 만드는 것이다.',
    '브랜드: ' + brandName,
    '포지셔닝: ' + positioning,
    '문체: ' + tone,
    '사실성 원칙: ' + factPolicy,
    '',
    '교정 원칙:',
    '1. 제공되지 않은 실제 수업 사례, 고객 반응, 기관 출강 사례, 효과, 후기, 재료, 운영 방식은 절대 만들어내지 않는다.',
    '2. 가격·시간·정원 같은 정보는 필요한 위치에서만 제시하고 반복을 줄인다.',
    '3. 지역 키워드는 자연스럽게 사용하고 문단마다 억지로 반복하지 않는다.',
    '4. "대구에서 찾는다면" 같은 검색엔진용 티가 나는 문장을 줄이고 사람 중심 문장으로 고친다.',
    '5. 공방 운영자가 직접 설명하는 듯 자연스럽고 신뢰감 있게 쓴다.',
    '6. 원문에 있는 유효한 사실과 CTA는 유지한다.',
    '7. 최종 본문은 네이버 블로그에서 읽기 쉽게 짧은 문단과 소제목을 사용한다.',
    '8. 과장 표현, 검증되지 않은 안전/건강 효능 표현은 넣지 않는다.',
    '9. 제목은 핵심키워드를 포함하되 과도한 키워드 나열을 피한다.',
    '10. 결과에는 제목 1개와 교정된 본문만 반환한다.'
  ].join('\n');

  var c = ctx.content || {};
  var input = [
    '[콘텐츠 기획]',
    'Pillar: ' + (c.pillar || ''),
    '콘텐츠유형: ' + (c.contentType || ''),
    '주제: ' + (c.topic || ''),
    '핵심키워드: ' + (c.primaryKeyword || ctx.keyword || ''),
    '보조키워드: ' + (c.secondaryKeywords || ''),
    '대상독자: ' + (c.audience || ''),
    '검색의도: ' + (c.searchIntent || ''),
    '콘텐츠목적: ' + (c.contentGoal || ''),
    'CTA: ' + (c.cta || ctx.cta || ''),
    '참고자료: ' + (c.references || ''),
    '사실확인메모: ' + (c.factNotes || ''),
    '',
    '[제목 후보]',
    '1: ' + ctx.title1,
    '2: ' + ctx.title2,
    '3: ' + ctx.title3,
    '현재 최종제목: ' + ctx.currentFinalTitle,
    '',
    '[현재 관리자수정본]',
    ctx.adminDraft
  ].join('\n');

  var payload = {
    model: model,
    instructions: instructions,
    input: input,
    reasoning: {
      effort: s.AI_REASONING_EFFORT || 'low'
    },
    max_output_tokens: Math.max(
      2500,
      Number(s.AI_MAX_OUTPUT_TOKENS || 5000)
    ),
    text: {
      format: {
        type: 'json_schema',
        name: 'naver_blog_review',
        strict: true,
        schema: {
          type: 'object',
          additionalProperties: false,
          properties: {
            finalTitle: { type: 'string' },
            revisedBody: { type: 'string' }
          },
          required: ['finalTitle', 'revisedBody']
        }
      }
    }
  };

  var response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var raw = response.getContentText();

  if (code < 200 || code >= 300) {
    var message = raw;
    try {
      var errJson = JSON.parse(raw);
      if (errJson && errJson.error && errJson.error.message) {
        message = errJson.error.message;
      }
    } catch (ignore) {}

    throw new Error('OpenAI API 오류 (' + code + '): ' + message);
  }

  var json = JSON.parse(raw);
  var outputText = ma_aiReviewExtractOutputText_(json);

  if (!outputText) {
    throw new Error('OpenAI 응답에서 output_text를 찾지 못했습니다.');
  }

  var result;
  try {
    result = JSON.parse(outputText);
  } catch (err) {
    throw new Error('AI 검수 JSON 파싱 실패: ' + outputText.substring(0, 300));
  }

  if (!result.finalTitle || !result.revisedBody) {
    throw new Error('AI 검수 결과에 finalTitle 또는 revisedBody가 없습니다.');
  }

  if (String(result.revisedBody).trim().length < 700) {
    throw new Error('AI 검수 본문이 지나치게 짧아 저장하지 않았습니다.');
  }

  return result;
}


function ma_aiReviewExtractOutputText_(json) {
  if (!json) return '';

  if (typeof json.output_text === 'string' && json.output_text.trim()) {
    return json.output_text.trim();
  }

  var output = json.output || [];
  for (var i = 0; i < output.length; i++) {
    var content = output[i] && output[i].content ? output[i].content : [];
    for (var j = 0; j < content.length; j++) {
      var item = content[j] || {};
      if (item.type === 'output_text' && typeof item.text === 'string') {
        return item.text.trim();
      }
    }
  }

  return '';
}


function ma_aiReviewText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}
