/**
 * Marketing Automation V1
 * Threads generation service
 *
 * Source of truth:
 * - Only finalized blog content (블로그초안 R: 최종본문)
 * - Final title (G)
 *
 * Generates exactly 3 variants:
 * A 정보형
 * B 공감형
 * C 질문·참여형
 *
 * Safety:
 * - No generation before blog 검수완료 + 최종본문 존재
 * - No duplicate Thread ID/content rows
 * - Existing thread rows are never overwritten
 * - I 최종본문 is intentionally left blank at generation time
 */

function ma_generateSelectedThreads() {
  var start = Date.now();
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var blogSheet = ss.getActiveSheet();

    if (blogSheet.getName() !== MA_CFG.SHEETS.BLOG) {
      throw new Error('블로그초안 시트에서 실행해 주세요.');
    }

    var row = blogSheet.getActiveRange().getRow();
    if (row < 2) {
      throw new Error('헤더가 아닌 블로그 행을 선택해 주세요.');
    }

    var h = ma_headerMap_(blogSheet);
    var v = blogSheet.getRange(row, 1, 1, blogSheet.getLastColumn()).getDisplayValues()[0];

    var blogId = ma_thrText_(v[h['Blog ID'] - 1]);
    var contentId = ma_thrText_(v[h['Content ID'] - 1]);
    var finalTitle = ma_thrText_(v[h['최종제목'] - 1]);
    var finalBody = ma_thrText_(v[h['최종본문'] - 1]);
    var reviewStatus = ma_thrText_(v[h['검수상태'] - 1]);
    var cta = ma_thrText_(v[h['CTA'] - 1]);

    if (!blogId || !contentId) {
      throw new Error('Blog ID 또는 Content ID가 없습니다.');
    }
    if (reviewStatus !== '검수완료') {
      throw new Error('블로그 검수상태가 검수완료가 아닙니다.');
    }
    if (!finalBody) {
      throw new Error('최종본문(R열)이 없습니다.');
    }
    if (!finalTitle) {
      throw new Error('최종제목(G열)이 없습니다.');
    }

    var threadSheet = ss.getSheetByName(MA_CFG.SHEETS.THREADS);
    if (!threadSheet) {
      throw new Error('Threads초안 시트를 찾을 수 없습니다.');
    }

    var existing = ma_findExistingThreads_(threadSheet, contentId, blogId);
    if (existing.length > 0) {
      throw new Error(
        '이미 생성된 Threads 초안이 있습니다: ' +
        existing.map(function(x){ return x.threadId; }).join(', ')
      );
    }

    var content = ma_thrLoadContent_(ss, contentId);
    var settings = ma_thrLoadSettings_(ss);

    var result = ma_thrCallOpenAI_({
      contentId: contentId,
      blogId: blogId,
      finalTitle: finalTitle,
      finalBody: finalBody,
      cta: cta,
      content: content,
      settings: settings
    });

    var variants = [
      { suffix: 'A', type: '정보형', data: result.info },
      { suffix: 'B', type: '공감형', data: result.empathy },
      { suffix: 'C', type: '질문·참여형', data: result.engagement }
    ];

    var created = [];
    var now = new Date();

    variants.forEach(function(item) {
      var threadId = ma_buildThreadId_(contentId, item.suffix);

      if (ma_threadIdExists_(threadSheet, threadId)) {
        throw new Error('중복 Thread ID가 이미 존재합니다: ' + threadId);
      }

      threadSheet.appendRow([
        threadId,                 // A Thread ID
        contentId,                // B Content ID
        blogId,                   // C Blog ID
        now,                      // D 생성일
        item.type,                // E Thread 유형
        item.data.hook,           // F Hook
        item.data.body,           // G 본문
        item.data.cta,            // H CTA
        '',                       // I 최종본문 - 검수 준비 전에는 비움
        '미검수',                 // J 검수상태
        '',                       // K 승인일
        '미준비',                 // L 발행상태
        '',                       // M 발행예정일
        '',                       // N 발행일
        '',                       // O 발행URL
        ''                        // P 관리자메모
      ]);

      created.push(threadId);
    });

    ma_log_(
      'THREADS_GENERATE',
      contentId,
      blogId,
      'SUCCESS',
      'INFO',
      'Threads 초안 3종 생성 완료: ' + created.join(', '),
      Date.now() - start,
      'menu'
    );

    ss.toast(
      'Threads 3종 생성 완료. 각 Thread를 검수본 준비 후 확인해 주세요.',
      'Marketing Automation',
      8
    );

    return created;

  } catch (err) {
    try {
      ma_log_(
        'THREADS_GENERATE',
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


function ma_thrCallOpenAI_(ctx) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = String(props.getProperty('OPENAI_API_KEY') || '').trim();

  if (!apiKey) {
    throw new Error('Script Properties에 OPENAI_API_KEY가 없습니다.');
  }

  var s = ctx.settings || {};
  var model = s.AI_MODEL || 'gpt-5.6-terra';
  var brandName = s.BRAND_NAME || '자연담은멍케이크';
  var tone = s.BLOG_TONE || '친근하지만 전문적인 공방 운영자';
  var factPolicy = s.AI_FACT_POLICY ||
    '제공된 사실만 사용하고 실제 사례가 없으면 만들지 않는다.';

  var instructions = [
    '너는 Threads용 짧은 콘텐츠 편집자다.',
    '블로그 최종본문을 바탕으로 서로 다른 목적의 Threads 글 3개를 만든다.',
    '브랜드: ' + brandName,
    '기본 톤: ' + tone,
    '사실성 원칙: ' + factPolicy,
    '',
    '중요 규칙:',
    '1. 블로그에 없는 실제 수업 사례, 고객 반응, 후기, 효과, 재료, 수치를 새로 만들지 않는다.',
    '2. 네이버 블로그 글을 그대로 축약하거나 문장을 복사하지 않는다.',
    '3. 정보형은 핵심 정보 전달 중심.',
    '4. 공감형은 보호자의 상황과 감정에 자연스럽게 공감하되 과장하지 않는다.',
    '5. 질문·참여형은 댓글이나 반응을 유도하는 질문으로 마무리하되 억지 참여 유도는 피한다.',
    '6. 각 글은 Hook + 본문 + CTA 구조.',
    '7. 해시태그 남발 금지. 필요하면 본문 말미에 0~3개 수준만 사용.',
    '8. 지역 키워드는 자연스럽게 사용하고 반복하지 않는다.',
    '9. 각 variant는 서로 문장 구조와 접근법이 분명히 달라야 한다.',
    '10. 최종 길이는 Threads에서 읽기 편한 짧은 분량으로 작성한다.'
  ].join('\n');

  var c = ctx.content || {};
  var input = [
    '[콘텐츠 정보]',
    'Content ID: ' + ctx.contentId,
    'Blog ID: ' + ctx.blogId,
    '최종제목: ' + ctx.finalTitle,
    'Pillar: ' + (c.pillar || ''),
    '콘텐츠유형: ' + (c.contentType || ''),
    '주제: ' + (c.topic || ''),
    '핵심키워드: ' + (c.primaryKeyword || ''),
    '보조키워드: ' + (c.secondaryKeywords || ''),
    '대상독자: ' + (c.audience || ''),
    'CTA: ' + (c.cta || ctx.cta || ''),
    '사실확인메모: ' + (c.factNotes || ''),
    '',
    '[블로그 최종본문]',
    ctx.finalBody
  ].join('\n');

  var payload = {
    model: model,
    instructions: instructions,
    input: input,
    reasoning: {
      effort: s.AI_REASONING_EFFORT || 'low'
    },
    max_output_tokens: 2200,
    text: {
      format: {
        type: 'json_schema',
        name: 'threads_variants',
        strict: true,
        schema: {
          type: 'object',
          additionalProperties: false,
          properties: {
            info: ma_thrVariantSchema_(),
            empathy: ma_thrVariantSchema_(),
            engagement: ma_thrVariantSchema_()
          },
          required: ['info', 'empathy', 'engagement']
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
  var outputText = ma_thrExtractOutputText_(json);
  if (!outputText) {
    throw new Error('OpenAI 응답에서 output_text를 찾지 못했습니다.');
  }

  var result;
  try {
    result = JSON.parse(outputText);
  } catch (err) {
    throw new Error('Threads JSON 파싱 실패: ' + outputText.substring(0, 300));
  }

  ['info', 'empathy', 'engagement'].forEach(function(key) {
    if (!result[key] ||
        !ma_thrText_(result[key].hook) ||
        !ma_thrText_(result[key].body) ||
        !ma_thrText_(result[key].cta)) {
      throw new Error('Threads 결과가 불완전합니다: ' + key);
    }
  });

  return result;
}


function ma_thrVariantSchema_() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      hook: { type: 'string' },
      body: { type: 'string' },
      cta: { type: 'string' }
    },
    required: ['hook', 'body', 'cta']
  };
}


function ma_thrExtractOutputText_(json) {
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


function ma_thrLoadContent_(ss, contentId) {
  var sheet = ss.getSheetByName(MA_CFG.SHEETS.CONTENT);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var h = ma_headerMap_(sheet);
  var ids = sheet
    .getRange(2, h['Content ID'], sheet.getLastRow() - 1, 1)
    .getDisplayValues();

  for (var i = 0; i < ids.length; i++) {
    if (ma_thrText_(ids[i][0]) === ma_thrText_(contentId)) {
      var row = i + 2;
      var v = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

      function get(name) {
        return h[name] ? ma_thrText_(v[h[name] - 1]) : '';
      }

      return {
        pillar: get('Pillar'),
        contentType: get('콘텐츠유형'),
        topic: get('주제'),
        primaryKeyword: get('핵심키워드'),
        secondaryKeywords: get('보조키워드'),
        audience: get('대상독자'),
        cta: get('CTA'),
        factNotes: get('사실확인메모')
      };
    }
  }

  return {};
}


function ma_thrLoadSettings_(ss) {
  var sheet = ss.getSheetByName(MA_CFG.SHEETS.SETTINGS);
  var out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
  values.forEach(function(r) {
    var key = ma_thrText_(r[0]);
    if (key) out[key] = ma_thrText_(r[1]);
  });

  return out;
}


function ma_findExistingThreads_(sheet, contentId, blogId) {
  var out = [];
  if (sheet.getLastRow() < 2) return out;

  var h = ma_headerMap_(sheet);
  var values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getDisplayValues();

  values.forEach(function(r) {
    var threadId = ma_thrText_(r[h['Thread ID'] - 1]);
    var cId = ma_thrText_(r[h['Content ID'] - 1]);
    var bId = ma_thrText_(r[h['Blog ID'] - 1]);

    if (threadId && cId === contentId && bId === blogId) {
      out.push({ threadId: threadId });
    }
  });

  return out;
}


function ma_threadIdExists_(sheet, threadId) {
  if (sheet.getLastRow() < 2) return false;

  var h = ma_headerMap_(sheet);
  var ids = sheet
    .getRange(2, h['Thread ID'], sheet.getLastRow() - 1, 1)
    .getDisplayValues();

  for (var i = 0; i < ids.length; i++) {
    if (ma_thrText_(ids[i][0]) === threadId) return true;
  }

  return false;
}


function ma_buildThreadId_(contentId, suffix) {
  var m = String(contentId || '').match(/^CNT-(\d{4})-(\d{4})$/);
  if (!m) {
    throw new Error('Content ID 형식이 올바르지 않습니다: ' + contentId);
  }
  return 'THR-' + m[1] + '-' + m[2] + '-' + suffix;
}


function ma_thrComposeFinal_(hook, body, cta) {
  return [hook, body, cta]
    .map(function(x) { return ma_thrText_(x); })
    .filter(function(x) { return !!x; })
    .join('\n\n');
}


function ma_thrText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}
