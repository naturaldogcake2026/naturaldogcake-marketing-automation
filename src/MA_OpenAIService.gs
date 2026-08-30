/**
 * Marketing Automation V1.1
 * OpenAI Responses API 연동
 *
 * 보안 원칙:
 * - OPENAI_API_KEY는 Script Properties에만 저장한다.
 * - 시트/로그/GitHub에 API 키를 기록하지 않는다.
 */

function ma_generateSelectedBlogDraft() {
  var start = Date.now();
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 자동화 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var contentSheet = ss.getActiveSheet();

    if (contentSheet.getName() !== MA_CFG.SHEETS.CONTENT) {
      throw new Error('콘텐츠DB 시트에서 실행해 주세요.');
    }

    var row = contentSheet.getActiveRange().getRow();
    if (row < 2) {
      throw new Error('헤더가 아닌 콘텐츠 행을 선택해 주세요.');
    }

    // Blog scaffold가 없으면 먼저 생성한다.
    var scaffold = ma_createBlogDraftForContentRow_(contentSheet, row);
    var contentId = scaffold.contentId;
    var blogId = scaffold.blogId;

    var blogSheet = ss.getSheetByName(MA_CFG.SHEETS.BLOG);
    var ch = ma_headerMap_(contentSheet);
    var bh = ma_headerMap_(blogSheet);
    var blogRow = ma_findRowByExactValue_(blogSheet, bh['Blog ID'], blogId);

    if (!blogRow) {
      throw new Error('생성된 Blog ID의 블로그초안 행을 찾을 수 없습니다: ' + blogId);
    }

    // 기존 AI 산출물이 있으면 실수로 덮어쓰지 않는다.
    var existingTitle = String(blogSheet.getRange(blogRow, bh['제목1']).getValue() || '').trim();
    var existingBody = String(blogSheet.getRange(blogRow, bh['본문']).getValue() || '').trim();

    if (existingTitle || existingBody) {
      throw new Error(
        '이미 AI 초안이 존재합니다. 기존 초안을 보호하기 위해 덮어쓰지 않았습니다: ' + blogId
      );
    }

    var contentValues = contentSheet
      .getRange(row, 1, 1, contentSheet.getLastColumn())
      .getDisplayValues()[0];

    var settings = ma_getSettings_();
    var requestData = ma_buildBlogAiInput_(contentValues, ch, settings);
    var draft = ma_callOpenAiForBlogDraft_(requestData, settings);

    ma_writeBlogAiDraft_(blogSheet, blogRow, bh, draft, requestData);

    // 실제 AI 본문이 생성된 뒤에만 검수대기로 이동한다.
    contentSheet.getRange(row, ch['콘텐츠상태']).setValue('검수대기');
    contentSheet.getRange(row, ch['최종수정일']).setValue(new Date());

    ma_log_(
      'BLOG_AI_GENERATE',
      contentId,
      blogId,
      'SUCCESS',
      'INFO',
      'AI 블로그 초안 생성 완료: ' + blogId,
      Date.now() - start,
      'menu'
    );

    ss.toast('AI 블로그 초안 생성 완료: ' + blogId, 'Marketing Automation', 7);

    return {
      ok: true,
      contentId: contentId,
      blogId: blogId
    };

  } catch (err) {
    var message = String(err && err.message ? err.message : err);

    try {
      ma_log_(
        'BLOG_AI_GENERATE',
        '',
        '',
        'FAIL',
        'ERROR',
        message,
        Date.now() - start,
        'menu'
      );
    } catch (ignore) {}

    throw err;

  } finally {
    lock.releaseLock();
  }
}

function ma_buildBlogAiInput_(rowValues, h, settings) {
  return {
    brandName: settings.BRAND_NAME || '자연담은멍케이크',
    regionPrimary: settings.REGION_PRIMARY || '경산',
    regionSecondary: settings.REGION_SECONDARY || '대구',
    positioning: settings.BLOG_POSITIONING || '',
    tone: settings.BLOG_TONE || '친근하지만 전문적인 공방 운영자',
    minLength: Number(settings.BLOG_MIN_LENGTH || 1500),
    factPolicy: settings.AI_FACT_POLICY ||
      '제공된 사실만 사용하고 실제 사례가 없으면 만들지 않는다',

    pillar: ma_text_(ma_valueByHeader_(rowValues, h, 'Pillar')),
    contentType: ma_text_(ma_valueByHeader_(rowValues, h, '콘텐츠유형')),
    topic: ma_text_(ma_valueByHeader_(rowValues, h, '주제')),
    primaryKeyword: ma_text_(ma_valueByHeader_(rowValues, h, '핵심키워드')),
    secondaryKeywords: ma_text_(ma_valueByHeader_(rowValues, h, '보조키워드')),
    audience: ma_text_(ma_valueByHeader_(rowValues, h, '대상독자')),
    searchIntent: ma_text_(ma_valueByHeader_(rowValues, h, '검색의도')),
    contentGoal: ma_text_(ma_valueByHeader_(rowValues, h, '콘텐츠목적')),
    cta: ma_text_(ma_valueByHeader_(rowValues, h, 'CTA')),
    references: ma_text_(ma_valueByHeader_(rowValues, h, '참고자료')),
    factNotes: ma_text_(ma_valueByHeader_(rowValues, h, '사실확인메모')),

    knownBusinessFacts: {
      classSnackPrice: settings.CLASS_SNACK_PRICE || '',
      classCakePrice: settings.CLASS_CAKE_PRICE || '',
      classDuration: settings.CLASS_DURATION || '',
      classCapacity: settings.CLASS_CAPACITY || '',
      classReservation: settings.CLASS_RESERVATION || ''
    }
  };
}

function ma_callOpenAiForBlogDraft_(inputData, settings) {
  var apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    throw new Error(
      'OPENAI_API_KEY가 설정되지 않았습니다. Apps Script 프로젝트 설정의 스크립트 속성에 먼저 등록해 주세요.'
    );
  }

  var model = String(settings.AI_MODEL || 'gpt-5.6-terra').trim();
  var maxTokens = Number(settings.AI_MAX_OUTPUT_TOKENS || 5000);
  var reasoningEffort = String(settings.AI_REASONING_EFFORT || 'low').trim();

  var systemInstruction = [
    '당신은 한국 네이버 블로그용 로컬 비즈니스 콘텐츠 에디터다.',
    '브랜드 포지셔닝과 검색 의도를 반영하되 키워드를 억지로 반복하지 않는다.',
    '제공되지 않은 실제 고객 사례, 출강 실적, 후기, 수치, 인증, 효능을 절대 만들어내지 않는다.',
    '제공된 사실과 factNotes를 최우선으로 따른다.',
    '본문은 한국어로 자연스럽고 읽기 쉽게 작성한다.',
    '과장광고, 허위 후기, 존재하지 않는 사례를 사용하지 않는다.',
    '네이버 블로그에서 읽기 편하도록 짧은 문단과 자연스러운 소제목을 사용한다.'
  ].join('\n');

  var userInstruction = [
    '아래 입력 데이터를 바탕으로 블로그 초안을 작성하세요.',
    '',
    JSON.stringify(inputData, null, 2),
    '',
    '작성 규칙:',
    '1. 제목 후보를 서로 다른 검색 의도로 3개 제안한다.',
    '2. 핵심키워드는 제목과 초반부에 자연스럽게 포함한다.',
    '3. 도입부는 독자의 상황이나 궁금증에서 시작한다.',
    '4. 본문은 최소 ' + inputData.minLength + '자 수준을 목표로 충분히 작성한다.',
    '5. 본문에 자연스러운 소제목을 넣되 마크다운 표는 쓰지 않는다.',
    '6. CTA는 입력된 CTA와 모순되지 않게 글 흐름에 자연스럽게 연결한다.',
    '7. 해시태그는 지역·서비스·검색의도를 반영해 8~15개를 한 문자열로 작성한다.',
    '8. 이미지요청은 실제 촬영 가능한 장면 중심으로 3~6개를 한 문자열로 작성한다.',
    '9. 실제 사례가 없는 경우 사례가 있는 것처럼 표현하지 않는다.',
    '10. 결과는 지정된 JSON 스키마에만 맞춰 반환한다.'
  ].join('\n');

  var payload = {
    model: model,
    instructions: systemInstruction,
    input: userInstruction,
    reasoning: {
      effort: reasoningEffort
    },
    max_output_tokens: maxTokens,
    text: {
      format: {
        type: 'json_schema',
        name: 'naver_blog_draft',
        strict: true,
        schema: {
          type: 'object',
          additionalProperties: false,
          properties: {
            title1: { type: 'string' },
            title2: { type: 'string' },
            title3: { type: 'string' },
            intro: { type: 'string' },
            body: { type: 'string' },
            hashtags: { type: 'string' },
            imageRequests: { type: 'string' }
          },
          required: [
            'title1',
            'title2',
            'title3',
            'intro',
            'body',
            'hashtags',
            'imageRequests'
          ]
        }
      }
    }
  };

  var response = UrlFetchApp.fetch(
    'https://api.openai.com/v1/responses',
    {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Bearer ' + apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  var status = response.getResponseCode();
  var raw = response.getContentText();

  if (status < 200 || status >= 300) {
    var apiMessage = ma_extractOpenAiError_(raw);
    throw new Error('OpenAI API 오류 (' + status + '): ' + apiMessage);
  }

  var data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    throw new Error('OpenAI 응답 JSON 해석 실패');
  }

  var outputText = ma_extractOpenAiOutputText_(data);

  if (!outputText) {
    throw new Error('OpenAI 응답에 출력 텍스트가 없습니다.');
  }

  var draft;
  try {
    draft = JSON.parse(outputText);
  } catch (e2) {
    throw new Error('OpenAI 구조화 출력 JSON 해석 실패');
  }

  ma_validateBlogDraft_(draft, inputData.minLength);
  return draft;
}

function ma_writeBlogAiDraft_(sheet, row, h, draft, requestData) {
  // AI 원본 컬럼만 기록한다. 관리자수정본(Q), 최종본문(R)은 건드리지 않는다.
  sheet.getRange(row, h['제목1']).setValue(draft.title1);
  sheet.getRange(row, h['제목2']).setValue(draft.title2);
  sheet.getRange(row, h['제목3']).setValue(draft.title3);
  sheet.getRange(row, h['핵심키워드']).setValue(requestData.primaryKeyword);
  sheet.getRange(row, h['도입부']).setValue(draft.intro);
  sheet.getRange(row, h['본문']).setValue(draft.body);
  sheet.getRange(row, h['CTA']).setValue(requestData.cta);
  sheet.getRange(row, h['해시태그']).setValue(draft.hashtags);
  sheet.getRange(row, h['이미지요청']).setValue(draft.imageRequests);

  sheet.getRange(row, h['사실검수']).setValue('미검수');
  sheet.getRange(row, h['문체검수']).setValue('미검수');
  sheet.getRange(row, h['SEO검수']).setValue('미검수');
  sheet.getRange(row, h['검수상태']).setValue('미검수');
  sheet.getRange(row, h['발행상태']).setValue('미준비');
}

function ma_validateBlogDraft_(draft, minLength) {
  var required = [
    'title1', 'title2', 'title3',
    'intro', 'body', 'hashtags', 'imageRequests'
  ];

  required.forEach(function(key) {
    if (!draft || typeof draft[key] !== 'string' || !draft[key].trim()) {
      throw new Error('AI 초안 필수 필드 누락: ' + key);
    }
  });

  // 너무 짧은 출력은 저장하지 않는다.
  var conservativeMin = Math.max(700, Math.floor(Number(minLength || 1500) * 0.6));
  if (draft.body.trim().length < conservativeMin) {
    throw new Error(
      'AI 본문이 너무 짧아 저장하지 않았습니다. length=' +
      draft.body.trim().length +
      ', required>=' + conservativeMin
    );
  }
}

function ma_extractOpenAiOutputText_(data) {
  if (data && typeof data.output_text === 'string' && data.output_text) {
    return data.output_text;
  }

  var output = data && Array.isArray(data.output) ? data.output : [];

  for (var i = 0; i < output.length; i++) {
    var content = Array.isArray(output[i].content) ? output[i].content : [];
    for (var j = 0; j < content.length; j++) {
      if (
        content[j] &&
        content[j].type === 'output_text' &&
        typeof content[j].text === 'string'
      ) {
        return content[j].text;
      }
    }
  }

  return '';
}

function ma_extractOpenAiError_(raw) {
  try {
    var parsed = JSON.parse(raw);
    if (parsed && parsed.error && parsed.error.message) {
      return String(parsed.error.message);
    }
  } catch (e) {}

  return String(raw || '알 수 없는 오류').substring(0, 1000);
}

function ma_findRowByExactValue_(sheet, col, value) {
  if (!value || sheet.getLastRow() < 2) return 0;

  var values = sheet
    .getRange(2, col, sheet.getLastRow() - 1, 1)
    .getDisplayValues();

  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === String(value).trim()) {
      return i + 2;
    }
  }

  return 0;
}

function ma_text_(value) {
  return String(value == null ? '' : value).trim();
}
