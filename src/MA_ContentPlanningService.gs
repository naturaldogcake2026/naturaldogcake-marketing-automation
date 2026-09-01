/**
 * Marketing Automation V1
 * AI 새 콘텐츠 기획
 *
 * 원칙:
 * - 콘텐츠DB에서 선택한 빈 행의 C:M(기획 영역)만 작성한다.
 * - Content ID, 상태, 날짜 등 시스템 관리 컬럼은 건드리지 않는다.
 * - 기존 콘텐츠 주제를 참고해 중복을 피한다.
 * - AI_FACT_POLICY를 적용하고 실제 사례/후기/효과를 만들지 않는다.
 */
function ma_planSelectedContentWithAi() {
  var start = Date.now();
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 자동화 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  var contentId = '';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

    if (sheet.getName() !== MA_CFG.SHEETS.CONTENT) {
      throw new Error('콘텐츠DB 시트에서 실행해 주세요.');
    }

    var row = sheet.getActiveRange().getRow();
    if (row < 2) {
      throw new Error('새 콘텐츠를 만들 빈 행을 선택해 주세요.');
    }

    var h = ma_headerMap_(sheet);
    var planningHeaders = [
      'Pillar', '콘텐츠유형', '주제', '핵심키워드', '보조키워드',
      '대상독자', '검색의도', '콘텐츠목적', 'CTA', '참고자료', '사실확인메모'
    ];

    planningHeaders.forEach(function(name) {
      if (!h[name]) throw new Error('콘텐츠DB 필수 헤더를 찾을 수 없습니다: ' + name);
    });

    contentId = String(
      h['Content ID'] ? sheet.getRange(row, h['Content ID']).getValue() : ''
    ).trim();

    // 사람 입력 또는 기존 AI 기획을 보호한다.
    var existing = planningHeaders.map(function(name) {
      return String(sheet.getRange(row, h[name]).getDisplayValue() || '').trim();
    });
    if (existing.some(function(v) { return v !== ''; })) {
      throw new Error(
        '선택 행의 기획 영역(C:M)에 이미 내용이 있습니다. 기존 내용을 보호하기 위해 덮어쓰지 않았습니다.'
      );
    }

    var settings = ma_getSettings_();
    var history = ma_planExistingContentHistory_(sheet, h, row);
    var input = ma_planBuildInput_(settings, history);
    var plan = ma_planCallOpenAi_(input, settings);

    // C:M에 해당하는 기획 컬럼만 기록한다.
    var valuesByHeader = {
      'Pillar': plan.pillar,
      '콘텐츠유형': plan.contentType,
      '주제': plan.topic,
      '핵심키워드': plan.primaryKeyword,
      '보조키워드': plan.secondaryKeywords,
      '대상독자': plan.audience,
      '검색의도': plan.searchIntent,
      '콘텐츠목적': plan.contentGoal,
      'CTA': plan.cta,
      '참고자료': plan.references,
      '사실확인메모': plan.factNotes
    };

    planningHeaders.forEach(function(name) {
      sheet.getRange(row, h[name]).setValue(valuesByHeader[name] || '');
    });

    ma_log_(
      'CONTENT_AI_PLAN',
      contentId,
      '',
      'SUCCESS',
      'INFO',
      'AI 새 콘텐츠 기획 완료: ' + plan.topic,
      Date.now() - start,
      'menu'
    );

    ss.toast(
      'AI 콘텐츠 기획 완료. 노란색 기획 칸을 확인한 뒤 1-1을 실행해 주세요.',
      'Marketing Automation',
      8
    );

    return { ok: true, row: row, topic: plan.topic };

  } catch (err) {
    var message = String(err && err.message ? err.message : err);
    try {
      ma_log_(
        'CONTENT_AI_PLAN',
        contentId,
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

function ma_planExistingContentHistory_(sheet, h, selectedRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getDisplayValues();

  var result = [];
  for (var i = 0; i < values.length; i++) {
    var rowNumber = i + 2;
    if (rowNumber === selectedRow) continue;

    var topic = ma_planValue_(values[i], h, '주제');
    if (!topic) continue;

    result.push({
      contentId: ma_planValue_(values[i], h, 'Content ID'),
      pillar: ma_planValue_(values[i], h, 'Pillar'),
      topic: topic,
      primaryKeyword: ma_planValue_(values[i], h, '핵심키워드'),
      contentGoal: ma_planValue_(values[i], h, '콘텐츠목적'),
      status: ma_planValue_(values[i], h, '콘텐츠상태')
    });
  }

  // 최근/현재 데이터가 많아져도 프롬프트가 과도하게 커지지 않게 제한한다.
  return result.slice(-80);
}

function ma_planBuildInput_(settings, history) {
  var pillars = {};
  for (var i = 1; i <= 8; i++) {
    var key = 'P' + String(i).padStart(2, '0');
    if (settings[key]) pillars[key] = String(settings[key]);
  }

  return {
    brandName: settings.BRAND_NAME || '자연담은멍케이크',
    regionPrimary: settings.REGION_PRIMARY || '경산',
    regionSecondary: settings.REGION_SECONDARY || '대구',
    positioning: settings.BLOG_POSITIONING || '',
    tone: settings.BLOG_TONE || '친근하지만 전문적인 공방 운영자',
    factPolicy: settings.AI_FACT_POLICY ||
      '제공된 사실만 사용하고 실제 사례가 없으면 만들지 않는다',
    pillars: pillars,
    threadCount: Number(settings.THREAD_COUNT || 3),
    knownBusinessFacts: ma_getBusinessFacts_(settings),
    existingContents: history
  };
}

function ma_planCallOpenAi_(inputData, settings) {
  var apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    throw new Error(
      'OPENAI_API_KEY가 설정되지 않았습니다. Apps Script 프로젝트 설정의 스크립트 속성을 확인해 주세요.'
    );
  }

  var model = String(settings.AI_MODEL || 'gpt-5.6-terra').trim();
  var reasoningEffort = String(settings.AI_REASONING_EFFORT || 'low').trim();

  var instructions = [
    '당신은 자연담은멍케이크의 콘텐츠 기획자다.',
    '목표는 경산·대구 지역에서 반려견 케이크 전문 공방으로 검색되고 신뢰를 쌓는 것이다.',
    'existingContents와 실질적으로 중복되는 주제는 제안하지 않는다.',
    '판매 글만 반복하지 말고 검색정보, 공방 전문성, 클래스, 지역검색, 브랜드 신뢰를 균형 있게 고려한다.',
    'Pillar는 입력된 pillars 중 가장 적합한 코드 하나를 선택한다.',
    '제공되지 않은 실제 고객 사례, 후기, 출강 실적, 수치, 효능, 인증을 만들어내지 않는다.',
    '사실확인메모에는 글 작성 시 지켜야 할 사실 경계나 확인할 운영정보를 간결하게 적는다.',
    '참고자료가 실제로 제공되지 않았다면 references는 빈 문자열로 반환한다.',
    '결과는 지정된 JSON 스키마에만 맞춰 반환한다.'
  ].join('\n');

  var prompt = [
    '다음 브랜드 정보와 기존 콘텐츠 목록을 보고 지금 만들 가치가 가장 높은 새 콘텐츠 1개를 기획하세요.',
    '',
    JSON.stringify(inputData, null, 2),
    '',
    '필드 작성 기준:',
    '- contentType: 정보형, 공방이야기, 클래스안내, 프로그램안내, FAQ 등 주제에 적합한 짧은 유형명',
    '- topic: 실제 게시물 한 편으로 바로 발전시킬 수 있는 구체적인 주제',
    '- primaryKeyword: 핵심 검색어 1개',
    '- secondaryKeywords: 관련 검색어 2~4개를 쉼표로 구분',
    '- audience: 이 글을 읽을 핵심 독자',
    '- searchIntent: 정보탐색, 구매전정보탐색, 지역검색 등 간결한 표현',
    '- contentGoal: 검색유입, 주문문의, 클래스예약, 출강문의, 브랜드신뢰 등 핵심 목적 1개',
    '- cta: 실제 다음 행동을 자연스럽게 유도',
    '- references: 실제 제공된 참고자료가 없으면 반드시 빈 문자열',
    '- factNotes: 허위 사례를 만들지 않도록 필요한 사실성 메모'
  ].join('\n');

  var payload = {
    model: model,
    instructions: instructions,
    input: prompt,
    reasoning: { effort: reasoningEffort },
    max_output_tokens: 1800,
    text: {
      format: {
        type: 'json_schema',
        name: 'content_plan',
        strict: true,
        schema: {
          type: 'object',
          additionalProperties: false,
          properties: {
            pillar: { type: 'string' },
            contentType: { type: 'string' },
            topic: { type: 'string' },
            primaryKeyword: { type: 'string' },
            secondaryKeywords: { type: 'string' },
            audience: { type: 'string' },
            searchIntent: { type: 'string' },
            contentGoal: { type: 'string' },
            cta: { type: 'string' },
            references: { type: 'string' },
            factNotes: { type: 'string' }
          },
          required: [
            'pillar', 'contentType', 'topic', 'primaryKeyword',
            'secondaryKeywords', 'audience', 'searchIntent',
            'contentGoal', 'cta', 'references', 'factNotes'
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
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  var status = response.getResponseCode();
  var raw = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(
      'OpenAI API 오류 (' + status + '): ' + ma_extractOpenAiError_(raw)
    );
  }

  var data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    throw new Error('OpenAI 응답 JSON 해석 실패');
  }

  var outputText = ma_extractOpenAiOutputText_(data);
  if (!outputText) throw new Error('OpenAI 응답에 출력 텍스트가 없습니다.');

  var plan;
  try {
    plan = JSON.parse(outputText);
  } catch (e2) {
    throw new Error('OpenAI 콘텐츠 기획 JSON 해석 실패');
  }

  ma_planValidate_(plan, inputData.pillars);
  return plan;
}

function ma_planValidate_(plan, pillars) {
  var required = [
    'pillar', 'contentType', 'topic', 'primaryKeyword',
    'secondaryKeywords', 'audience', 'searchIntent',
    'contentGoal', 'cta', 'references', 'factNotes'
  ];

  required.forEach(function(key) {
    if (!plan || typeof plan[key] !== 'string') {
      throw new Error('AI 콘텐츠 기획 필드 누락: ' + key);
    }
  });

  var pillarKeys = Object.keys(pillars || {});
  if (pillarKeys.length && pillarKeys.indexOf(plan.pillar.trim()) < 0) {
    throw new Error('AI가 설정에 없는 Pillar를 반환했습니다: ' + plan.pillar);
  }

  ['contentType', 'topic', 'primaryKeyword', 'audience',
   'searchIntent', 'contentGoal', 'cta', 'factNotes'].forEach(function(key) {
    if (!plan[key].trim()) {
      throw new Error('AI 콘텐츠 기획 필수값이 비어 있습니다: ' + key);
    }
  });
}

function ma_planValue_(rowValues, h, header) {
  var col = h[header];
  if (!col) return '';
  return String(rowValues[col - 1] || '').trim();
}
