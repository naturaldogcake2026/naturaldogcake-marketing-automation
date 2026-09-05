/**
 * Marketing Automation V1.6.1
 * 로드맵 우선 AI 콘텐츠 기획
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
    if (row < 2) throw new Error('새 콘텐츠를 만들 빈 행을 선택해 주세요.');

    var h = ma_headerMap_(sheet);
    var planningHeaders = [
      'Pillar','콘텐츠유형','주제','핵심키워드','보조키워드',
      '대상독자','검색의도','콘텐츠목적','CTA','참고자료','사실확인메모'
    ];
    planningHeaders.forEach(function(name) {
      if (!h[name]) throw new Error('콘텐츠DB 필수 헤더를 찾을 수 없습니다: ' + name);
    });

    contentId = String(h['Content ID'] ? sheet.getRange(row,h['Content ID']).getValue() : '').trim();

    var existing = planningHeaders.map(function(name) {
      return String(sheet.getRange(row,h[name]).getDisplayValue() || '').trim();
    });
    if (existing.some(function(v){ return v !== ''; })) {
      throw new Error('선택 행의 기획 영역(C:M)에 이미 내용이 있습니다. 기존 내용을 보호하기 위해 덮어쓰지 않았습니다.');
    }

    var settings = ma_getSettings_();
    var history = ma_planExistingContentHistory_(sheet,h,row);
    var input = ma_planBuildInput_(settings,history);
    var plan = ma_planCallOpenAi_(input,settings);

    var valuesByHeader = {
      'Pillar':plan.pillar, '콘텐츠유형':plan.contentType, '주제':plan.topic,
      '핵심키워드':plan.primaryKeyword, '보조키워드':plan.secondaryKeywords,
      '대상독자':plan.audience, '검색의도':plan.searchIntent,
      '콘텐츠목적':plan.contentGoal, 'CTA':plan.cta,
      '참고자료':plan.references, '사실확인메모':plan.factNotes
    };
    planningHeaders.forEach(function(name){
      sheet.getRange(row,h[name]).setValue(valuesByHeader[name] || '');
    });

    ma_v161MarkRoadmapPlanned_(input.nextRoadmapItem, plan.topic, contentId);

    ma_log_('CONTENT_AI_PLAN_V161',contentId,'','SUCCESS','INFO',
      'V1.6.1 로드맵 우선 기획 완료: ' + plan.topic,Date.now()-start,'menu');

    ss.toast('V1.6.1 콘텐츠 기획 완료. 로드맵 우선순위를 반영했습니다.',
      'Marketing Automation',8);
    return {ok:true,row:row,topic:plan.topic};
  } catch(err) {
    var message=String(err && err.message ? err.message : err);
    try { ma_log_('CONTENT_AI_PLAN_V161',contentId,'','FAIL','ERROR',message,Date.now()-start,'menu'); }
    catch(ignore){}
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function ma_planExistingContentHistory_(sheet,h,selectedRow) {
  var lastRow=sheet.getLastRow();
  if(lastRow<2) return [];
  var values=sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getDisplayValues();
  var result=[];
  for(var i=0;i<values.length;i++){
    if(i+2===selectedRow) continue;
    var topic=ma_planValue_(values[i],h,'주제');
    if(!topic) continue;
    result.push({
      contentId:ma_planValue_(values[i],h,'Content ID'),
      pillar:ma_planValue_(values[i],h,'Pillar'),
      contentType:ma_planValue_(values[i],h,'콘텐츠유형'),
      topic:topic,
      primaryKeyword:ma_planValue_(values[i],h,'핵심키워드'),
      searchIntent:ma_planValue_(values[i],h,'검색의도'),
      contentGoal:ma_planValue_(values[i],h,'콘텐츠목적'),
      status:ma_planValue_(values[i],h,'콘텐츠상태')
    });
  }
  return result.slice(-80);
}

function ma_planBuildInput_(settings,history) {
  var pillars={};
  for(var i=1;i<=8;i++){
    var key='P'+String(i).padStart(2,'0');
    if(settings[key]) pillars[key]=String(settings[key]);
  }

  return {
    brandName:settings.BRAND_NAME || '자연담은멍케이크',
    regionPrimary:settings.REGION_PRIMARY || '경산',
    regionSecondary:settings.REGION_SECONDARY || '대구',
    positioning:settings.BLOG_POSITIONING || '',
    tone:settings.BLOG_TONE || '친근하지만 전문적인 공방 운영자',
    factPolicy:settings.AI_FACT_POLICY || '제공된 사실만 사용하고 실제 사례가 없으면 만들지 않는다',
    pillars:pillars,
    knownBusinessFacts:ma_getBusinessFacts_(settings),
    existingContents:history,
    actualBlogReferences:(typeof ma_v16GetBlogReferenceHistory_==='function') ? ma_v16GetBlogReferenceHistory_() : [],
    blogStrategy:(typeof ma_v16GetBlogStrategy_==='function') ? ma_v16GetBlogStrategy_() : {},
    roadmap:(typeof ma_v161GetRoadmap_==='function') ? ma_v161GetRoadmap_() : [],
    nextRoadmapItem:(typeof ma_v161GetNextRoadmapItem_==='function') ? ma_v161GetNextRoadmapItem_() : null
  };
}

function ma_planCallOpenAi_(inputData,settings) {
  var apiKey=PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if(!apiKey) throw new Error('OPENAI_API_KEY가 설정되지 않았습니다. Apps Script 프로젝트 설정의 스크립트 속성을 확인해 주세요.');

  var model=String(settings.AI_MODEL || 'gpt-5.6-terra').trim();
  var reasoningEffort=String(settings.AI_REASONING_EFFORT || 'low').trim();

  var instructions=[
    '당신은 자연담은멍케이크의 네이버 블로그 콘텐츠 기획자다.',
    '가장 중요한 규칙은 blogStrategy의 roadmap_first 원칙이다.',
    'nextRoadmapItem이 있으면 그것을 이번 콘텐츠의 기본 주제로 사용한다.',
    '로드맵 제목을 기계적으로 복사할 필요는 없지만 목적과 핵심 주제는 바꾸지 않는다.',
    'actualBlogReferences와 existingContents를 읽고 이미 다룬 내용과 겹치는 부분은 차별화한다.',
    'nextRoadmapItem이 null일 때만 전체 전략과 이력을 바탕으로 새로운 주제를 제안한다.',
    '자료대기 또는 실제자료 필요=예인 로드맵 항목은 실제 자료 확보 전 자동 선택하지 않는다.',
    '경산·대구 검색어는 실제 공방/제품/제작/클래스 내용에 자연스럽게 결합한다.',
    '제공되지 않은 고객 사례, 후기, 출강 실적, 수치, 효능, 인증, 고객 선호, 주문 절차를 만들지 않는다.',
    '기관·단체 프로그램은 실제 출강 사례가 아니라 제안 가능한 프로그램 안내로만 표현한다.',
    'Pillar는 입력된 pillars 중 하나만 선택한다.',
    'references는 실제 제공된 참고자료가 없으면 빈 문자열이다.',
    '결과는 지정 JSON 스키마만 반환한다.'
  ].join('\n');

  var prompt=[
    '다음 데이터를 사용해 이번 블로그 콘텐츠 1편을 기획하세요.',
    '',
    JSON.stringify(inputData,null,2),
    '',
    '우선 판단:',
    '- nextRoadmapItem이 있으면 해당 항목을 최우선으로 기획',
    '- 기존 발행/예약 글과 동일한 질문을 반복하지 말 것',
    '- 로드맵의 목적을 유지하면서 검색 유입과 실제 공방 신뢰를 함께 고려',
    '',
    '필드 기준:',
    '- contentType: 제작이야기, 제품이야기, 공방이야기, 클래스이야기, 프로그램안내, 정보형, FAQ 등',
    '- topic: 한 편의 글로 바로 발전 가능한 구체적 주제',
    '- primaryKeyword: 핵심 검색어 1개',
    '- secondaryKeywords: 2~4개 쉼표 구분',
    '- audience: 핵심 독자',
    '- searchIntent: 간결한 검색의도',
    '- contentGoal: 핵심 목적 1개',
    '- cta: 자연스러운 다음 행동',
    '- references: 없으면 빈 문자열',
    '- factNotes: 사실 경계 및 미확인 정보 금지 메모'
  ].join('\n');

  var payload={
    model:model,
    instructions:instructions,
    input:prompt,
    reasoning:{effort:reasoningEffort},
    max_output_tokens:1800,
    text:{format:{
      type:'json_schema',name:'content_plan',strict:true,
      schema:{
        type:'object',additionalProperties:false,
        properties:{
          pillar:{type:'string'},contentType:{type:'string'},topic:{type:'string'},
          primaryKeyword:{type:'string'},secondaryKeywords:{type:'string'},
          audience:{type:'string'},searchIntent:{type:'string'},contentGoal:{type:'string'},
          cta:{type:'string'},references:{type:'string'},factNotes:{type:'string'}
        },
        required:['pillar','contentType','topic','primaryKeyword','secondaryKeywords',
          'audience','searchIntent','contentGoal','cta','references','factNotes']
      }
    }}
  };

  var response=UrlFetchApp.fetch('https://api.openai.com/v1/responses',{
    method:'post',contentType:'application/json',
    headers:{Authorization:'Bearer '+apiKey},
    payload:JSON.stringify(payload),muteHttpExceptions:true
  });
  var status=response.getResponseCode();
  var raw=response.getContentText();
  if(status<200 || status>=300) throw new Error('OpenAI API 오류 ('+status+'): '+ma_extractOpenAiError_(raw));

  var data;
  try{data=JSON.parse(raw);}catch(e){throw new Error('OpenAI 응답 JSON 해석 실패');}
  var outputText=ma_extractOpenAiOutputText_(data);
  if(!outputText) throw new Error('OpenAI 응답에 출력 텍스트가 없습니다.');

  var plan;
  try{plan=JSON.parse(outputText);}catch(e2){throw new Error('OpenAI 콘텐츠 기획 JSON 해석 실패');}
  ma_planValidate_(plan,inputData.pillars);
  return plan;
}

function ma_v161MarkRoadmapPlanned_(item,topic,contentId) {
  if(!item || !item.roadmapId) return;
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName(MA_V161_BLOG_ROADMAP_SHEET);
  if(!sheet || sheet.getLastRow()<2) return;

  var values=sheet.getRange(2,1,sheet.getLastRow()-1,7).getDisplayValues();
  for(var i=0;i<values.length;i++){
    if(String(values[i][1]||'').trim()===item.roadmapId){
      var row=i+2;
      if(String(sheet.getRange(row,5).getDisplayValue()||'').trim()==='기획대기'){
        sheet.getRange(row,5).setValue('기획완료');
        var old=String(sheet.getRange(row,7).getDisplayValue()||'').trim();
        var note='자동기획: '+(contentId || 'ID미부여')+' / '+topic;
        sheet.getRange(row,7).setValue(old ? old+' | '+note : note);
      }
      return;
    }
  }
}

function ma_planValidate_(plan,pillars) {
  var required=['pillar','contentType','topic','primaryKeyword','secondaryKeywords',
    'audience','searchIntent','contentGoal','cta','references','factNotes'];
  required.forEach(function(key){
    if(!plan || typeof plan[key]!=='string') throw new Error('AI 콘텐츠 기획 필드 누락: '+key);
  });
  var pillarKeys=Object.keys(pillars || {});
  if(pillarKeys.length && pillarKeys.indexOf(plan.pillar.trim())<0)
    throw new Error('AI가 설정에 없는 Pillar를 반환했습니다: '+plan.pillar);
  ['contentType','topic','primaryKeyword','audience','searchIntent','contentGoal','cta','factNotes']
    .forEach(function(key){
      if(!plan[key].trim()) throw new Error('AI 콘텐츠 기획 필수값이 비어 있습니다: '+key);
    });
}

function ma_planValue_(rowValues,h,header) {
  var col=h[header];
  if(!col) return '';
  return String(rowValues[col-1] || '').trim();
}
