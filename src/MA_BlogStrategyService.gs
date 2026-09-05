/**
 * Marketing Automation V1.6.1
 * 실제 블로그 이력 + 합의된 콘텐츠 로드맵
 */

var MA_V16_BLOG_STRATEGY_SHEET = '블로그전략';
var MA_V161_BLOG_ROADMAP_SHEET = '블로그로드맵';

var MA_V16_BLOG_STRATEGY_HEADERS = [
  '기준 ID', '글 제목', '상태', '콘텐츠분류', '발행일', '메모'
];

var MA_V16_SEED_BLOG_REFERENCES = [
  ['BLOGREF-0001', '경산 반려견 케이크 공방, 자연담은멍케이크를 소개합니다', '발행완료', '공방소개', '2026-08-25', '기존 네이버 블로그'],
  ['BLOGREF-0002', '경산 강아지 케이크, 사진 속 우리 아이를 전신케이크로 만드는 과정', '발행완료', '제작이야기', '2026-08-28', '기존 네이버 블로그'],
  ['BLOGREF-0003', '경산 반려견 케이크 공방, 자연담은멍케이크 공간을 소개합니다', '발행완료', '공간소개', '2026-09-01', '기존 네이버 블로그'],
  ['BLOGREF-0004', '경산 강아지 케이크 원데이클래스, 우리 아이를 직접 만들어봤어요', '발행완료', '원데이클래스', '2026-09-04', '기존 네이버 블로그'],
  ['BLOGREF-0005', '경산 강아지 케이크｜작지만 특별한 머핀케이크 이야기', '예약발행', '제품이야기', '', '예약발행 예정'],
  ['BLOGREF-0006', '경산 반려견 원데이클래스, 케이크 만들기가 처음이어도 괜찮을까요?', '예약발행', '원데이클래스', '', '예약발행 예정']
];

var MA_V161_ROADMAP_HEADERS = [
  '우선순위', '로드맵 ID', '추천 글', '목적', '상태', '실제자료 필요', '운영메모'
];

var MA_V161_ROADMAP_SEEDS = [
  [1, 'ROAD-0001', '강아지 생일케이크 주문 전 알아두면 좋은 것들', '실제 구매 검색 유입', '기획대기', '아니오', '다음 신규 글 우선 후보. 제공된 주문 사실만 사용'],
  [2, 'ROAD-0002', '반려견과 함께하는 특별한 체험, 자연담은멍케이크 원데이클래스', '클래스 자체를 정식 상품처럼 소개', '기획대기', '아니오', '기존 실제 클래스 글과 중복되지 않게 상품 구성·예약 관점으로 차별화'],
  [3, 'ROAD-0003', '기관·단체 반려견 프로그램은 어떻게 진행될까요?', '관공서·복지관·문화센터 담당자용', '기획대기', '아니오', '실제 출강 실적처럼 쓰지 말고 제안 가능한 프로그램 안내로 작성'],
  [4, 'ROAD-0004', '새로운 실제 케이크 제작 사례', '포트폴리오 축적', '자료대기', '예', '실제 새 케이크 사진·사례 확보 후에만 기획 가능. 자료 없으면 다음 순위로 건너뜀'],
  [5, 'ROAD-0005', '자연담은멍케이크에서는 왜 ‘우리 아이’를 표현할까요?', '브랜드 철학', '기획대기', '아니오', '사진 참고 제작과 우리 아이를 표현하는 브랜드 방향 중심']
];

function ma_installV16BlogStrategy() {
  return ma_installV161BlogRoadmap();
}

function ma_installV161BlogRoadmap() {
  var start = Date.now();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    throw new Error('다른 마케팅 자동화 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var refAdded = ma_v161EnsureSheetWithSeeds_(
      ss,
      MA_V16_BLOG_STRATEGY_SHEET,
      MA_V16_BLOG_STRATEGY_HEADERS,
      MA_V16_SEED_BLOG_REFERENCES,
      2
    );
    var roadmapAdded = ma_v161EnsureSheetWithSeeds_(
      ss,
      MA_V161_BLOG_ROADMAP_SHEET,
      MA_V161_ROADMAP_HEADERS,
      MA_V161_ROADMAP_SEEDS,
      2
    );

    var roadmapSheet = ss.getSheetByName(MA_V161_BLOG_ROADMAP_SHEET);
    var rows = Math.max(roadmapSheet.getMaxRows() - 1, 1);
    roadmapSheet.getRange(2, 3, rows, 5).setBackground('#fff2cc');

    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['기획대기', '자료대기', '기획완료', '보류', '완료'], true)
      .setAllowInvalid(true)
      .build();
    roadmapSheet.getRange(2, 5, rows, 1).setDataValidation(statusRule);

    roadmapSheet.setColumnWidth(1, 80);
    roadmapSheet.setColumnWidth(2, 110);
    roadmapSheet.setColumnWidth(3, 500);
    roadmapSheet.setColumnWidth(4, 240);
    roadmapSheet.setColumnWidth(5, 100);
    roadmapSheet.setColumnWidth(6, 100);
    roadmapSheet.setColumnWidth(7, 420);

    ma_log_(
      'BLOG_ROADMAP_V161_INSTALL', '', '', 'SUCCESS', 'INFO',
      'V1.6.1 적용: 블로그기준 신규=' + refAdded + '건, 로드맵 신규=' + roadmapAdded + '건',
      Date.now() - start, 'menu'
    );

    ss.toast(
      'V1.6.1 블로그 로드맵 적용 완료. 로드맵 신규 ' + roadmapAdded + '건',
      'Marketing Automation', 8
    );
    return { referenceAdded: refAdded, roadmapAdded: roadmapAdded };
  } finally {
    lock.releaseLock();
  }
}

function ma_v161EnsureSheetWithSeeds_(ss, sheetName, headers, seeds, uniqueCol) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  var current = headerRange.getDisplayValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var actual = String(current[i] || '').trim();
    if (actual && actual !== headers[i]) {
      throw new Error(sheetName + ' 헤더 불일치 ' + ma_v16ColumnLetter_(i + 1) +
        '1: [' + actual + '] / 예상 [' + headers[i] + ']');
    }
  }
  headerRange.setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  var existing = {};
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, uniqueCol, sheet.getLastRow() - 1, 1)
      .getDisplayValues().forEach(function(r) {
        var key = String(r[0] || '').trim();
        if (key) existing[key] = true;
      });
  }

  var add = [];
  seeds.forEach(function(seed) {
    var key = String(seed[uniqueCol - 1] || '').trim();
    if (!existing[key]) {
      add.push(seed);
      existing[key] = true;
    }
  });
  if (add.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, add.length, headers.length).setValues(add);
  }
  return add.length;
}

function ma_v16GetBlogReferenceHistory_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MA_V16_BLOG_STRATEGY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    return MA_V16_SEED_BLOG_REFERENCES.map(function(r) {
      return {referenceId:r[0], title:r[1], status:r[2], contentType:r[3], publishDate:r[4], note:r[5]};
    });
  }
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getDisplayValues()
    .filter(function(r){ return String(r[1] || '').trim(); })
    .map(function(r){
      return {
        referenceId:String(r[0]||'').trim(), title:String(r[1]||'').trim(),
        status:String(r[2]||'').trim(), contentType:String(r[3]||'').trim(),
        publishDate:String(r[4]||'').trim(), note:String(r[5]||'').trim()
      };
    }).slice(-100);
}

function ma_v161GetRoadmap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MA_V161_BLOG_ROADMAP_SHEET);
  var rows;

  if (!sheet || sheet.getLastRow() < 2) {
    rows = MA_V161_ROADMAP_SEEDS;
  } else {
    rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues();
  }

  return rows.filter(function(r) {
    return String(r[2] || '').trim();
  }).map(function(r) {
    return {
      priority:Number(r[0] || 999),
      roadmapId:String(r[1] || '').trim(),
      recommendedTopic:String(r[2] || '').trim(),
      purpose:String(r[3] || '').trim(),
      status:String(r[4] || '').trim(),
      requiresActualMaterial:String(r[5] || '').trim(),
      operationNote:String(r[6] || '').trim()
    };
  }).sort(function(a,b){ return a.priority - b.priority; });
}

function ma_v161GetNextRoadmapItem_() {
  var roadmap = ma_v161GetRoadmap_();
  for (var i = 0; i < roadmap.length; i++) {
    var item = roadmap[i];
    if (item.status !== '기획대기') continue;
    if (item.requiresActualMaterial === '예') continue;
    return item;
  }
  return null;
}

function ma_v16GetBlogStrategy_() {
  return {
    version:'V1.6.1',
    mode:'roadmap_first',
    currentStage:'초기 블로그 포트폴리오 구축 단계',
    targetMix:{portfolioAndRealWorkshop:70, generalSearchInformation:30},
    rules:[
      '합의된 블로그로드맵의 기획대기 항목을 우선한다.',
      '자료대기 또는 실제자료 필요=예인 항목은 실제 자료가 확보되기 전 자동 기획하지 않는다.',
      '기존 발행·예약 글과 주제 또는 검색의도가 실질적으로 겹치면 차별화하거나 다음 항목을 선택한다.',
      '정보형 SEO 글을 2편 이상 연속 기획하지 않는다.',
      '검색 키워드는 실제 공방·제품·제작·클래스 이야기에 자연스럽게 포함한다.',
      '기관·단체 프로그램은 실제 사례가 생기기 전 제안 가능한 프로그램 안내로만 다룬다.'
    ]
  };
}

function ma_v16ColumnLetter_(col) {
  var result = '';
  while (col > 0) {
    var mod = (col - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    col = Math.floor((col - mod) / 26);
  }
  return result;
}
