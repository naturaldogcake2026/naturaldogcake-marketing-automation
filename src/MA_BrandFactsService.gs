/**
 * 자연담은멍케이크 Marketing Automation V1.4
 * Brand Facts: 설정 시트를 AI 콘텐츠 운영 사실의 단일 기준으로 사용한다.
 */
var MA_V14_FACT_DEFAULTS = {
  ORDER_CAKE_TYPES: ['전신케이크, 머핀케이크', '주문제작 케이크 유형'],
  ORDER_MIN_LEAD_TIME: ['최소 3일 전 예약', '주문제작 케이크 최소 예약 리드타임'],
  ORDER_FULFILLMENT: ['픽업만 가능', '현재 주문 수령 방식'],
  ORDER_REFERENCE: ['보내주신 반려견 사진을 참고하여 제작', '주문제작 시 참고 기준'],
  CONTENT_FACTS_VERSION: ['V1.4', 'AI 콘텐츠 사실정보 기준 버전']
};

function ma_getBusinessFacts_(settings) {
  settings = settings || {};
  return {
    cakeTypes: settings.ORDER_CAKE_TYPES || '',
    minimumLeadTime: settings.ORDER_MIN_LEAD_TIME || '',
    fulfillment: settings.ORDER_FULFILLMENT || '',
    orderReference: settings.ORDER_REFERENCE || '',
    classSnackPrice: settings.CLASS_SNACK_PRICE || '',
    classCakePrice: settings.CLASS_CAKE_PRICE || '',
    classDuration: settings.CLASS_DURATION || '',
    classCapacity: settings.CLASS_CAPACITY || '',
    classReservation: settings.CLASS_RESERVATION || ''
  };
}

function ma_businessFactsText_(settings) {
  var f = ma_getBusinessFacts_(settings);
  return [
    '케이크 유형: ' + (f.cakeTypes || '(설정 없음)'),
    '최소 예약기간: ' + (f.minimumLeadTime || '(설정 없음)'),
    '수령 방식: ' + (f.fulfillment || '(설정 없음)'),
    '제작 참고: ' + (f.orderReference || '(설정 없음)'),
    '간식 원데이클래스 가격: ' + (f.classSnackPrice || '(설정 없음)'),
    '케이크 원데이클래스 가격: ' + (f.classCakePrice || '(설정 없음)'),
    '원데이클래스 소요시간: ' + (f.classDuration || '(설정 없음)'),
    '원데이클래스 정원: ' + (f.classCapacity || '(설정 없음)'),
    '원데이클래스 예약: ' + (f.classReservation || '(설정 없음)')
  ].join('\n');
}

function ma_installV14BusinessFacts() {
  var start = Date.now();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MA_CFG.SHEETS.SETTINGS);
  if (!sheet) throw new Error('설정 시트를 찾을 수 없습니다.');

  var lastRow = sheet.getLastRow();
  var values = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, Math.min(3, sheet.getLastColumn())).getValues()
    : [];
  var rowByKey = {};
  values.forEach(function(r, i) {
    var key = String(r[0] || '').trim();
    if (key) rowByKey[key] = i + 2;
  });

  var added = 0;
  Object.keys(MA_V14_FACT_DEFAULTS).forEach(function(key) {
    if (rowByKey[key]) return; // 기존 운영자가 수정한 값은 보호한다.
    var item = MA_V14_FACT_DEFAULTS[key];
    sheet.appendRow([key, item[0], item[1]]);
    added++;
  });

  ma_log_(
    'BUSINESS_FACTS_V14_INSTALL', '', '', 'SUCCESS', 'INFO',
    'V1.4 사실정보 설정 적용: 신규=' + added + '건, 기존값 보존',
    Date.now() - start, 'menu'
  );
  ss.toast(
    'V1.4 사실정보 설정 적용 완료: 신규 ' + added + '건',
    'Marketing Automation',
    6
  );
}
