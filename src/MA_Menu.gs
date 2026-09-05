function onOpen() {
  var ui = SpreadsheetApp.getUi();

  var createMenu = ui.createMenu('① 새 콘텐츠 만들기')
    .addItem('1-0. AI 새 콘텐츠 기획', 'ma_planSelectedContentWithAi')
    .addItem('1-1. 선택 행 ID/기본값 적용', 'ma_applySelectedContentDefaults')
    .addItem('1-2. 블로그 초안 틀 생성', 'ma_createSelectedBlogDraft')
    .addItem('1-3. AI 블로그 작성', 'ma_generateSelectedBlogDraft');

  var reviewMenu = ui.createMenu('② 검수·확정')
    .addItem('2-1. 블로그 검수본 준비', 'ma_prepareSelectedBlogReview')
    .addItem('2-2. 블로그 AI 검수·교정', 'ma_aiReviewSelectedBlog')
    .addItem('2-3. 블로그 최종본 확정', 'ma_finalizeSelectedBlog')
    .addSeparator()
    .addItem('2-4. Threads 3종 생성', 'ma_generateSelectedThreads')
    .addItem('2-5. Thread 검수본 준비', 'ma_prepareSelectedThreadReview')
    .addItem('2-6. Thread 최종본 확정', 'ma_finalizeSelectedThread');

  var publishMenu = ui.createMenu('③ 발행관리')
    .addItem('3-1. 발행대기 생성', 'ma_syncPublishQueue')
    .addItem('3-2. 오늘 발행할 콘텐츠 확인', 'ma_showTodayPublishItems')
    .addItem('3-3. 선택 행 발행일정 동기화', 'ma_syncSelectedPublishSchedule')
    .addItem('3-4. 선택 행 수동 발행완료', 'ma_markSelectedPublishComplete')
    .addItem('3-5. 전체 발행일정 정리·재동기화 V1.3.3', 'ma_reconcileAllPublishSchedules')
    .addItem('3-6. 발행완료 누락정보 복구 V1.5.2', 'ma_repairCompletedPublishMetadataV152');

  var performanceMenu = ui.createMenu('④ 성과관리')
    .addItem('4-1. V1.5 성과관리 시트 설치', 'ma_installV15PerformanceSheet')
    .addItem('4-2. 선택 발행건 성과 기록행 추가', 'ma_addSelectedPerformanceSnapshot')
    .addItem('4-3. 오늘 성과 측정대상 확인', 'ma_showTodayPerformanceTargets')
    .addItem('4-4. 오늘 측정대상 기록행 생성', 'ma_createTodayPerformanceSnapshots');

  var systemMenu = ui.createMenu('⑤ 시스템관리')
    .addItem('5-1. 상태 동기화', 'ma_syncStates')
    .addItem('5-2. 상태 점검', 'ma_runHealthCheck')
    .addItem('5-3. 전체 동기화', 'ma_syncAll')
    .addItem('5-4. V1.4 사실정보 설정 적용', 'ma_installV14BusinessFacts')
    .addItem('5-5. V1.6.1 블로그 로드맵 적용', 'ma_installV161BlogRoadmap');

  ui.createMenu('자연담은멍케이크 콘텐츠')
    .addSubMenu(createMenu)
    .addSubMenu(reviewMenu)
    .addSubMenu(publishMenu)
    .addSubMenu(performanceMenu)
    .addSubMenu(systemMenu)
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  if (row < 2) return;

  try {
    if (sheet.getName() === MA_CFG.SHEETS.CONTENT) {
      ma_applyContentDefaultsForRow_(sheet, row);
    } else if (sheet.getName() === MA_CFG.SHEETS.BLOG) {
      ma_recalcBlogReviewState_(sheet, row);
    } else if (sheet.getName() === MA_CFG.SHEETS.PUBLISH) {
      ma_recalcPublishState_(sheet, row);
    }
  } catch (err) {
    ma_log_('ON_EDIT','','','ERROR','ERROR',
      String(err && err.message ? err.message : err),0,'onEdit');
  }
}
