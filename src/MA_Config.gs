var MA_CFG = {
  VERSION: 'MA-V1.0.0',
  SHEETS: {
    CONTENT: '肄섑뀗痢잻B',
    CALENDAR: '肄섑뀗痢좎틮由곕뜑',
    BLOG: '釉붾줈洹몄큹??,
    THREADS: 'Threads珥덉븞',
    PUBLISH: '諛쒗뻾愿由?,
    SETTINGS: '?ㅼ젙',
    LOG: '?먮룞?붾줈洹?
  },
  HEADERS: {
    CONTENT: [
      'Content ID','?깅줉??,'Pillar','肄섑뀗痢좎쑀??,'二쇱젣','?듭떖?ㅼ썙??,'蹂댁“?ㅼ썙??,'??곷룆??,
      '寃?됱쓽??,'肄섑뀗痢좊ぉ??,'CTA','李멸퀬?먮즺','?ъ떎?뺤씤硫붾え','肄섑뀗痢좎긽??,'?곗꽑?쒖쐞',
      '紐⑺몴諛쒗뻾??,'釉붾줈洹퇙D','Threads ?앹꽦??,'愿由ъ옄硫붾え','理쒖쥌?섏젙??
    ],
    BLOG: [
      'Blog ID','Content ID','?앹꽦??,'?쒕ぉ1','?쒕ぉ2','?쒕ぉ3','理쒖쥌?쒕ぉ','?듭떖?ㅼ썙??,
      '?꾩엯遺','蹂몃Ц','CTA','?댁떆?쒓렇','?대?吏?붿껌','?ъ떎寃??,'臾몄껜寃??,'SEO寃??,
      '愿由ъ옄?섏젙蹂?,'理쒖쥌蹂몃Ц','寃?섏긽??,'?뱀씤??,'諛쒗뻾?곹깭','諛쒗뻾??,'諛쒗뻾URL'
    ],
    THREADS: [
      'Thread ID','Content ID','Blog ID','?앹꽦??,'Thread ?좏삎','Hook','蹂몃Ц','CTA',
      '理쒖쥌蹂몃Ц','寃?섏긽??,'?뱀씤??,'諛쒗뻾?곹깭','諛쒗뻾?덉젙??,'諛쒗뻾??,'諛쒗뻾URL','愿由ъ옄硫붾え'
    ],
    PUBLISH: [
      'Publish ID','Content ID','肄섑뀗痢쟅D','?뚮옯??,'理쒖쥌?쒕ぉ','理쒖쥌蹂몃Ц 議댁옱','寃?섏긽??,
      '?뱀씤?щ?','諛쒗뻾?덉젙??,'諛쒗뻾?덉젙?쒓컙','諛쒗뻾?곹깭','諛쒗뻾??,'諛쒗뻾URL',
      '?ㅻ쪟?щ?','?ㅻ쪟硫붿떆吏','?ъ떆?꾪슏??
    ],
    LOG: ['?ㅽ뻾?쒓컖','Run ID','湲곕뒫','Content ID','??갏D','?ㅽ뻾寃곌낵','?섏?','硫붿떆吏','泥섎━?쒓컙','?ㅽ뻾二쇱껜']
  }
};

function ma_getSettings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MA_CFG.SHEETS.SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(3, sheet.getLastColumn())).getValues();
  var out = {};
  values.forEach(function(r) {
    var key = String(r[0] || '').trim();
    if (key) out[key] = r[1];
  });
  return out;
}

function ma_headerMap_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error('?ㅻ뜑媛 ?놁뒿?덈떎: ' + sheet.getName());
  var row = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var map = {};
  row.forEach(function(v, i) {
    var key = String(v || '').trim();
    if (key) map[key] = i + 1;
  });
  return map;
}

function ma_valueByHeader_(row, headerMap, header) {
  var col = headerMap[header];
  return col ? row[col - 1] : '';
}

function ma_colLetter_(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function ma_escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

