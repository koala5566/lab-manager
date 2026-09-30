/**
 * 實驗室藥品器材耗材管理系統 — 資料安全與美化（第一批優化）
 *
 * 選單「🔧 維護 → 首頁、美化與資料保護（一次設定）」執行一次就好，會：
 *   1. 建立「首頁」儀表板（品項數、需補充、最近盤點、盤點進度…＋各分頁連結）
 *   2. 分頁上色、排順序
 *   3. 「品項」標色：低於安全存量整列淡紅、已淘汰灰字、隔行底色
 *   4. 資料保護（警告模式）：改到盤點紀錄、設定、自動欄位時先跳出確認
 *   5. 每週五 17:00 自動備份（完成盤點後也會自動備份），備份放在雲端硬碟「實驗室管理備份」資料夾，保留最近 10 份
 *   6. 刪除偵測：盤點紀錄或品項被刪列、重要分頁被刪時，右下角警告並記在「操作紀錄」
 * 選單「🔧 維護 → 立即備份」：馬上備份一份。
 * 需要「01_基礎」～「06_補充查詢」。
 */

const BACKUP_FOLDER = '實驗室管理備份';
const BACKUP_KEEP = 10;
const WATCH_SHEETS = ['品項', '盤點紀錄'];
const KEY_SHEETS = ['設定', '品項', '盤點紀錄', '盤點表', '異動紀錄'];
const PROTECT_TAG = '實驗室系統保護：';

// ---------------------------------------------------------------- 一次設定

function setupEnhancements() {
  const ss = SpreadsheetApp.getActive();
  const done = [];
  const step = function (name, fn) {
    try { fn(); done.push('✔ ' + name); } catch (e) { done.push('✘ ' + name + '：' + e.message); }
  };
  step('請購清單、借用紀錄、實驗排程、準備事項工作表、實驗室使用設定', function () {
    ['ensureLabSettings_', 'ensureMeasureCols_', 'setupPurchaseSheet_', 'setupLoanSheet_', 'setupExperimentSheet_', 'setupTodoSheet_'].forEach(function (fn) {
      if (typeof globalThis[fn] === 'function') globalThis[fn]();
    });
  });
  step('首頁儀表板', setupDashboard_);
  step('分頁顏色與順序', arrangeTabs_);
  step('品項標色（低於安全存量、已淘汰、隔行底色）', styleItemSheet_);
  step('資料保護（警告模式）', protectSheets_);
  step('每週自動備份與刪除偵測', installTriggers_);
  step('第一次備份', function () { backupNow('設定完成', true); });
  recordBaseline_();
  ss.setActiveSheet(ss.getSheetByName('首頁'));
  SpreadsheetApp.getUi().alert('設定完成', done.join('\n') +
    '\n\n之後不用再執行；新增工作表或欄位後想重新套用，可以再按一次。', SpreadsheetApp.getUi().ButtonSet.OK);
}

function goHome() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('首頁');
  if (!sh) { SpreadsheetApp.getUi().alert('還沒有首頁。請先執行「🔧 維護 → 首頁、美化與資料保護（一次設定）」。'); return; }
  ss.setActiveSheet(sh);
}

// ---------------------------------------------------------------- 備份

/**
 * 複製整份試算表（含程式）到雲端硬碟「實驗室管理備份」資料夾，只保留最近 10 份。
 * quiet=true 時不跳提示（給完成盤點、每週排程用）。回傳備份檔名。
 */
function backupNow(reason, quiet) {
  if (typeof reason !== 'string') reason = '手動';
  const ss = SpreadsheetApp.getActive();
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(BACKUP_FOLDER);
  const stamp = Utilities.formatDate(new Date(), SCHOOL_TZ, 'yyyy-MM-dd_HHmm');
  const name = '備份_' + stamp + '_' + reason;
  DriveApp.getFileById(ss.getId()).makeCopy(name, folder);

  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf('備份_') === 0) files.push(f);
  }
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(BACKUP_KEEP).forEach(function (f) { f.setTrashed(true); });

  if (!quiet) {
    SpreadsheetApp.getUi().alert('備份完成', '已備份到雲端硬碟「' + BACKUP_FOLDER + '」資料夾：\n' + name +
      '\n\n（只保留最近 ' + BACKUP_KEEP + ' 份，較舊的會移到垃圾桶。）', SpreadsheetApp.getUi().ButtonSet.OK);
  }
  return name;
}

/** 每週排程呼叫。 */
function weeklyBackup() {
  backupNow('每週自動', true);
}

// ---------------------------------------------------------------- 排程與刪除偵測

function installTriggers_() {
  const ss = SpreadsheetApp.getActive();
  const have = {};
  ScriptApp.getProjectTriggers().forEach(function (t) { have[t.getHandlerFunction()] = true; });
  if (!have.weeklyBackup) {
    ScriptApp.newTrigger('weeklyBackup').timeBased().everyWeeks(1)
      .onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(17).inTimezone(SCHOOL_TZ).create();
  }
  if (!have.onChangeWatch) {
    ScriptApp.newTrigger('onChangeWatch').forSpreadsheet(ss).onChange().create();
  }
}

/** 記下目前各重要工作表的列數，作為之後比較的基準。 */
function recordBaseline_() {
  const ss = SpreadsheetApp.getActive();
  const base = {};
  WATCH_SHEETS.forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh) base[n] = Math.max(0, sh.getLastRow() - 1);
  });
  PropertiesService.getDocumentProperties().setProperty('baseline', JSON.stringify(base));
}

/** 試算表結構改變時（刪列、刪分頁…）由 Google 自動呼叫。 */
function onChangeWatch(e) {
  const type = e && e.changeType;
  // 其他變動（新增列、編輯…）只更新基準，這樣之後刪列才算得準
  if (type !== 'REMOVE_ROW' && type !== 'REMOVE_GRID') { recordBaseline_(); return; }
  const ss = SpreadsheetApp.getActive();
  const base = JSON.parse(PropertiesService.getDocumentProperties().getProperty('baseline') || '{}');
  const notes = [];
  WATCH_SHEETS.forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (!sh || !(n in base)) return;
    const now = Math.max(0, sh.getLastRow() - 1);
    if (now < base[n]) {
      notes.push([n, base[n], now, base[n] - now,
        n === '品項' ? '品項不要刪列，請把「狀態」改成「已淘汰」' : '如果是誤刪，請馬上按 Ctrl+Z 復原']);
    }
  });
  KEY_SHEETS.forEach(function (n) {
    if (!ss.getSheetByName(n)) notes.push([n, '', '', '整個分頁', '重要分頁不見了！請馬上按 Ctrl+Z，或從「檔案 → 版本記錄」還原']);
  });
  if (!notes.length) { recordBaseline_(); return; }

  const log = ss.getSheetByName('操作紀錄') || ss.insertSheet('操作紀錄');
  if (log.getLastRow() === 0) {
    log.appendRow(['時間', '工作表', '原本列數', '現在列數', '減少', '說明']);
    log.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#DDEBF7');
    log.setFrozenRows(1);
    log.setTabColor('#9AA0A6');
  }
  const stamp = Utilities.formatDate(new Date(), SCHOOL_TZ, 'yyyy/MM/dd HH:mm');
  notes.forEach(function (x) { log.appendRow([stamp].concat(x)); });
  ss.toast(notes.map(function (x) { return x[0] + '：少了 ' + x[3] + (typeof x[3] === 'number' ? ' 列' : '') + '。' + x[4]; })
    .join('\n'), '⚠ 偵測到資料被刪除', 20);
  recordBaseline_();
}

// ---------------------------------------------------------------- 資料保護（警告模式）

/**
 * 警告模式：您仍然可以修改，但會先跳出「確定要編輯嗎？」，防止手滑。程式本身的寫入不受影響。
 * 重複執行會先移除舊的保護再重建。
 */
function protectSheets_() {
  const ss = SpreadsheetApp.getActive();
  const clearOld = function (sh) {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).concat(sh.getProtections(SpreadsheetApp.ProtectionType.RANGE))
      .forEach(function (p) { if (String(p.getDescription()).indexOf(PROTECT_TAG) === 0) p.remove(); });
  };
  const warnSheet = function (name, desc, openA1s) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    clearOld(sh);
    const p = sh.protect().setDescription(PROTECT_TAG + desc).setWarningOnly(true);
    if (openA1s && openA1s.length) p.setUnprotectedRanges(openA1s.map(function (a) { return sh.getRange(a); }));
  };

  warnSheet('盤點紀錄', '盤點紀錄是歷史資料，請用「完成盤點」寫入');
  warnSheet('設定', '設定會影響整個系統');
  warnSheet('首頁', '首頁是自動產生的');
  warnSheet('需補充清單', '需補充清單是公式自動產生的');
  warnSheet('查詢', '只在黃色格子輸入條件', ['A4:G4']);
  const count = ss.getSheetByName('盤點表');
  if (count) {
    const k = COUNT_HEADER.indexOf('本次數量') + 1;
    const L = colLetter_(k), R = colLetter_(k + 1);
    warnSheet('盤點表', '只在黃色的本次數量、本次說明欄輸入', [L + '2:' + R]);
  }

  // 品項：一般欄位可自由編輯；只保護標題列與程式自動填的欄位
  const items = getTable_('品項');
  clearOld(items.sheet);
  const warnRange = function (a1, desc) {
    items.sheet.getRange(a1).protect().setDescription(PROTECT_TAG + desc).setWarningOnly(true);
  };
  warnRange('1:1', '欄位名稱不要修改，程式靠名稱找欄位');
  ['編號', '最新盤點日期', '最新數量', '最新數量說明'].forEach(function (n) {
    if (!(n in items.col)) return;
    const L = colLetter_(items.col[n] + 1);
    warnRange(L + '2:' + L, n + '由程式自動填寫');
  });
}

// ---------------------------------------------------------------- 分頁顏色與順序

function arrangeTabs_() {
  const ss = SpreadsheetApp.getActive();
  const order = [
    ['首頁', '#1A73E8'],
    ['盤點表', '#FBBC04'], ['實驗排程', '#FBBC04'], ['準備事項', '#FBBC04'], ['查詢', '#FBBC04'], ['需補充清單', '#34A853'],
    ['請購清單', '#FF6D01'], ['借用紀錄', '#FF6D01'],
    ['品項', '#4285F4'], ['盤點紀錄', '#4285F4'], ['異動紀錄', '#4285F4'],
    ['實驗套組', '#A142F4'], ['玻片需求', '#A142F4'],
    ['設定', '#9AA0A6'], ['轉入檢查', '#9AA0A6'], ['操作紀錄', '#9AA0A6'], ['安全存量建議', '#9AA0A6'], ['危險分類建議', '#9AA0A6'],
  ];
  let pos = 1;
  order.forEach(function (x) {
    const sh = ss.getSheetByName(x[0]);
    if (!sh) return;
    sh.setTabColor(x[1]);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(pos++);
  });
}

// ---------------------------------------------------------------- 品項標色

function styleItemSheet_() {
  const items = getTable_('品項');
  const sh = items.sheet, c = items.col;
  const lastCol = sh.getLastColumn(), rows = sh.getMaxRows() - 1;
  const range = sh.getRange(2, 1, rows, lastCol);
  const L = function (n) { return '$' + colLetter_(c[n] + 1); };

  const retired = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=' + L('狀態') + '2="已淘汰"')
    .setFontColor('#9AA0A6').setItalic(true).setRanges([range]).build();
  const low = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(' + L('安全存量') + '2<>"",' + L('最新數量') + '2<>"",' + L('最新數量') + '2<' +
      L('安全存量') + '2,' + L('狀態') + '2<>"已淘汰")')
    .setBackground('#FCE8E6').setFontColor('#B3261E').setRanges([range]).build();
  const others = sh.getConditionalFormatRules().filter(function (r) {
    const f = r.getBooleanCondition();
    const v = f && f.getCriteriaValues()[0];
    return !(v && /已淘汰|最新數量|安全存量/.test(String(v)) && /^=/.test(String(v)));
  });
  sh.setConditionalFormatRules([low, retired].concat(others));

  // 隔行底色（自動欄位原本的灰底會蓋過隔行色，所以先清掉那些欄的底色）
  ['編號', '最新盤點日期', '最新數量', '最新數量說明'].forEach(function (n) {
    if (n in c) sh.getRange(2, c[n] + 1, rows, 1).setBackground(null).setFontColor('#1A73E8');
  });
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.getRange(1, 1, rows + 1, lastCol).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  sh.getRange(1, 1, 1, lastCol).setBackground('#DDEBF7').setFontWeight('bold');

  const recs = SpreadsheetApp.getActive().getSheetByName('盤點紀錄');
  if (recs && !recs.getBandings().length) {
    recs.getRange(1, 1, recs.getMaxRows(), recs.getLastColumn()).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    recs.getRange(1, 1, 1, recs.getLastColumn()).setBackground('#DDEBF7').setFontWeight('bold');
  }
}

// ---------------------------------------------------------------- 首頁儀表板

function setupDashboard_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('首頁');
  if (sh) {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
    sh.clear();
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  } else {
    sh = ss.insertSheet('首頁', 0);
  }
  const items = getTable_('品項');
  const recs = getTable_('盤點紀錄');
  const s = getSettings_();
  const I = function (n) { const L = colLetter_(items.col[n] + 1); return "'品項'!" + L + '2:' + L; };
  const R = function (n) { const L = colLetter_(recs.col[n] + 1); return "'盤點紀錄'!" + L + '2:' + L; };
  const termCell = "'設定'!B" + (Object.keys(s.params).indexOf('學年度學期') + 2);
  const cq = colLetter_(COUNT_HEADER.indexOf('本次數量') + 1), cd = colLetter_(COUNT_HEADER.indexOf('已盤') + 1),
    cc = colLetter_(COUNT_HEADER.indexOf('編號') + 1);

  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 24);
  for (let i = 2; i <= 7; i++) sh.setColumnWidth(i, 170);
  sh.setColumnWidth(8, 24);

  // 標題
  sh.getRange('B2:G2').merge().setValue('🧪 實驗室藥品器材耗材管理系統')
    .setFontSize(22).setFontWeight('bold').setFontColor('#1A73E8');
  sh.getRange('B3:G3').merge().setFormula('="' + s.params['學校名稱'] + '　教務處設備組　｜　目前學期："&' + termCell + '')
    .setFontSize(12).setFontColor('#5F6368');
  sh.setRowHeight(2, 44);

  // 數字卡片：[標題, 公式, 說明, 顏色]
  const restockCount = "IF(OR('需補充清單'!A4=\"\",LEFT('需補充清單'!A4,4)=\"目前沒有\"),0,COUNTA('需補充清單'!A4:A))";
  const cards = [
    ['使用中品項', '=COUNTIFS(' + I('狀態') + ',"使用中")',
      '="藥品 "&COUNTIFS(' + I('類別') + ',"藥品",' + I('狀態') + ',"使用中")&"　器材 "&COUNTIFS(' + I('類別') + ',"器材",' +
      I('狀態') + ',"使用中")&"　耗材 "&COUNTIFS(' + I('類別') + ',"耗材",' + I('狀態') + ',"使用中")', '#1A73E8'],
    ['需補充', '=' + restockCount, '低於安全存量的品項', '#D93025'],
    ['最近盤點', '=IFERROR((YEAR(MAX(' + R('盤點日期') + '))-1911)&TEXT(MAX(' + R('盤點日期') + '),".mm.dd"),"—")',
      '="共 "&COUNTUNIQUE(' + R('盤點日期') + ')&" 次盤點紀錄"', '#188038'],
    ['盤點表進度', "=IF(COUNTA('盤點表'!" + cc + '2:' + cc + ')=0,"—",COUNTIF(\'盤點表\'!' + cd + '2:' + cd + ',"✔")&"／"&COUNTA(\'盤點表\'!' + cc + '2:' + cc + '))',
      "=IF(COUNTA('盤點表'!" + cc + '2:' + cc + ')=0,"目前沒有進行中的盤點","已填／共幾列")', '#E37400'],
    ['危險物品', '=COUNTIFS(' + I('危險物品') + ',"是",' + I('狀態') + ',"使用中")', '「品項」危險物品＝是', '#A142F4'],
    ['未設安全存量', '=COUNTIFS(' + I('類別') + ',"<>器材",' + I('狀態') + ',"使用中",' + I('安全存量') + ',"")',
      '藥品＋耗材中還沒填的', '#5F6368'],
  ];
  const drawCards = function (top, list) {
    list.forEach(function (card, i) {
      const col = 2 + i;
      sh.getRange(top, col).setValue(card[0]).setFontSize(11).setFontColor('#5F6368');
      sh.getRange(top + 1, col).setFormula(card[1]).setFontSize(26).setFontWeight('bold').setFontColor(card[3]);
      const note = sh.getRange(top + 2, col);
      if (String(card[2]).charAt(0) === '=') note.setFormula(card[2]); else note.setValue(card[2]);
      note.setFontSize(9).setFontColor('#80868B').setWrap(true);
      sh.getRange(top, col, 3, 1).setBackground('#F8F9FA').setHorizontalAlignment('center')
        .setBorder(true, true, true, true, null, null, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);
    });
    sh.setRowHeight(top + 1, 48);
    sh.setRowHeight(top + 2, 34);
  };
  drawCards(5, cards);

  // 第二排：請購、實驗、借用（工作表存在才算）
  const col = function (sheet, name) {
    const x = ss.getSheetByName(sheet);
    if (!x) return null;
    const head = x.getRange(1, 1, 1, Math.max(1, x.getLastColumn())).getValues()[0].map(String);
    const i = head.indexOf(name);
    return i < 0 ? null : "'" + sheet + "'!" + colLetter_(i + 1) + '2:' + colLetter_(i + 1);
  };
  const cards2 = [];
  const rq = col('請購清單', '狀態');
  if (rq) cards2.push(['待處理請購', '=COUNTIF(' + rq + ',"待處理")', '="已請購 "&COUNTIF(' + rq + ',"已請購")&" 項等待到貨"', '#FF6D01']);
  const ed = col('實驗排程', '日期'), es = col('實驗排程', '狀態');
  if (ed && es) cards2.push(['7 天內實驗室使用', '=COUNTIFS(' + ed + ',">="&TODAY(),' + ed + ',"<="&(TODAY()+7),' + es + ',"<>取消",' + es + ',"<>已歸還")',
    '="其中待準備 "&COUNTIFS(' + ed + ',">="&TODAY(),' + ed + ',"<="&(TODAY()+7),' + es + ',"待準備")&" 個"', '#E37400']);
  const ts = col('準備事項', '狀態'), tdd = col('準備事項', '需要日期');
  if (ts && tdd) cards2.push(['待準備事項', '=COUNTIF(' + ts + ',"待準備")',
    '=IF(COUNTIFS(' + ts + ',"待準備",' + tdd + ',"<"&TODAY(),' + tdd + ',"<>")>0,"⚠ 過期 "&COUNTIFS(' + ts + ',"待準備",' + tdd +
    ',"<"&TODAY(),' + tdd + ',"<>")&" 項","今明兩天 "&COUNTIFS(' + ts + ',"待準備",' + tdd + ',">="&TODAY(),' + tdd + ',"<="&(TODAY()+1))&" 項")', '#1A73E8']);
  const ls = col('借用紀錄', '狀態'), ld = col('借用紀錄', '預計歸還');
  if (ls && ld) cards2.push(['借出中', '=COUNTIF(' + ls + ',"借出中")',
    '=IF(COUNTIFS(' + ls + ',"借出中",' + ld + ',"<"&TODAY(),' + ld + ',"<>")>0,"⚠ 逾期 "&COUNTIFS(' + ls + ',"借出中",' + ld +
    ',"<"&TODAY(),' + ld + ',"<>")&" 筆","沒有逾期")', '#A142F4']);
  let top = 9;
  if (cards2.length) { drawCards(9, cards2); top = 13; }

  // 今天／下一個上課日的課、待準備事項（公式，每天自己更新）
  const E = function (n) { return col('實驗排程', n); };
  if (E('日期') && E('節次') && E('教室') && E('狀態') && E('實驗名稱')) {
    const D = E('日期'), P = E('節次'), L = E('教室'), S = E('狀態'), N = E('實驗名稱'), C = E('班級'), TC = E('教師');
    const TY = E('用途類型') || N;
    const WD = function (d) { return '"（"&CHOOSE(WEEKDAY(' + d + ',2),"一","二","三","四","五","六","日")&"）"'; };
    const dayList = function (day) {
      return '=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(QUERY(SORT(FILTER({IFERROR(MATCH(LEFT(' + P + '&"",1),{"1","2","3","4","午","5","6","7"},0),9),' +
        'IF(' + P + '&""="午","中午",' + P + '&""),REGEXREPLACE(REGEXREPLACE(' + L + '&"","學?實驗室",""),"能?教室",""),' +
        'IF(' + N + '="",' + TY + '&"",' + N + '&"")&IF(' + C + '="",""," "&' + C + ')&IF(' + TC + '="",""," "&' + TC + ')},' +
        D + '=' + day + ',' + S + '<>"取消"),1,TRUE),"select Col2,Col3,Col4",0),14,3),"（沒有登記）"))';
    };
    const r0 = top;
    [['B', 'TODAY()', '📅 今天 '], ['E', 'WORKDAY(TODAY(),1)', '📅 下一個上課日 ']].forEach(function (x) {
      const c0 = x[0] === 'B' ? 2 : 5;
      sh.getRange(r0, c0, 1, 3).merge().setFormula('="' + x[2] + '"&TEXT(' + x[1] + ',"m/d")&' + WD(x[1]))
        .setFontSize(13).setFontWeight('bold').setFontColor('#1A73E8');
      sh.getRange(r0 + 1, c0, 1, 3).setValues([['節次', '實驗室', '內容・班級・老師']]).setFontWeight('bold').setBackground('#F1F3F4').setFontSize(10);
      sh.getRange(r0 + 2, c0).setFormula(dayList(x[1]));
      sh.getRange(r0 + 2, c0, 14, 3).setFontSize(10).setVerticalAlignment('top');
      sh.getRange(r0 + 16, c0, 1, 3).merge().setFormula('=IF(COUNTIFS(' + D + ',' + x[1] + ',' + S + ',"<>取消")>14,"…還有 "&(COUNTIFS(' + D + ',' + x[1] + ',' +
        S + ',"<>取消")-14)&" 筆，請看「🗓 實驗室使用 → 實驗室使用一覽」","")').setFontSize(9).setFontColor('#80868B');
      sh.getRange(r0 + 1, c0, 16, 3).setBorder(true, true, true, true, null, null, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);
    });
    top = r0 + 18;
    const TD = function (n) { return col('準備事項', n); };
    if (TD('需要日期') && TD('狀態') && TD('事項')) {
      const d = TD('需要日期'), st = TD('狀態'), w = TD('事項'), q = TD('數量'), u = TD('單位'), lab = TD('實驗室'), pp = TD('節次'),
        cl = TD('班級'), tc = TD('教師'), nt = TD('備註');
      sh.getRange(top, 2, 1, 6).merge().setValue('📝 待準備事項（過期＋到下一個上課日）').setFontSize(13).setFontWeight('bold').setFontColor('#1A73E8');
      sh.getRange(top + 1, 2, 1, 5).setValues([['日期', '要準備的東西', '實驗室・節次・班級', '老師', '備註']]).setFontWeight('bold').setBackground('#F1F3F4').setFontSize(10);
      sh.getRange(top + 2, 2).setFormula('=ARRAYFORMULA(IFERROR(ARRAY_CONSTRAIN(QUERY(SORT(FILTER({' + d + ',IF(' + d + '<TODAY(),"⚠ ","")&TEXT(' + d + ',"m/d"),' +
        w + '&IF(' + q + '="",""," × "&' + q + '&" "&' + u + '),TRIM(REGEXREPLACE(REGEXREPLACE(' + lab + '&"","學?實驗室",""),"能?教室","")&" "&' + pp + '&" "&' + cl + '),' + tc + '&"",' + nt + '&""},' +
        st + '="待準備",' + d + '<>"",' + d + '<=WORKDAY(TODAY(),1)),1,TRUE),"select Col2,Col3,Col4,Col5,Col6",0),8,5),"（沒有 👍）"))');
      sh.getRange(top + 2, 2, 8, 5).setFontSize(10).setVerticalAlignment('top');
      sh.getRange(top + 10, 2, 1, 5).merge().setFormula('=IF(COUNTIFS(' + st + ',"待準備",' + d + ',"<="&WORKDAY(TODAY(),1),' + d + ',"<>")>8,"…還有更多，請看「🗓 實驗室使用 → 📝 準備事項待辦清單」","")')
        .setFontSize(9).setFontColor('#80868B');
      sh.getRange(top + 1, 2, 10, 5).setBorder(true, true, true, true, null, null, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);
      top += 12;
    }
  }

  // 常用操作
  sh.getRange(top, 2, 1, 6).merge().setValue('常用操作（上方選單「🧪 實驗室管理」）').setFontSize(13).setFontWeight('bold');
  const howto = [
    ['盤點', '📋 盤點 → 產生盤點表 → 電腦填黃色欄位，或手機開盤點網頁 → 📋 盤點 → 完成盤點'],
    ['申報', '🖨 列印 → 申報清單（期初／期末）→ 選那次盤點 → 🖨 列印；簽稿文字在預覽上方'],
    ['請購', '🛒 請購 → 新增請購需求／從需補充清單加入 → 狀態改「已請購」→ 到貨時 🛒 請購 → 到貨入庫（自動加庫存）'],
    ['實驗室使用', '🗓 實驗室使用 → 登記使用（點課表格子）→ 列印門口海報／本週課表；使用一覽看今天誰在用'],
    ['實驗準備', '🗓 實驗室使用 → 列印實驗準備單（需要總數對照庫存，附準備／歸還打勾欄）'],
    ['準備事項', '老師口頭、便條交代的：🗓 實驗室使用 → 新增準備事項；待辦清單打勾＝已準備（手機「待辦」分頁也可以）'],
    ['新增／異動', '📦 品項 → 新增品項；新購、領用、報廢、移位用 📦 品項 → 登記異動（不要刪列）'],
    ['借用', '📦 品項 → 借出登記／歸還登記；逾期會在首頁與「借用紀錄」標紅'],
    ['備份', '每週五 17:00 自動備份；完成盤點後也會備份；要馬上備份：🔧 維護 → 立即備份'],
  ];
  howto.forEach(function (h, i) {
    sh.getRange(top + 1 + i, 2).setValue(h[0]).setFontWeight('bold').setFontColor('#1A73E8');
    sh.getRange(top + 1 + i, 3, 1, 5).merge().setValue(h[1]).setFontColor('#3C4043').setWrap(true);
  });

  // 分頁連結
  const linkRow = top + 1 + howto.length + 1;
  sh.getRange(linkRow, 2, 1, 6).merge().setValue('前往分頁').setFontSize(13).setFontWeight('bold');
  const tabs = ['盤點表', '查詢', '需補充清單', '請購清單', '實驗排程', '準備事項', '借用紀錄', '品項', '盤點紀錄', '異動紀錄', '實驗套組',
    '玻片需求', '設定'];
  tabs.forEach(function (n, i) {
    const t = ss.getSheetByName(n);
    if (!t) return;
    const cell = sh.getRange(linkRow + 1 + Math.floor(i / 6), 2 + (i % 6));
    cell.setFormula('=HYPERLINK("#gid=' + t.getSheetId() + '","→ ' + n + '")').setFontColor('#1A73E8');
  });
  sh.getRange(linkRow + 3 + Math.ceil(tabs.length / 6) - 2, 2, 1, 6).merge()
    .setValue('這一頁是自動產生的，數字會自己更新。要重建：🔧 維護 → 首頁、美化與資料保護（一次設定）。')
    .setFontSize(9).setFontColor('#9AA0A6');
  sh.setFrozenRows(0);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
}
