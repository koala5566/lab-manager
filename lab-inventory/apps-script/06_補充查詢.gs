/**
 * 實驗室藥品器材耗材管理系統 — 需補充清單、查詢（第 7 步）
 *
 * 選單「更新需補充清單」：在「需補充清單」「查詢」兩個工作表放好公式（第一次執行一次即可，之後公式自己會更新）。
 *   ・需補充清單：「品項」有填安全存量、且最新數量低於安全存量的，自動列出（依清單分區、品名排序）
 *   ・查詢：在黃色格子輸入品名關鍵字、選類別／科別／教室…，下方自動列出符合的品項
 * 選單「列印需補充清單」：開啟列印預覽（A4、每頁表頭與核章欄）。
 * 手機 Google 試算表 App 也能看這兩個工作表（查詢也可在手機網頁的「查詢」分頁）。
 * 需要「01_基礎」「02_盤點」「04_列印」。
 */

const RESTOCK_COLS = ['編號', '類別', '清單分區', '品名', '化學式或規格', '教室', '櫃別', '排序位置', '單位',
  '最新盤點日期', '最新數量', '最新數量說明', '安全存量'];
const QUERY_COLS = ['編號', '類別', '科別', '清單分區', '品名', '化學式或規格', '教室', '櫃別', '排序位置', '分處存放',
  '單位', '最新盤點日期', '最新數量', '最新數量說明', '安全存量', '狀態', '備註'];

// ---------------------------------------------------------------- 選單：放好兩張表的公式

function refreshRestock() {
  // 公式放好之後會自己更新；只有公式不見（第一次、或被刪掉）才重建，避免清掉查詢條件
  const ss = SpreadsheetApp.getActive();
  const has = function (name, a1) {
    const sh = ss.getSheetByName(name);
    return sh && /FILTER\(/i.test(sh.getRange(a1).getFormula());
  };
  if (!has('需補充清單', 'A4') || !has('查詢', 'A7')) {
    const items = getTable_('品項');
    need_(items, QUERY_COLS);
    if (!has('需補充清單', 'A4')) setupRestockSheet_(items);
    if (!has('查詢', 'A7')) setupQuerySheet_(items);
  }
  const n = restockItems_().length;
  SpreadsheetApp.getActive().toast(n ? '目前有 ' + n + ' 個品項低於安全存量。' :
    '目前沒有低於安全存量的品項（或還沒填安全存量）。', '需補充清單', 8);
}

/** '品項'!E2:E 這種範圍字串 */
function itemRange_(items, name) {
  const L = colLetter_(items.col[name] + 1);
  return "'品項'!" + L + '2:' + L;
}

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function prepSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  return sh;
}

function headerRow_(sh, row, cols) {
  if (sh.getMaxColumns() < cols.length + 1) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length + 1 - sh.getMaxColumns());
  sh.getRange(row, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#DDEBF7')
    .setHorizontalAlignment('center').setWrap(true);
}

function setupRestockSheet_(items) {
  const sh = prepSheet_('需補充清單');
  const r = function (n) { return itemRange_(items, n); };
  sh.getRange('A1').setValue('需補充清單：「品項」有填安全存量、且最新數量低於安全存量的品項（自動更新，請勿在下方輸入）')
    .setFontWeight('bold');
  sh.getRange('A2').setValue('安全存量在「品項」的「安全存量」欄填寫；空白＝不檢查。已淘汰的不列入。').setFontColor('#666666');
  const cols = RESTOCK_COLS.concat(['建議補充']);
  headerRow_(sh, 3, cols);
  const pick = RESTOCK_COLS.map(r).join(',') + ',' + r('安全存量') + '-' + r('最新數量');
  const formula = '=IFERROR(SORT(FILTER({' + pick + '},' +
    r('編號') + '<>"",' + r('安全存量') + '<>"",' + r('最新數量') + '<>"",' +
    r('最新數量') + '<' + r('安全存量') + ',' + r('狀態') + '<>"已淘汰"),3,TRUE,4,TRUE),"目前沒有低於安全存量的品項")';
  sh.getRange(4, 1).setFormula(formula);
  sh.getRange(4, 1, sh.getMaxRows() - 3, cols.length).setBackground('#FFF8F0');
  sh.setFrozenRows(3);
  [8, 6, 11, 24, 16, 12, 11, 7, 6, 11, 8, 16, 8, 8].forEach(function (w, i) { sh.setColumnWidth(i + 1, w * 8); });
}

function setupQuerySheet_(items) {
  const sh = prepSheet_('查詢');
  const settings = getSettings_();
  const r = function (n) { return itemRange_(items, n); };
  sh.getRange('A1').setValue('查詢：在黃色格子輸入或選擇條件（空白＝不限），下方自動列出符合的品項。清除條件：選取黃色格子按 Delete。')
    .setFontWeight('bold');
  const labels = ['品名／化學式／編號 關鍵字', '類別', '科別', '清單分區', '教室', '櫃別', '狀態'];
  sh.getRange(3, 1, 1, labels.length).setValues([labels]).setFontWeight('bold').setHorizontalAlignment('center')
    .setWrap(true);
  const input = sh.getRange(4, 1, 1, labels.length);
  input.setBackground('#FFF2CC').setBorder(true, true, true, true, true, true);
  sh.getRange(4, 7).setValue('使用中');

  function listRule(listName) {
    const c = settings.header.indexOf(listName);
    return SpreadsheetApp.newDataValidation()
      .requireValueInRange(settings.sheet.getRange(2, c + 1, 60, 1), true).setAllowInvalid(true).build();
  }
  sh.getRange(4, 2).setDataValidation(listRule('類別'));
  sh.getRange(4, 3).setDataValidation(listRule('科別'));
  sh.getRange(4, 4).setDataValidation(listRule('清單分區'));
  sh.getRange(4, 5).setDataValidation(listRule('教室'));
  sh.getRange(4, 7).setDataValidation(listRule('狀態'));

  headerRow_(sh, 6, QUERY_COLS);
  // 每個條件：格子空白就不限；關鍵字與教室、櫃別用「包含」，其他用「等於」
  const cond = [
    '(($A$4="")+ISNUMBER(SEARCH($A$4,' + r('品名') + '&" "&' + r('化學式或規格') + '&" "&' + r('編號') + ')))>0',
    '(($B$4="")+(' + r('類別') + '=$B$4))>0',
    '(($C$4="")+(' + r('科別') + '=$C$4))>0',
    '(($D$4="")+(' + r('清單分區') + '=$D$4))>0',
    '(($E$4="")+ISNUMBER(SEARCH($E$4,' + r('教室') + '&" "&' + r('分處存放') + ')))>0',
    '(($F$4="")+ISNUMBER(SEARCH($F$4,' + r('櫃別') + '&"")))>0',
    '(($G$4="")+(' + r('狀態') + '=$G$4))>0',
    r('編號') + '<>""',
  ];
  const formula = '=IFERROR(FILTER({' + QUERY_COLS.map(r).join(',') + '},' + cond.join(',') + '),"找不到符合條件的品項")';
  sh.getRange(7, 1).setFormula(formula);
  sh.getRange(5, 1).setFormula('=IFERROR("共 "&COUNTA(A7:A)&" 筆","")').setFontColor('#1a73e8');
  sh.setFrozenRows(6);
  [8, 6, 8, 11, 24, 16, 12, 11, 7, 14, 6, 11, 8, 16, 8, 7, 20].forEach(function (w, i) { sh.setColumnWidth(i + 1, w * 8); });
}

// ---------------------------------------------------------------- 套用安全存量建議（一次性）

/**
 * 把「安全存量建議」工作表（由 安全存量建議.xlsx 匯入）裡「採用」為 TRUE 的列，
 * 填進「品項」的「安全存量」。品項已經有填安全存量的不覆蓋。
 * 執行方式：Apps Script 上方函式選單選 applySafetyStock → ▶ 執行。
 */
function applySafetyStock() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName('安全存量建議')) {
    throw new Error('找不到「安全存量建議」工作表。請先「檔案 → 匯入」安全存量建議.xlsx，匯入位置選「插入新工作表」。');
  }
  const sug = getTable_('安全存量建議');
  need_(sug, ['採用', '編號', '建議安全存量']);
  const items = getTable_('品項');
  need_(items, ['編號', '安全存量']);
  const rowOf = {};
  items.rows.forEach(function (r, i) { rowOf[String(r[items.col['編號']]).trim()] = i; });
  const col = items.col['安全存量'];
  const values = items.rows.map(function (r) { return [r[col]]; });
  let set = 0, kept = 0, skipped = 0, missing = 0;
  sug.rows.forEach(function (r) {
    const on = r[sug.col['採用']];
    const code = String(r[sug.col['編號']]).trim();
    const v = r[sug.col['建議安全存量']];
    if (!code) return;
    if (!(on === true || /^(TRUE|是|✔|V|Y)$/i.test(String(on).trim())) || !isNumber_(v)) { skipped++; return; }
    if (!(code in rowOf)) { missing++; return; }
    const cur = values[rowOf[code]][0];
    if (String(cur).trim() !== '') { kept++; return; }
    values[rowOf[code]][0] = Number(v);
    set++;
  });
  if (items.rows.length) items.sheet.getRange(2, col + 1, items.rows.length, 1).setValues(values);
  const n = restockItems_().length;
  const msg = '已填入 ' + set + ' 個品項的安全存量。' +
    (kept ? '\n原本已有安全存量、沒有覆蓋：' + kept + ' 個。' : '') +
    (skipped ? '\n不採用（FALSE）或沒有數字：' + skipped + ' 個。' : '') +
    (missing ? '\n品項找不到編號：' + missing + ' 個。' : '') +
    '\n\n目前低於安全存量：' + n + ' 個（見「需補充清單」）。\n「安全存量建議」工作表確認完可以刪除。';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return msg;
}

// ---------------------------------------------------------------- 列印需補充清單

function printRestockList() {
  showPrintDialog({ page: 'restock' });
}

/** 低於安全存量的品項（依清單分區列印順序、品名排序）。 */
function restockItems_() {
  const settings = getSettings_();
  const order = {};
  settings.zones.forEach(function (z, i) { order[String(z['清單分區'])] = i; });
  return printData_().items.filter(function (it) {
    return isNumber_(it['安全存量']) && String(it['安全存量']).trim() !== '' &&
      isNumber_(it['最新數量']) && String(it['最新數量']).trim() !== '' &&
      Number(it['最新數量']) < Number(it['安全存量']);
  }).sort(function (a, b) {
    return (order[a['清單分區']] || 99) - (order[b['清單分區']] || 99) || naturalCompare_(a['品名'], b['品名']);
  });
}

function page_restock() {
  const settings = getSettings_();
  const list = restockItems_();
  const rows = list.map(function (it, i) {
    // 藥品印到排序位置（A 1-1）；器材耗材的排序位置只是清單序號，不印
    const loc = [it['教室'], it['櫃別'], it['清單分區'] === '藥品清單' ? it['排序位置'] : '']
      .filter(function (x) { return String(x).trim(); }).join(' ');
    return '<tr><td class="c">' + (i + 1) + '</td><td class="c">' + esc_(it['清單分區']) + '</td><td>' + esc_(it['品名']) +
      (it['化學式或規格'] ? '<br><small>' + esc_(it['化學式或規格']) + '</small>' : '') + '</td><td class="c">' +
      esc_(loc) + '</td><td class="c">' + esc_(it['單位']) + '</td><td class="c">' +
      esc_(it['最新數量說明'] || numText_(it['最新數量'])) + '<br><small>' + esc_(it['最新盤點日期']) + '</small></td>' +
      '<td class="c">' + esc_(numText_(it['安全存量'])) + '</td><td class="c">' +
      esc_(numText_(round_(Number(it['安全存量']) - Number(it['最新數量'])))) + '</td><td></td></tr>';
  });
  if (!rows.length) rows.push('<tr><td colspan="9" class="c">目前沒有低於安全存量的品項</td></tr>');
  const today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  const body = reportTable_(settings, {
    title: '實驗室需補充品項清單', sub: '最新數量低於安全存量', place: '', updated: today,
    widths: [5, 12, 26, 17, 6, 11, 8, 8, 7],
    heads: ['序', '清單分區', '品名', '存放位置', '單位', '最新數量', '安全<br>存量', '建議<br>補充', '備註'], rows: rows,
  });
  return HtmlService.createHtmlOutput(printShell_('需補充清單', body)).setTitle('需補充清單（列印）');
}
