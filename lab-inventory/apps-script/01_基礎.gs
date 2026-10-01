/**
 * 實驗室藥品器材耗材管理系統 — 基礎功能（第 2 步）
 * 臺北市立中山女子高級中學 教務處設備組
 *
 * 內容：自訂選單、讀取設定、自動產生編號、計算最新數量、套用格式與下拉選單。
 * 其他步驟的程式會用到這裡的共用函式，請不要刪除這個檔案。
 */

// 自訂選單：[子選單名稱, [[項目, 函式名稱], …]]；只有一項的直接放在主選單。
// 後面步驟的功能「程式存在才會顯示」，所以這裡先全部列好。null 代表分隔線。
const MENU = [
  ['🏠 回到首頁', 'goHome'],
  ['📋 盤點', [
    ['產生盤點表', 'makeCountSheet'],
    ['完成盤點（存入盤點紀錄）', 'finishCount'],
    ['開啟手機盤點網址', 'showMobileLink'],
  ]],
  ['📦 品項', [
    ['新增品項', 'addItemDialog'],
    ['登記異動（新購／領用／報廢／移位）', 'movementDialog'],
    null,
    ['借出登記', 'loanDialog'],
    ['歸還登記', 'returnDialog'],
  ]],
  ['🛒 請購', [
    ['新增請購需求', 'purchaseDialog'],
    ['從需補充清單加入', 'addRestockToPurchase'],
    ['到貨入庫', 'receiveDialog'],
    null,
    ['列印請購清單', 'printPurchaseList'],
  ]],
  ['🗓 實驗室使用', [
    ['登記使用（一週課表）', 'bookDialog'],
    ['實驗室使用一覽', 'usageBoard'],
    ['✏️ 補實驗名稱（一次填好幾堂）', 'nameDialog'],
    ['匯入實驗室課表（老師的 Excel）', 'importScheduleDialog'],
    null,
    ['列印門口海報（A4 橫式）', 'printPosterDialog'],
    ['列印本週課表（A4 直式）', 'printLabWeekDialog'],
    null,
    ['📝 新增準備事項（老師交代的）', 'todoAddDialog'],
    ['📝 準備事項待辦清單', 'todoDialog'],
    ['列印準備事項', 'printTodoDialog'],
    null,
    ['📝 依實驗套組產生準備事項（一段期間）', 'kitTodoDialog'],
    ['列印實驗準備單', 'printPrepDialog'],
    ['實驗套組庫存檢核', 'printKitCheck'],
  ]],
  ['🖨 列印', [
    ['藥品清單', 'printDrugList'],
    ['器材耗材清單', 'printEquipmentList'],
    ['申報清單（期初／期末）', 'printDeclaration'],
    null,
    ['藥品櫃標示', 'printCabinetLabels'],
    ['需補充清單', 'printRestockList'],
  ]],
  ['📊 統計報表', [
    ['各學期耗用量', 'printUsage'],
    ['危險物品統計', 'printHazard'],
    ['永久玻片需求對照', 'printSlides'],
  ]],
  ['🔧 維護', [
    ['立即備份', 'backupNow'],
    ['📦 學期末永久備份（不會被自動刪除）', 'backupForever'],
    ['📧 每天早上 email 提醒（設定）', 'mailSetupDialog'],
    ['首頁、美化與資料保護（一次設定）', 'setupEnhancements'],
    null,
    ['重新計算最新數量', 'updateLatest'],
    ['補上缺少的編號', 'fillMissingIds'],
    ['更新需補充清單', 'refreshRestock'],
    ['套用安全存量建議', 'applySafetyStock'],
    ['套用危險分類建議', 'applyHazardSuggestions'],
    ['產生大包裝建議清單', 'makeBulkSuggestions'],
    ['套用大包裝建議', 'applyBulkSuggestions'],
    ['套用工作表格式與下拉選單', 'setupSheets'],
  ]],
];

const PREFIX = { '藥品': '藥', '器材': '器', '耗材': '耗' };

// ---------------------------------------------------------------- 選單

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const exists = function (fn) { return typeof globalThis[fn] === 'function'; };
  const fill = function (menu, list) {
    let count = 0, lastWasSeparator = true;
    list.forEach(function (m) {
      if (!m) {
        if (!lastWasSeparator) menu.addSeparator();
        lastWasSeparator = true;
      } else if (exists(m[1])) {
        menu.addItem(m[0], m[1]);
        lastWasSeparator = false;
        count++;
      }
    });
    return count;
  };
  const menu = ui.createMenu('🧪 實驗室管理');
  MENU.forEach(function (m) {
    if (Array.isArray(m[1])) {
      const sub = ui.createMenu(m[0]);
      if (fill(sub, m[1])) menu.addSubMenu(sub);
    } else if (exists(m[1])) {
      menu.addItem(m[0], m[1]);
    }
  });
  menu.addToUi();
  checkTimeZone_();
}

const SCHOOL_TZ = 'Asia/Taipei';

/**
 * 試算表時區不是台北時，自動改成台北（從 Excel 轉來的檔案常是別的時區，會讓時間差好幾小時）。
 * 已經存在的日期格子是「日曆日期」，改時區不會讓它們變動。
 */
function checkTimeZone_() {
  const ss = SpreadsheetApp.getActive();
  if (ss.getSpreadsheetTimeZone() === SCHOOL_TZ) return;
  const old = ss.getSpreadsheetTimeZone();
  ss.setSpreadsheetTimeZone(SCHOOL_TZ);
  tzCache_ = null;
  ss.toast('試算表時區原本是「' + old + '」，已自動改為台北時間。', '時區已修正', 10);
}

// ---------------------------------------------------------------- 共用：讀工作表

/**
 * 只讀不寫的動作（開網頁、寄提醒信）期間，同一張工作表只讀一次，加快速度。
 * 用法：withReadMemo_(function () { …只讀取的程式… })。有寫入（ensureSheet_、appendRow_）時會自動清掉。
 */
let readMemo_ = null;
function withReadMemo_(fn) {
  const outer = readMemo_;
  if (!outer) readMemo_ = {};
  try { return fn(); } finally { if (!outer) readMemo_ = null; }
}

/** 讀整張工作表，依第一列標題找欄位，欄位順序日後調整也不受影響。 */
function getTable_(name) {
  if (readMemo_ && readMemo_['t:' + name]) return readMemo_['t:' + name];
  const t = getTable__(name);
  if (readMemo_) readMemo_['t:' + name] = t;
  return t;
}
function getTable__(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error('找不到工作表「' + name + '」，請確認分頁名稱沒有被改掉。');
  const values = sheet.getDataRange().getValues();
  const header = values.shift().map(function (h) { return String(h).trim(); });
  const col = {};
  header.forEach(function (h, i) { if (h && !(h in col)) col[h] = i; });
  return { sheet: sheet, header: header, col: col, rows: values };
}

/** 確認必要欄位都在，缺了就用中文告訴使用者。 */
function need_(table, names) {
  const missing = names.filter(function (n) { return !(n in table.col); });
  if (missing.length) {
    throw new Error('工作表「' + table.sheet.getName() + '」缺少欄位：' + missing.join('、'));
  }
}

/** 讀「設定」：參數（A、B 欄）、各下拉選單清單、清單分區表。 */
function getSettings_() {
  if (readMemo_ && readMemo_.settings) return readMemo_.settings;
  const st = getSettings__();
  if (readMemo_) readMemo_.settings = st;
  return st;
}
function getSettings__() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('設定');
  if (!sheet) throw new Error('找不到工作表「設定」。');
  const values = sheet.getDataRange().getValues();
  const header = values[0].map(function (h) { return String(h).trim(); });

  const params = {};
  for (let r = 1; r < values.length; r++) {
    const k = String(values[r][0]).trim();
    if (k) params[k] = values[r][1];
  }

  const lists = {};
  header.forEach(function (h, c) {
    if (!h || c < 2) return;
    lists[h] = [];
    for (let r = 1; r < values.length; r++) {
      const v = String(values[r][c]).trim();
      if (v) lists[h].push(v);
    }
  });

  const zones = [];
  const zc = header.indexOf('清單分區');
  if (zc >= 0) {
    const names = header.slice(zc, zc + 7);
    for (let r = 1; r < values.length; r++) {
      if (!String(values[r][zc]).trim()) continue;
      const z = {};
      names.forEach(function (n, i) { z[n] = values[r][zc + i]; });
      zones.push(z);
    }
    zones.sort(function (a, b) { return (Number(a['列印順序']) || 99) - (Number(b['列印順序']) || 99); });
  }
  return { params: params, lists: lists, zones: zones, sheet: sheet, header: header };
}

// ---------------------------------------------------------------- 共用：日期、數字、排序

let tzCache_ = null;
function tz_() {
  if (!tzCache_) tzCache_ = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return tzCache_;
}

function pad2_(n) { return (n < 10 ? '0' : '') + n; }

/**
 * 把儲存格的日期轉成「2026-08-26」這種字串（依試算表的時區），方便比較先後。
 * 可接受：日期、「115.08.26」（民國）、「2026/08/26」。認不出來回傳 ''。
 */
function dateKey_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  // 不存在的日期（例：2/30、13 月）回傳 ''
  const ok = function (y, mo, d) { const t = new Date(y, mo - 1, d); return t.getFullYear() === y && t.getMonth() === mo - 1 && t.getDate() === d; };
  let m = s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
  if (m) return ok(+m[1], +m[2], +m[3]) ? m[1] + '-' + pad2_(Number(m[2])) + '-' + pad2_(Number(m[3])) : '';
  m = s.match(/^(\d{2,3})[.\/-](\d{1,2})[.\/-](\d{1,2})$/);
  if (m) return ok(+m[1] + 1911, +m[2], +m[3]) ? (Number(m[1]) + 1911) + '-' + pad2_(Number(m[2])) + '-' + pad2_(Number(m[3])) : '';
  return '';
}

/** 日期（或 dateKey_ 字串）→ 民國文字「115.08.26」。 */
function rocText_(v) {
  const k = /^\d{4}-\d{2}-\d{2}$/.test(String(v)) ? String(v) : dateKey_(v);
  if (!k) return String(v || '');
  return (Number(k.slice(0, 4)) - 1911) + '.' + k.slice(5, 7) + '.' + k.slice(8, 10);
}

/**
 * 要寫進儲存格的文字：開頭是 = + @（或 - 但不是數字）時，試算表會當成公式（例：「+886」變 #ERROR!）。
 * 前面加 ' 讓它當純文字（儲存格裡看不到這個 '）。
 */
function safeCell_(v) {
  if (typeof v !== 'string') return v;
  if (/^[=+@]/.test(v) || (/^-/.test(v) && isNaN(Number(v)))) return "'" + v;
  return v;
}

/** 全形數字、符號轉半形：「３－４」→「3-4」 */
function halfWidth_(s) {
  return String(s == null ? '' : s).replace(/[０-９Ａ-Ｚａ-ｚ－，．]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
}

function isNumber_(v) {
  return v !== '' && v !== null && !isNaN(Number(v));
}

function round_(n) {
  return Math.round(n * 1000) / 1000;
}

/** 自然排序：「3-10」排在「3-9」後面、「器材區2-A」排在「器材區10-A」前面。 */
function naturalCompare_(a, b) {
  const pa = String(a).match(/\d+|\D+/g) || [];
  const pb = String(b).match(/\d+|\D+/g) || [];
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const x = pa[i], y = pb[i];
    if (x === undefined) return -1;
    if (y === undefined) return 1;
    const nx = /^\d+$/.test(x), ny = /^\d+$/.test(y);
    if (nx && ny && Number(x) !== Number(y)) return Number(x) - Number(y);
    if (x !== y) return x < y ? -1 : 1;
  }
  return 0;
}

/** 「生物實驗室一、生物實驗室二」→ ['生物實驗室一', '生物實驗室二'] */
function splitPlaces_(s) {
  return String(s || '').split(/[、,，]/).map(function (x) { return x.trim(); }).filter(String);
}

// ---------------------------------------------------------------- 最新數量

/**
 * 依「盤點紀錄」重新計算每個品項的最新盤點日期、最新數量（多處存放會加總）、最新數量說明。
 * 完成盤點時程式會自動執行；手動改過盤點紀錄後，也可以從選單執行。
 */
function updateLatest() {
  const items = getTable_('品項');
  const recs = getTable_('盤點紀錄');
  need_(items, ['編號', '最新盤點日期', '最新數量', '最新數量說明']);
  need_(recs, ['盤點日期', '編號', '存放處', '數量', '數量說明']);
  const rc = recs.col;

  const latest = {};
  recs.rows.forEach(function (r) {
    const code = String(r[rc['編號']]).trim();
    const k = dateKey_(r[rc['盤點日期']]);
    if (!code || !k) return;
    const cur = latest[code];
    if (!cur || k > cur.k) latest[code] = { k: k, rows: [r] };
    else if (k === cur.k) cur.rows.push(r);
  });

  const dates = [], qtys = [], notes = [];
  items.rows.forEach(function (r) {
    const L = latest[String(r[items.col['編號']]).trim()];
    if (!L) { dates.push(['']); qtys.push(['']); notes.push(['']); return; }
    const nums = L.rows.map(function (x) { return x[rc['數量']]; }).filter(isNumber_);
    const total = nums.length ? round_(nums.reduce(function (a, b) { return a + Number(b); }, 0)) : '';
    let note = '';
    if (L.rows.length === 1) {
      note = String(L.rows[0][rc['數量說明']] || '');
    } else if (L.rows.some(function (x) { return String(x[rc['數量說明']]).trim() !== ''; })) {
      note = L.rows.map(function (x) {
        const place = String(x[rc['存放處']] || '').trim();
        const text = String(x[rc['數量說明']]).trim() || String(x[rc['數量']]);
        return (place ? place + ' ' : '') + text;
      }).join('、');
    }
    dates.push([rocText_(L.k)]);
    qtys.push([total]);
    notes.push([note]);
  });

  const n = items.rows.length;
  if (n) {
    const sh = items.sheet;
    const dateRange = sh.getRange(2, items.col['最新盤點日期'] + 1, n, 1);
    dateRange.setNumberFormat('@').setValues(dates);   // 文字格式，避免「115.08.26」被當成數字
    sh.getRange(2, items.col['最新數量'] + 1, n, 1).setValues(qtys);
    sh.getRange(2, items.col['最新數量說明'] + 1, n, 1).setValues(notes);
  }
  SpreadsheetApp.getActive().toast('已更新 ' + n + ' 個品項的最新數量。', '完成', 5);
}

// ---------------------------------------------------------------- 編號

/** 找出各類別目前用到的最大號碼。 */
function maxIds_(items) {
  const max = {};
  items.rows.forEach(function (r) {
    const m = String(r[items.col['編號']]).match(/^(\D+)(\d+)$/);
    if (m) max[m[1]] = Math.max(max[m[1]] || 0, Number(m[2]));
  });
  return max;
}

/** 在指定列補上編號（有品名、有類別、還沒有編號的才補）。回傳補了幾筆。 */
function assignIds_(items, rowIndexes) {
  need_(items, ['編號', '類別', '品名', '狀態']);
  const max = maxIds_(items);
  let count = 0;
  rowIndexes.forEach(function (i) {
    const r = items.rows[i];
    if (!r || String(r[items.col['編號']]).trim()) return;
    if (!String(r[items.col['品名']]).trim()) return;
    const prefix = PREFIX[String(r[items.col['類別']]).trim()];
    if (!prefix) return;
    max[prefix] = (max[prefix] || 0) + 1;
    const code = prefix + ('000' + max[prefix]).slice(-4);
    items.sheet.getRange(i + 2, items.col['編號'] + 1).setValue(code);
    r[items.col['編號']] = code;
    if (!String(r[items.col['狀態']]).trim()) {
      items.sheet.getRange(i + 2, items.col['狀態'] + 1).setValue('使用中');
    }
    count++;
  });
  return count;
}

/** 選單：把所有缺編號的品項補上編號。 */
function fillMissingIds() {
  const items = getTable_('品項');
  const count = assignIds_(items, items.rows.map(function (_, i) { return i; }));
  SpreadsheetApp.getActive().toast(count ? '補上了 ' + count + ' 個編號。' : '沒有缺編號的品項。', '完成', 5);
}

/**
 * 在「品項」新增一列、填好「類別」和「品名」後，自動給編號、狀態預設「使用中」。
 * （這是 Google 的「簡易觸發條件」，不用另外設定；手機 App 上編輯也會觸發。）
 */
function onEdit(e) {
  if (!e || !e.range) return;
  if (typeof logEdit_ === 'function') logEdit_(e);   // 修改紀錄（07_安全美化）
  const sheet = e.range.getSheet();
  if (sheet.getName() !== '品項' || e.range.getLastRow() < 2) return;
  const first = Math.max(e.range.getRow(), 2);
  const last = Math.min(e.range.getLastRow(), first + 199);
  const items = getTable_('品項');
  const rows = [];
  for (let r = first; r <= last; r++) rows.push(r - 2);
  assignIds_(items, rows);
}

// ---------------------------------------------------------------- 格式與下拉選單

/**
 * 重新套用各工作表的下拉選單、凍結標題列、數字格式、自動欄位灰底。
 * 第一次安裝、或下拉選單的選項（在「設定」）有增減後執行。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActive();
  const settings = getSettings_();

  function listRange(name) {
    const c = settings.header.indexOf(name);
    if (c < 0) throw new Error('「設定」找不到「' + name + '」這一欄。');
    return settings.sheet.getRange(2, c + 1, 60, 1);
  }
  function applyList(table, colName, listName, strict) {
    if (!(colName in table.col)) return;
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(listRange(listName), true)
      .setAllowInvalid(!strict)
      .setHelpText(strict ? '請從下拉選單選擇' : '建議從下拉選單選擇，也可以自行輸入')
      .build();
    table.sheet.getRange(2, table.col[colName] + 1, table.sheet.getMaxRows() - 1, 1).setDataValidation(rule);
  }
  function styleHeader(sheet) {
    const last = sheet.getLastColumn();
    if (!last) return;
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, last).setFontWeight('bold').setBackground('#DDEBF7')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  }

  // 品項
  const items = getTable_('品項');
  styleHeader(items.sheet);
  applyList(items, '類別', '類別', true);
  applyList(items, '科別', '科別', true);
  applyList(items, '清單分區', '清單分區', true);
  applyList(items, '單位', '單位', false);
  applyList(items, '教室', '教室', false);
  applyList(items, 'SDS', '有無', true);
  applyList(items, '危險物品', '是否', true);
  applyList(items, '狀態', '狀態', true);
  const rows = items.sheet.getMaxRows() - 1;
  ['排序位置', '化學式或規格', '最新盤點日期'].forEach(function (n) {
    if (n in items.col) items.sheet.getRange(2, items.col[n] + 1, rows, 1).setNumberFormat('@');
  });
  ['編號', '最新盤點日期', '最新數量', '最新數量說明'].forEach(function (n) {
    if (!(n in items.col)) return;
    items.sheet.getRange(2, items.col[n] + 1, rows, 1).setBackground('#F2F2F2');
    items.sheet.getRange(1, items.col[n] + 1).setNote('程式自動填寫，請勿手動修改。');
  });

  // 盤點紀錄
  const recs = getTable_('盤點紀錄');
  styleHeader(recs.sheet);
  if ('盤點日期' in recs.col) {
    recs.sheet.getRange(2, recs.col['盤點日期'] + 1, recs.sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd');
  }

  // 異動紀錄
  const moves = ss.getSheetByName('異動紀錄') ? getTable_('異動紀錄') : null;
  if (moves) {
    styleHeader(moves.sheet);
    applyList(moves, '類型', '異動類型', true);
    if ('日期' in moves.col) {
      moves.sheet.getRange(2, moves.col['日期'] + 1, moves.sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd');
    }
  }

  ['盤點表', '實驗套組', '玻片需求', '轉入檢查'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastColumn()) styleHeader(sh);
  });

  ss.toast('格式與下拉選單已套用完成。', '完成', 5);
}
