/**
 * 實驗室藥品器材耗材管理系統 — 網頁工作台（手機、平板、電腦共用）
 *
 * 開網頁（部署的網址）就是這個畫面；舊版手機網頁在網址後面加 ?page=old。
 *   ・今天：現在第幾節、哪幾間在上課、今天和下一個上課日的課、待準備、需補充（電腦、平板是儀表板）
 *   ・課表：各實驗室單日／整週／月曆／清單；點一堂課看詳細（器材夠不夠、準備事項、補實驗名稱、準備單）
 *   ・待辦：準備事項打勾、新增（可對到某堂課、品項搜尋）
 *   ・盤點：一次一項、大按鈕 −／＋、存好自動跳下一項
 *   ・更多：需補充、品項查詢、外觀（淺色／深色）
 * 畫面程式由 tools/webapp/ 的原始檔產生（build_webapp.py），請不要直接改下面的 APP_HTML。
 * 需要「01」～「15」；「18」的一鍵產生準備事項也會用到這裡。
 */

function page_app() {
  return HtmlService.createHtmlOutput(APP_HTML)
    .setTitle('實驗室管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

/** 舊版手機網頁（?page=old） */
function page_old(p) { return oldMobilePage_(p); }

/** 一開網頁要的資料，一次給齊（之後只在需要時補抓）。同一張工作表只讀一次（withReadMemo_）。 */
function appInit() { return withReadMemo_(appInit_); }
function appInit_() {
  const cfg = labConfig_();
  const s = getSettings_();
  const today = today_(), next = nextSchoolDay_(today);
  const books = bookings_(today, next, cfg);
  const byDay = function (k) {
    return books.filter(function (b) { return b.date === k; }).sort(function (a, b) {
      return (a.pis[0] === undefined ? 99 : a.pis[0]) - (b.pis[0] === undefined ? 99 : b.pis[0]) || naturalCompare_(a.lab, b.lab);
    });
  };
  const td = typeof todoData === 'function' ? todoData() : { list: [], items: [] };
  let restock = null;
  try { if (typeof restockItems_ === 'function') restock = restockItems_().length; } catch (e) { restock = null; }
  const count = appCountInfo_();
  return {
    school: String(s.params['學校名稱'] || ''), term: String(s.params['學年度學期'] || ''),
    rooms: cfg.rooms, periods: cfg.periods, colors: cfg.colors, types: cfg.types, classNames: cfg.classNames,
    today: { key: today, roc: rocText_(today), wd: weekdayOf_(today), list: byDay(today) },
    next: { key: next, roc: rocText_(next), wd: weekdayOf_(next), list: byDay(next) },
    todos: td.list, items: td.items,
    kpi: { restock: restock, loans: appSheetCount_(LOAN_SHEET, '狀態', ['借出中'], '預計歸還'),
      purchase: appSheetCount_(REQ_SHEET, '狀態', ['待處理', '已請購']), count: count },
    url: ScriptApp.getService().getUrl() || '',
  };
}

/** 某工作表某欄是某些值的列數；dueCol 有給時也算逾期幾筆。 */
function appSheetCount_(name, col, values, dueCol) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return { n: 0, late: 0 };
  const t = getTable_(name), c = t.col;
  if (!(col in c)) return { n: 0, late: 0 };
  const today = today_();
  let n = 0, late = 0;
  t.rows.forEach(function (r) {
    if (values.indexOf(String(r[c[col]]).trim()) < 0) return;
    n++;
    if (dueCol && dueCol in c) { const d = dateKey_(r[c[dueCol]]); if (d && d < today) late++; }
  });
  return { n: n, late: late };
}

/** 盤點表進度：{name, date, total, done} */
function appCountInfo_() {
  const info = JSON.parse(PropertiesService.getDocumentProperties().getProperty('count') || '{}');
  const sh = SpreadsheetApp.getActive().getSheetByName('盤點表');
  if (!sh || sh.getLastRow() < 2) return { name: info.name || '', total: 0, done: 0 };
  const t = getTable_('盤點表'), c = t.col;
  let total = 0, done = 0;
  t.rows.forEach(function (r) {
    if (!String(r[c['編號']]).trim()) return;
    total++;
    if (String(r[c['本次數量']]).trim() !== '' || String(r[c['本次說明']]).trim() !== '') done++;
  });
  return { name: info.name || '', date: info.date ? rocText_(info.date) : '', total: total, done: done };
}

/** 一堂課的器材檢查（實驗名稱要有對應的實驗套組）。 */
function appLessonKit(name, groups) {
  if (!name || typeof experimentCheck_ !== 'function') return null;
  const ck = experimentCheck_(name, Number(groups) || 0);
  if (!ck) return null;
  return { groups: ck.groups, rows: ck.rows.map(function (x) {
    return { name: x.name, per: x.per, need: x.need, have: x.haveText, unit: x.unit, loc: x.loc, short: x.short, note: x.note,
      state: x.state || (x.short ? 'bad' : 'ok'), bulk: !!x.bulk };
  }) };
}

/** 重新整理用：今天頁、待辦的資料 */
function appRefresh() { return appInit(); }
