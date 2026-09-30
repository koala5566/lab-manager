/**
 * 實驗室藥品器材耗材管理系統 — 計量方式（大包裝）與「夠不夠」判斷
 *
 * 「品項」多兩欄：
 *   ・計量方式：一般（預設）／大包裝（一包很多、不會去數的：塑膠滴管、夾鏈袋、牙籤…）
 *   ・每包約：知道大概才填，例：「100 支」「約 50 個」；不知道就空白
 * 準備事項、實驗套組檢核都用 stockJudge_ 判斷：
 *   ・單位相同 → 比數量（夠／不夠）
 *   ・有「每包約」而且單位對得上 → 換算後比（只用來估計，不會改庫存）
 *   ・大包裝 → 不比數量，看「有沒有、快不快沒了」（低於安全存量＝快沒了）
 *   ・一般品項單位不同 → 提醒「單位不同，請自行確認」，不亂判不夠
 * 庫存永遠用品項自己的單位顯示（2.5 包，不會變成 2.5 支）。
 * 🔧 維護 → 產生大包裝建議清單 → 勾選 → 套用大包裝建議。
 */

const MEASURE_OPTIONS = ['一般', '大包裝'];
const BULK_SHEET = '大包裝建議';

/** 「品項」沒有「計量方式」「每包約」時補上（加在最右邊），並設下拉選單。 */
function ensureMeasureCols_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('品項');
  if (!sh) return;
  const head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(function (h) { return String(h).trim(); });
  ['計量方式', '每包約'].forEach(function (name) {
    if (head.indexOf(name) >= 0) return;
    const col = head.length + 1;
    sh.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#DDEBF7').setHorizontalAlignment('center');
    head.push(name);
  });
  const mc = head.indexOf('計量方式') + 1, pc = head.indexOf('每包約') + 1, n = sh.getMaxRows() - 1;
  sh.getRange(2, mc, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MEASURE_OPTIONS, true)
    .setAllowInvalid(true).build());
  sh.getRange(1, mc).setNote('一般：數得出來，準備時比數量。\n大包裝：一包很多不會去數（塑膠滴管、夾鏈袋、牙籤…），準備時只看有沒有、快不快沒了。\n空白＝一般。');
  sh.getRange(1, pc).setNote('知道一包大概有幾個才填，例：100 支、約 50 個。只用來估計夠不夠，不會改庫存。');
  sh.getRange(2, pc, n, 1).setNumberFormat('@');
}

/** 「每包約」→ {per: 100, perUnit: '支'}；空白或看不懂回傳 per: 0 */
function parsePer_(v) {
  const m = halfWidth_(v).replace(/^約/, '').trim().match(/^([\d.]+)\s*(.*)$/);
  if (!m || !Number(m[1])) return { per: 0, perUnit: '' };
  return { per: Number(m[1]), perUnit: m[2].replace(/[\/／].*$/, '').trim() };
}

/** 單位同義字：支＝隻＝枝＝根、個＝顆＝只、g＝公克… */
function unitKey_(u) {
  const s = halfWidth_(u).trim().toLowerCase();
  if (!s) return '';
  const groups = [['支', '隻', '枝', '根'], ['個', '顆', '只'], ['g', '公克', '克'], ['kg', '公斤'], ['ml', '毫升', 'cc', 'c.c.'], ['l', '公升'], ['張', '枚']];
  for (let i = 0; i < groups.length; i++) if (groups[i].indexOf(s) >= 0) return groups[i][0];
  return s;
}

/**
 * 夠不夠？s＝品項 {qty, note, unit, measure, per, perUnit, safe}；need＝要幾個；needUnit＝要的單位。
 * 回傳 {state: 'ok'|'warn'|'bad'|'', text: 給人看的一句話, stockText: 「2.5 包」}
 */
function stockJudge_(s, need, needUnit) {
  if (!s) return { state: '', text: '', stockText: '' };
  const u = String(s.unit || '').trim();
  const q = isNumber_(s.qty) && String(s.qty).trim() !== '' ? Number(s.qty) : null;
  const stockText = String(s.note || '').trim() || (q === null ? '—' : numText_(q) + (u ? ' ' + u : ''));
  const r = function (state, text) { return { state: state, text: text, stockText: stockText }; };
  if (q === null) return r('warn', '庫存數量不明（' + stockText + '），請自行確認');
  if (q <= 0) return r('bad', '沒有庫存');
  const n = isNumber_(need) && String(need).trim() !== '' ? Number(need) : null;
  const nu = unitKey_(needUnit), su = unitKey_(u), pu = unitKey_(s.perUnit);
  const safe = isNumber_(s.safe) && String(s.safe).trim() !== '' ? Number(s.safe) : null;
  const low = safe !== null && q < safe;
  const lowText = '快沒了（剩 ' + stockText + (safe !== null ? '，安全存量 ' + numText_(safe) + (u ? ' ' + u : '') : '') + '）';
  const needText = n === null ? '' : numText_(n) + (needUnit ? ' ' + needUnit : '');
  // 單位一樣（或沒寫單位）→ 直接比
  if (n !== null && (!nu || nu === su)) {
    if (q < n) return r('bad', '不夠（要 ' + needText + '，庫存 ' + stockText + '）');
    return low ? r('warn', lowText) : r('ok', '夠（庫存 ' + stockText + '）');
  }
  // 有「每包約」、單位對得上 → 換算估計
  if (n !== null && s.per && (nu === pu || !pu)) {
    const total = Math.round(q * s.per * 10) / 10;
    const tText = '約 ' + numText_(total) + ' ' + (s.perUnit || needUnit || '');
    if (total < n) return r('bad', '可能不夠（要 ' + needText + '，庫存 ' + stockText + '，' + tText + '）');
    return low ? r('warn', lowText + '，' + tText) : r('ok', '夠（' + stockText + '，' + tText + '）');
  }
  // 大包裝 → 不數，看有沒有、快不快沒了
  if (s.measure === '大包裝') {
    if (low) return r('warn', lowText);
    if (safe === null && q < 1) return r('warn', '剩不到 1 ' + (u || '包') + '（' + stockText + '），請確認');
    return r('ok', '有庫存（' + stockText + '，大包裝不細數）');
  }
  if (n === null) return low ? r('warn', lowText) : r('ok', '庫存 ' + stockText);
  return r('warn', '單位不同（要 ' + needText + '、庫存 ' + stockText + '），請自行確認');
}

/** 從品項一列取出計量資訊 */
function measureOf_(r, ic) {
  const p = '每包約' in ic ? parsePer_(r[ic['每包約']]) : { per: 0, perUnit: '' };
  return { measure: '計量方式' in ic && String(r[ic['計量方式']]).trim() === '大包裝' ? '大包裝' : '一般',
    per: p.per, perUnit: p.perUnit, safe: '安全存量' in ic ? r[ic['安全存量']] : '' };
}

// ---------------------------------------------------------------- 大包裝建議清單

const BULK_UNITS = ['包', '小包', '盒', '袋', '捲', '卷', '罐', '箱', '小箱', '打', '捆', '串'];

/** 產生「大包裝建議」工作表：耗材＋單位是包／盒／袋…，預設勾選；「每包約」從規格猜（例：20包/箱）。 */
function makeBulkSuggestions() {
  ensureMeasureCols_();
  const items = getTable_('品項'), ic = items.col;
  const out = [];
  items.rows.forEach(function (r) {
    const code = String(r[ic['編號']]).trim();
    if (!code || String(r[ic['狀態']]).trim() === '已淘汰') return;
    if (String(r[ic['計量方式']]).trim()) return;   // 已經設過的不再建議
    const cat = String(r[ic['類別']]).trim(), unit = String(r[ic['單位']]).trim(), spec = String(r[ic['化學式或規格']] || '').trim();
    if (cat !== '耗材' || BULK_UNITS.indexOf(unit) < 0) return;
    const m = halfWidth_(spec).match(/(\d+)\s*([^\d\s*×x\/／]+)\s*[\/／]\s*(箱|盒|包|袋|罐)/);
    out.push([true, code, String(r[ic['品名']]), spec, unit, String(r[ic['最新數量說明']] || '') || r[ic['最新數量']],
      m ? m[1] + ' ' + m[2] : '', '耗材、以「' + unit + '」計']);
  });
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(BULK_SHEET);
  if (sh) sh.clear(); else sh = ss.insertSheet(BULK_SHEET);
  const head = ['採用', '編號', '品名', '規格', '單位', '最新數量', '每包約（知道才填）', '理由'];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#DDEBF7');
  if (out.length) {
    sh.getRange(2, 7, out.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, out.length, head.length).setValues(out);
    sh.getRange(2, 1, out.length, 1).insertCheckboxes();
    sh.getRange(2, 1, out.length, 1).setBackground('#FFF2CC');
    sh.getRange(2, 7, out.length, 1).setBackground('#FFF2CC');
  }
  [6, 9, 22, 18, 6, 12, 18, 20].forEach(function (w, i) { sh.setColumnWidth(i + 1, w * 8); });
  sh.setFrozenRows(1);
  sh.setTabColor('#9AA0A6');
  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert('大包裝建議', '找到 ' + out.length + ' 項可能是「大包裝」的耗材（都預設勾選）。\n\n' +
    '1. 不是大包裝的（例：一盒一盒數得出來、準備時會用盒為單位的）把「採用」取消勾選\n' +
    '2. 知道一包大概幾個的，可以在「每包約」填，例：100 支\n' +
    '3. 好了之後按「🔧 維護 → 套用大包裝建議」', SpreadsheetApp.getUi().ButtonSet.OK);
}

/** 把「大包裝建議」勾選的寫進品項 */
function applyBulkSuggestions() {
  const sh = SpreadsheetApp.getActive().getSheetByName(BULK_SHEET);
  if (!sh) throw new Error('找不到「' + BULK_SHEET + '」工作表，請先按「產生大包裝建議清單」。');
  ensureMeasureCols_();
  const sug = sh.getDataRange().getValues().slice(1);
  const items = getTable_('品項'), ic = items.col;
  const byCode = {};
  items.rows.forEach(function (r, i) { byCode[String(r[ic['編號']]).trim()] = i + 2; });
  let n = 0, per = 0;
  sug.forEach(function (s) {
    if (s[0] !== true) return;
    const row = byCode[String(s[1]).trim()];
    if (!row) return;
    items.sheet.getRange(row, ic['計量方式'] + 1).setValue('大包裝');
    const p = String(s[6] || '').trim();
    if (p) { items.sheet.getRange(row, ic['每包約'] + 1).setValue(safeCell_(p)); per++; }
    n++;
  });
  SpreadsheetApp.getUi().alert('套用完成', '已把 ' + n + ' 項設為「大包裝」' + (per ? '（其中 ' + per + ' 項有填每包約）' : '') + '。\n' +
    '「大包裝建議」工作表可以刪掉了。之後要改，直接改「品項」的「計量方式」「每包約」兩欄。', SpreadsheetApp.getUi().ButtonSet.OK);
  return n;
}
