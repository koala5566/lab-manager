/**
 * 實驗室藥品器材耗材管理系統 — 盤點功能（第 3 步）
 *
 * 選單「產生盤點表」：選範圍與日期 → 產生「盤點表」（附上次數量、空白的本次欄）
 * 選單「完成盤點」：把已填的數量存進「盤點紀錄」，存好的列從盤點表移除，沒填的留著下次繼續。
 * 需要「01_基礎」裡的共用函式。
 */

const COUNT_HEADER = ['教室', '櫃別', '排序位置', '存放處', '品名', '化學式或規格', '單位',
  '上次日期', '上次數量', '上次說明', '本次數量', '本次說明', '已盤', '編號', '清單分區'];
const COUNT_INPUT = ['本次數量', '本次說明'];
const SUM_MARK = '（上次為各處合計）';

// ---------------------------------------------------------------- 產生盤點表

function makeCountSheet() {
  const settings = getSettings_();
  const pending = countPending_();
  const now = new Date();
  const data = {
    lists: {
      '類別': settings.lists['類別'] || [],
      '科別': settings.lists['科別'] || [],
      '清單分區': settings.zones.map(function (z) { return String(z['清單分區']); }),
      '教室': settings.lists['教室'] || [],
    },
    date: Utilities.formatDate(now, tz_(), 'yyyy-MM-dd'),
    name: defaultCountName_(settings, now),
    pending: pending,
  };
  const html = DIALOG_STYLE + `
    <label>盤點範圍</label>
    <div class="row">
      <select id="type" onchange="fillValues()">
        <option value="全部">全部品項</option>
        <option>類別</option><option>科別</option><option>清單分區</option><option>教室</option>
      </select>
      <select id="val"></select>
    </div>
    <label>盤點日期</label><input type="date" id="date">
    <label>盤點名稱</label><input id="name">
    <div id="warn" class="warn"></div>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()">產生盤點表</button></div>
    <script>
      const D = __DATA__;
      document.getElementById('date').value = D.date;
      document.getElementById('name').value = D.name;
      if (D.pending.total) {
        document.getElementById('warn').innerHTML = '盤點表目前還有 ' + D.pending.total + ' 列未儲存（其中 ' +
          D.pending.filled + ' 列已填數量）。產生新的盤點表會把它們清除。<br>' +
          '<label class="inline"><input type="checkbox" id="ok"> 我確定要清除</label>';
      }
      function fillValues() {
        const t = document.getElementById('type').value, v = document.getElementById('val');
        v.innerHTML = '';
        v.style.display = t === '全部' ? 'none' : '';
        (D.lists[t] || []).forEach(function (x) { const o = document.createElement('option'); o.text = x; v.add(o); });
      }
      fillValues();
      function go() {
        const ok = document.getElementById('ok');
        if (ok && !ok.checked) { alert('請先勾選「我確定要清除」，或按取消後先執行「完成盤點」。'); return; }
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '產生中…';
        google.script.run
          .withSuccessHandler(function () { google.script.host.close(); })
          .withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '產生盤點表'; })
          .createCountSheet({ type: document.getElementById('type').value, value: document.getElementById('val').value,
            date: document.getElementById('date').value, name: document.getElementById('name').value });
      }
    </script>`;
  showDialog_(html, data, '產生盤點表', 420);
}

/** 對話框按下「產生盤點表」後執行。 */
function createCountSheet(opts) {
  if (!opts.date) throw new Error('請選擇盤點日期。');
  const items = getTable_('品項');
  need_(items, ['編號', '類別', '科別', '清單分區', '品名', '化學式或規格', '單位', '教室', '櫃別',
    '排序位置', '分處存放', '狀態']);
  const ic = items.col;
  const rooms = getSettings_().lists['教室'] || [];
  const last = lastCounts_();

  const rows = [];
  items.rows.forEach(function (r) {
    const code = String(r[ic['編號']]).trim();
    if (!code || String(r[ic['狀態']]).trim() === '已淘汰') return;
    if (opts.type !== '全部' && opts.type !== '教室' && String(r[ic[opts.type]]).trim() !== opts.value) return;
    const places = splitPlaces_(r[ic['分處存放']]);
    const L = last[code] || { byPlace: {} };
    let hinted = false;
    (places.length ? places : ['']).forEach(function (place) {
      // 存放處是教室名稱（如 生物實驗室一）就放在教室欄排序；否則（如 木櫃、冰箱）放在櫃別
      const isRoom = place && rooms.some(function (x) { return place.indexOf(x) === 0; });
      const room = isRoom ? place : r[ic['教室']];
      const cab = place && !isRoom ? place : r[ic['櫃別']];
      // 依教室盤點時，分處存放的品項只列出在這間教室的那一處
      if (opts.type === '教室' && String(room).indexOf(opts.value) !== 0) return;
      let prev = L.byPlace[place];
      if (!prev && place && !hinted && L.byPlace['']) {
        prev = { q: L.byPlace[''].q, note: SUM_MARK };   // 上次只有合計，放在列出的第一處提示
        hinted = true;
      }
      rows.push([room, cab, String(r[ic['排序位置']]), place, r[ic['品名']], r[ic['化學式或規格']], r[ic['單位']],
        prev ? rocText_(L.k) : '', prev ? prev.q : '', prev ? prev.note : '', '', '', '',
        code, r[ic['清單分區']]]);
    });
  });
  if (!rows.length) throw new Error('這個範圍沒有「使用中」的品項。');
  // 教室依「設定」的教室順序（化學準備室、化學實驗室一…），其餘依櫃別、排序位置
  const roomRank = function (room) {
    for (let i = 0; i < rooms.length; i++) if (String(room).indexOf(rooms[i]) === 0) return i;
    return rooms.length;
  };
  rows.sort(function (a, b) {
    return roomRank(a[0]) - roomRank(b[0]) || naturalCompare_(a[0], b[0]) || naturalCompare_(a[1], b[1]) ||
      naturalCompare_(a[2], b[2]) || naturalCompare_(a[13], b[13]) || naturalCompare_(a[3], b[3]);
  });

  const scope = opts.type === '全部' ? '全部品項' : opts.type + '：' + opts.value;
  PropertiesService.getDocumentProperties().setProperty('count',
    JSON.stringify({ date: opts.date, name: opts.name, scope: scope }));
  writeCountSheet_(rows);
  const sh = SpreadsheetApp.getActive().getSheetByName('盤點表');
  SpreadsheetApp.getActive().setActiveSheet(sh);
  SpreadsheetApp.getActive().toast('已產生 ' + rows.length + ' 列。黃色欄位填數量，填完從選單按「完成盤點」。',
    opts.name + '（' + scope + '）', 10);
}

/** 每個品項最近一次盤點：{編號: {k: 日期, byPlace: {存放處: {q, note}}}} */
function lastCounts_() {
  const recs = getTable_('盤點紀錄');
  need_(recs, ['盤點日期', '編號', '存放處', '數量', '數量說明']);
  const rc = recs.col, out = {};
  recs.rows.forEach(function (r) {
    const code = String(r[rc['編號']]).trim(), k = dateKey_(r[rc['盤點日期']]);
    if (!code || !k) return;
    if (!out[code] || k > out[code].k) out[code] = { k: k, byPlace: {} };
    if (k === out[code].k) {
      out[code].byPlace[String(r[rc['存放處']]).trim()] = { q: r[rc['數量']], note: String(r[rc['數量說明']]) };
    }
  });
  return out;
}

/** 把列資料寫進「盤點表」，並設定顏色、格式。 */
function writeCountSheet_(rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('盤點表') || ss.insertSheet('盤點表');
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations();
  const n = COUNT_HEADER.length;
  sh.getRange(1, 1, 1, n).setValues([COUNT_HEADER]).setFontWeight('bold').setBackground('#DDEBF7')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setFrozenRows(1);
  const info = JSON.parse(PropertiesService.getDocumentProperties().getProperty('count') || '{}');
  sh.getRange(1, 1).setNote(info.date ? '本次盤點：' + (info.name || '') + '　日期：' + rocText_(info.date) +
    '　範圍：' + (info.scope || '') : '');
  if (!rows.length) return sh;

  const col = function (name) { return COUNT_HEADER.indexOf(name) + 1; };
  if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
  sh.getRange(2, col('排序位置'), rows.length, 1).setNumberFormat('@');
  sh.getRange(2, col('上次日期'), rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, n).setValues(rows);
  sh.getRange(2, col('已盤'), rows.length, 1).setFormulaR1C1('=IF(AND(RC[-2]="",RC[-1]=""),"","✔")');
  sh.getRange(2, 1, rows.length, n).setBackground('#F2F2F2').setVerticalAlignment('middle');
  sh.getRange(2, col('本次數量'), rows.length, 2).setBackground('#FFF2CC');
  sh.getRange(1, col('本次數量'), 1, 2).setBackground('#FFD966');

  const done = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$M2="✔"').setBackground('#E2EFDA')
    .setRanges([sh.getRange(2, 1, rows.length, n)]).build();
  sh.setConditionalFormatRules([done]);

  [10, 12, 8, 12, 22, 16, 5, 9, 8, 14, 9, 14, 5, 8, 10].forEach(function (w, i) {
    sh.setColumnWidth(i + 1, w * 8);
  });
  return sh;
}

// ---------------------------------------------------------------- 完成盤點

function finishCount() {
  const info = JSON.parse(PropertiesService.getDocumentProperties().getProperty('count') || '{}');
  const p = countPending_();
  if (!p.total) {
    SpreadsheetApp.getUi().alert('盤點表是空的。請先從選單「產生盤點表」。');
    return;
  }
  const data = { date: info.date || Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'),
    name: info.name || '', scope: info.scope || '', total: p.total, filled: p.filled, carry: p.carry };
  const html = DIALOG_STYLE + `
    <p>範圍：<b id="scope"></b><br>已填 <b id="filled"></b> 列／共 <b id="total"></b> 列</p>
    <label>盤點日期</label><input type="date" id="date">
    <label>盤點名稱</label><input id="name">
    <label>還沒填的列</label>
    <label class="inline"><input type="radio" name="mode" value="keep" checked> 留在盤點表，下次繼續盤</label>
    <label class="inline"><input type="radio" name="mode" value="carry"> 沿用上次數量一起存（表示沒有變動）<span id="carry"></span></label>
    <p class="hint">同一天、同一品項、同一存放處若已經有紀錄，會以這次為準（覆蓋）。</p>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()">存入盤點紀錄</button></div>
    <script>
      const D = __DATA__;
      document.getElementById('scope').textContent = D.scope;
      document.getElementById('filled').textContent = D.filled;
      document.getElementById('total').textContent = D.total;
      document.getElementById('date').value = D.date;
      document.getElementById('name').value = D.name;
      document.getElementById('carry').textContent = '（可沿用 ' + D.carry + ' 列）';
      function go() {
        const mode = document.querySelector('input[name=mode]:checked').value;
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '儲存中…';
        google.script.run
          .withSuccessHandler(function (msg) { alert(msg); google.script.host.close(); })
          .withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '存入盤點紀錄'; })
          .saveCount({ date: document.getElementById('date').value, name: document.getElementById('name').value, mode: mode });
      }
    </script>`;
  showDialog_(html, data, '完成盤點', 440);
}

/** 對話框按下「存入盤點紀錄」後執行。回傳給使用者看的結果訊息。 */
function saveCount(opts) {
  if (!opts.date) throw new Error('請選擇盤點日期。');
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const t = getTable_('盤點表');
    need_(t, COUNT_HEADER);
    const c = t.col;
    const dateValue = Utilities.parseDate(opts.date, tz_(), 'yyyy-MM-dd');
    const stamp = Utilities.formatDate(new Date(), SCHOOL_TZ, 'yyyy/MM/dd HH:mm');   // 登錄時間一律用台灣時間

    const save = [], keep = [];
    t.rows.forEach(function (r) {
      if (!String(r[c['編號']]).trim()) return;
      let q = parseQty_(r[c['本次數量']]);
      const note = String(r[c['本次說明']]).trim();
      if (note) q = { n: q.n !== '' ? q.n : parseQty_(note).n, note: note };
      if (q.n === '' && !q.note && opts.mode === 'carry' && String(r[c['上次日期']]).trim() &&
          String(r[c['上次說明']]) !== SUM_MARK) {
        q = { n: r[c['上次數量']], note: String(r[c['上次說明']]) };
      }
      if (q.n === '' && !q.note) { keep.push(r.slice(0, COUNT_HEADER.length)); return; }
      save.push([dateValue, opts.name, String(r[c['編號']]).trim(), r[c['品名']], String(r[c['存放處']]).trim(),
        q.n, q.note, stamp]);
    });
    if (!save.length) throw new Error('沒有已填數量的列可以儲存。');

    const replaced = appendRecords_(save, opts.date);
    keep.forEach(function (r) { r[COUNT_HEADER.indexOf('已盤')] = ''; });
    if (!keep.length) PropertiesService.getDocumentProperties().deleteProperty('count');
    writeCountSheet_(keep);
    updateLatest();
    if (typeof refreshRestock === 'function') refreshRestock();

    return '已存入 ' + save.length + ' 筆盤點紀錄' + (replaced ? '（覆蓋同日舊紀錄 ' + replaced + ' 筆）' : '') +
      '。\n' + (keep.length ? '盤點表還有 ' + keep.length + ' 列沒填，已保留。' : '本次盤點全部完成！');
  } finally {
    lock.releaseLock();
  }
}

/** 加到「盤點紀錄」；同日、同編號、同存放處的舊紀錄先移除。回傳移除筆數。 */
function appendRecords_(save, dateKey) {
  const recs = getTable_('盤點紀錄');
  need_(recs, ['盤點日期', '盤點名稱', '編號', '品名', '存放處', '數量', '數量說明', '登錄時間']);
  const rc = recs.col, sh = recs.sheet;
  const newKeys = {};
  save.forEach(function (s) { newKeys[s[2] + '|' + s[4]] = true; });
  const dupRows = [];
  recs.rows.forEach(function (r, i) {
    if (dateKey_(r[rc['盤點日期']]) === dateKey &&
        newKeys[String(r[rc['編號']]).trim() + '|' + String(r[rc['存放處']]).trim()]) dupRows.push(i + 2);
  });
  for (let i = dupRows.length - 1; i >= 0; i--) sh.deleteRow(dupRows[i]);

  const order = ['盤點日期', '盤點名稱', '編號', '品名', '存放處', '數量', '數量說明', '登錄時間'];
  const width = recs.header.length;
  const out = save.map(function (s) {
    const row = new Array(width).fill('');
    order.forEach(function (name, j) { row[rc[name]] = s[j]; });
    return row;
  });
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, out.length, width).setValues(out);
  sh.getRange(start, rc['盤點日期'] + 1, out.length, 1).setNumberFormat('yyyy/mm/dd');
  return dupRows.length;
}

// ---------------------------------------------------------------- 共用

/** 盤點表目前的狀況：共幾列、已填幾列、可沿用幾列。 */
function countPending_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('盤點表');
  if (!sh || sh.getLastRow() < 2) return { total: 0, filled: 0, carry: 0 };
  const t = getTable_('盤點表');
  if (!('編號' in t.col) || !('本次數量' in t.col)) return { total: 0, filled: 0, carry: 0 };
  let total = 0, filled = 0, carry = 0;
  t.rows.forEach(function (r) {
    if (!String(r[t.col['編號']]).trim()) return;
    total++;
    const has = String(r[t.col['本次數量']]).trim() || String(r[t.col['本次說明']]).trim();
    if (has) filled++;
    else if (String(r[t.col['上次日期']]).trim() && String(r[t.col['上次說明']]) !== SUM_MARK) carry++;
  });
  return { total: total, filled: filled, carry: carry };
}

/**
 * 數量文字 → {n: 數字, note: 原文}。純數字時 note 為空。
 * 「3箱+5串」→ 3；「100g*2+50g」→ 250；「1.5粉狀 6.5固體顆粒」→ 8（規則同舊資料轉入）。
 */
function parseQty_(v) {
  if (v === null || v === undefined || String(v).trim() === '') return { n: '', note: '' };
  if (isNumber_(v)) return { n: Number(v), note: '' };
  const text = String(v).replace(/\s+/g, ' ').trim();
  const re = /(\d+(?:\.\d+)?)\s*([^\d\s+*()（）]*)\s*(?:\*\s*(\d+(?:\.\d+)?))?/g;
  const units = [], vals = [];
  let m;
  while ((m = re.exec(text)) !== null) {
    units.push(m[2]);
    vals.push(Number(m[1]) * (m[3] ? Number(m[3]) : 1));
  }
  if (!vals.length) return { n: '', note: text };
  const forms = ['粉狀', '固體顆粒', '固', '液'];
  const same = units.every(function (u) { return u === units[0]; }) ||
    units.every(function (u) { return forms.indexOf(u) >= 0; });
  const n = same ? vals.reduce(function (a, b) { return a + b; }, 0) : vals[0];
  return { n: round_(n), note: text };
}

/** 依今天日期猜盤點名稱，例如 115-1期中。 */
function defaultCountName_(settings, d) {
  const m = d.getMonth() + 1;
  const stage = { 8: '期初', 9: '期初', 10: '期中', 11: '期中', 12: '期中', 1: '期末',
    2: '期初', 3: '期初', 4: '期中', 5: '期中', 6: '期末', 7: '期末' }[m];
  return String(settings.params['學年度學期'] || '') + stage;
}

const DIALOG_STYLE = `<style>
  body { font-family: "Noto Sans TC", "Microsoft JhengHei", sans-serif; font-size: 14px; color: #222; }
  label { display: block; margin: 12px 0 4px; font-weight: bold; }
  label.inline { font-weight: normal; margin: 6px 0; }
  select, input:not([type=radio]):not([type=checkbox]) { width: 100%; padding: 6px; font-size: 14px; box-sizing: border-box; }
  .row { display: flex; gap: 8px; }
  .warn { color: #b00020; margin-top: 12px; }
  .hint { color: #666; font-size: 12px; }
  .btns { margin-top: 18px; text-align: right; }
  button { padding: 8px 16px; font-size: 14px; margin-left: 8px; }
  button.primary { background: #1a73e8; color: #fff; border: none; border-radius: 4px; }
</style>`;

/** 顯示對話框，data 以 JSON 放進頁面的 __DATA__。 */
function showDialog_(html, data, title, height) {
  const json = JSON.stringify(data).replace(/</g, '\\u003c');
  const out = HtmlService.createHtmlOutput(html.replace('__DATA__', function () { return json; }))
    .setWidth(420).setHeight(height);
  SpreadsheetApp.getUi().showModalDialog(out, title);
}
