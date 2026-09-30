/**
 * 實驗室藥品器材耗材管理系統 — 請購清單、借用紀錄（第三批）
 *
 * 🛒 請購
 *   ・新增請購需求：老師請買的、不在庫存的都可以登記（可連結現有品項，也可以只寫品名）
 *   ・從需補充清單加入：把低於安全存量、還沒登記請購的品項一次加進來
 *   ・到貨入庫：勾選到貨的項目 → 狀態改「已到貨」；有對應品項的自動登記「新購」並加庫存
 *   ・列印請購清單：待處理＋已請購，附預估金額合計
 *   狀態流程：待處理 → 已請購 → 已到貨（或取消）；「已請購」直接在工作表改狀態即可
 * 🤝 借用
 *   ・借出登記／歸還登記；逾期未還在首頁與「借用紀錄」標紅
 *   ・借出不會改變庫存數量（東西還是學校的），盤點時記得把借出的也算進去
 * 需要「01_基礎」「02_盤點」「04_列印」「08_品項異動」。
 */

const REQ_SHEET = '請購清單';
const REQ_COLS = ['登記日期', '需求來源', '品名', '規格', '數量', '單位', '用途／課程', '需要日期', '對應品項編號',
  '狀態', '請購日期', '到貨日期', '預估單價', '預估金額', '廠商', '備註'];
const REQ_STATUS = ['待處理', '已請購', '已到貨', '取消'];

const LOAN_SHEET = '借用紀錄';
const LOAN_COLS = ['借出日期', '借用人', '品項編號', '品名', '數量', '用途', '預計歸還', '歸還日期', '狀態', '備註'];
const LOAN_STATUS = ['借出中', '已歸還'];

// ---------------------------------------------------------------- 共用：建立工作表、品項搜尋

/** 工作表不存在就建立（含標題、凍結）；存在就補上缺少的欄位。回傳 getTable_ 結果。 */
function ensureSheet_(name, cols, widths, tabColor) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]);
    if (tabColor) sh.setTabColor(tabColor);
    (widths || []).forEach(function (w, i) { sh.setColumnWidth(i + 1, w * 8); });
  } else {
    const head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String);
    cols.forEach(function (c) {
      if (head.indexOf(c) < 0) { sh.getRange(1, head.length + 1).setValue(c); head.push(c); }
    });
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#DDEBF7')
    .setHorizontalAlignment('center').setWrap(true);
  return getTable_(name);
}

/** 同一時間只讓一個存檔動作寫入（手機、電腦同時存時排隊，避免寫到同一列）。 */
function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) throw new Error('系統正在存別的資料，請等幾秒再按一次。');
  try { return fn(); } finally { lock.releaseLock(); }
}

function listRule_(values) {
  return SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(false).build();
}

function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }

function dateValue_(key) { return key ? Utilities.parseDate(key, tz_(), 'yyyy-MM-dd') : ''; }

/** 對話框用的品項清單（搜尋用）。 */
function pickerItems_() {
  const items = getTable_('品項');
  const ic = items.col;
  return items.rows.filter(function (r) { return String(r[ic['編號']]).trim(); }).map(function (r) {
    return {
      code: String(r[ic['編號']]).trim(), name: String(r[ic['品名']]), spec: String(r[ic['化學式或規格']]),
      room: String(r[ic['教室']]), cab: String(r[ic['櫃別']]), pos: String(r[ic['排序位置']]), unit: String(r[ic['單位']]),
      qty: String(r[ic['最新數量']]), note: String(r[ic['最新數量說明']]), status: String(r[ic['狀態']]),
    };
  });
}

/** 對話框裡的品項搜尋元件：需要頁面有 #kw、#res、#picked，並定義 onPick(item)。 */
const PICKER_HTML = `
  <style> #res { max-height: 150px; overflow: auto; border: 1px solid #dadce0; border-radius: 6px; margin-top: 4px; }
    #res:empty { display: none; } #res div { padding: 7px 10px; cursor: pointer; border-bottom: 1px solid #f1f3f4; }
    #res div:hover { background: #e8f0fe; } #res small, #picked small { color: #5f6368; }
    #picked { background: #e8f0fe; padding: 8px 10px; border-radius: 6px; margin-top: 6px; } #picked:empty { display: none; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0 12px; } .grid label { margin-top: 10px; }
    .full { grid-column: 1 / 3; } .req::after { content: " *"; color: #d93025; } </style>
  <script>
    function itemLine(x) { return x.code + '　' + [x.room, x.cab, x.pos].filter(String).join(' ') + '　目前 ' +
      (x.note || x.qty || '—') + (x.unit ? ' ' + x.unit : '') + (x.status === '已淘汰' ? '（已淘汰）' : ''); }
    function search() {
      const k = document.getElementById('kw').value.trim().toLowerCase(), res = document.getElementById('res');
      res.innerHTML = '';
      if (!k) return;
      D.items.filter(function (x) { return (x.name + ' ' + x.spec + ' ' + x.code).toLowerCase().indexOf(k) >= 0; })
        .slice(0, 30).forEach(function (x) {
          const d = document.createElement('div');
          d.appendChild(document.createTextNode(x.name + ' '));
          const sm = document.createElement('small'); sm.textContent = itemLine(x); d.appendChild(sm);
          d.onclick = function () { res.innerHTML = ''; showPicked(x); onPick(x); };
          res.appendChild(d);
        });
    }
    function showPicked(x) {
      const p = document.getElementById('picked'); p.innerHTML = '';
      if (!x) return;
      p.appendChild(document.createTextNode('✔ ' + x.name + ' '));
      const sm = document.createElement('small'); sm.textContent = itemLine(x); p.appendChild(sm);
    }
  </script>`;

// ---------------------------------------------------------------- 🛒 請購：工作表

function setupPurchaseSheet_() {
  const t = ensureSheet_(REQ_SHEET, REQ_COLS, [11, 10, 20, 12, 6, 6, 16, 11, 10, 8, 11, 11, 8, 9, 10, 20], '#FF6D01');
  const sh = t.sheet, c = t.col, n = sh.getMaxRows() - 1;
  sh.getRange(2, c['狀態'] + 1, n, 1).setDataValidation(listRule_(REQ_STATUS));
  ['登記日期', '需要日期', '請購日期', '到貨日期'].forEach(function (k) {
    sh.getRange(2, c[k] + 1, n, 1).setNumberFormat('yyyy/mm/dd');
  });
  // 預估金額＝數量×預估單價（公式放在標題格，整欄自動算）
  const L = function (k) { return colLetter_(c[k] + 1); };
  sh.getRange(1, c['預估金額'] + 1).setFormula('={"預估金額";ARRAYFORMULA(IF((' + L('數量') + '2:' + L('數量') + '="")+(' +
    L('預估單價') + '2:' + L('預估單價') + '=""),"",' + L('數量') + '2:' + L('數量') + '*' + L('預估單價') + '2:' + L('預估單價') + '))}');
  const all = sh.getRange(2, 1, n, sh.getLastColumn());
  const S = '$' + L('狀態'), need = '$' + L('需要日期');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=' + S + '2="已到貨"').setFontColor('#9AA0A6').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=' + S + '2="取消"').setFontColor('#9AA0A6').setStrikethrough(true).setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(' + S + '2="待處理",' + need + '2<>"",' + need + '2-TODAY()<=7)')
      .setBackground('#FCE8E6').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=' + S + '2="已請購"').setBackground('#E8F0FE').setRanges([all]).build(),
  ]);
  return getTable_(REQ_SHEET);
}

/**
 * 在工作表最後一筆資料下面加一列。用第一欄找最後一筆（「預估金額」整欄是公式，getLastRow 會不準）。
 * 公式欄（預估金額）寫空白，不影響整欄公式。回傳寫入的列號。
 */
function appendRow_(t, obj) {
  const row = new Array(t.header.length).fill('');
  Object.keys(obj).forEach(function (k) { if (k in t.col && k !== '預估金額') row[t.col[k]] = safeCell_(obj[k]); });
  const vals = t.sheet.getRange(2, 1, Math.max(1, t.sheet.getMaxRows() - 1), 1).getValues();
  let last = vals.length;
  while (last > 0 && String(vals[last - 1][0]).trim() === '') last--;
  const r = last + 2;
  if (r > t.sheet.getMaxRows()) t.sheet.insertRowsAfter(t.sheet.getMaxRows(), 50);
  t.sheet.getRange(r, 1, 1, row.length).setValues([row]);
  return r;
}

// ---------------------------------------------------------------- 🛒 請購：新增

function purchaseDialog() {
  setupPurchaseSheet_();
  const data = { items: pickerItems_(), today: today_() };
  const html = DIALOG_STYLE + PICKER_HTML + `
    <label>連結現有品項（可不選；新東西直接填下面品名）</label>
    <input id="kw" oninput="search()" placeholder="搜尋品名、化學式或編號"><div id="res"></div><div id="picked"></div>
    <div class="grid">
      <div class="full"><label class="req">品名</label><input id="品名"></div>
      <div><label>規格</label><input id="規格"></div>
      <div><label>單位</label><input id="單位"></div>
      <div><label class="req">數量</label><input id="數量" inputmode="decimal"></div>
      <div><label>預估單價</label><input id="預估單價" inputmode="decimal"></div>
      <div><label>需求來源（老師）</label><input id="需求來源" placeholder="例：王老師"></div>
      <div><label>需要日期</label><input type="date" id="需要日期"></div>
      <div class="full"><label>用途／課程</label><input id="用途／課程" placeholder="例：高二選修 酸鹼滴定"></div>
      <div><label>廠商</label><input id="廠商"></div>
      <div><label>備註</label><input id="備註"></div>
    </div>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()">登記</button></div>
    <script>
      const D = __DATA__; let cur = null;
      function onPick(x) { cur = x; document.getElementById('品名').value = x.name;
        document.getElementById('規格').value = x.spec; document.getElementById('單位').value = x.unit; }
      function go() {
        const f = { code: cur && cur.name === document.getElementById('品名').value.trim() ? cur.code : '' };
        ['品名','規格','單位','數量','預估單價','需求來源','需要日期','用途／課程','廠商','備註']
          .forEach(function (k) { f[k] = document.getElementById(k).value.trim(); });
        if (!f['品名']) { alert('請填品名。'); return; }
        if (!f['數量'] || isNaN(Number(f['數量']))) { alert('數量請填數字。'); return; }
        if (f['預估單價'] && isNaN(Number(f['預估單價']))) { alert('預估單價請填數字。'); return; }
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '登記中…';
        google.script.run.withSuccessHandler(showDone)
          .withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '登記'; }).addPurchase(f);
      }
    </script>`;
  showDialog_(html, data, '🛒 新增請購需求', 600);
}

function addPurchase(f) { return withLock_(function () { return addPurchase__(f); }); }
function addPurchase__(f) {
  const t = setupPurchaseSheet_();
  const r = appendRow_(t, {
    '登記日期': dateValue_(today_()), '需求來源': f['需求來源'], '品名': f['品名'], '規格': f['規格'],
    '數量': Number(f['數量']), '單位': f['單位'], '用途／課程': f['用途／課程'], '需要日期': dateValue_(f['需要日期']),
    '對應品項編號': f.code || '', '狀態': '待處理', '預估單價': f['預估單價'] ? Number(f['預估單價']) : '',
    '廠商': f['廠商'], '備註': f['備註'],
  });
  return '已登記請購：「' + f['品名'] + '」× ' + f['數量'] + (f['單位'] ? ' ' + f['單位'] : '') +
    (f.code ? '（對應品項 ' + f.code + '，到貨時會自動加庫存）' : '（新品項，到貨後請用「新增品項」建檔）') +
    '\n在「請購清單」第 ' + r + ' 列。';
}

/** 把需補充清單中、還沒有「待處理／已請購」的品項加入請購清單。 */
function addRestockToPurchase() { return withLock_(addRestockToPurchase__); }
function addRestockToPurchase__() {
  if (typeof restockItems_ !== 'function') throw new Error('需要先安裝「06_補充查詢」。');
  const t = setupPurchaseSheet_();
  const c = t.col;
  const open = {};
  t.rows.forEach(function (r) {
    const st = String(r[c['狀態']]).trim(), code = String(r[c['對應品項編號']]).trim();
    if (code && (st === '待處理' || st === '已請購')) open[code] = true;
  });
  const add = restockItems_().filter(function (it) { return !open[it['編號']]; });
  add.forEach(function (it) {
    const need = Math.max(1, Math.ceil(Number(it['安全存量']) - Number(it['最新數量'])));
    appendRow_(getTable_(REQ_SHEET), {
      '登記日期': dateValue_(today_()), '需求來源': '需補充清單', '品名': it['品名'], '規格': it['化學式或規格'],
      '數量': need, '單位': it['單位'], '用途／課程': '補足安全存量（目前 ' + (it['最新數量說明'] || numText_(it['最新數量'])) +
        '，安全存量 ' + numText_(it['安全存量']) + '）', '對應品項編號': it['編號'], '狀態': '待處理',
    });
  });
  SpreadsheetApp.getActive().setActiveSheet(t.sheet);
  SpreadsheetApp.getUi().alert(add.length ? '已加入 ' + add.length + ' 項到「請購清單」（數量＝補到安全存量，可自行修改）。' :
    '需補充清單的品項都已經在請購清單裡了（或目前沒有需補充的品項）。');
}

// ---------------------------------------------------------------- 🛒 請購：到貨入庫

function receiveDialog() {
  const t = setupPurchaseSheet_();
  const c = t.col;
  const rows = [];
  t.rows.forEach(function (r, i) {
    const st = String(r[c['狀態']]).trim();
    if (!String(r[c['品名']]).trim() || (st !== '待處理' && st !== '已請購')) return;
    rows.push({ row: i + 2, name: String(r[c['品名']]), spec: String(r[c['規格']]), qty: r[c['數量']], unit: String(r[c['單位']]),
      code: String(r[c['對應品項編號']]).trim(), status: st, who: String(r[c['需求來源']]) });
  });
  if (!rows.length) { SpreadsheetApp.getUi().alert('請購清單沒有「待處理」或「已請購」的項目。'); return; }
  const html = DIALOG_STYLE + `
    <style> table { width: 100%; border-collapse: collapse; font-size: 14px; } td, th { padding: 6px 4px; border-bottom: 1px solid #eee; }
      th { text-align: left; color: #5f6368; font-weight: normal; } td input[type=text] { width: 60px; padding: 4px; }
      small { color: #5f6368; } .new { color: #e37400; } </style>
    <p class="hint">勾選已到貨的項目，數量可改成實際到貨數。有對應品項的會自動登記「新購」並加庫存；
      <span class="new">橘字</span>是還沒建檔的新東西，到貨後請用「新增品項」建檔。</p>
    <table><tr><th></th><th>品名</th><th>到貨數</th></tr><tbody id="tb"></tbody></table>
    <label>到貨日期</label><input type="date" id="date">
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()">到貨入庫</button></div>
    <script>
      const D = __DATA__;
      document.getElementById('date').value = D.today;
      D.rows.forEach(function (r, i) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td><input type="checkbox" id="c' + i + '"></td><td></td><td><input type="text" id="q' + i + '"></td>';
        const td = tr.children[1];
        const nm = document.createElement('span'); nm.textContent = r.name + (r.spec ? '（' + r.spec + '）' : '');
        if (!r.code) nm.className = 'new';
        td.appendChild(nm); td.appendChild(document.createElement('br'));
        const sm = document.createElement('small'); sm.textContent = r.status + '　' + (r.who || '') + '　' + (r.code || '未建檔'); td.appendChild(sm);
        tr.querySelector('#q' + i).value = r.qty;
        document.getElementById('tb').appendChild(tr);
      });
      function go() {
        const list = [];
        D.rows.forEach(function (r, i) {
          if (document.getElementById('c' + i).checked) list.push({ row: r.row, name: r.name, qty: document.getElementById('q' + i).value.trim() });
        });
        if (!list.length) { alert('請勾選到貨的項目。'); return; }
        if (list.some(function (x) { return x.qty === '' || isNaN(Number(x.qty)); })) { alert('到貨數請填數字。'); return; }
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '入庫中…';
        google.script.run.withSuccessHandler(showDone)
          .withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '到貨入庫'; })
          .receivePurchases(list, document.getElementById('date').value);
      }
    </script>`;
  showDialog_(html, { rows: rows, today: today_() }, '📦 到貨入庫', 200 + rows.length * 52);
}

/**
 * 到貨入庫。為了避免「重複入庫」和「列號跑掉入到別的東西」：
 *   ・那一列的品名要對得上、狀態要還是「待處理／已請購」才入庫；
 *   ・列號對不上（中間有人刪列）就用品名找還沒到貨的那一筆（只有一筆時才用）。
 */
function receivePurchases(list, date) {
  date = date || today_();
  const t = getTable_(REQ_SHEET);
  const c = t.col;
  const msgs = [], newOnes = [], skipped = [];
  const open = function (r) { const s = String(r[c['狀態']]).trim(); return s === '待處理' || s === '已請購' || s === ''; };
  const nameOf = function (r) { return String(r[c['品名']]).trim(); };
  let done = 0;
  list.forEach(function (x) {
    let rowNo = Number(x.row), r = t.rows[rowNo - 2];
    const want = String(x.name || '').trim();
    if (!r || (want && nameOf(r) !== want) || !open(r)) {
      if (r && want && nameOf(r) === want && !open(r)) { skipped.push(want + '（已經是「' + String(r[c['狀態']]).trim() + '」，沒有再入庫）'); return; }
      const hits = [];
      t.rows.forEach(function (rr, i) { if (want && nameOf(rr) === want && open(rr)) hits.push(i + 2); });
      if (hits.length !== 1) { skipped.push((want || '第 ' + rowNo + ' 列') + '（找不到還沒到貨的這一筆，請重新開啟到貨入庫）'); return; }
      rowNo = hits[0]; r = t.rows[rowNo - 2];
    }
    x = Object.assign({}, x, { row: rowNo });
    r[c['狀態']] = '已到貨';
    done++;
    const name = String(r[c['品名']]), code = String(r[c['對應品項編號']]).trim();
    t.sheet.getRange(x.row, c['狀態'] + 1).setValue('已到貨');
    t.sheet.getRange(x.row, c['到貨日期'] + 1).setValue(dateValue_(date));
    if (!String(r[c['請購日期']]).trim()) t.sheet.getRange(x.row, c['請購日期'] + 1).setValue(dateValue_(date));
    if (code) {
      const res = recordMovement({ code: code, type: '新購', qty: x.qty, upd: true, date: date,
        note: '請購到貨' + (r[c['用途／課程']] ? '：' + r[c['用途／課程']] : ''), who: String(r[c['需求來源']]) });
      msgs.push('✔ ' + name + ' ' + x.qty + '　' + res.split('\n').slice(1).join(' '));
    } else {
      newOnes.push(name);
    }
  });
  return '到貨入庫完成：' + done + ' 項。\n' + msgs.join('\n') +
    (newOnes.length ? '\n\n還沒建檔的新東西（請用「📦 品項 → 新增品項」建檔並填數量）：\n' + newOnes.join('、') : '') +
    (skipped.length ? '\n\n⚠ 略過：\n' + skipped.join('\n') : '');
}

// ---------------------------------------------------------------- 🛒 請購：列印

function printPurchaseList() {
  setupPurchaseSheet_();
  showPrintDialog({ page: 'purchase' });
}

function page_purchase() {
  const settings = getSettings_();
  const t = getTable_(REQ_SHEET);
  const c = t.col;
  let total = 0, n = 0;
  const groups = { '待處理': [], '已請購': [] };
  t.rows.forEach(function (r) {
    const st = String(r[c['狀態']]).trim();
    if (!String(r[c['品名']]).trim() || !(st in groups)) return;
    groups[st].push(r);
  });
  const rows = [];
  Object.keys(groups).forEach(function (st) {
    if (!groups[st].length) return;
    rows.push('<tr class="grp"><td colspan="9"><b>' + st + '</b>（' + groups[st].length + ' 項）</td></tr>');
    groups[st].forEach(function (r) {
      n++;
      const amt = isNumber_(r[c['數量']]) && isNumber_(r[c['預估單價']]) && String(r[c['預估單價']]).trim() !== ''
        ? Number(r[c['數量']]) * Number(r[c['預估單價']]) : '';
      if (amt !== '') total += amt;
      rows.push('<tr><td class="c">' + n + '</td><td>' + esc_(r[c['品名']]) + (r[c['規格']] ? '<br><small>' + esc_(r[c['規格']]) + '</small>' : '') +
        '</td><td class="c">' + esc_(numText_(r[c['數量']])) + ' ' + esc_(r[c['單位']]) + '</td><td class="s">' + esc_(r[c['用途／課程']]) +
        '</td><td class="c">' + esc_(r[c['需求來源']]) + '</td><td class="c">' + esc_(r[c['需要日期']] ? rocText_(r[c['需要日期']]) : '') +
        '</td><td class="c">' + esc_(numText_(r[c['預估單價']])) + '</td><td class="c">' + (amt === '' ? '' : esc_(amt.toLocaleString())) +
        '</td><td class="s">' + esc_([r[c['廠商']], r[c['備註']]].filter(String).join('；')) + '</td></tr>');
    });
  });
  if (!n) rows.push('<tr><td colspan="9" class="c">目前沒有待處理或已請購的項目</td></tr>');
  else if (total) rows.push('<tr><td colspan="7" class="c"><b>預估金額合計</b></td><td class="c"><b>' + total.toLocaleString() + '</b></td><td></td></tr>');
  const body = reportTable_(settings, {
    title: '實驗室請購需求清單', sub: '', place: '', updated: today_(),
    widths: [5, 22, 10, 20, 9, 9, 7, 8, 10],
    heads: ['序', '品名', '數量', '用途／課程', '需求<br>來源', '需要<br>日期', '預估<br>單價', '預估<br>金額', '廠商／備註'], rows: rows,
  });
  return HtmlService.createHtmlOutput(printShell_('請購清單', body)).setTitle('請購清單（列印）');
}

// ---------------------------------------------------------------- 🤝 借用

function setupLoanSheet_() {
  const t = ensureSheet_(LOAN_SHEET, LOAN_COLS, [11, 10, 9, 20, 6, 16, 11, 11, 8, 20], '#FF6D01');
  const sh = t.sheet, c = t.col, n = sh.getMaxRows() - 1;
  sh.getRange(2, c['狀態'] + 1, n, 1).setDataValidation(listRule_(LOAN_STATUS));
  ['借出日期', '預計歸還', '歸還日期'].forEach(function (k) { sh.getRange(2, c[k] + 1, n, 1).setNumberFormat('yyyy/mm/dd'); });
  const all = sh.getRange(2, 1, n, sh.getLastColumn());
  const S = '$' + colLetter_(c['狀態'] + 1), due = '$' + colLetter_(c['預計歸還'] + 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=' + S + '2="已歸還"').setFontColor('#9AA0A6').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(' + S + '2="借出中",' + due + '2<>"",' + due + '2<TODAY())')
      .setBackground('#FCE8E6').setFontColor('#B3261E').setRanges([all]).build(),
  ]);
  return getTable_(LOAN_SHEET);
}

function loanDialog() {
  setupLoanSheet_();
  const html = DIALOG_STYLE + PICKER_HTML + `
    <label class="req">借出的品項</label>
    <input id="kw" oninput="search()" placeholder="搜尋品名、化學式或編號"><div id="res"></div><div id="picked"></div>
    <div class="grid">
      <div><label class="req">數量</label><input id="qty" inputmode="decimal"></div>
      <div><label class="req">借用人／單位</label><input id="who" placeholder="例：多功能教室4、李老師"></div>
      <div><label>借出日期</label><input type="date" id="date"></div>
      <div><label>預計歸還</label><input type="date" id="due"></div>
      <div class="full"><label>用途</label><input id="use"></div>
      <div class="full"><label>備註</label><input id="note"></div>
    </div>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()">借出</button></div>
    <script>
      const D = __DATA__; let cur = null;
      document.getElementById('date').value = D.today;
      function onPick(x) { cur = x; }
      function go() {
        if (!cur) { alert('請先搜尋並點選品項。'); return; }
        const f = { code: cur.code, name: cur.name, qty: document.getElementById('qty').value.trim(),
          who: document.getElementById('who').value.trim(), date: document.getElementById('date').value,
          due: document.getElementById('due').value, use: document.getElementById('use').value.trim(),
          note: document.getElementById('note').value.trim() };
        if (!f.qty || isNaN(Number(f.qty))) { alert('數量請填數字。'); return; }
        if (!f.who) { alert('請填借用人。'); return; }
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '登記中…';
        google.script.run.withSuccessHandler(showDone)
          .withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '借出'; }).addLoan(f);
      }
    </script>`;
  showDialog_(html, { items: pickerItems_(), today: today_() }, '🤝 借出登記', 480);
}

function addLoan(f) { return withLock_(function () { return addLoan__(f); }); }
function addLoan__(f) {
  const t = setupLoanSheet_();
  appendRow_(t, { '借出日期': dateValue_(f.date || today_()), '借用人': f.who, '品項編號': f.code, '品名': f.name,
    '數量': Number(f.qty), '用途': f.use, '預計歸還': dateValue_(f.due), '狀態': '借出中', '備註': f.note });
  return '已登記借出：「' + f.name + '」× ' + f.qty + ' 給 ' + f.who + (f.due ? '，預計 ' + rocText_(f.due) + ' 歸還' : '') +
    '。\n（借出不會改變庫存數量；盤點時請把借出的也算進去。）';
}

function returnDialog() {
  const t = setupLoanSheet_();
  const c = t.col;
  const rows = [];
  t.rows.forEach(function (r, i) {
    if (String(r[c['狀態']]).trim() !== '借出中') return;
    rows.push({ row: i + 2, name: String(r[c['品名']]), qty: numText_(r[c['數量']]), who: String(r[c['借用人']]),
      date: rocText_(r[c['借出日期']]), due: r[c['預計歸還']] ? rocText_(r[c['預計歸還']]) : '',
      late: r[c['預計歸還']] instanceof Date && dateKey_(r[c['預計歸還']]) < today_() });
  });
  if (!rows.length) { SpreadsheetApp.getUi().alert('目前沒有借出中的項目。'); return; }
  const html = DIALOG_STYLE + `
    <style> .it { padding: 8px 4px; border-bottom: 1px solid #eee; } .it small { color: #5f6368; } .late { color: #d93025; } </style>
    <p class="hint">勾選已經歸還的項目：</p><div id="list"></div>
    <label>歸還日期</label><input type="date" id="date">
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()">歸還</button></div>
    <script>
      const D = __DATA__;
      document.getElementById('date').value = D.today;
      D.rows.forEach(function (r, i) {
        const l = document.createElement('label'); l.className = 'inline it';
        const c = document.createElement('input'); c.type = 'checkbox'; c.id = 'c' + i;
        l.appendChild(c); l.appendChild(document.createTextNode(' ' + r.name + ' × ' + r.qty + '　' + r.who + ' '));
        const sm = document.createElement('small'); sm.textContent = r.date + ' 借出' + (r.due ? '，預計 ' + r.due + ' 還' : '') + (r.late ? '（逾期）' : '');
        if (r.late) sm.className = 'late';
        l.appendChild(sm); document.getElementById('list').appendChild(l);
      });
      function go() {
        const rows = D.rows.filter(function (r, i) { return document.getElementById('c' + i).checked; }).map(function (r) { return { row: r.row, name: r.name }; });
        if (!rows.length) { alert('請勾選歸還的項目。'); return; }
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '登記中…';
        google.script.run.withSuccessHandler(showDone)
          .withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '歸還'; })
          .returnLoans(rows, document.getElementById('date').value);
      }
    </script>`;
  showDialog_(html, { rows: rows, today: today_() }, '🤝 歸還登記', 200 + rows.length * 44);
}

/** 歸還：品名要對得上、狀態要是「借出中」；列號跑掉時用品名找（只有一筆借出中才用）。 */
function returnLoans(rows, date) {
  return withLock_(function () {
    const t = getTable_(LOAN_SHEET);
    const c = t.col;
    const isOut = function (r) { return String(r[c['狀態']]).trim() === '借出中'; };
    let n = 0; const skipped = [];
    rows.forEach(function (x) {
      if (typeof x !== 'object') x = { row: x, name: '' };
      let rowNo = Number(x.row), r = t.rows[rowNo - 2];
      const want = String(x.name || '').trim();
      if (!r || !isOut(r) || (want && String(r[c['品名']]).trim() !== want)) {
        const hits = [];
        t.rows.forEach(function (rr, i) { if (want && String(rr[c['品名']]).trim() === want && isOut(rr)) hits.push(i + 2); });
        if (hits.length !== 1) { skipped.push(want || '第 ' + rowNo + ' 列'); return; }
        rowNo = hits[0]; r = t.rows[rowNo - 2];
      }
      r[c['狀態']] = '已歸還';
      t.sheet.getRange(rowNo, c['狀態'] + 1).setValue('已歸還');
      t.sheet.getRange(rowNo, c['歸還日期'] + 1).setValue(dateValue_(date || today_()));
      n++;
    });
    return '已登記歸還 ' + n + ' 項。' + (skipped.length ? '\n⚠ 略過（已經歸還或找不到）：' + skipped.join('、') : '');
  });
}
