/**
 * 實驗室藥品器材耗材管理系統 — 新增品項、登記異動（第二批優化）
 *
 * 選單「📦 品項 → 新增品項」：表單填寫，自動編號、檢查必填、重複品名提醒；可順便填目前數量。
 * 選單「📦 品項 → 登記異動」：新購／報廢／領用／移位
 *   ・一律記進「異動紀錄」
 *   ・新購／領用／報廢：可勾「同時更新目前數量」（在盤點紀錄加一筆：上次數量 ± 數量）
 *   ・移位：直接改「品項」的教室、櫃別、排序位置
 *   ・報廢：可勾「整個品項淘汰」（狀態改為已淘汰）
 * 需要「01_基礎」「02_盤點」。
 */

// ---------------------------------------------------------------- 新增品項

function addItemDialog() {
  const s = getSettings_();
  const items = getTable_('品項');
  const ic = items.col;
  const cabs = {};
  const names = [];
  items.rows.forEach(function (r) {
    const c = String(r[ic['櫃別']]).trim();
    if (c) cabs[c] = true;
    names.push(String(r[ic['品名']]).trim());
  });
  const data = {
    lists: {
      '類別': s.lists['類別'] || [], '科別': s.lists['科別'] || [], '單位': s.lists['單位'] || [],
      '教室': s.lists['教室'] || [], '清單分區': s.zones.map(function (z) { return String(z['清單分區']); }),
    },
    cabs: Object.keys(cabs).sort(naturalCompare_), names: names,
    today: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'),
  };
  const html = DIALOG_STYLE + `
    <style> .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0 12px; } .grid label { margin-top: 10px; }
      .req::after { content: " *"; color: #d93025; } .full { grid-column: 1 / 3; } </style>
    <div class="grid">
      <div><label class="req">類別</label><select id="類別"></select></div>
      <div><label>科別</label><select id="科別"></select></div>
      <div class="full"><label class="req">清單分區</label><select id="清單分區"></select></div>
      <div class="full"><label class="req">品名</label><input id="品名" oninput="dupCheck()"><div id="dup" class="hint"></div></div>
      <div><label>化學式或規格</label><input id="化學式或規格"></div>
      <div><label>單位</label><input id="單位" list="units"></div>
      <div><label>教室</label><select id="教室"></select></div>
      <div><label>櫃別</label><input id="櫃別" list="cabs"></div>
      <div><label>排序位置</label><input id="排序位置" placeholder="例：3-2"></div>
      <div><label>安全存量</label><input id="安全存量" inputmode="decimal"></div>
      <div class="full"><label>分處存放（放好幾處時才填，用頓號分隔）</label><input id="分處存放" placeholder="例：化學準備室、生物準備室冰箱"></div>
      <div><label>SDS</label><select id="SDS"><option></option><option>有</option><option>無</option></select></div>
      <div><label>危險物品</label><select id="危險物品"><option></option><option>是</option><option>否</option></select></div>
      <div class="full"><label>備註（會列印）</label><input id="備註"></div>
      <div class="full"><label>自己筆記（不列印）</label><input id="自己筆記"></div>
      <div><label>目前數量（可空白）</label><input id="數量" placeholder="例：5 或 3箱+5串"></div>
      <div><label>數量日期</label><input type="date" id="日期"></div>
    </div>
    <datalist id="units"></datalist><datalist id="cabs"></datalist>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()">新增</button></div>
    <script>
      const D = __DATA__;
      function fill(id, list, blank) { const s = document.getElementById(id); if (blank) s.add(new Option('', ''));
        list.forEach(function (x) { s.add(new Option(x, x)); }); }
      fill('類別', D.lists['類別']); fill('科別', D.lists['科別'], true); fill('清單分區', D.lists['清單分區']);
      fill('教室', D.lists['教室'], true);
      D.lists['單位'].forEach(function (x) { const o = document.createElement('option'); o.value = x; document.getElementById('units').appendChild(o); });
      D.cabs.forEach(function (x) { const o = document.createElement('option'); o.value = x; document.getElementById('cabs').appendChild(o); });
      document.getElementById('日期').value = D.today;
      function dupCheck() {
        const v = document.getElementById('品名').value.trim();
        const hit = v && D.names.filter(function (n) { return n === v; }).length;
        document.getElementById('dup').textContent = hit ? '⚠ 已經有同名的品項（' + hit + ' 筆），確定要再新增一筆嗎？' : '';
      }
      function go() {
        const f = {};
        ['類別','科別','清單分區','品名','化學式或規格','單位','教室','櫃別','排序位置','安全存量','分處存放','SDS','危險物品',
         '備註','自己筆記','數量','日期'].forEach(function (k) { f[k] = document.getElementById(k).value.trim(); });
        if (!f['品名']) { alert('請填品名。'); return; }
        if (f['安全存量'] && isNaN(Number(f['安全存量']))) { alert('安全存量請填數字。'); return; }
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '新增中…';
        google.script.run.withSuccessHandler(showDone)
          .withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '新增'; }).addItem(f);
      }
    </script>`;
  showDialog_(html, data, '➕ 新增品項', 640);
}

/** 表單送出：寫進「品項」最後一列，自動編號；有填數量就加一筆盤點紀錄。 */
function addItem(f) {
  if (!f['品名'] || !f['類別'] || !f['清單分區']) throw new Error('類別、清單分區、品名都要填。');
  const prefix = PREFIX[f['類別']];
  if (!prefix) throw new Error('類別只能是藥品、器材或耗材。');
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const items = getTable_('品項');
    need_(items, ['編號', '類別', '品名', '狀態']);
    const max = maxIds_(items);
    const code = prefix + ('000' + ((max[prefix] || 0) + 1)).slice(-4);
    const row = new Array(items.header.length).fill('');
    const put = function (name, v) { if (name in items.col) row[items.col[name]] = v; };
    ['類別', '科別', '清單分區', '品名', '化學式或規格', '單位', '教室', '櫃別', '排序位置', '分處存放', 'SDS', '危險物品',
      '備註', '自己筆記'].forEach(function (k) { put(k, f[k] || ''); });
    put('編號', code);
    put('狀態', '使用中');
    if (f['安全存量'] !== '') put('安全存量', Number(f['安全存量']));
    const r = items.rows.length + 2;
    ['排序位置', '化學式或規格'].forEach(function (k) {
      if (k in items.col) items.sheet.getRange(r, items.col[k] + 1).setNumberFormat('@');
    });
    items.sheet.getRange(r, 1, 1, row.length).setValues([row]);

    let qtyMsg = '';
    if (f['數量']) {
      const q = parseQty_(f['數量']);
      const date = f['日期'] || Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
      appendRecords_([[Utilities.parseDate(date, tz_(), 'yyyy-MM-dd'), '新增品項', code, f['品名'], '', q.n, q.note,
        Utilities.formatDate(new Date(), SCHOOL_TZ, 'yyyy/MM/dd HH:mm')]], date);
      updateLatest();
      qtyMsg = '\n目前數量 ' + f['數量'] + ' 已記入盤點紀錄（' + rocText_(date) + '）。';
    }
    if (typeof recordBaseline_ === 'function') recordBaseline_();
    return '已新增「' + f['品名'] + '」，編號 ' + code + '（品項第 ' + r + ' 列）。' + qtyMsg;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- 登記異動

function movementDialog() {
  const items = getTable_('品項');
  const ic = items.col;
  const list = items.rows.filter(function (r) { return String(r[ic['編號']]).trim(); }).map(function (r) {
    return {
      code: String(r[ic['編號']]).trim(), name: String(r[ic['品名']]), spec: String(r[ic['化學式或規格']]),
      room: String(r[ic['教室']]), cab: String(r[ic['櫃別']]), pos: String(r[ic['排序位置']]), unit: String(r[ic['單位']]),
      qty: String(r[ic['最新數量']]), note: String(r[ic['最新數量說明']]), status: String(r[ic['狀態']]),
    };
  });
  const s = getSettings_();
  const data = { items: list, rooms: s.lists['教室'] || [], today: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd') };
  const html = DIALOG_STYLE + `
    <style> #res { max-height: 170px; overflow: auto; border: 1px solid #dadce0; border-radius: 6px; margin-top: 4px; }
      #res div { padding: 7px 10px; cursor: pointer; border-bottom: 1px solid #f1f3f4; } #res div:hover { background: #e8f0fe; }
      #res small, #picked small { color: #5f6368; } #picked { background: #e8f0fe; padding: 8px 10px; border-radius: 6px; margin-top: 6px; }
      .types { display: flex; gap: 6px; } .types label { flex: 1; text-align: center; border: 1px solid #dadce0; border-radius: 6px;
        padding: 8px 0; margin: 0; font-weight: normal; cursor: pointer; } .types input { display: none; }
      .types label.on { background: #1a73e8; color: #fff; border-color: #1a73e8; } .hide { display: none; } </style>
    <label>品項（輸入品名、化學式或編號搜尋）</label><input id="kw" oninput="search()" placeholder="例：酒精、NaCl、藥0004">
    <div id="res"></div><div id="picked" class="hide"></div>
    <label>類型</label>
    <div class="types" id="types"></div>
    <div id="qtyBox"><label>數量</label><input id="qty" inputmode="decimal" placeholder="例：6"></div>
    <div id="moveBox" class="hide">
      <label>新的教室</label><select id="room"></select>
      <div class="row"><div style="flex:1"><label>新的櫃別</label><input id="cab"></div>
        <div style="flex:1"><label>新的排序位置</label><input id="pos"></div></div>
    </div>
    <label class="inline" id="updBox"><input type="checkbox" id="upd" checked> 同時更新目前數量（盤點紀錄加一筆）</label>
    <label class="inline hide" id="retBox"><input type="checkbox" id="ret"> 整個品項淘汰（狀態改為「已淘汰」，不再盤點、列印）</label>
    <div class="row"><div style="flex:1"><label>日期</label><input type="date" id="date"></div>
      <div style="flex:1"><label>經手人</label><input id="who"></div></div>
    <label>說明</label><input id="note" placeholder="例：115.10 請購、實驗打破…">
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()">登記</button></div>
    <script>
      const D = __DATA__; let cur = null, type = '新購';
      document.getElementById('date').value = D.today;
      D.rooms.forEach(function (x) { document.getElementById('room').add(new Option(x, x)); });
      ['新購', '領用', '報廢', '移位'].forEach(function (t) {
        const l = document.createElement('label'); l.textContent = t; l.onclick = function () { setType(t); };
        l.id = 't_' + t; document.getElementById('types').appendChild(l);
      });
      function setType(t) {
        type = t;
        document.querySelectorAll('.types label').forEach(function (l) { l.className = l.id === 't_' + t ? 'on' : ''; });
        document.getElementById('moveBox').className = t === '移位' ? '' : 'hide';
        document.getElementById('qtyBox').className = t === '移位' ? 'hide' : '';
        document.getElementById('updBox').className = t === '移位' ? 'inline hide' : 'inline';
        document.getElementById('retBox').className = t === '報廢' ? 'inline' : 'inline hide';
      }
      setType('新購');
      function search() {
        const k = document.getElementById('kw').value.trim().toLowerCase(), res = document.getElementById('res');
        res.innerHTML = '';
        if (!k) return;
        D.items.filter(function (x) { return (x.name + ' ' + x.spec + ' ' + x.code).toLowerCase().indexOf(k) >= 0; })
          .slice(0, 30).forEach(function (x) {
            const d = document.createElement('div');
            d.appendChild(document.createTextNode(x.name + ' '));
            const sm = document.createElement('small');
            sm.textContent = x.code + '　' + [x.room, x.cab, x.pos].filter(String).join(' ') + '　目前 ' + (x.note || x.qty || '—') +
              (x.status === '已淘汰' ? '（已淘汰）' : '');
            d.appendChild(sm); d.onclick = function () { pick(x); }; res.appendChild(d);
          });
      }
      function pick(x) {
        cur = x;
        document.getElementById('res').innerHTML = '';
        const p = document.getElementById('picked'); p.className = ''; p.innerHTML = '';
        p.appendChild(document.createTextNode('✔ ' + x.name + ' '));
        const sm = document.createElement('small');
        sm.textContent = x.code + '　' + [x.room, x.cab, x.pos].filter(String).join(' ') + '　目前 ' + (x.note || x.qty || '—') + ' ' + x.unit;
        p.appendChild(sm);
        document.getElementById('room').value = x.room; document.getElementById('cab').value = x.cab;
        document.getElementById('pos').value = x.pos;
      }
      function go() {
        if (!cur) { alert('請先搜尋並點選品項。'); return; }
        const m = { code: cur.code, type: type, qty: document.getElementById('qty').value.trim(),
          room: document.getElementById('room').value, cab: document.getElementById('cab').value.trim(),
          pos: document.getElementById('pos').value.trim(), upd: document.getElementById('upd').checked,
          retire: document.getElementById('ret').checked, date: document.getElementById('date').value,
          who: document.getElementById('who').value.trim(), note: document.getElementById('note').value.trim() };
        if (type !== '移位' && (m.qty === '' || isNaN(Number(m.qty)))) { alert('請填數量（數字）。'); return; }
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '登記中…';
        google.script.run.withSuccessHandler(showDone)
          .withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '登記'; }).recordMovement(m);
      }
    </script>`;
  showDialog_(html, data, '📝 登記異動', 640);
}

/** 對話框送出：寫異動紀錄，依類型更新品項或數量。回傳結果訊息。 */
function recordMovement(m) {
  if (!m.code || !m.type) throw new Error('請選擇品項與類型。');
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const items = getTable_('品項');
    const ic = items.col;
    const i = items.rows.findIndex(function (r) { return String(r[ic['編號']]).trim() === m.code; });
    if (i < 0) throw new Error('品項找不到編號 ' + m.code + '。');
    const r = items.rows[i], rowNo = i + 2;
    const name = String(r[ic['品名']]);
    const date = m.date || Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    const dateValue = Utilities.parseDate(date, tz_(), 'yyyy-MM-dd');
    const msgs = [];
    let desc = m.note || '';

    if (m.type === '移位') {
      const from = [r[ic['教室']], r[ic['櫃別']], r[ic['排序位置']]].filter(function (x) { return String(x).trim(); }).join(' ');
      const to = [m.room, m.cab, m.pos].filter(String).join(' ');
      if (from === to) throw new Error('新位置和原本一樣。');
      items.sheet.getRange(rowNo, ic['排序位置'] + 1).setNumberFormat('@');
      items.sheet.getRange(rowNo, ic['教室'] + 1).setValue(m.room);
      items.sheet.getRange(rowNo, ic['櫃別'] + 1).setValue(m.cab);
      items.sheet.getRange(rowNo, ic['排序位置'] + 1).setValue(m.pos);
      desc = '由「' + from + '」移到「' + to + '」' + (desc ? '；' + desc : '');
      msgs.push('位置已改為：' + to + '（藥品記得重印櫃標示）');
    } else {
      const n = Number(m.qty);
      if (m.upd) {
        const latestNote = String(r[ic['最新數量說明']]).trim();
        const latest = r[ic['最新數量']];
        if (latestNote) {
          msgs.push('⚠ 目前數量是文字寫法（' + latestNote + '），沒有自動加減，請下次盤點時更新。');
        } else if (!isNumber_(latest) || String(latest).trim() === '') {
          msgs.push('⚠ 這個品項還沒有盤點數量，沒有自動加減。');
        } else {
          const after = Math.max(0, round_(Number(latest) + (m.type === '新購' ? n : -n)));
          replaceDayRecord_(m.code, name, date, dateValue, after, '異動：' + m.type);
          updateLatest();
          msgs.push('目前數量：' + numText_(latest) + ' → ' + numText_(after));
        }
      }
      if (m.type === '報廢' && m.retire) {
        items.sheet.getRange(rowNo, ic['狀態'] + 1).setValue('已淘汰');
        msgs.push('狀態已改為「已淘汰」');
      }
    }

    const mv = getTable_('異動紀錄');
    need_(mv, ['日期', '編號', '品名', '類型', '數量', '說明', '經手人']);
    const out = new Array(mv.header.length).fill('');
    const put = function (k, v) { out[mv.col[k]] = v; };
    put('日期', dateValue); put('編號', m.code); put('品名', name); put('類型', m.type);
    put('數量', m.type === '移位' ? '' : Number(m.qty)); put('說明', desc); put('經手人', m.who || '');
    const at = mv.sheet.getLastRow() + 1;
    mv.sheet.getRange(at, 1, 1, out.length).setValues([out]);
    mv.sheet.getRange(at, mv.col['日期'] + 1).setNumberFormat('yyyy/mm/dd');
    if (typeof recordBaseline_ === 'function') recordBaseline_();
    return '已登記「' + name + '」' + m.type + (m.type === '移位' ? '' : ' ' + m.qty) + '。\n' + msgs.join('\n');
  } finally {
    lock.releaseLock();
  }
}

/** 異動後的數量：同一天這個品項的舊紀錄（含各存放處）先移除，再寫一筆總數，避免重複加總。 */
function replaceDayRecord_(code, name, dateKey, dateValue, qty, label) {
  const recs = getTable_('盤點紀錄');
  const rc = recs.col, sh = recs.sheet;
  const del = [];
  recs.rows.forEach(function (r, i) {
    if (String(r[rc['編號']]).trim() === code && dateKey_(r[rc['盤點日期']]) === dateKey) del.push(i + 2);
  });
  for (let i = del.length - 1; i >= 0; i--) sh.deleteRow(del[i]);
  const row = new Array(recs.header.length).fill('');
  row[rc['盤點日期']] = dateValue; row[rc['盤點名稱']] = label; row[rc['編號']] = code; row[rc['品名']] = name;
  row[rc['存放處']] = ''; row[rc['數量']] = qty; row[rc['數量說明']] = '';
  row[rc['登錄時間']] = Utilities.formatDate(new Date(), SCHOOL_TZ, 'yyyy/MM/dd HH:mm');
  const at = sh.getLastRow() + 1;
  sh.getRange(at, 1, 1, row.length).setValues([row]);
  sh.getRange(at, rc['盤點日期'] + 1).setNumberFormat('yyyy/mm/dd');
}
