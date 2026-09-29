/**
 * 實驗室藥品器材耗材管理系統 — 手機盤點網頁（第 4 步）
 *
 * 部署成「網頁應用程式」後，手機用瀏覽器開網址：
 *   ・盤點：選教室、櫃別 → 逐項輸入數量，輸入完自動存進「盤點表」
 *   ・查詢：依品名、化學式、類別、教室找品項，看存放位置與最新數量
 * 「完成盤點」仍在電腦上按（選單），存進盤點紀錄。
 * 需要「01_基礎」「02_盤點」。
 */

/**
 * 網頁入口（Google 規定的名稱，不能改）。
 * 網址後面加 ?page=xxx 時，交給其他步驟的 page_xxx 函式（例如列印頁 page_print）；否則顯示手機盤點。
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.page && /^\w+$/.test(p.page) && typeof globalThis['page_' + p.page] === 'function') {
    return globalThis['page_' + p.page](p);
  }
  return HtmlService.createHtmlOutput(MOBILE_HTML)
    .setTitle('實驗室盤點')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 手機載入盤點表。 */
function mobileData() {
  const info = JSON.parse(PropertiesService.getDocumentProperties().getProperty('count') || '{}');
  const sh = SpreadsheetApp.getActive().getSheetByName('盤點表');
  const out = { name: info.name || '', date: info.date ? rocText_(info.date) : '', scope: info.scope || '', rows: [] };
  if (!sh || sh.getLastRow() < 2) return out;
  const t = getTable_('盤點表');
  need_(t, COUNT_HEADER);
  const c = t.col;
  t.rows.forEach(function (r) {
    if (!String(r[c['編號']]).trim()) return;
    out.rows.push({
      code: String(r[c['編號']]).trim(), place: String(r[c['存放處']]).trim(),
      room: String(r[c['教室']]), cab: String(r[c['櫃別']]), pos: String(r[c['排序位置']]),
      name: String(r[c['品名']]), spec: String(r[c['化學式或規格']]), unit: String(r[c['單位']]),
      lastDate: String(r[c['上次日期']]), lastQty: String(r[c['上次數量']]), lastNote: String(r[c['上次說明']]),
      qty: String(r[c['本次數量']]), note: String(r[c['本次說明']]),
    });
  });
  return out;
}

/** 手機存一筆：依編號＋存放處找到盤點表那一列，寫入本次數量與說明。 */
function mobileSave(item) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const t = getTable_('盤點表');
    need_(t, COUNT_HEADER);
    const c = t.col;
    const i = t.rows.findIndex(function (r) {
      return String(r[c['編號']]).trim() === item.code && String(r[c['存放處']]).trim() === item.place;
    });
    if (i < 0) throw new Error('盤點表找不到「' + item.code + '」，可能已在電腦上完成盤點或重新產生。請重新整理頁面。');
    const qty = String(item.qty || '').trim();
    const value = qty !== '' && isNumber_(qty) ? Number(qty) : qty;
    t.sheet.getRange(i + 2, c['本次數量'] + 1, 1, 2).setValues([[value, String(item.note || '').trim()]]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** 手機查詢品項（最多 60 筆）。 */
function mobileSearch(q) {
  const items = getTable_('品項');
  const ic = items.col;
  const kw = String(q.keyword || '').trim().toLowerCase();
  const out = [];
  items.rows.forEach(function (r) {
    if (out.length >= 60 || !String(r[ic['編號']]).trim()) return;
    if (q.category && String(r[ic['類別']]) !== q.category) return;
    if (q.room && String(r[ic['教室']]).indexOf(q.room) < 0 && String(r[ic['分處存放']]).indexOf(q.room) < 0) return;
    const text = [r[ic['品名']], r[ic['化學式或規格']], r[ic['編號']], r[ic['櫃別']]].join(' ').toLowerCase();
    if (kw && text.indexOf(kw) < 0) return;
    out.push({
      code: String(r[ic['編號']]), name: String(r[ic['品名']]), spec: String(r[ic['化學式或規格']]),
      room: String(r[ic['教室']]), cab: String(r[ic['櫃別']]), pos: String(r[ic['排序位置']]),
      places: String(r[ic['分處存放']]), unit: String(r[ic['單位']]), status: String(r[ic['狀態']]),
      date: String(r[ic['最新盤點日期']]), qty: String(r[ic['最新數量']]), note: String(r[ic['最新數量說明']]),
      zone: String(r[ic['清單分區']]),
    });
  });
  const s = getSettings_();
  return { rows: out, more: out.length >= 60, categories: s.lists['類別'] || [], rooms: s.lists['教室'] || [] };
}

/** 網頁應用程式網址（結尾 /exec）。找不到時提醒使用者並回傳 ''。 */
function webAppUrl_() {
  const s = getSettings_();
  let url = String(s.params['手機網頁網址'] || '').trim();
  if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) {
    url = ScriptApp.getService().getUrl() || '';
  }
  if (!/\/exec$/.test(url)) {
    SpreadsheetApp.getUi().alert('還沒有部署成網頁應用程式，或網址不正確。\n\n' +
      '請依安裝說明第 4 步「部署」，把部署後的網址（結尾是 /exec）貼到「設定」工作表的「手機網頁網址」。');
    return '';
  }
  return url;
}

/** 選單：顯示手機網址與 QR code。 */
function showMobileLink() {
  const url = webAppUrl_();
  if (!url) return;
  const html = `<div style="font-family:sans-serif;text-align:center">
    <div id="qr" style="display:inline-block;margin:8px"></div>
    <p style="font-size:13px;color:#555">用手機相機掃描 QR code 開啟，<br>再用瀏覽器選單「加到主畫面」。</p>
    <input id="u" readonly style="width:100%;font-size:12px;padding:6px" value="">
    <p><button onclick="var u=document.getElementById('u');u.select();document.execCommand('copy');this.textContent='已複製';">複製網址</button></p>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
    <script>
      var url = __DATA__;
      document.getElementById('u').value = url;
      new QRCode(document.getElementById('qr'), { text: url, width: 200, height: 200 });
    </script></div>`;
  showDialog_(html, url, '手機盤點網址', 400);
}

// ---------------------------------------------------------------- 手機頁面

const MOBILE_HTML = `<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<style>
  * { box-sizing: border-box; }
  body { margin: 0; font-family: -apple-system, "Noto Sans TC", "PingFang TC", sans-serif; font-size: 16px; background: #f3f5f7; color: #1f2328; }
  header { position: sticky; top: 0; z-index: 5; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,.12); padding: 8px 12px 10px; }
  .tabs { display: flex; gap: 6px; margin-bottom: 8px; }
  .tabs button { flex: 1; padding: 9px; font-size: 16px; border: 1px solid #c9d1d9; background: #fff; border-radius: 8px; }
  .tabs button.on { background: #1a73e8; color: #fff; border-color: #1a73e8; }
  .title { font-size: 14px; color: #555; margin-bottom: 6px; }
  .filters { display: flex; gap: 6px; }
  select, input { font-size: 16px; padding: 9px; border: 1px solid #c9d1d9; border-radius: 8px; background: #fff; width: 100%; }
  .bar { height: 8px; background: #e6e9ec; border-radius: 4px; margin-top: 8px; overflow: hidden; }
  .bar div { height: 100%; background: #34a853; width: 0; transition: width .3s; }
  .meta { display: flex; justify-content: space-between; font-size: 13px; color: #555; margin-top: 4px; }
  main { padding: 10px 12px 80px; }
  h3 { margin: 16px 2px 6px; font-size: 15px; color: #1a73e8; }
  .card { background: #fff; border-radius: 10px; padding: 10px 12px; margin-bottom: 8px; border-left: 5px solid #fbbc04; }
  .card.done { border-left-color: #34a853; }
  .name { font-size: 18px; font-weight: 600; }
  .spec { font-size: 14px; color: #555; margin-left: 4px; }
  .loc, .last { font-size: 13px; color: #666; margin-top: 2px; }
  .last b { color: #1f2328; }
  .row { display: flex; gap: 6px; margin-top: 8px; align-items: center; }
  .row input.q { flex: 1; font-size: 20px; text-align: center; }
  .row button { padding: 10px 10px; font-size: 14px; border: 1px solid #c9d1d9; background: #f6f8fa; border-radius: 8px; white-space: nowrap; }
  input.n { margin-top: 6px; font-size: 15px; }
  .st { font-size: 13px; min-height: 18px; margin-top: 4px; }
  .st.ok { color: #188038; } .st.err { color: #d93025; } .st.ing { color: #888; }
  .empty { text-align: center; color: #666; padding: 40px 10px; line-height: 1.8; }
  label.chk { display: flex; align-items: center; gap: 6px; font-size: 14px; margin-top: 8px; }
  label.chk input { width: auto; }
</style></head>
<body>
<header>
  <div class="tabs"><button id="tCount" class="on" onclick="tab('count')">盤點</button><button id="tFind" onclick="tab('find')">查詢</button></div>
  <div id="hCount">
    <div class="title" id="title">載入中…</div>
    <div class="filters"><select id="room" onchange="onRoom()"></select><select id="cab" onchange="render()"></select></div>
    <label class="chk"><input type="checkbox" id="todo" onchange="render()"> 只看還沒盤的</label>
    <div class="bar"><div id="bar"></div></div>
    <div class="meta"><span id="prog"></span><a href="#" onclick="load();return false;">重新整理</a></div>
  </div>
  <div id="hFind" style="display:none">
    <input id="kw" placeholder="品名、化學式或編號" onkeydown="if(event.key==='Enter')find()">
    <div class="filters" style="margin-top:6px"><select id="fCat"></select><select id="fRoom"></select></div>
    <div class="row"><button style="flex:1;background:#1a73e8;color:#fff;border:none" onclick="find()">查詢</button></div>
  </div>
</header>
<main id="mCount"></main>
<main id="mFind" style="display:none"></main>
<script>
var D = null;
function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
function uniq(a) { var s = {}, o = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; o.push(x); } }); return o; }
function fail(e) { document.getElementById('mCount').innerHTML = ''; document.getElementById('mCount').appendChild(el('div', 'empty', '讀取失敗：' + (e && e.message ? e.message : e))); }

function tab(t) {
  var c = t === 'count';
  document.getElementById('tCount').className = c ? 'on' : '';
  document.getElementById('tFind').className = c ? '' : 'on';
  document.getElementById('hCount').style.display = c ? '' : 'none';
  document.getElementById('mCount').style.display = c ? '' : 'none';
  document.getElementById('hFind').style.display = c ? 'none' : '';
  document.getElementById('mFind').style.display = c ? 'none' : '';
  if (!c && !document.getElementById('fCat').options.length) find();
}

function load() {
  document.getElementById('title').textContent = '載入中…';
  google.script.run.withSuccessHandler(function (d) { D = d; init(); }).withFailureHandler(fail).mobileData();
}

function init() {
  var m = document.getElementById('mCount');
  m.innerHTML = '';
  if (!D.rows.length) {
    document.getElementById('title').textContent = '目前沒有盤點表';
    m.appendChild(el('div', 'empty', '請先在電腦上從選單「產生盤點表」，再回來按「重新整理」。'));
    progress();
    return;
  }
  document.getElementById('title').textContent = D.name + '　' + D.date + '　' + D.scope;
  var room = document.getElementById('room'), keep = room.value;
  room.innerHTML = '';
  uniq(D.rows.map(function (r) { return r.room; })).forEach(function (x) { room.add(new Option(x || '（未填教室）', x)); });
  if (keep && Array.prototype.some.call(room.options, function (o) { return o.value === keep; })) room.value = keep;
  onRoom();
}

function onRoom() {
  var room = document.getElementById('room').value, cab = document.getElementById('cab');
  cab.innerHTML = '';
  cab.add(new Option('全部櫃別', ''));
  uniq(D.rows.filter(function (r) { return r.room === room; }).map(function (r) { return r.cab; }))
    .forEach(function (x) { cab.add(new Option(x || '（未填櫃別）', x)); });
  render();
}

function filled(r) { return String(r.qty).trim() !== '' || String(r.note).trim() !== ''; }

function render() {
  var m = document.getElementById('mCount');
  m.innerHTML = '';
  var room = document.getElementById('room').value, cab = document.getElementById('cab').value;
  var todo = document.getElementById('todo').checked, lastCab = null, shown = 0;
  D.rows.forEach(function (r, i) {
    if (r.room !== room || (cab !== '' && r.cab !== cab) || (todo && filled(r))) return;
    if (r.cab !== lastCab) { m.appendChild(el('h3', '', r.cab || '（未填櫃別）')); lastCab = r.cab; }
    m.appendChild(card(r, i));
    shown++;
  });
  if (!shown) m.appendChild(el('div', 'empty', todo ? '這裡都盤完了 👍' : '沒有品項'));
  progress();
}

function card(r, i) {
  var c = el('div', 'card' + (filled(r) ? ' done' : ''));
  var t = el('div'); t.appendChild(el('span', 'name', r.name)); if (r.spec) t.appendChild(el('span', 'spec', r.spec));
  c.appendChild(t);
  var loc = [r.cab, r.pos].filter(String).join(' ');
  if (r.place) loc += '　存放處：' + r.place;
  c.appendChild(el('div', 'loc', loc + (r.unit ? '　單位：' + r.unit : '')));
  var last = el('div', 'last');
  if (r.lastDate) {
    last.appendChild(document.createTextNode('上次 ' + r.lastDate + '：'));
    last.appendChild(el('b', '', r.lastQty));
    if (r.lastNote) last.appendChild(document.createTextNode('（' + r.lastNote + '）'));
  } else last.textContent = '沒有上次紀錄';
  c.appendChild(last);

  var row = el('div', 'row');
  var q = el('input', 'q'); q.value = r.qty; q.placeholder = '本次數量'; q.setAttribute('inputmode', 'decimal');
  var same = el('button', '', '＝上次');
  var more = el('button', '', '說明');
  row.appendChild(q);
  if (r.lastDate && r.lastQty !== '' && r.lastNote.indexOf('合計') < 0) row.appendChild(same);
  row.appendChild(more);
  c.appendChild(row);
  var n = el('input', 'n'); n.placeholder = '數量說明（例：3箱+5串），沒有可空白'; n.value = r.note;
  n.style.display = r.note ? '' : 'none';
  c.appendChild(n);
  var st = el('div', 'st'); c.appendChild(st);

  function save() {
    r.qty = q.value.trim(); r.note = n.value.trim();
    c.className = 'card' + (filled(r) ? ' done' : '');
    st.className = 'st ing'; st.textContent = '儲存中…';
    google.script.run
      .withSuccessHandler(function () { st.className = 'st ok'; st.textContent = '✔ 已存'; progress(); })
      .withFailureHandler(function (e) { st.className = 'st err'; st.textContent = '⚠ 沒存到，點這裡重試：' + (e.message || e); st.onclick = save; })
      .mobileSave({ code: r.code, place: r.place, qty: r.qty, note: r.note });
  }
  q.addEventListener('change', save);
  n.addEventListener('change', save);
  q.addEventListener('keydown', function (e) { if (e.key === 'Enter') { q.blur(); nextInput(c); } });
  same.onclick = function () { q.value = r.lastQty; if (r.lastNote) { n.value = r.lastNote; n.style.display = ''; } save(); nextInput(c); };
  more.onclick = function () { n.style.display = ''; n.focus(); };
  return c;
}

function nextInput(c) {
  var nx = c.nextElementSibling;
  while (nx && !nx.classList.contains('card')) nx = nx.nextElementSibling;
  if (nx) { var q = nx.querySelector('input.q'); if (q) q.focus(); }
}

function progress() {
  var total = D ? D.rows.length : 0, done = D ? D.rows.filter(filled).length : 0;
  document.getElementById('bar').style.width = total ? (done * 100 / total) + '%' : '0';
  document.getElementById('prog').textContent = total ? '已盤 ' + done + '／' + total : '';
}

function find() {
  var m = document.getElementById('mFind');
  m.innerHTML = ''; m.appendChild(el('div', 'empty', '查詢中…'));
  google.script.run.withSuccessHandler(function (res) {
    var fc = document.getElementById('fCat'), fr = document.getElementById('fRoom');
    if (!fc.options.length) {
      fc.add(new Option('全部類別', '')); res.categories.forEach(function (x) { fc.add(new Option(x, x)); });
      fr.add(new Option('全部教室', '')); res.rooms.forEach(function (x) { fr.add(new Option(x, x)); });
    }
    m.innerHTML = '';
    if (!res.rows.length) { m.appendChild(el('div', 'empty', '找不到符合的品項')); return; }
    res.rows.forEach(function (r) {
      var c = el('div', 'card' + (r.status === '已淘汰' ? '' : ' done'));
      var t = el('div'); t.appendChild(el('span', 'name', r.name)); if (r.spec) t.appendChild(el('span', 'spec', r.spec));
      c.appendChild(t);
      c.appendChild(el('div', 'loc', r.code + '　' + r.zone + (r.status === '已淘汰' ? '　（已淘汰）' : '')));
      c.appendChild(el('div', 'loc', '位置：' + [r.room, r.cab, r.pos].filter(String).join(' ') + (r.places ? '（分處：' + r.places + '）' : '')));
      var last = el('div', 'last');
      last.appendChild(document.createTextNode('最新 ' + (r.date || '—') + '：'));
      last.appendChild(el('b', '', (r.note || r.qty || '—') + (r.unit && !r.note ? ' ' + r.unit : '')));
      c.appendChild(last);
      m.appendChild(c);
    });
    if (res.more) m.appendChild(el('div', 'empty', '只顯示前 60 筆，請輸入更精確的關鍵字'));
  }).withFailureHandler(function (e) { m.innerHTML = ''; m.appendChild(el('div', 'empty', '查詢失敗：' + (e.message || e))); })
    .mobileSearch({ keyword: document.getElementById('kw').value,
      category: document.getElementById('fCat').value, room: document.getElementById('fRoom').value });
}

load();
</script>
</body></html>`;
