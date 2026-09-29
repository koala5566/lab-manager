/**
 * 實驗室藥品器材耗材管理系統 — 列印報表（第 5 步）
 *
 * 選單「列印藥品清單」「列印器材耗材清單」→ 在新分頁開啟列印頁 → 按「🖨 列印」。
 * A4 直式、每頁重複表頭（學校、清單名稱、學年度學期、更新日期、欄位名稱），
 * 每頁頁尾核章欄（依「設定」的核章欄）。「自己筆記」不會印出。
 * 列印頁透過第 4 步部署的網頁應用程式開啟（網址 ?page=print）。
 * 需要「01_基礎」「03_手機盤點」。
 */

// ---------------------------------------------------------------- 選單

function printDrugList() {
  const url = webAppUrl_();
  if (!url) return;
  const html = DIALOG_STYLE + `
    <p>按下按鈕，會在瀏覽器新分頁開啟「藥品清單」列印頁（約 5～10 秒）。</p>
    <p class="hint">開啟後按頁面上方的「🖨 列印」。</p>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" onclick="window.open(D, '_blank'); google.script.host.close();">開啟列印頁</button></div>
    <p class="hint">沒有反應？<a id="a" target="_blank">點這裡開啟</a></p>
    <script>const D = __DATA__; document.getElementById('a').href = D;</script>`;
  showDialog_(html, url + '?page=print&r=drug', '列印藥品清單', 220);
}

function printEquipmentList() {
  const url = webAppUrl_();
  if (!url) return;
  const zones = getSettings_().zones
    .filter(function (z) { return String(z['列印格式']) !== '藥品'; })
    .map(function (z) { return String(z['清單分區']); });
  const html = DIALOG_STYLE + `
    <p>勾選要列印的清單分區（每個分區各自從新的一頁開始）：</p>
    <label class="inline"><input type="checkbox" id="all" checked onchange="toggle(this.checked)"> <b>全選</b></label>
    <div id="list"></div>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" onclick="go()">開啟列印頁</button></div>
    <script>
      const D = __DATA__;
      D.zones.forEach(function (z) {
        const l = document.createElement('label'); l.className = 'inline';
        const c = document.createElement('input'); c.type = 'checkbox'; c.checked = true; c.value = z;
        l.appendChild(c); l.appendChild(document.createTextNode(' ' + z));
        document.getElementById('list').appendChild(l);
      });
      function toggle(on) { document.querySelectorAll('#list input').forEach(function (c) { c.checked = on; }); }
      function go() {
        const z = Array.prototype.filter.call(document.querySelectorAll('#list input'), function (c) { return c.checked; })
          .map(function (c) { return c.value; });
        if (!z.length) { alert('請至少勾選一個分區。'); return; }
        window.open(D.url + '?page=print&r=equip&z=' + encodeURIComponent(z.join('|')), '_blank');
        google.script.host.close();
      }
    </script>`;
  showDialog_(html, { url: url, zones: zones }, '列印器材耗材清單', 120 + zones.length * 30);
}

// ---------------------------------------------------------------- 申報清單（期初／期末）

const KIT_NAME = '實驗套組';

/**
 * 申報時的清單順序，比照簽文：科學館共用 → 藥品 → 生物（器材、耗材、顯微鏡、解剖、套組）→ 化學。
 * 「設定」新增的其他分區排在最後。
 */
function declOrder_(settings) {
  const fixed = ['科學館共用', '藥品清單', '生物器材', '生物耗材', '顯微鏡觀察', '解剖用具', KIT_NAME, '永久玻片',
    '化準器材區', '化準綜合區', '化3後櫃', '化1化2'];
  const names = settings.zones.map(function (z) { return String(z['清單分區']); });
  return fixed.filter(function (x) { return x === KIT_NAME || names.indexOf(x) >= 0; })
    .concat(names.filter(function (x) { return fixed.indexOf(x) < 0; }));
}

/** 「115-1期初」→「115學年度第1學期期初」 */
function termFromName_(name) {
  const m = String(name || '').match(/(\d{2,3})-(\d)\s*(期初|期中|期末)?/);
  return m ? m[1] + '學年度第' + m[2] + '學期' + (m[3] || '') : String(name || '');
}

function printDeclaration() {
  const url = webAppUrl_();
  if (!url) return;
  const settings = getSettings_();
  const recs = getTable_('盤點紀錄');
  need_(recs, ['盤點日期', '盤點名稱']);
  const names = {};
  recs.rows.forEach(function (r) {
    const k = dateKey_(r[recs.col['盤點日期']]);
    if (k && !names[k]) names[k] = String(r[recs.col['盤點名稱']] || '').trim();
  });
  const events = Object.keys(names).sort().reverse().slice(0, 20).map(function (k) {
    return { k: k, label: rocText_(k) + '　' + names[k], term: termFromName_(names[k]) };
  });
  const lists = declOrder_(settings).map(function (x) { return { name: x, on: x !== '永久玻片' }; });
  const html = DIALOG_STYLE + `
    <label>申報哪一次盤點</label><select id="ev" onchange="setTerm()"></select>
    <label>學期標示（會印在每頁表頭，可修改）</label><input id="term">
    <p class="hint">每個品項只印一欄數量：該次盤點的數字（那次沒點到的，印之前最近一次）。</p>
    <label>要印的清單（依簽文順序）</label><div id="list"></div>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" onclick="go()">開啟列印頁</button></div>
    <script>
      const D = __DATA__;
      const ev = document.getElementById('ev');
      D.events.forEach(function (e, i) { ev.add(new Option(e.label, i)); });
      function setTerm() { document.getElementById('term').value = D.events[ev.value].term; }
      setTerm();
      D.lists.forEach(function (z) {
        const l = document.createElement('label'); l.className = 'inline';
        const c = document.createElement('input'); c.type = 'checkbox'; c.checked = z.on; c.value = z.name;
        l.appendChild(c); l.appendChild(document.createTextNode(' ' + z.name));
        document.getElementById('list').appendChild(l);
      });
      function go() {
        const z = Array.prototype.filter.call(document.querySelectorAll('#list input'), function (c) { return c.checked; })
          .map(function (c) { return c.value; });
        if (!z.length) { alert('請至少勾選一份清單。'); return; }
        window.open(D.url + '?page=print&r=decl&d=' + D.events[ev.value].k + '&t=' +
          encodeURIComponent(document.getElementById('term').value) + '&z=' + encodeURIComponent(z.join('|')), '_blank');
        google.script.host.close();
      }
    </script>`;
  if (!events.length) {
    SpreadsheetApp.getUi().alert('盤點紀錄是空的，沒有可以申報的盤點。');
    return;
  }
  showDialog_(html, { url: url, events: events, lists: lists }, '列印申報清單（期初／期末）', 330 + lists.length * 28);
}

/** 每個品項只留「截至 asOf 那天」最近一次的盤點（日期統一標成 asOf）。 */
function asOfView_(items, asOf) {
  return items.map(function (it) {
    const keys = Object.keys(it.counts).filter(function (k) { return k <= asOf; }).sort();
    const o = Object.assign({}, it);
    o.counts = {};
    if (keys.length) o.counts[asOf] = it.counts[keys[keys.length - 1]];
    return o;
  });
}

/** 實驗套組清單（讀「實驗套組」工作表，依套組分段）。 */
function kitTable_(settings, opts, asOf) {
  const sh = SpreadsheetApp.getActive().getSheetByName(KIT_NAME);
  if (!sh || sh.getLastRow() < 2) return '';
  const t = getTable_(KIT_NAME);
  need_(t, ['套組名稱', '課程年級', '組數', '品名', '每組數量', '每組數量說明', '備註', '存放位置']);
  const c = t.col, rows = [];
  let last = null, seq = 0;
  t.rows.forEach(function (r) {
    const kit = String(r[c['套組名稱']]).trim();
    if (!kit || !String(r[c['品名']]).trim()) return;
    if (kit !== last) {
      rows.push('<tr class="grp"><td colspan="5"><b>' + esc_(r[c['課程年級']]) + '：' + esc_(kit) +
        '</b>　每組器材清單（共 ' + esc_(r[c['組數']]) + ' 組）</td></tr>');
      last = kit;
      seq = 0;
    }
    seq++;
    rows.push('<tr><td class="c">' + seq + '</td><td>' + esc_(r[c['品名']]) + '</td><td class="c">' +
      esc_(r[c['每組數量說明']] || numText_(r[c['每組數量']])) + '</td><td class="s">' + esc_(r[c['備註']]) +
      '</td><td class="c">' + esc_(r[c['存放位置']]) + '</td></tr>');
  });
  if (!rows.length) return '';
  return reportTable_(settings, {
    title: '生物實驗套組', sub: '', place: '', updated: asOf, term: opts.term,
    widths: [8, 42, 14, 18, 18], heads: ['序', '品項', '每組數量', '備註', '存放位置'], rows: rows,
  });
}

/** 列印頁工具列上的簽稿文字（不會印出），方便複製到公文系統。 */
function memoBox_(term) {
  const text = '主旨：有關本校' + term + '生物及化學實驗室各類清單，請核示。\n' +
    '說明：\n' +
    '一、依據110年8月13日訂定高級中等學校實驗室管理指引辦理。\n' +
    '二、生物及化學實驗室各類清單如下：\n' +
    '(一)科學館(含探究與實作課程)共用清單。\n' +
    '(二)藥品清單。\n' +
    '(三)生物實驗室清單：分為(玻璃)器材及物品、耗材、顯微鏡觀察、解剖用具、實驗套組。主要存放於生物實驗室(一)(二)及生物準備室。\n' +
    '(四)化學實驗室清單：化學準備室器材區及綜合區(含耗材、文具、工具區)及配藥區清單、化學實驗室(三)物品耗材清單及化學實驗室(一)(二)物品與耗材清單。\n' +
    '擬辦：奉核後，依各清單庫存情形及依照教師課程提出之需求進行請購。';
  return '<details style="flex-basis:100%"><summary style="cursor:pointer"><b>📄 簽稿文字</b>（點開複製，貼到公文系統；可自行修改）</summary>' +
    '<textarea id="memo" style="width:100%;height:190px;font-size:14px;margin-top:6px">' + esc_(text) + '</textarea>' +
    '<button onclick="var m=document.getElementById(\'memo\');m.select();document.execCommand(\'copy\');this.textContent=\'已複製\';">複製簽稿文字</button></details>';
}

// ---------------------------------------------------------------- 列印頁（網址 ?page=print）

function page_print(p) {
  const settings = getSettings_();
  const data = printData_();
  let body = '', title = '', extra = '';
  if (p.r === 'decl') {
    // 申報：每個品項只印「截至所選盤點日期」的那一次數量
    const asOf = dateKey_(p.d);
    const term = String(p.t || '').trim();
    const opts = { n: 1, term: term };
    const view = { items: asOfView_(data.items, asOf) };
    const want = String(p.z || '').split('|').filter(String);
    title = term + '申報清單';
    declOrder_(settings).forEach(function (name) {
      if (want.length && want.indexOf(name) < 0) return;
      if (name === KIT_NAME) { body += kitTable_(settings, opts, asOf); return; }
      const z = settings.zones.filter(function (x) { return String(x['清單分區']) === name; })[0];
      body += String(z['列印格式']) === '藥品' ? drugTable_(settings, view, opts) : equipTable_(settings, view, z, opts);
    });
    extra = memoBox_(term);
  } else if (p.r === 'drug') {
    title = '藥品清單';
    body = drugTable_(settings, data);
  } else {
    title = '器材耗材清單';
    const want = String(p.z || '').split('|').filter(String);
    settings.zones.forEach(function (z) {
      const name = String(z['清單分區']);
      if (String(z['列印格式']) === '藥品') return;
      if (want.length && want.indexOf(name) < 0) return;
      body += equipTable_(settings, data, z);
    });
  }
  if (!body) body = '<p style="padding:20px">沒有可列印的品項。</p>';
  return HtmlService.createHtmlOutput(printShell_(title, body, extra)).setTitle(title + '（列印）');
}

/** 讀品項與盤點紀錄，整理成列印要用的樣子。 */
function printData_() {
  const items = getTable_('品項');
  need_(items, ['編號', '清單分區', '品名', '化學式或規格', '單位', '教室', '櫃別', '排序位置', '分處存放',
    '狀態', '備註']);
  const recs = getTable_('盤點紀錄');
  need_(recs, ['盤點日期', '編號', '存放處', '數量', '數量說明']);
  const rc = recs.col;
  const byCode = {};   // 編號 → {日期: [紀錄]}
  recs.rows.forEach(function (r) {
    const code = String(r[rc['編號']]).trim(), k = dateKey_(r[rc['盤點日期']]);
    if (!code || !k) return;
    const place = String(r[rc['存放處']]).trim();
    ((byCode[code] = byCode[code] || {})[k] = byCode[code][k] || []).push({
      place: place, q: r[rc['數量']], note: String(r[rc['數量說明']]).trim(),
    });
  });
  const ic = items.col;
  const list = items.rows.filter(function (r) {
    return String(r[ic['編號']]).trim() && String(r[ic['狀態']]).trim() !== '已淘汰';
  }).map(function (r) {
    const o = {};
    Object.keys(ic).forEach(function (h) { o[h] = r[ic[h]]; });
    o['編號'] = String(o['編號']).trim();
    o.counts = byCode[o['編號']] || {};
    return o;
  });
  return { items: list };
}

/** 這些品項用到的盤點日期，取最近 n 次，由舊到新。 */
function recentDates_(items, n) {
  const set = {};
  items.forEach(function (it) { Object.keys(it.counts).forEach(function (k) { set[k] = true; }); });
  return Object.keys(set).sort().slice(-Math.max(1, n));
}

/** 某次盤點要印的數量文字：有說明印說明，沒有就印數字；多處存放時加總或逐處列出。 */
function qtyText_(recs) {
  if (!recs || !recs.length) return '';
  if (recs.length === 1) return recs[0].note || numText_(recs[0].q);
  if (recs.some(function (x) { return x.note; })) {
    return recs.map(function (x) { return (x.place ? x.place + ' ' : '') + (x.note || numText_(x.q)); }).join('、');
  }
  const nums = recs.map(function (x) { return x.q; }).filter(isNumber_);
  return nums.length ? numText_(round_(nums.reduce(function (a, b) { return a + Number(b); }, 0))) : '';
}

function numText_(v) {
  if (v === '' || v === null || v === undefined) return '';
  return isNumber_(v) ? String(round_(Number(v))) : String(v);
}

function esc_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** 「115-1」→「115學年度第1學期」 */
function termText_(settings) {
  const t = String(settings.params['學年度學期'] || '');
  const m = t.match(/^(\d+)-(\d)$/);
  return m ? m[1] + '學年度第' + m[2] + '學期' : t;
}

// ---------------------------------------------------------------- 藥品清單

function drugTable_(settings, data, opts) {
  opts = opts || {};
  const items = data.items.filter(function (it) { return String(it['清單分區']) === '藥品清單'; });
  // 櫃別：A～F 等單一字母的櫃子在前，其次 2-B、2-D、小冰箱…；再依排序位置
  const cabKey = function (c) { return /^[A-Z]$/i.test(String(c).trim()) ? 0 : 1; };
  items.sort(function (a, b) {
    return cabKey(a['櫃別']) - cabKey(b['櫃別']) || naturalCompare_(a['櫃別'], b['櫃別']) ||
      naturalCompare_(a['排序位置'], b['排序位置']) || naturalCompare_(a['編號'], b['編號']);
  });
  const n = opts.n || Number(settings.params['藥品清單列印盤點次數']) || 4;
  const dates = recentDates_(items, n);
  const qw = 8.5;
  const widths = [5, 16, 14, 10, 6, 6].concat(dates.map(function () { return qw; }));
  widths.push(Math.max(8, 100 - widths.reduce(function (a, b) { return a + b; }, 0)));
  const heads = ['序號', '中文名稱', '化學式', '教室', '存放<br>位置<br>（櫃）', '存放<br>位置<br>（排序）']
    .concat(dates.map(function (k) { return '盤點數量<br>' + rocText_(k); })).concat(['備註']);
  const rows = items.map(function (it, i) {
    return '<tr><td class="c">' + (i + 1) + '</td><td>' + esc_(it['品名']) + '</td><td>' + esc_(it['化學式或規格']) +
      '</td><td class="c">' + esc_(it['教室']) + '</td><td class="c">' + esc_(it['櫃別']) + '</td><td class="c">' +
      esc_(it['排序位置']) + '</td>' +
      dates.map(function (k) { return '<td class="c">' + esc_(qtyText_(it.counts[k])) + '</td>'; }).join('') +
      '<td class="s">' + esc_(it['備註']) + '</td></tr>';
  });
  return reportTable_(settings, {
    title: '化學與生物實驗室藥品清單', sub: '', place: '', updated: dates[dates.length - 1],
    widths: widths, heads: heads, rows: rows, term: opts.term,
  });
}

// ---------------------------------------------------------------- 器材耗材清單（依清單分區）

function equipTable_(settings, data, zone, opts) {
  opts = opts || {};
  const name = String(zone['清單分區']);
  const items = data.items.filter(function (it) { return String(it['清單分區']) === name; });
  if (!items.length) return '';
  items.sort(function (a, b) {
    return naturalCompare_(a['排序位置'], b['排序位置']) || naturalCompare_(a['編號'], b['編號']);
  });
  const n = opts.n || Number(settings.params['器材耗材清單列印盤點次數']) || 1;
  const dates = recentDates_(items, n);
  const latest = dates[dates.length - 1];
  const rooms = {};
  items.forEach(function (it) { rooms[String(it['教室'])] = true; });
  const multiRoom = Object.keys(rooms).length > 1;

  let heads, widths, rows;
  const placeCols = String(zone['列印格式']) === '分處' ? placeColumns_(settings, zone, items, latest) : [];
  const qtyHeads = dates.map(function (k) { return (placeCols.length ? '總數量' : '數量') + '<br>' + rocText_(k); });
  if (placeCols.length) {
    heads = ['序', '品項名稱', '單位'].concat(placeCols.map(function (c) { return esc_(c.label) + '<br>' + rocText_(latest); }))
      .concat(qtyHeads).concat(['備註']);
    widths = [5, 30, 7].concat(placeCols.map(function () { return 9; })).concat(dates.map(function () { return 9; }));
  } else {
    heads = ['序', '存放位置', '品項', '單位'].concat(qtyHeads).concat(['備註']);
    widths = [5, 17, 32, 7].concat(dates.map(function () { return 9; }));
  }
  widths.push(Math.max(12, 100 - widths.reduce(function (a, b) { return a + b; }, 0)));

  rows = items.map(function (it, i) {
    let tds = '<td class="c">' + (i + 1) + '</td>';
    if (placeCols.length) {
      tds += '<td>' + esc_(it['品名']) + esc_(it['化學式或規格'] ? '（' + it['化學式或規格'] + '）' : '') +
        '</td><td class="c">' + esc_(it['單位']) + '</td>';
      const recs = it.counts[latest] || [];
      tds += placeCols.map(function (c) {
        const hit = recs.filter(function (x) { return c.match(x.place || it['櫃別'] || it['教室']); });
        return '<td class="c">' + esc_(qtyText_(hit)) + '</td>';
      }).join('');
    } else {
      const loc = (multiRoom || !String(it['櫃別']).trim() ? String(it['教室']) + ' ' : '') + String(it['櫃別']);
      tds += '<td class="c">' + esc_(loc.trim()) + '</td><td>' + esc_(it['品名']) +
        esc_(it['化學式或規格'] ? '（' + it['化學式或規格'] + '）' : '') + '</td><td class="c">' + esc_(it['單位']) + '</td>';
    }
    tds += dates.map(function (k) { return '<td class="c">' + esc_(qtyText_(it.counts[k])) + '</td>'; }).join('');
    tds += '<td class="s">' + esc_(it['備註']) + '</td>';
    return '<tr>' + tds + '</tr>';
  });
  return reportTable_(settings, {
    title: String(zone['列印標題'] || name), sub: String(zone['列印副標題'] || ''),
    place: String(zone['存放地點'] || ''), updated: latest, widths: widths, heads: heads, rows: rows,
    term: opts.term,
  });
}

/**
 * 分處欄位：「設定」清單分區表的「分處欄位」（例：生1、生2 或 木櫃、防潮櫃、冰箱）。
 * 欄名可以寫舊簡稱（生1），會對應到教室全名（生物實驗室一）。沒設定時依品項的分處存放自動產生。
 */
function placeColumns_(settings, zone, items, latest) {
  const full = {};   // 舊簡稱 → 教室全名
  const rooms = settings.lists['教室'] || [], shorts = settings.lists['舊簡稱'] || [];
  shorts.forEach(function (s, i) { if (rooms[i]) full[s] = rooms[i]; });
  let labels = splitPlaces_(zone['分處欄位']);
  if (!labels.length) {
    const seen = {};
    items.forEach(function (it) {
      splitPlaces_(it['分處存放']).forEach(function (x) { if (!seen[x]) { seen[x] = true; labels.push(x); } });
    });
  }
  return labels.map(function (label) {
    const names = [label, full[label] || label];
    return { label: label, match: function (place) { return names.indexOf(String(place).trim()) >= 0; } };
  });
}

// ---------------------------------------------------------------- 版面

/** 一份清單＝一個表格：表頭與頁尾核章欄在每一頁重複。 */
function reportTable_(settings, o) {
  const cols = o.heads.length;
  const signs = String(settings.params['核章欄'] || '實驗室管理員,設備組長,主任,校長').split(/[,，、]/)
    .map(function (x) { return x.trim(); }).filter(String);
  const sub = [o.sub, o.place ? '存放地點：' + o.place : ''].filter(String).join('　');
  return '<section class="zone"><table class="r"><colgroup>' +
    o.widths.map(function (w) { return '<col style="width:' + w.toFixed(1) + '%">'; }).join('') + '</colgroup>' +
    '<thead><tr class="h"><th colspan="' + cols + '">' +
    '<div class="school">' + esc_(settings.params['學校名稱']) + '</div>' +
    '<div class="ttl">' + esc_(o.title) + '</div>' + (sub ? '<div class="sub2">' + esc_(sub) + '</div>' : '') +
    '<div class="sub"><span>' + esc_(o.term || termText_(settings)) + '</span><span>更新日期：' +
    esc_(o.updated ? rocText_(o.updated) : '') + '</span></div></th></tr>' +
    '<tr class="cols">' + o.heads.map(function (h) { return '<th>' + h + '</th>'; }).join('') + '</tr></thead>' +
    '<tfoot><tr><td class="f" colspan="' + cols + '"><table class="sign"><tr>' +
    signs.map(function (s) { return '<td>' + esc_(s) + '</td>'; }).join('') + '</tr><tr class="stamp">' +
    signs.map(function () { return '<td></td>'; }).join('') + '</tr></table></td></tr></tfoot>' +
    '<tbody>' + o.rows.join('') + '</tbody></table></section>';
}

function printShell_(title, body, extra) {
  return `<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><title>${esc_(title)}</title>
<style>
  @page { size: A4 portrait; margin: 10mm 9mm 12mm;
    @bottom-center { content: "第 " counter(page) " 頁"; font-size: 9pt; } }
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; box-sizing: border-box; }
  body { margin: 0; color: #000; background: #e9ecef;
    font-family: "Noto Sans TC", "Microsoft JhengHei", "PMingLiU", "PingFang TC", sans-serif; font-size: 10pt; }
  .bar { position: sticky; top: 0; z-index: 9; background: #fff8e1; border-bottom: 1px solid #e0c97a;
    padding: 10px 16px; font-size: 14px; display: flex; gap: 16px; align-items: center; flex-wrap: wrap; }
  .bar button { font-size: 16px; padding: 8px 20px; background: #1a73e8; color: #fff; border: none; border-radius: 6px; cursor: pointer; }
  .paper { width: 192mm; margin: 16px auto; background: #fff; padding: 8mm 0; box-shadow: 0 1px 4px rgba(0,0,0,.2); }
  section.zone + section.zone { break-before: page; page-break-before: always; margin-top: 16mm; }
  table.r { width: 100%; border-collapse: collapse; table-layout: fixed; }
  table.r thead { display: table-header-group; } table.r tfoot { display: table-footer-group; }
  table.r tbody tr { break-inside: avoid; page-break-inside: avoid; }
  table.r th, table.r td { border: 0.6pt solid #000; padding: 2px 3px; vertical-align: middle;
    overflow-wrap: anywhere; line-height: 1.3; }
  tr.h th { border: none; padding: 0 0 4px; font-weight: normal; }
  .school { font-size: 14pt; font-weight: bold; text-align: center; }
  .ttl { font-size: 13pt; font-weight: bold; text-align: center; }
  .sub2 { font-size: 10pt; text-align: center; }
  .sub { display: flex; justify-content: space-between; font-size: 10pt; margin-top: 2px; }
  tr.cols th { background: #e7e6e6; font-size: 9pt; text-align: center; font-weight: bold; }
  td.c { text-align: center; } td.s { font-size: 8.5pt; }
  tr.grp td { background: #f2f2f2; padding: 4px; }
  td.f { border: none !important; padding: 6px 0 0 !important; }
  table.sign { width: 100%; border-collapse: collapse; table-layout: fixed; }
  table.sign td { border: 0.6pt solid #000; text-align: center; font-size: 10pt; padding: 2px; }
  table.sign tr.stamp td { height: 15mm; }
  @media print {
    body { background: #fff; }
    .bar { display: none; }
    .paper { width: auto; margin: 0; padding: 0; box-shadow: none; }
    section.zone + section.zone { margin-top: 0; }
  }
</style></head><body>
<div class="bar"><button onclick="window.print()">🖨 列印</button>
  <span>列印設定：目的地選印表機或「另存為 PDF」、紙張 A4、版面配置「直向」；「更多設定」裡<b>取消勾選「頁首和頁尾」</b>、勾選「背景圖形」。</span>
  ${extra || ''}</div>
<div class="paper">${body}</div>
</body></html>`;
}
