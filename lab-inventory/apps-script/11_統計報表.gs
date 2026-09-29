/**
 * 實驗室藥品器材耗材管理系統 — 統計報表（第三批）
 *
 * 📊 統計報表
 *   ・各學期耗用量：藥品、耗材每學期用掉多少（依盤點紀錄＋異動紀錄的新購推算），由多到少
 *   ・危險物品統計：「品項」危險物品＝是的，依危險分類、教室統計並列出清單
 *   ・永久玻片需求對照：「玻片需求」的需求數 vs 目前可用數，不足標紅
 * 🔧 維護 → 套用危險分類建議：把「危險分類建議」分頁（由 危險分類建議.xlsx 匯入）勾選的填進品項
 * 需要「01_基礎」「04_列印」「10_實驗準備」（查庫存）。
 */

function printUsage() { showPrintDialog({ page: 'usage' }); }
function printHazard() { showPrintDialog({ page: 'hazard' }); }
function printSlides() { showPrintDialog({ page: 'slides' }); }

/** 日期 → 學期「115-1」：8～12 月、1 月算上學期，2～7 月算下學期。 */
function semOf_(k) {
  const y = Number(k.slice(0, 4)) - 1911, m = Number(k.slice(5, 7));
  if (m >= 8) return y + '-1';
  if (m === 1) return (y - 1) + '-1';
  return (y - 1) + '-2';
}

// ---------------------------------------------------------------- 各學期耗用量

/**
 * 每兩次盤點之間的用量＝上次數量＋期間新購－這次數量（小於 0 當 0），算在「這次」所屬的學期。
 * 多處存放的品項用同一天的加總。只看藥品、耗材。
 */
function usageData_() {
  const items = getTable_('品項');
  const ic = items.col;
  const recs = getTable_('盤點紀錄');
  const rc = recs.col;
  const totals = {};   // 編號 → {日期: 數量加總}
  recs.rows.forEach(function (r) {
    const code = String(r[rc['編號']]).trim(), k = dateKey_(r[rc['盤點日期']]), q = r[rc['數量']];
    if (!code || !k || !isNumber_(q) || String(q).trim() === '') return;
    const t = (totals[code] = totals[code] || {});
    t[k] = (t[k] || 0) + Number(q);
  });
  const buys = {};     // 編號 → [{k, q}]
  const mv = SpreadsheetApp.getActive().getSheetByName('異動紀錄') ? getTable_('異動紀錄') : null;
  if (mv) mv.rows.forEach(function (r) {
    const code = String(r[mv.col['編號']]).trim(), k = dateKey_(r[mv.col['日期']]);
    if (code && k && String(r[mv.col['類型']]).trim() === '新購' && isNumber_(r[mv.col['數量']])) {
      (buys[code] = buys[code] || []).push({ k: k, q: Number(r[mv.col['數量']]) });
    }
  });
  const sems = {}, out = [];
  items.rows.forEach(function (r) {
    const code = String(r[ic['編號']]).trim();
    if (!code || String(r[ic['類別']]).trim() === '器材' || !totals[code]) return;
    const keys = Object.keys(totals[code]).sort();
    const per = {};
    for (let i = 1; i < keys.length; i++) {
      const a = keys[i - 1], b = keys[i];
      const bought = (buys[code] || []).filter(function (x) { return x.k > a && x.k <= b; })
        .reduce(function (s, x) { return s + x.q; }, 0);
      const used = totals[code][a] + bought - totals[code][b];
      if (used > 0) { const s = semOf_(b); per[s] = round_((per[s] || 0) + used); sems[s] = true; }
    }
    const sum = Object.keys(per).reduce(function (s, k) { return s + per[k]; }, 0);
    if (sum > 0) out.push({ code: code, cat: String(r[ic['類別']]), name: String(r[ic['品名']]), unit: String(r[ic['單位']]) ||
      (String(r[ic['類別']]) === '藥品' ? '瓶' : ''), per: per, sum: round_(sum), latest: r[ic['最新數量說明']] || numText_(r[ic['最新數量']]) });
  });
  out.sort(function (a, b) { return a.cat === b.cat ? b.sum - a.sum : (a.cat === '藥品' ? -1 : 1); });
  return { items: out, sems: Object.keys(sems).sort(function (a, b) { return naturalCompare_(a, b); }).slice(-4) };
}

function page_usage() {
  const s = getSettings_();
  const d = usageData_();
  const rows = [];
  let lastCat = '';
  d.items.forEach(function (it, i) {
    if (it.cat !== lastCat) { rows.push('<tr class="grp"><td colspan="' + (5 + d.sems.length) + '"><b>' + esc_(it.cat) + '</b></td></tr>'); lastCat = it.cat; }
    rows.push('<tr><td class="c">' + (i + 1) + '</td><td>' + esc_(it.name) + '</td><td class="c">' + esc_(it.unit) + '</td>' +
      d.sems.map(function (sm) { return '<td class="c">' + (it.per[sm] ? esc_(numText_(it.per[sm])) : '') + '</td>'; }).join('') +
      '<td class="c"><b>' + esc_(numText_(it.sum)) + '</b></td><td class="c">' + esc_(it.latest) + '</td></tr>');
  });
  if (!rows.length) rows.push('<tr><td colspan="' + (5 + d.sems.length) + '" class="c">盤點紀錄還不夠算出耗用量</td></tr>');
  const qw = 9;
  const body = reportTable_(s, {
    title: '藥品、耗材各學期耗用量', sub: '用量＝上次盤點＋期間新購－本次盤點（依盤點紀錄推算，僅供參考）', place: '',
    updated: today_(), widths: [6, 30, 8].concat(d.sems.map(function () { return qw; }))
      .concat([9, Math.max(9, 100 - 6 - 30 - 8 - 9 - d.sems.length * qw)]),
    heads: ['序', '品名', '單位'].concat(d.sems.map(function (x) { return x + '<br>學期'; })).concat(['合計', '目前<br>數量']), rows: rows,
  });
  return HtmlService.createHtmlOutput(printShell_('各學期耗用量', body)).setTitle('各學期耗用量（列印）');
}

// ---------------------------------------------------------------- 危險物品統計

function page_hazard() {
  const s = getSettings_();
  const items = getTable_('品項');
  const ic = items.col;
  const list = items.rows.filter(function (r) {
    return String(r[ic['編號']]).trim() && String(r[ic['危險物品']]).trim() === '是' && String(r[ic['狀態']]).trim() !== '已淘汰';
  });
  const byCls = {}, byRoom = {};
  list.forEach(function (r) {
    const cls = splitPlaces_(r[ic['危險分類']]);
    (cls.length ? cls : ['未分類']).forEach(function (c) {
      (byCls[c] = byCls[c] || []).push(r);
      const room = String(r[ic['教室']]).split(/[\/／]/)[0] || '（未填教室）';
      byRoom[room] = byRoom[room] || {};
      byRoom[room][c] = (byRoom[room][c] || 0) + 1;
    });
  });
  const classes = Object.keys(byCls).sort(function (a, b) { return a === '未分類' ? 1 : b === '未分類' ? -1 : byCls[b].length - byCls[a].length; });
  let body;
  if (!list.length) {
    body = '<p style="padding:20px">「品項」還沒有標記危險物品。可以匯入「危險分類建議.xlsx」後，執行「🔧 維護 → 套用危險分類建議」；' +
      '或直接在「品項」的「危險物品」選「是」、「危險分類」填分類（例：易燃液體、腐蝕性、氧化性、毒性、禁水性）。</p>';
  } else {
    const rooms = Object.keys(byRoom).sort(naturalCompare_);
    const sumRows = classes.map(function (c) {
      return '<tr><td>' + esc_(c) + '</td>' + rooms.map(function (r) { return '<td class="c">' + (byRoom[r][c] || '') + '</td>'; }).join('') +
        '<td class="c"><b>' + byCls[c].length + '</b></td></tr>';
    });
    const summary = reportTable_(s, {
      title: '危險物品分類統計', sub: '依「品項」危險物品＝是（使用中）；一個品項有多個分類時各算一次', place: '', updated: today_(),
      widths: [22].concat(rooms.map(function () { return Math.floor(66 / Math.max(1, rooms.length)); })).concat([12]),
      heads: ['危險分類'].concat(rooms.map(esc_)).concat(['合計']), rows: sumRows,
    });
    const detailRows = [];
    let n = 0;
    classes.forEach(function (c) {
      detailRows.push('<tr class="grp"><td colspan="7"><b>' + esc_(c) + '</b>（' + byCls[c].length + ' 項）</td></tr>');
      byCls[c].forEach(function (r) {
        n++;
        detailRows.push('<tr><td class="c">' + n + '</td><td>' + esc_(r[ic['品名']]) + '</td><td>' + esc_(r[ic['化學式或規格']]) +
          '</td><td class="c">' + esc_([r[ic['教室']], r[ic['櫃別']], r[ic['排序位置']]].filter(function (x) { return String(x).trim(); }).join(' ')) +
          '</td><td class="c">' + esc_(r[ic['最新數量說明']] || numText_(r[ic['最新數量']])) + '</td><td class="c">' + esc_(r[ic['SDS']]) +
          '</td><td class="s">' + esc_(r[ic['危險分類']]) + '</td></tr>');
      });
    });
    const detail = reportTable_(s, {
      title: '危險物品清單', sub: '', place: '', updated: today_(), widths: [6, 22, 18, 20, 10, 6, 18],
      heads: ['序', '品名', '化學式', '存放位置', '最新數量', 'SDS', '危險分類'], rows: detailRows,
    });
    body = summary + detail;
  }
  return HtmlService.createHtmlOutput(printShell_('危險物品統計', body)).setTitle('危險物品統計（列印）');
}

/** 把「危險分類建議」分頁中採用＝TRUE 的，填進品項（危險分類已有值的不覆蓋）。 */
function applyHazardSuggestions() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName('危險分類建議')) {
    throw new Error('找不到「危險分類建議」工作表。請先「檔案 → 匯入」危險分類建議.xlsx，匯入位置選「插入新工作表」。');
  }
  const sug = getTable_('危險分類建議');
  need_(sug, ['採用', '編號', '建議危險分類']);
  const items = getTable_('品項');
  need_(items, ['編號', '危險物品', '危險分類']);
  const idx = {};
  items.rows.forEach(function (r, i) { idx[String(r[items.col['編號']]).trim()] = i; });
  const hz = items.rows.map(function (r) { return [r[items.col['危險物品']]]; });
  const cls = items.rows.map(function (r) { return [r[items.col['危險分類']]]; });
  let set = 0, kept = 0;
  sug.rows.forEach(function (r) {
    const on = r[sug.col['採用']], code = String(r[sug.col['編號']]).trim(), v = String(r[sug.col['建議危險分類']]).trim();
    if (!code || !v || !(on === true || /^(TRUE|是|✔|V|Y)$/i.test(String(on).trim())) || !(code in idx)) return;
    const i = idx[code];
    if (String(cls[i][0]).trim()) { kept++; return; }
    cls[i][0] = v; hz[i][0] = '是'; set++;
  });
  const n = items.rows.length;
  if (n) {
    items.sheet.getRange(2, items.col['危險物品'] + 1, n, 1).setValues(hz);
    items.sheet.getRange(2, items.col['危險分類'] + 1, n, 1).setValues(cls);
  }
  const msg = '已標記 ' + set + ' 個品項為危險物品並填入分類。' + (kept ? '\n原本已有危險分類、沒有覆蓋：' + kept + ' 個。' : '') +
    '\n\n可到「📊 統計報表 → 危險物品統計」查看。「危險分類建議」分頁確認完可以刪除。';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return msg;
}

// ---------------------------------------------------------------- 永久玻片需求對照

function page_slides() {
  const s = getSettings_();
  const sh = SpreadsheetApp.getActive().getSheetByName('玻片需求');
  const rows = [];
  let short = 0;
  if (sh && sh.getLastRow() > 1) {
    const t = getTable_('玻片需求');
    const c = t.col;
    const stock = stockIndex_();
    let lastCourse = '', n = 0;
    t.rows.forEach(function (r) {
      const name = String(r[c['品名']]).trim();
      if (!name) return;
      const course = [r[c['年級']], r[c['課本／課程']]].filter(function (x) { return String(x).trim(); }).join(' ');
      if (course && course !== lastCourse) { rows.push('<tr class="grp"><td colspan="7"><b>' + esc_(course) + '</b></td></tr>'); lastCourse = course; }
      const need = r[c['需求數量(112年)']];
      const st = stock(String(r[c['編號']] || '').trim(), name);
      const have = st && isNumber_(st.qty) && String(st.qty).trim() !== '' ? Number(st.qty) : '';
      const hasNeed = isNumber_(need) && String(need).trim() !== '';
      const diff = hasNeed && have !== '' ? have - Number(need) : '';
      if (diff !== '' && diff < 0) short++;
      n++;
      rows.push('<tr' + (diff !== '' && diff < 0 ? ' class="short"' : '') + '><td class="c">' + n + '</td><td>' + esc_(name) +
        '</td><td class="c">' + esc_(hasNeed ? numText_(need) : '') + '</td><td class="c">' + esc_(have === '' ? (st ? '—' : '未建檔') : numText_(have)) +
        '</td><td class="c"><b>' + esc_(diff === '' ? '' : (diff > 0 ? '+' : '') + numText_(diff)) + '</b></td><td class="c">' +
        esc_(numText_(r[c['預計購買數量']])) + '</td><td class="s">' + esc_(r[c['備註']]) + '</td></tr>');
    });
  }
  if (!rows.length) rows.push('<tr><td colspan="7" class="c">「玻片需求」工作表沒有資料</td></tr>');
  const body = reportTable_(s, {
    title: '永久玻片需求與可用數量對照', sub: short ? '不足 ' + short + ' 種（紅底）' : '', place: '', updated: today_(),
    widths: [6, 24, 10, 10, 10, 10, 30], heads: ['序', '玻片', '需求數', '可用數', '差額', '預計<br>購買', '備註'], rows: rows,
  });
  return HtmlService.createHtmlOutput(printShell_('永久玻片需求對照', body).replace('</style>', 'tr.short td { background: #fce8e6; }</style>'))
    .setTitle('永久玻片需求對照（列印）');
}
