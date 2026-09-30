/**
 * 實驗室藥品器材耗材管理系統 — 實驗準備（第三批）
 *
 * 🧪 實驗準備
 *   ・新增實驗排程：哪天、第幾節、哪一班、哪位老師、哪間教室、做哪個實驗（實驗套組）、幾組
 *   ・列印實驗準備單：選日期範圍 → 每個實驗一張：需要總數＝每組數量×組數，對照目前庫存，不足標紅；附「準備／歸還」打勾欄
 *   ・實驗套組庫存檢核：所有套組依預設組數，檢查庫存夠不夠
 * 實驗內容在「實驗套組」工作表維護：套組名稱、課程年級、組數、編號（可空白，會用品名對）、品名、每組數量…
 * 需要「01_基礎」「04_列印」「09_請購借用」。
 */

const EXP_SHEET = '實驗排程';
const EXP_COLS = ['日期', '節次', '班級', '教師', '教室', '實驗名稱', '組數', '狀態', '備註', '用途類型'];
const EXP_STATUS = ['待準備', '已準備', '已歸還', '取消'];

function setupExperimentSheet_() {
  const t = ensureSheet_(EXP_SHEET, EXP_COLS, [11, 7, 8, 9, 13, 20, 6, 8, 24, 10], '#FBBC04');
  const sh = t.sheet, c = t.col, n = sh.getMaxRows() - 1;
  sh.getRange(2, c['狀態'] + 1, n, 1).setDataValidation(listRule_(EXP_STATUS));
  sh.getRange(2, c['日期'] + 1, n, 1).setNumberFormat('yyyy/mm/dd');
  if (typeof fixPeriodCol_ === 'function') fixPeriodCol_(sh, c['節次'] + 1);   // 「3-4」不要變成日期
  // 實驗名稱可以自由填（社團、考試、多元選修…也放這欄），不設下拉選單；舊版設過的清掉
  sh.getRange(2, c['實驗名稱'] + 1, n, 1).clearDataValidations();
  const st = getSettings_();
  const rooms = st.lists['上課教室'] || st.lists['教室'] || [];
  if (rooms.length) sh.getRange(2, c['教室'] + 1, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(rooms, true).setAllowInvalid(true).build());
  const types = st.lists['用途類型'] || [];
  if (types.length && '用途類型' in c) sh.getRange(2, c['用途類型'] + 1, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(types, true).setAllowInvalid(true).build());
  const all = sh.getRange(2, 1, n, sh.getLastColumn());
  const S = '$' + colLetter_(c['狀態'] + 1), D = '$' + colLetter_(c['日期'] + 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=OR(' + S + '2="已歸還",' + S + '2="取消")')
      .setFontColor('#9AA0A6').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(' + S + '2="待準備",' + D + '2<>"",' + D + '2-TODAY()<=3)')
      .setBackground('#FEF7E0').setFontColor('#B06000').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=' + S + '2="已準備"').setBackground('#E6F4EA').setRanges([all]).build(),
  ]);
  return getTable_(EXP_SHEET);
}

/** 讀「實驗套組」：{套組名稱: {name, grade, groups, place, items: [{code, name, per, perText, note}]}}，依出現順序。 */
function kits_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('實驗套組');
  if (!sh || sh.getLastRow() < 2) return { order: [], map: {} };
  const t = getTable_('實驗套組');
  const c = t.col, map = {}, order = [];
  t.rows.forEach(function (r) {
    const k = String(r[c['套組名稱']]).trim();
    if (!k || !String(r[c['品名']]).trim()) return;
    if (!map[k]) {
      map[k] = { name: k, grade: String(r[c['課程年級']] || ''), groups: Number(r[c['組數']]) || 0, items: [], link: '' };
      order.push(k);
    }
    if (!map[k].link && '講義連結' in c && String(r[c['講義連結']]).trim()) map[k].link = String(r[c['講義連結']]).trim();
    if (!map[k].groups && Number(r[c['組數']])) map[k].groups = Number(r[c['組數']]);
    const per = isNumber_(r[c['每組數量']]) ? Number(r[c['每組數量']]) : parseQty_(r[c['每組數量說明']]).n;
    map[k].items.push({ code: String(r[c['編號']] || '').trim(), name: String(r[c['品名']]).trim(), per: per === '' ? '' : Number(per),
      perText: String(r[c['每組數量說明']] || ''), note: String(r[c['備註']] || ''), place: String(r[c['存放位置']] || '') });
  });
  return { order: order, map: map };
}

/** 品項查表：用編號找，找不到用品名找。 */
function stockIndex_() {
  const items = getTable_('品項');
  const ic = items.col, byCode = {}, byName = {};
  items.rows.forEach(function (r) {
    const o = { code: String(r[ic['編號']]).trim(), name: String(r[ic['品名']]).trim(), unit: String(r[ic['單位']]),
      qty: r[ic['最新數量']], note: String(r[ic['最新數量說明']]), status: String(r[ic['狀態']]),
      loc: [r[ic['教室']], r[ic['櫃別']]].filter(function (x) { return String(x).trim(); }).join(' ') };
    if (typeof measureOf_ === 'function') Object.assign(o, measureOf_(r, ic));   // 計量方式、每包約、安全存量
    if (o.code) byCode[o.code] = o;
    if (o.name && o.status !== '已淘汰' && !byName[o.name]) byName[o.name] = o;
  });
  return function (code, name) { return byCode[code] || byName[String(name).trim()] || null; };
}

// ---------------------------------------------------------------- 新增實驗排程

function experimentDialog() {
  setupExperimentSheet_();
  const k = kits_();
  const data = { kits: k.order.map(function (n) { return { name: n, groups: k.map[n].groups, grade: k.map[n].grade }; }),
    rooms: getSettings_().lists['教室'] || [], today: today_() };
  const html = DIALOG_STYLE + `
    <style> .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0 12px; } .grid label { margin-top: 10px; }
      .full { grid-column: 1 / 3; } .req::after { content: " *"; color: #d93025; } </style>
    <div class="grid">
      <div><label class="req">日期</label><input type="date" id="日期"></div>
      <div><label>節次</label><input id="節次" placeholder="例：3-4"></div>
      <div><label>班級</label><input id="班級" placeholder="例：101"></div>
      <div><label>教師</label><input id="教師"></div>
      <div class="full"><label>教室</label><select id="教室"></select></div>
      <div class="full"><label class="req">實驗（實驗套組）</label><select id="kit" onchange="pickKit()"></select>
        <input id="other" placeholder="其他實驗名稱" style="display:none;margin-top:6px"></div>
      <div><label>組數</label><input id="組數" inputmode="numeric"></div>
      <div><label>備註</label><input id="備註"></div>
    </div>
    <p class="hint">實驗內容（每組需要哪些器材）在「實驗套組」工作表維護；選「其他」的實驗不會自動算器材。</p>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()">新增</button></div>
    <script>
      const D = __DATA__;
      document.getElementById('日期').value = D.today;
      const room = document.getElementById('教室'); room.add(new Option('', ''));
      D.rooms.forEach(function (x) { room.add(new Option(x, x)); });
      const kit = document.getElementById('kit');
      D.kits.forEach(function (x) { kit.add(new Option(x.name + (x.grade ? '（' + x.grade + '）' : ''), x.name)); });
      kit.add(new Option('其他（自行輸入）', '__other'));
      function pickKit() {
        const v = kit.value, k = D.kits.filter(function (x) { return x.name === v; })[0];
        document.getElementById('other').style.display = v === '__other' ? '' : 'none';
        if (k && k.groups) document.getElementById('組數').value = k.groups;
      }
      pickKit();
      function go() {
        const f = {};
        ['日期','節次','班級','教師','教室','組數','備註'].forEach(function (k) { f[k] = document.getElementById(k).value.trim(); });
        f['實驗名稱'] = kit.value === '__other' ? document.getElementById('other').value.trim() : kit.value;
        if (!f['日期'] || !f['實驗名稱']) { alert('請填日期與實驗。'); return; }
        if (f['組數'] && isNaN(Number(f['組數']))) { alert('組數請填數字。'); return; }
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '新增中…';
        google.script.run.withSuccessHandler(showDone)
          .withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '新增'; }).addExperiment(f);
      }
    </script>`;
  showDialog_(html, data, '🧪 新增實驗排程', 520);
}

/** 節次文字（有 12_實驗室使用 時會把被變成日期的轉回來） */
function pv_(v) { return typeof periodVal_ === 'function' ? periodVal_(v) : String(v == null ? '' : v).trim(); }

function addExperiment(f) { return withLock_(function () { return addExperiment__(f); }); }
function addExperiment__(f) {
  const t = setupExperimentSheet_();
  appendRow_(t, { '日期': dateValue_(f['日期']), '節次': f['節次'], '班級': f['班級'], '教師': f['教師'], '教室': f['教室'],
    '實驗名稱': f['實驗名稱'], '組數': f['組數'] ? Number(f['組數']) : '', '狀態': '待準備', '備註': f['備註'], '用途類型': '實驗課' });
  const check = experimentCheck_(f['實驗名稱'], Number(f['組數']) || 0);
  const short = check ? check.rows.filter(function (x) { return x.short; }) : [];
  return '已新增：' + rocText_(f['日期']) + ' ' + (f['班級'] || '') + ' ' + f['實驗名稱'] + (f['組數'] ? '（' + f['組數'] + ' 組）' : '') +
    (check ? (short.length ? '\n\n⚠ 庫存不足：' + short.map(function (x) { return x.name + '（需 ' + x.need + '，有 ' + x.have + '）'; }).join('、') :
      '\n\n✔ 器材數量足夠') : '\n\n（這個實驗在「實驗套組」沒有內容，無法自動檢查器材）');
}

/** 某實驗、某組數需要的器材與庫存比對。找不到套組回傳 null。 */
function experimentCheck_(kitName, groups, stock) {
  const k = kits_().map[String(kitName).trim()];
  if (!k) return null;
  stock = stock || stockIndex_();
  const g = groups || k.groups || 0;
  return {
    kit: k, groups: g,
    rows: k.items.map(function (it) {
      const s = stock(it.code, it.name);
      const need = it.per === '' ? '' : round_(it.per * g);
      const have = s && isNumber_(s.qty) && String(s.qty).trim() !== '' ? Number(s.qty) : '';
      // 每組寫「8片」、庫存單位是「盒」這種單位不同的，不能直接比，改成提醒自行確認
      const perUnit = String(it.perText).replace(/^[\d.\s]+/, '').trim();
      let short, warn = '', state = '';
      if (typeof stockJudge_ === 'function') {
        // 單位相同比數量；有「每包約」換算；大包裝看有沒有、快沒了；單位不同提醒自行確認
        const j = s ? stockJudge_(s, need, perUnit) : { state: '', text: '' };
        state = j.state; short = j.state === 'bad'; if (j.state === 'warn') warn = j.text;
      } else {
        const mismatch = !!(perUnit && s && s.unit && perUnit !== s.unit);
        short = !mismatch && need !== '' && have !== '' && have < need;
        if (mismatch) warn = '單位不同（每組以「' + perUnit + '」計、庫存以「' + s.unit + '」計），請自行確認';
      }
      return { name: it.name, per: it.perText || numText_(it.per), need: need, have: have === '' ? (s ? (s.note || '—') : '未建檔') : have,
        haveText: s ? (s.note || numText_(s.qty)) : '未建檔', unit: s && !s.note ? s.unit : '', loc: it.place || (s ? s.loc : ''),
        note: [it.note, warn].filter(String).join('；'), state: s ? state : 'warn', bulk: !!(s && s.measure === '大包裝'),
        short: short };
    }),
  };
}

// ---------------------------------------------------------------- 列印實驗準備單

function printPrepDialog() {
  setupExperimentSheet_();
  const html = DIALOG_STYLE + `
    <p>列印這段期間「待準備」「已準備」的實驗（每個實驗一張準備單）：</p>
    <div class="row"><div style="flex:1"><label>從</label><input type="date" id="from"></div>
      <div style="flex:1"><label>到</label><input type="date" id="to"></div></div>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" onclick="openPrint({ page: 'prep', from: document.getElementById('from').value, to: document.getElementById('to').value }, this)">開啟列印頁</button></div>
    <script>${OPEN_PRINT_JS}
      const D = __DATA__; document.getElementById('from').value = D.from; document.getElementById('to').value = D.to;</script>`;
  const d = new Date();
  d.setDate(d.getDate() + 7);
  showDialog_(html, { from: today_(), to: Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') }, '🖨 列印實驗準備單', 260);
}

function printKitCheck() {
  showPrintDialog({ page: 'prep', kits: '1' });
}

function page_prep(p) {
  const stock = stockIndex_();
  let blocks = [];
  let title = '實驗準備單';
  if (p.kits) {
    title = '實驗套組庫存檢核';
    const k = kits_();
    blocks = k.order.map(function (n) {
      return { head: esc_(k.map[n].grade ? k.map[n].grade + '：' : '') + '<b>' + esc_(n) + '</b>　預設 ' + k.map[n].groups + ' 組',
        check: experimentCheck_(n, k.map[n].groups, stock), checkbox: false };
    });
  } else {
    const t = setupExperimentSheet_();
    const c = t.col;
    const from = dateKey_(p.from) || today_(), to = dateKey_(p.to) || '9999-12-31';
    t.rows.filter(function (r) {
      const k = dateKey_(r[c['日期']]), st = String(r[c['狀態']]).trim();
      const type = '用途類型' in c ? String(r[c['用途類型']]).trim() : '';
      // 準備單只印實驗課、補做實驗（社團、自主學習等不用準備器材）
      return k && k >= from && k <= to && (st === '待準備' || st === '已準備' || st === '') &&
        (!type || type === '實驗課' || type === '補做實驗');
    }).sort(function (a, b) {
      return dateKey_(a[c['日期']]) < dateKey_(b[c['日期']]) ? -1 : dateKey_(a[c['日期']]) > dateKey_(b[c['日期']]) ? 1 :
        naturalCompare_(pv_(a[c['節次']]), pv_(b[c['節次']]));
    }).forEach(function (r) {
      const g = Number(r[c['組數']]) || 0;
      blocks.push({
        head: '<b>' + esc_(rocText_(r[c['日期']])) + '</b>　' + esc_(pv_(r[c['節次']]) ? '第 ' + pv_(r[c['節次']]) + ' 節　' : '') +
          esc_(r[c['班級']] || '') + '　' + esc_(r[c['教師']] || '') + '　' + esc_(r[c['教室']] || '') +
          '<br><span style="font-size:13pt"><b>' + esc_(r[c['實驗名稱']]) + '</b></span>' + (g ? '　' + g + ' 組' : '') +
          (r[c['備註']] ? '　<small>' + esc_(r[c['備註']]) + '</small>' : ''),
        check: experimentCheck_(r[c['實驗名稱']], g, stock), checkbox: true,
      });
    });
  }
  const body = blocks.map(function (b) {
    let table;
    if (!b.check) {
      table = '<p class="nokit">「實驗套組」沒有這個實驗的內容，請到「實驗套組」工作表建立（套組名稱要和這裡一樣）。</p>';
    } else {
      const rows = b.check.rows.map(function (x, i) {
        return '<tr' + (x.short ? ' class="short"' : '') + '><td class="c">' + (i + 1) + '</td><td>' + esc_(x.name) +
          (x.note ? '<br><small>' + esc_(x.note) + '</small>' : '') + '</td><td class="c">' + esc_(x.per) + '</td><td class="c"><b>' +
          esc_(numText_(x.need)) + '</b></td><td class="c">' + esc_(x.haveText) + (x.unit ? ' ' + esc_(x.unit) : '') +
          (x.short ? '<br><b>不足</b>' : '') + '</td><td class="c">' + esc_(x.loc) + '</td>' +
          (b.checkbox ? '<td class="c box">☐</td><td class="c box">☐</td>' : '') + '</tr>';
      }).join('');
      const short = b.check.rows.filter(function (x) { return x.short; }).length;
      table = '<table class="r"><colgroup><col style="width:6%"><col style="width:' + (b.checkbox ? 32 : 40) +
        '%"><col style="width:10%"><col style="width:10%"><col style="width:14%"><col style="width:' + (b.checkbox ? 16 : 20) + '%">' +
        (b.checkbox ? '<col style="width:6%"><col style="width:6%">' : '') + '</colgroup><thead><tr class="cols"><th>序</th><th>品項</th>' +
        '<th>每組</th><th>需要總數</th><th>目前庫存</th><th>存放位置</th>' + (b.checkbox ? '<th>準備</th><th>歸還</th>' : '') +
        '</tr></thead><tbody>' + rows + '</tbody></table>' +
        (short ? '<p class="warn">⚠ ' + short + ' 項庫存不足，請調度或登記請購。</p>' : '<p class="ok">✔ 庫存足夠</p>');
    }
    return '<section class="exp"><div class="eh">' + b.head + '</div>' + table + '</section>';
  }).join('') || '<p style="padding:20px">這段期間沒有待準備的實驗。可以用「🧪 實驗準備 → 新增實驗排程」登記。</p>';
  const s = getSettings_();
  const html = printShell_(title, '<div class="dochead"><div class="school">' + esc_(s.params['學校名稱']) + '</div><div class="ttl">' +
    esc_(title) + '</div><div class="sub"><span>' + esc_(termText_(s)) + '</span><span>列印日期：' + esc_(rocText_(today_())) +
    '</span></div></div>' + body).replace('</style>', PREP_CSS + '</style>');
  return HtmlService.createHtmlOutput(html).setTitle(title + '（列印）');
}

const PREP_CSS = `
  .dochead { margin-bottom: 4mm; }
  section.exp { break-inside: avoid; page-break-inside: avoid; margin: 0 0 7mm; }
  .eh { background: #f1f3f4; border: 0.6pt solid #000; border-bottom: none; padding: 5px 8px; font-size: 11pt; line-height: 1.5; }
  tr.short td { background: #fce8e6; }
  td.box { font-size: 14pt; }
  p.warn { color: #b3261e; margin: 2px 0 0; font-size: 9.5pt; } p.ok { color: #188038; margin: 2px 0 0; font-size: 9.5pt; }
  p.nokit { border: 0.6pt solid #000; padding: 8px; margin: 0; color: #b06000; }
`;
