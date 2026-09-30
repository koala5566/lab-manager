/**
 * 實驗室藥品器材耗材管理系統 — 準備事項（待辦）
 *
 * 老師口頭交代、寫便條紙要準備的東西，自己輸入成待辦：
 *   ・新增：哪天要、（可選）對到當天哪一堂課，要準備什麼、幾個；品名可以從品項搜尋（會顯示庫存），也可以直接打
 *   ・待辦清單：依日期排，打勾＝已準備，用完收回可按「已歸還」；過期沒準備的標紅
 *   ・列印準備事項：選日期範圍，印出打勾用的清單（含存放位置、庫存）
 *   ・手機網頁也有「待辦」分頁，可以在走廊聽到就記、在準備室邊拿邊打勾
 * 資料存在「準備事項」工作表（一列＝一樣東西）。
 * 需要「01_基礎」「04_列印」「09_請購借用」「12_實驗室使用」。
 */

const TODO_SHEET = '準備事項';
const TODO_COLS = ['ID', '需要日期', '節次', '實驗室', '班級', '教師', '事項', '數量', '單位', '編號', '狀態', '備註', '登錄時間', '完成時間'];
const TODO_STATUS = ['待準備', '已準備', '已歸還', '取消'];

function setupTodoSheet_() {
  const t = ensureSheet_(TODO_SHEET, TODO_COLS, [10, 11, 7, 13, 8, 8, 24, 6, 6, 9, 8, 24, 15, 15], '#FBBC04');
  const sh = t.sheet, c = t.col, n = sh.getMaxRows() - 1;
  sh.getRange(2, c['狀態'] + 1, n, 1).setDataValidation(listRule_(TODO_STATUS));
  sh.getRange(2, c['需要日期'] + 1, n, 1).setNumberFormat('yyyy/mm/dd');
  sh.hideColumns(c['ID'] + 1);
  const S = '$' + colLetter_(c['狀態'] + 1), D = '$' + colLetter_(c['需要日期'] + 1);
  const all = sh.getRange(2, 1, n, sh.getLastColumn());
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(' + S + '2="待準備",' + D + '2<>"",' + D + '2<TODAY())')
      .setBackground('#FCE8E6').setFontColor('#B3261E').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(' + S + '2="待準備",' + D + '2<>"",' + D + '2<=TODAY()+1)')
      .setBackground('#FEF7E0').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=OR(' + S + '2="已準備",' + S + '2="已歸還",' + S + '2="取消")')
      .setFontColor('#9AA0A6').setRanges([all]).build(),
  ]);
  return t;
}

function todoTable_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TODO_SHEET);
  if (sh) {
    const t = getTable_(TODO_SHEET);
    if (TODO_COLS.every(function (x) { return x in t.col; })) return t;
  }
  return setupTodoSheet_();
}

// ---------------------------------------------------------------- 資料（電腦對話框、手機共用）

/** 待辦清單：所有待準備＋最近 7 天內完成的；另附品項（搜尋用）與實驗室清單。 */
function todoData() {
  const cfg = labConfig_();
  const t = todoTable_(), c = t.col;
  const today = today_(), since = addDays_(today, -7);
  const stock = {};
  const items = pickerItems_().filter(function (x) { return x.status !== '已淘汰'; });
  items.forEach(function (x) { stock[x.code] = x; });
  const list = [];
  t.rows.forEach(function (r) {
    const id = String(r[c['ID']]).trim();
    const st = String(r[c['狀態']]).trim() || '待準備';
    const d = dateKey_(r[c['需要日期']]);
    if (!id || !String(r[c['事項']]).trim() || st === '取消') return;
    if (st !== '待準備' && (!d || d < since)) return;
    const code = String(r[c['編號']]).trim(), s = stock[code];
    const need = Number(r[c['數量']]);
    list.push({
      id: id, date: d, roc: d ? rocText_(d) : '', wd: d ? weekdayOf_(d) : '', period: String(r[c['節次']]).trim(),
      lab: String(r[c['實驗室']]).trim(), cls: classLabel_(r[c['班級']], cfg.classNames), teacher: String(r[c['教師']]).trim(),
      what: String(r[c['事項']]).trim(), qty: String(r[c['數量']]).trim(), unit: String(r[c['單位']]).trim(), code: code,
      status: st, note: String(r[c['備註']]).trim(),
      place: s ? [s.room, s.cab, s.pos].filter(String).join(' ') : '',
      have: s ? (s.note || s.qty) : '', short: !!(s && need && isNumber_(s.qty) && Number(s.qty) < need),
    });
  });
  list.sort(function (a, b) {
    return (a.date || '9999') < (b.date || '9999') ? -1 : (a.date || '9999') > (b.date || '9999') ? 1 :
      naturalCompare_(a.lab, b.lab) || naturalCompare_(a.period, b.period);
  });
  return {
    list: list, today: today, rooms: cfg.rooms,
    items: items.map(function (x) { return [x.code, x.name, x.spec, x.unit, x.note || x.qty, [x.room, x.cab].filter(String).join(' ')]; }),
  };
}

/** 新增：p = {date, period, lab, cls, teacher, note, items: [{what, qty, unit, code}]} */
function todoAdd(p) {
  const list = (p.items || []).filter(function (x) { return String(x.what || '').trim(); });
  if (!list.length) throw new Error('請至少填一樣要準備的東西。');
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const t = todoTable_();
    const stamp = Utilities.formatDate(new Date(), SCHOOL_TZ, 'yyyy-MM-dd HH:mm');
    const base = Date.now().toString(36);
    list.forEach(function (x, i) {
      appendRow_(t, { 'ID': 'T' + base + i, '需要日期': p.date ? dateValue_(p.date) : '', '節次': p.period || '', '實驗室': p.lab || '',
        '班級': p.cls || '', '教師': p.teacher || '', '事項': String(x.what).trim(), '數量': isNumber_(x.qty) ? Number(x.qty) : (x.qty || ''),
        '單位': x.unit || '', '編號': x.code || '', '狀態': '待準備', '備註': p.note || '', '登錄時間': stamp, '完成時間': '' });
    });
  } finally {
    lock.releaseLock();
  }
  return '已新增 ' + list.length + ' 項準備事項' + (p.date ? '（' + rocText_(p.date) + '（' + weekdayOf_(dateKey_(p.date)) + '）要用）' : '') + '：\n' +
    list.map(function (x) { return '・' + x.what + (x.qty ? ' × ' + x.qty + (x.unit || '') : ''); }).join('\n');
}

/** 改狀態：ids 陣列，status＝已準備／已歸還／待準備／取消 */
function todoSet(ids, status) {
  if (TODO_STATUS.indexOf(status) < 0) throw new Error('狀態不對：' + status);
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const t = todoTable_(), c = t.col;
    const stamp = status === '待準備' ? '' : Utilities.formatDate(new Date(), SCHOOL_TZ, 'yyyy-MM-dd HH:mm');
    let n = 0;
    t.rows.forEach(function (r, i) {
      if (ids.indexOf(String(r[c['ID']]).trim()) < 0) return;
      t.sheet.getRange(i + 2, c['狀態'] + 1).setValue(status);
      t.sheet.getRange(i + 2, c['完成時間'] + 1).setValue(stamp);
      n++;
    });
    if (!n) throw new Error('找不到這筆準備事項，可能已被刪除。請重新整理。');
    return n;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- 電腦：對話框

function todoDialog() { todoShow_(false); }
function todoAddDialog() { todoShow_(true); }

function todoShow_(add) {
  const html = DIALOG_STYLE + '<div id="todo"></div><script>' + todoJs_() + 'todoInit("todo", { print: true, add: ' + (add ? 'true' : 'false') + ' });</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(820).setHeight(760),
    add ? '📝 新增準備事項' : '📝 準備事項（待辦）');
}

/** 待辦畫面程式（電腦對話框、手機網頁共用）。呼叫 todoInit(容器id, {print, add}) 開始。 */
function todoJs_() {
  return `
  (function () {
    const css = document.createElement('style');
    css.textContent = '.td-bar{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:4px 0 10px}.td-bar button,.td-f button{padding:7px 12px;font-size:14px;border:1px solid #dadce0;border-radius:6px;background:#fff;cursor:pointer}' +
      '.td-bar button.on{background:#1a73e8;color:#fff;border-color:#1a73e8}.td-bar .add{background:#188038;color:#fff;border-color:#188038}.td-sp{margin-left:auto}' +
      '.td-day{font-weight:bold;margin:12px 0 4px;color:#1a73e8}.td-day.late{color:#d93025}.td-day.soon{color:#e37400}' +
      '.td-it{display:flex;gap:10px;align-items:flex-start;background:#fff;border:1px solid #e0e3e7;border-left:5px solid #fbbc04;border-radius:8px;padding:8px 10px;margin:5px 0}' +
      '.td-it.done{border-left-color:#34a853;opacity:.65}.td-it.back{border-left-color:#9aa0a6;opacity:.55}.td-it input[type=checkbox]{width:22px;height:22px;margin-top:2px;flex:none}' +
      '.td-w{font-size:16px;font-weight:600}.td-s{font-size:12.5px;color:#5f6368;margin-top:2px}.td-short{color:#d93025;font-weight:600}.td-it .td-a{margin-left:auto;display:flex;gap:4px;flex:none}' +
      '.td-it .td-a button{font-size:12px;padding:4px 8px;border:1px solid #dadce0;border-radius:5px;background:#fff;cursor:pointer}.td-empty{color:#666;text-align:center;padding:30px}' +
      '.td-f{background:#f8f9fa;border:1px solid #dadce0;border-radius:10px;padding:10px 12px;margin-bottom:12px}.td-f label{display:block;font-size:13px;color:#5f6368;margin:8px 0 2px}' +
      '.td-f input,.td-f select{width:100%;font-size:15px;padding:7px;border:1px solid #c9d1d9;border-radius:6px;box-sizing:border-box;background:#fff}' +
      '.td-g{display:grid;grid-template-columns:1fr 1fr;gap:0 10px}.td-row{display:grid;grid-template-columns:1fr 70px 60px 30px;gap:6px;margin-top:6px;position:relative}' +
      '.td-row .x{padding:0;border:none;background:none;font-size:18px;color:#999;cursor:pointer}.td-sug{position:absolute;top:38px;left:0;right:0;z-index:5;background:#fff;border:1px solid #dadce0;border-radius:6px;max-height:180px;overflow:auto;box-shadow:0 2px 6px rgba(0,0,0,.15)}' +
      '.td-sug div{padding:7px 9px;border-bottom:1px solid #f1f3f4;cursor:pointer;font-size:14px}.td-sug div:hover{background:#e8f0fe}.td-sug small{color:#5f6368}' +
      '.td-info{font-size:12px;color:#5f6368;grid-column:1/5}.td-msg{margin-top:8px;font-size:14px;white-space:pre-line}.td-msg.ok{color:#188038}.td-msg.err{color:#d93025}' +
      '@media print{.td-bar,.td-f,.td-a{display:none}}';
    document.head.appendChild(css);
    function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
    function key(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
    window.todoInit = function (id, opt) {
      const box = document.getElementById(id);
      let T = null, view = 'todo', showAdd = !!(opt && opt.add), rows = [{}], lessons = [];
      function load() {
        box.innerHTML = '<p style="color:#666">讀取中…</p>';
        google.script.run.withSuccessHandler(function (d) { T = d; draw(); if (showAdd) fetchLessons(); }).withFailureHandler(function (e) { box.textContent = '讀取失敗：' + (e.message || e); }).todoData();
      }
      function dayTitle(x) {
        if (!x.date) return { t: '沒有指定日期', c: '' };
        const diff = Math.round((new Date(x.date) - new Date(T.today)) / 86400000);
        const tag = diff === 0 ? '　今天' : diff === 1 ? '　明天' : diff === 2 ? '　後天' : diff < 0 ? '　已過 ' + (-diff) + ' 天' : '';
        return { t: x.roc + '（' + x.wd + '）' + tag, c: diff < 0 && x.status === '待準備' ? 'late' : diff <= 1 ? 'soon' : '' };
      }
      function draw() {
        const todo = T.list.filter(function (x) { return x.status === '待準備'; });
        const done = T.list.filter(function (x) { return x.status !== '待準備'; });
        let h = '<div class="td-bar"><button class="add" data-a="add">＋ 新增</button>' +
          '<button data-a="todo" class="' + (view === 'todo' ? 'on' : '') + '">待準備 ' + todo.length + '</button>' +
          '<button data-a="done" class="' + (view === 'done' ? 'on' : '') + '">最近完成 ' + done.length + '</button>' +
          '<span class="td-sp"></span><button data-a="reload">重新整理</button>' + (opt && opt.print ? '<button data-a="print">🖨 列印</button>' : '') + '</div>';
        if (showAdd) h += form();
        const list = view === 'todo' ? todo : done.slice().reverse();
        if (!list.length) h += '<div class="td-empty">' + (view === 'todo' ? '沒有待準備的事項 👍' : '最近 7 天沒有完成的事項') + '</div>';
        let last = null;
        list.forEach(function (x) {
          const dt = dayTitle(x);
          if (dt.t !== last) { h += '<div class="td-day ' + dt.c + '">' + esc(dt.t) + '</div>'; last = dt.t; }
          const where = [x.lab, x.period ? (x.period === '午' ? '中午' : '第' + x.period + '節') : '', x.cls, x.teacher ? x.teacher + '老師' : ''].filter(String).join(' ');
          h += '<div class="td-it ' + (x.status === '已準備' ? 'done' : x.status === '已歸還' ? 'back' : '') + '">' +
            '<input type="checkbox" data-id="' + esc(x.id) + '"' + (x.status !== '待準備' ? ' checked' : '') + '>' +
            '<div><div class="td-w">' + esc(x.what) + (x.qty ? ' × ' + esc(x.qty) + ' ' + esc(x.unit) : '') + '</div>' +
            (where ? '<div class="td-s">' + esc(where) + '</div>' : '') +
            (x.place || x.have ? '<div class="td-s">📍 ' + esc(x.place || '—') + (x.have !== '' ? '　庫存 ' + esc(x.have) + ' ' + esc(x.unit) : '') +
              (x.short ? '　<span class="td-short">⚠ 可能不夠</span>' : '') + '</div>' : '') +
            (x.note ? '<div class="td-s">📝 ' + esc(x.note) + '</div>' : '') + '</div>' +
            '<div class="td-a">' + (x.status === '已準備' ? '<button data-r="已歸還" data-id="' + esc(x.id) + '">已歸還</button>' : '') +
            (x.status === '待準備' ? '<button data-r="取消" data-id="' + esc(x.id) + '">取消</button>' : '') +
            (x.status !== '待準備' ? '<span style="font-size:12px;color:#5f6368">' + esc(x.status) + '</span>' : '') + '</div></div>';
        });
        box.innerHTML = h;
        box.querySelectorAll('.td-bar button').forEach(function (b) {
          b.onclick = function () {
            const a = b.getAttribute('data-a');
            if (a === 'add') { showAdd = !showAdd; if (showAdd) { rows = [{}]; lessons = []; } draw(); if (showAdd) fetchLessons(); }
            else if (a === 'reload') load();
            else if (a === 'print') window.print();
            else { view = a; draw(); }
          };
        });
        box.querySelectorAll('.td-it input[type=checkbox]').forEach(function (c) {
          c.onchange = function () { set([c.getAttribute('data-id')], c.checked ? '已準備' : '待準備'); };
        });
        box.querySelectorAll('.td-a button').forEach(function (b) {
          b.onclick = function () {
            const s = b.getAttribute('data-r');
            // 取消要按兩次（避免誤按；不用 confirm，瀏覽器會跳出一長串網址）
            if (s === '取消' && b.textContent !== '確定取消？') { b.textContent = '確定取消？'; b.style.color = '#d93025'; return; }
            set([b.getAttribute('data-id')], s);
          };
        });
        if (showAdd) bindForm();
      }
      function set(ids, status) {
        T.list.forEach(function (x) { if (ids.indexOf(x.id) >= 0) x.status = status; });
        if (status === '取消') T.list = T.list.filter(function (x) { return ids.indexOf(x.id) < 0; });
        draw();
        google.script.run.withFailureHandler(function (e) { alert('沒有存到：' + (e.message || e)); load(); }).todoSet(ids, status);
      }
      // ---- 新增表單
      function form() {
        const f = window.__tdf || {};
        let h = '<div class="td-f"><b>新增準備事項</b>' +
          '<div class="td-g"><div><label>哪天要用</label><input type="date" id="tdDate" value="' + esc(f.date || nextSchoolDay()) + '"></div>' +
          '<div><label>對到當天的課（可不選）</label><select id="tdLesson"><option value="">— 不對應，自己填 —</option>' +
          lessons.map(function (l, i) { return '<option value="' + i + '">' + esc(l.label) + '</option>'; }).join('') + '</select></div>' +
          '<div><label>實驗室</label><select id="tdLab"><option value=""></option>' + T.rooms.map(function (r) { return '<option' + (r === f.lab ? ' selected' : '') + '>' + esc(r) + '</option>'; }).join('') + '<option value="準備室">準備室</option></select></div>' +
          '<div><label>節次（例：3、3-4、午）</label><input id="tdPeriod" value="' + esc(f.period || '') + '"></div>' +
          '<div><label>班級</label><input id="tdCls" value="' + esc(f.cls || '') + '"></div>' +
          '<div><label>老師</label><input id="tdTeacher" value="' + esc(f.teacher || '') + '"></div></div>' +
          '<label>要準備的東西（品名可以搜尋品項，找不到直接打）</label><div id="tdRows">';
        rows.forEach(function (r, i) {
          h += '<div class="td-row"><input data-i="' + i + '" class="tdWhat" placeholder="例：燒杯 250mL、剪刀、冰塊" value="' + esc(r.what || '') + '" autocomplete="off">' +
            '<input data-i="' + i + '" class="tdQty" placeholder="數量" inputmode="decimal" value="' + esc(r.qty || '') + '">' +
            '<input data-i="' + i + '" class="tdUnit" placeholder="單位" value="' + esc(r.unit || '') + '">' +
            '<button class="x" data-i="' + i + '" title="刪掉這一行">✕</button>' +
            (r.info ? '<div class="td-info">' + esc(r.info) + '</div>' : '') + '</div>';
        });
        h += '</div><button id="tdMore" style="margin-top:6px">＋ 再一樣</button>' +
          '<label>備註（例：便條內容、老師要自己來拿）</label><input id="tdNote" value="' + esc(f.note || '') + '">' +
          '<div style="display:flex;gap:8px;margin-top:10px"><button id="tdCancel">收起</button><button id="tdSave" style="flex:1;background:#1a73e8;color:#fff;border-color:#1a73e8">儲存</button></div>' +
          '<div id="tdMsg" class="td-msg"></div></div>';
        return h;
      }
      function nextSchoolDay() { const d = new Date(); d.setDate(d.getDate() + 1); while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1); return key(d); }
      function keep() {
        const g = function (i) { const e = document.getElementById(i); return e ? e.value : ''; };
        window.__tdf = { date: g('tdDate'), lab: g('tdLab'), period: g('tdPeriod'), cls: g('tdCls'), teacher: g('tdTeacher'), note: g('tdNote') };
        box.querySelectorAll('.tdWhat').forEach(function (e) { rows[+e.getAttribute('data-i')].what = e.value; });
        box.querySelectorAll('.tdQty').forEach(function (e) { rows[+e.getAttribute('data-i')].qty = e.value; });
        box.querySelectorAll('.tdUnit').forEach(function (e) { rows[+e.getAttribute('data-i')].unit = e.value; });
      }
      function fetchLessons() {
        const d = document.getElementById('tdDate'); if (!d || !d.value) return;
        google.script.run.withSuccessHandler(function (u) {
          lessons = u.bookings.filter(function (b) { return b.type !== '放假' && b.type !== '考試'; }).map(function (b) {
            return { lab: b.lab, period: b.periodText, cls: b.cls, teacher: b.teacher,
              label: b.lab + ' ' + (b.periodText === '午' ? '中午' : '第' + b.periodText + '節') + ' ' + [b.cls, b.teacher, b.content || b.type].filter(String).join(' ') };
          });
          lessons.sort(function (a, b) { return a.label < b.label ? -1 : 1; });
          const s = document.getElementById('tdLesson'); if (!s) return;
          s.innerHTML = '<option value="">— ' + (lessons.length ? '不對應，自己填' : '這天沒有登記的課') + ' —</option>' +
            lessons.map(function (l, i) { return '<option value="' + i + '">' + esc(l.label) + '</option>'; }).join('');
        }).usageData(d.value, 1);
      }
      function bindForm() {
        document.getElementById('tdDate').onchange = function () { keep(); fetchLessons(); };
        document.getElementById('tdLesson').onchange = function () {
          const l = lessons[+this.value]; if (!l) return;
          document.getElementById('tdLab').value = l.lab; document.getElementById('tdPeriod').value = l.period;
          document.getElementById('tdCls').value = l.cls; document.getElementById('tdTeacher').value = l.teacher;
        };
        document.getElementById('tdMore').onclick = function () { keep(); rows.push({}); draw(); const w = box.querySelectorAll('.tdWhat'); w[w.length - 1].focus(); };
        document.getElementById('tdCancel').onclick = function () { showAdd = false; window.__tdf = null; draw(); };
        document.getElementById('tdSave').onclick = save;
        box.querySelectorAll('.td-row .x').forEach(function (b) {
          b.onclick = function () { keep(); rows.splice(+b.getAttribute('data-i'), 1); if (!rows.length) rows.push({}); draw(); };
        });
        box.querySelectorAll('.tdWhat').forEach(function (inp) {
          inp.oninput = function () { const i = +inp.getAttribute('data-i'); rows[i].code = ''; rows[i].info = ''; suggest(inp); };
          inp.onblur = function () { setTimeout(function () { const s = inp.parentNode.querySelector('.td-sug'); if (s) s.remove(); }, 200); };
        });
      }
      function suggest(inp) {
        const old = inp.parentNode.querySelector('.td-sug'); if (old) old.remove();
        const k = inp.value.trim().toLowerCase(); if (!k) return;
        const hit = T.items.filter(function (x) { return (x[1] + ' ' + x[2] + ' ' + x[0]).toLowerCase().indexOf(k) >= 0; }).slice(0, 12);
        if (!hit.length) return;
        const s = document.createElement('div'); s.className = 'td-sug';
        hit.forEach(function (x) {
          const d = document.createElement('div');
          d.innerHTML = esc(x[1]) + ' <small>' + esc(x[2]) + '　' + esc(x[5]) + '　庫存 ' + esc(x[4] || '—') + ' ' + esc(x[3]) + '</small>';
          d.onmousedown = function (e) {
            e.preventDefault(); keep();
            const i = +inp.getAttribute('data-i');
            rows[i].what = x[1] + (x[2] ? ' ' + x[2] : ''); rows[i].code = x[0]; rows[i].unit = rows[i].unit || x[3];
            rows[i].info = '📍 ' + (x[5] || '—') + '　庫存 ' + (x[4] || '—') + ' ' + x[3];
            draw();
          };
          s.appendChild(d);
        });
        inp.parentNode.appendChild(s);
      }
      function save() {
        keep();
        const f = window.__tdf, items = rows.filter(function (r) { return (r.what || '').trim(); });
        const msg = document.getElementById('tdMsg');
        if (!items.length) { msg.className = 'td-msg err'; msg.textContent = '請至少填一樣要準備的東西。'; return; }
        const b = document.getElementById('tdSave'); b.disabled = true; b.textContent = '儲存中…';
        google.script.run.withSuccessHandler(function (m) {
          rows = [{}]; window.__tdf = { date: f.date }; showAdd = true;
          google.script.run.withSuccessHandler(function (d) {
            T = d; view = 'todo'; draw(); fetchLessons();
            const mm = document.getElementById('tdMsg'); mm.className = 'td-msg ok'; mm.textContent = '✔ ' + m;
          }).todoData();
        }).withFailureHandler(function (e) { b.disabled = false; b.textContent = '儲存'; msg.className = 'td-msg err'; msg.textContent = e.message || e; })
          .todoAdd({ date: f.date, lab: f.lab, period: f.period, cls: f.cls, teacher: f.teacher, note: f.note,
            items: items.map(function (r) { return { what: r.what, qty: r.qty, unit: r.unit, code: r.code || '' }; }) });
      }
      load();
    };
  })();`;
}

// ---------------------------------------------------------------- 列印

function printTodoDialog() {
  const html = DIALOG_STYLE + `
    <p class="hint">印出這段日期的準備事項（依日期、實驗室排），邊準備邊打勾。</p>
    <label>從</label><input type="date" id="from">
    <label>到</label><input type="date" id="to">
    <label class="inline"><input type="checkbox" id="done"> 已準備的也印</label>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" onclick="openPrint({ page: 'todo', from: document.getElementById('from').value, to: document.getElementById('to').value, done: document.getElementById('done').checked ? '1' : '' }, this)">開啟列印頁</button></div>
    <script>${OPEN_PRINT_JS}
      const D = __DATA__; document.getElementById('from').value = D.from; document.getElementById('to').value = D.to;</script>`;
  showDialog_(html, { from: today_(), to: addDays_(today_(), 6) }, '🖨 列印準備事項', 300);
}

function page_todo(p) {
  const cfg = labConfig_();
  const from = dateKey_(p.from) || today_(), to = dateKey_(p.to) || addDays_(from, 6);
  const t = todoTable_(), c = t.col;
  const stock = {};
  pickerItems_().forEach(function (x) { stock[x.code] = x; });
  const list = t.rows.filter(function (r) {
    const st = String(r[c['狀態']]).trim() || '待準備', d = dateKey_(r[c['需要日期']]);
    if (!String(r[c['事項']]).trim() || st === '取消' || st === '已歸還') return false;
    if (st === '已準備' && !p.done) return false;
    return d ? d >= from && d <= to : st === '待準備';
  }).sort(function (a, b) {
    const x = dateKey_(a[c['需要日期']]) || '9999', y = dateKey_(b[c['需要日期']]) || '9999';
    return x < y ? -1 : x > y ? 1 : naturalCompare_(a[c['實驗室']], b[c['實驗室']]) || naturalCompare_(a[c['節次']], b[c['節次']]);
  });
  let body = '', last = null;
  list.forEach(function (r) {
    const d = dateKey_(r[c['需要日期']]);
    const head = d ? rocText_(d) + '（' + weekdayOf_(d) + '）' : '沒有指定日期';
    if (head !== last) {
      if (last !== null) body += '</tbody></table>';
      body += '<h3>' + esc_(head) + '</h3><table class="td"><colgroup><col style="width:9%"><col style="width:9%"><col style="width:30%"><col style="width:10%">' +
        '<col style="width:24%"><col style="width:18%"></colgroup><thead><tr><th>準備</th><th>歸還</th><th>要準備的東西</th><th>數量</th><th>實驗室・節次・班級・老師</th><th>存放位置／庫存</th></tr></thead><tbody>';
      last = head;
    }
    const s = stock[String(r[c['編號']]).trim()];
    const per = String(r[c['節次']]).trim();
    const done = String(r[c['狀態']]).trim() === '已準備';
    body += '<tr><td class="box">' + (done ? '☑' : '☐') + '</td><td class="box">☐</td><td>' + esc_(r[c['事項']]) +
      (String(r[c['備註']]).trim() ? '<div class="nt">' + esc_(r[c['備註']]) + '</div>' : '') + '</td><td>' +
      esc_(String(r[c['數量']]) + ' ' + String(r[c['單位']])) + '</td><td>' +
      esc_([r[c['實驗室']], per ? (per === '午' ? '中午' : '第' + per + '節') : '', classLabel_(r[c['班級']], cfg.classNames), r[c['教師']]].filter(String).join(' ')) + '</td><td>' +
      (s ? esc_([s.room, s.cab, s.pos].filter(String).join(' ')) + '<div class="nt">庫存 ' + esc_((s.note || s.qty) + ' ' + s.unit) + '</div>' : '') + '</td></tr>';
  });
  if (last !== null) body += '</tbody></table>';
  if (!body) body = '<p style="padding:20px">' + rocText_(from) + '～' + rocText_(to) + ' 沒有待準備的事項。</p>';
  body = '<div class="hd"><div class="t">準備事項清單</div><div>' + esc_(rocText_(from) + '～' + rocText_(to)) + '　列印：' + esc_(rocText_(today_())) + '</div></div>' + body;
  return HtmlService.createHtmlOutput(labShell_('準備事項清單', '<section>' + body + '</section>', false, TODO_CSS)).setTitle('準備事項清單（列印）');
}

const TODO_CSS = `
  .hd { display: flex; justify-content: space-between; align-items: baseline; border-bottom: 2pt solid #000; padding-bottom: 2mm; margin-bottom: 3mm; }
  .hd .t { font-size: 20pt; font-weight: 900; }
  h3 { font-size: 13pt; margin: 5mm 0 1.5mm; }
  table.td { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 11pt; }
  table.td th, table.td td { border: 0.8pt solid #000; padding: 1.5mm 2mm; vertical-align: middle; }
  table.td th { background: #e8eaed; font-size: 10pt; } table.td td.box { text-align: center; font-size: 16pt; }
  table.td tr { break-inside: avoid; } .nt { font-size: 9pt; color: #5f6368; }
`;
