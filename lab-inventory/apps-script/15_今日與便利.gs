/**
 * 實驗室藥品器材耗材管理系統 — 今日總覽與便利功能
 *
 *   ・補實驗名稱：選期間、實驗室、老師 → 勾選那幾堂實驗課 → 一次填好實驗名稱（和組數）
 *   ・點一堂課（六間使用一覽、手機）：看詳細、補實驗名稱（可同週同老師一起補）、新增這堂課的準備事項、開器材準備單
 *   ・手機「今天」分頁：今天和下一個上課日的課、待準備事項（可打勾）、需補充幾項
 * 需要「12_實驗室使用」「14_準備事項」。
 */

// ---------------------------------------------------------------- 共用：找「實驗排程」的某一列

/** 一列的辨識碼：日期｜實驗室｜節次｜班級（班級用班名） */
function lessonKey_(r, c, cfg) {
  return [dateKey_(r[c['日期']]), String(r[c['教室']]).trim(), String(r[c['節次']]).trim(), classLabel_(r[c['班級']], cfg.classNames)].join('|');
}

/** 依列號找，列號對不上（中間有刪列）就用辨識碼找。回傳列號（從 2 起）或 -1 */
function findLessonRow_(t, cfg, row, key) {
  const c = t.col;
  const i = Number(row) - 2;
  if (i >= 0 && i < t.rows.length && lessonKey_(t.rows[i], c, cfg) === key) return Number(row);
  for (let j = 0; j < t.rows.length; j++) if (lessonKey_(t.rows[j], c, cfg) === key) return j + 2;
  return -1;
}

function isPrepRow_(r, c) {
  const type = String(r[c['用途類型']] || '').trim() || '實驗課';
  return PREP_TYPES.indexOf(type) >= 0 && String(r[c['狀態']]).trim() !== '取消';
}

/** 寫入實驗名稱（和組數）；狀態空白的補「待準備」 */
function setLessonName_(t, rowNo, name, groups) {
  const c = t.col, sh = t.sheet;
  sh.getRange(rowNo, c['實驗名稱'] + 1).setValue(name);
  if (groups !== '' && groups !== undefined && groups !== null && isNumber_(groups)) sh.getRange(rowNo, c['組數'] + 1).setValue(Number(groups));
  if (!String(t.rows[rowNo - 2][c['狀態']]).trim()) sh.getRange(rowNo, c['狀態'] + 1).setValue('待準備');
}

// ---------------------------------------------------------------- A. 補實驗名稱

function nameDialog() {
  const html = DIALOG_STYLE + `
    <style>
      body { font-size: 14px; } .f { display: grid; grid-template-columns: repeat(4, 1fr); gap: 0 10px; } .f label { margin-top: 6px; }
      table.n { border-collapse: collapse; width: 100%; margin-top: 8px; } table.n th, table.n td { border: 1px solid #dadce0; padding: 4px 6px; font-size: 13px; text-align: left; }
      table.n th { background: #f1f3f4; position: sticky; top: 0; } table.n tr.wk td { background: #e8f0fe; font-weight: bold; }
      #list { max-height: 330px; overflow: auto; } .ap { background: #f8f9fa; border: 1px solid #dadce0; border-radius: 8px; padding: 10px; margin-top: 10px; }
      .ap .row2 { display: grid; grid-template-columns: 3fr 1fr 2fr; gap: 10px; align-items: end; } .cnt { color: #5f6368; margin-top: 6px; }
      td .nm { color: #188038; } td .no { color: #b3261e; }
    </style>
    <div class="f">
      <div><label>從</label><input type="date" id="from" onchange="load()"></div>
      <div><label>到</label><input type="date" id="to" onchange="load()"></div>
      <div><label>實驗室</label><select id="lab" onchange="draw()"><option value="">全部</option></select></div>
      <div><label>老師</label><select id="tch" onchange="draw()"><option value="">全部</option></select></div>
    </div>
    <label class="inline"><input type="checkbox" id="blank" checked onchange="draw()"> 只看還沒填名稱的</label>
    <div class="cnt" id="cnt"></div>
    <div id="list"></div>
    <div class="ap"><div class="row2">
      <div><label>實驗名稱（可從實驗套組選，也可以直接打）</label><input id="name" list="kits"></div>
      <div><label>組數（可空白）</label><input id="groups" inputmode="numeric"></div>
      <div><button class="primary" id="go" onclick="apply()" style="width:100%">套用到勾選的</button></div>
    </div><datalist id="kits"></datalist></div>
    <div class="btns"><button onclick="google.script.host.close()">關閉</button></div>
    <script>
      const D = __DATA__; let R = [];
      const $ = function (id) { return document.getElementById(id); };
      function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
      D.rooms.forEach(function (x) { $('lab').add(new Option(x, x)); });
      D.kits.forEach(function (k) { const o = document.createElement('option'); o.value = k; $('kits').appendChild(o); });
      $('from').value = D.from; $('to').value = D.to;
      function load() {
        $('list').innerHTML = '<p class="hint">讀取中…</p>';
        google.script.run.withSuccessHandler(function (rows) {
          R = rows;
          const t = $('tch'), cur = t.value; t.innerHTML = '<option value="">全部</option>';
          rows.map(function (r) { return r.teacher; }).filter(function (x, i, a) { return x && a.indexOf(x) === i; }).sort()
            .forEach(function (x) { t.add(new Option(x, x)); });
          t.value = cur; draw();
        }).withFailureHandler(function (e) { alert(e.message); }).nameRows($('from').value, $('to').value);
      }
      function shown() {
        return R.filter(function (r) {
          return (!$('lab').value || r.lab === $('lab').value) && (!$('tch').value || r.teacher === $('tch').value) && (!$('blank').checked || !r.name);
        });
      }
      function draw() {
        const list = shown();
        $('cnt').textContent = '共 ' + list.length + ' 堂' + (list.length ? '（預設全部勾選，不要的取消勾選）' : '');
        if (!list.length) { $('list').innerHTML = '<p class="hint">沒有符合的實驗課。</p>'; return; }
        let h = '<table class="n"><tr><th style="width:34px"><input type="checkbox" id="all" checked onchange="toggle(this.checked)"></th><th>日期</th><th>節次</th><th>實驗室</th><th>班級</th><th>老師</th><th>目前實驗名稱</th></tr>';
        let wk = '';
        list.forEach(function (r) {
          if (r.week !== wk) { wk = r.week; h += '<tr class="wk"><td colspan="7">' + esc(r.week) + ' 那一週</td></tr>'; }
          h += '<tr><td><input type="checkbox" class="c" data-i="' + R.indexOf(r) + '" checked></td><td>' + esc(r.roc.slice(4) + '（' + r.wd + '）') + '</td><td>' + esc(r.period) +
            '</td><td>' + esc(r.lab) + '</td><td>' + esc(r.cls) + '</td><td>' + esc(r.teacher) + '</td><td>' +
            (r.name ? '<span class="nm">' + esc(r.name) + '</span>' : '<span class="no">（未填）</span>') + '</td></tr>';
        });
        $('list').innerHTML = h + '</table>';
      }
      function toggle(on) { document.querySelectorAll('.c').forEach(function (c) { c.checked = on; }); }
      function apply() {
        const pick = Array.prototype.filter.call(document.querySelectorAll('.c'), function (c) { return c.checked; })
          .map(function (c) { const r = R[+c.getAttribute('data-i')]; return { row: r.row, key: r.key }; });
        const name = $('name').value.trim();
        if (!pick.length) { alert('請勾選要填的實驗課。'); return; }
        if (!name) { alert('請填實驗名稱。'); return; }
        const b = $('go'); b.disabled = true; b.textContent = '寫入中…';
        google.script.run.withSuccessHandler(function (m) {
          b.disabled = false; b.textContent = '套用到勾選的'; $('name').value = ''; $('groups').value = '';
          $('cnt').textContent = '✔ ' + m; load();
        }).withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '套用到勾選的'; })
          .applyNames({ items: pick, name: name, groups: $('groups').value.trim() });
      }
      load();
    </script>`;
  const cfg = labConfig_();
  const mon = mondayOf_(today_());
  const json = JSON.stringify({ rooms: cfg.rooms, kits: kits_().order, from: mon, to: addDays_(mon, 13) }).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html.replace('__DATA__', function () { return json; }))
    .setWidth(860).setHeight(720), '✏️ 補實驗名稱');
}

/** 期間內的實驗課、補做實驗 */
function nameRows(fromKey, toKey) {
  const cfg = labConfig_();
  fromKey = dateKey_(fromKey) || today_();
  toKey = dateKey_(toKey) || addDays_(fromKey, 13);
  const t = expTable_(), c = t.col;
  const out = [];
  t.rows.forEach(function (r, i) {
    const d = dateKey_(r[c['日期']]);
    if (!d || d < fromKey || d > toKey || !isPrepRow_(r, c)) return;
    const mon = mondayOf_(d);
    out.push({ row: i + 2, key: lessonKey_(r, c, cfg), date: d, roc: rocText_(d), wd: weekdayOf_(d),
      week: rocText_(mon).slice(4) + '（一）～' + rocText_(addDays_(mon, 4)).slice(4) + '（五）',
      period: String(r[c['節次']]).trim(), lab: String(r[c['教室']]).trim(), cls: classLabel_(r[c['班級']], cfg.classNames),
      teacher: String(r[c['教師']]).trim(), name: String(r[c['實驗名稱']]).trim() });
  });
  out.sort(function (a, b) {
    return a.date < b.date ? -1 : a.date > b.date ? 1 : naturalCompare_(a.period, b.period) || naturalCompare_(a.lab, b.lab);
  });
  return out;
}

/** p = {items: [{row, key}], name, groups} */
function applyNames(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('請填實驗名稱。');
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const cfg = labConfig_(), t = expTable_();
    let n = 0, miss = 0;
    (p.items || []).forEach(function (x) {
      const row = findLessonRow_(t, cfg, x.row, x.key);
      if (row < 0) { miss++; return; }
      setLessonName_(t, row, name, p.groups);
      n++;
    });
    return '已填 ' + n + ' 堂「' + name + '」' + (miss ? '（' + miss + ' 堂找不到，可能被刪了）' : '') + '。';
  } finally {
    lock.releaseLock();
  }
}

/** 從一堂課補名稱；same＝同一週、同一間、同一位老師、還沒填名稱的實驗課也一起填。p = {row, key, name, groups, same} */
function applyLessonName(p) {
  const name = String(p.name || '').trim();
  if (!name) throw new Error('請填實驗名稱。');
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const cfg = labConfig_(), t = expTable_(), c = t.col;
    const row = findLessonRow_(t, cfg, p.row, p.key);
    if (row < 0) throw new Error('找不到這堂課，可能已被刪除或改過。請重新整理。');
    const me = t.rows[row - 2];
    const rows = [row];
    if (p.same) {
      const mon = mondayOf_(dateKey_(me[c['日期']])), lab = String(me[c['教室']]).trim(), tch = String(me[c['教師']]).trim();
      t.rows.forEach(function (r, i) {
        const d = dateKey_(r[c['日期']]);
        if (i + 2 === row || !d || !isPrepRow_(r, c) || String(r[c['實驗名稱']]).trim()) return;
        if (mondayOf_(d) === mon && String(r[c['教室']]).trim() === lab && String(r[c['教師']]).trim() === tch) rows.push(i + 2);
      });
    }
    rows.forEach(function (r) { setLessonName_(t, r, name, p.groups); });
    let msg = '已填 ' + rows.length + ' 堂「' + name + '」。';
    const check = typeof experimentCheck_ === 'function' ? experimentCheck_(name, Number(p.groups) || Number(me[c['組數']]) || 0) : null;
    if (check) {
      const short = check.rows.filter(function (x) { return x.short; });
      msg += short.length ? '\n⚠ 器材不足：' + short.map(function (x) { return x.name + '（需 ' + x.need + '，有 ' + x.have + '）'; }).join('、') : '\n✔ 器材數量足夠';
    }
    return msg;
  } finally {
    lock.releaseLock();
  }
}

function lessonKits() { return kits_().order; }

// ---------------------------------------------------------------- C. 點一堂課

/** 點課的面板（電腦使用一覽、手機共用）。window.lessonPanel(b, U, opt, reload) */
function lessonJs_() {
  return `
  (function () {
    const css = document.createElement('style');
    css.textContent = '.lp-bg{position:fixed;inset:0;background:rgba(0,0,0,.35);z-index:50;overflow:auto;padding:16px 10px}' +
      '.lp{background:#fff;max-width:760px;margin:0 auto;border-radius:12px;padding:14px 16px;box-shadow:0 4px 20px rgba(0,0,0,.25)}' +
      '.lp h3{margin:0 30px 4px 0;font-size:18px}.lp .x{float:right;border:none;background:none;font-size:22px;cursor:pointer;color:#5f6368}' +
      '.lp .m{color:#3c4043;font-size:14px;margin:2px 0}.lp .tg{display:inline-block;padding:1px 8px;border-radius:8px;font-size:12px;margin-right:6px}' +
      '.lp .sec{border-top:1px solid #e0e3e7;margin-top:12px;padding-top:10px}.lp .sec>b{display:block;margin-bottom:6px}' +
      '.lp .nm{display:grid;grid-template-columns:3fr 1fr auto;gap:8px;align-items:center}.lp input{font-size:15px;padding:7px;border:1px solid #c9d1d9;border-radius:6px;width:100%;box-sizing:border-box}' +
      '.lp button.b{padding:8px 14px;font-size:14px;border:1px solid #1a73e8;background:#1a73e8;color:#fff;border-radius:6px;cursor:pointer}' +
      '.lp button.g{padding:8px 14px;font-size:14px;border:1px solid #dadce0;background:#fff;border-radius:6px;cursor:pointer}' +
      '.lp .msg{font-size:13px;margin-top:6px;white-space:pre-line}.lp .ok{color:#188038}.lp .err{color:#d93025}';
    document.head.appendChild(css);
    function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
    let kits = null;
    window.lessonPanel = function (b, U, opt, reload) {
      const prep = b.type === '實驗課' || b.type === '補做實驗';
      const bg = document.createElement('div'); bg.className = 'lp-bg';
      const w = '日一二三四五六'.charAt(new Date(b.date.replace(/-/g, '/')).getDay());
      const per = b.periodText === '午' ? '中午' : '第 ' + b.periodText + ' 節';
      let h = '<div class="lp"><button class="x" title="關閉">✕</button><h3>' + esc(b.content || b.type) + '</h3>' +
        '<div class="m"><span class="tg" style="background:' + ((U.colors || {})[b.type] || '#eee') + '">' + esc(b.type) + '</span>' +
        esc((+b.date.slice(0, 4) - 1911) + '.' + b.date.slice(5, 7) + '.' + b.date.slice(8) + '（' + w + '）' + per + '　' + b.lab) + '</div>' +
        '<div class="m">' + esc([b.cls, b.teacher ? b.teacher + '老師' : '', b.groups ? b.groups + ' 組' : '', b.status].filter(String).join('　')) + '</div>' +
        (b.note ? '<div class="m">📝 ' + esc(b.note) + '</div>' : '');
      if (prep) {
        h += '<div class="sec"><b>實驗名稱</b><div class="nm"><input class="lpName" list="lpKits" placeholder="例：酸鹼滴定" value="' + esc(b.content) + '">' +
          '<input class="lpGroups" placeholder="組數" inputmode="numeric" value="' + esc(b.groups || '') + '"><button class="b lpSave">儲存</button></div>' +
          '<label style="display:flex;gap:6px;align-items:center;font-size:13px;margin-top:6px"><input type="checkbox" class="lpSame" checked style="width:auto"> 同一週、同一間、' +
          esc(b.teacher || '同一位') + '老師還沒填名稱的實驗課也一起填</label><datalist id="lpKits"></datalist><div class="msg lpMsg"></div>' +
          (opt.print && b.content ? '<button class="g lpPrep" style="margin-top:8px">🖨 這天的器材準備單</button>' : '') + '</div>';
      }
      h += '<div class="sec"><b>這堂課的準備事項</b><div class="lpTodo"></div></div></div>';
      bg.innerHTML = h;
      document.body.appendChild(bg);
      let changed = false;
      const close = function () { bg.remove(); if (changed && reload) reload(); };
      bg.querySelector('.x').onclick = close;
      bg.onclick = function (e) { if (e.target === bg) close(); };
      if (prep) {
        const fill = function (list) { const d = bg.querySelector('#lpKits'); list.forEach(function (k) { const o = document.createElement('option'); o.value = k; d.appendChild(o); }); };
        if (kits) fill(kits); else google.script.run.withSuccessHandler(function (k) { kits = k; fill(k); }).lessonKits();
        bg.querySelector('.lpSave').onclick = function () {
          const btn = this, msg = bg.querySelector('.lpMsg'), name = bg.querySelector('.lpName').value.trim();
          if (!name) { msg.className = 'msg lpMsg err'; msg.textContent = '請填實驗名稱。'; return; }
          btn.disabled = true; btn.textContent = '儲存中…';
          google.script.run.withSuccessHandler(function (m) {
            btn.disabled = false; btn.textContent = '儲存'; msg.className = 'msg lpMsg ok'; msg.textContent = '✔ ' + m; changed = true;
          }).withFailureHandler(function (e) { btn.disabled = false; btn.textContent = '儲存'; msg.className = 'msg lpMsg err'; msg.textContent = e.message || e; })
            .applyLessonName({ row: b.row, key: [b.date, b.lab, String(b.periodText).trim(), b.cls].join('|'), name: name,
              groups: bg.querySelector('.lpGroups').value.trim(), same: bg.querySelector('.lpSame').checked });
        };
        const pb = bg.querySelector('.lpPrep');
        if (pb) pb.onclick = function () { pb.disabled = true; pb.textContent = '開啟中…'; google.script.run.showPrintDialog({ page: 'prep', from: b.date, to: b.date }); };
      }
      if (window.todoInit) {
        const box = bg.querySelector('.lpTodo'); box.id = 'lpTodo' + Date.now();
        window.todoInit(box.id, { add: true, lesson: { date: b.date, lab: b.lab, period: b.periodText, cls: b.cls, teacher: b.teacher },
          onSaved: function () { changed = true; } });
      } else bg.querySelector('.lpTodo').textContent = '（安裝「14_準備事項」後可以在這裡新增）';
    };
  })();`;
}

// ---------------------------------------------------------------- F. 手機「今天」

function nextSchoolDay_(key) {
  let k = addDays_(key, 1);
  while (weekdayOf_(k) === '六' || weekdayOf_(k) === '日') k = addDays_(k, 1);
  return k;
}

/** 今天、下一個上課日的課；待準備事項（過期＋到下一個上課日）；需補充幾項 */
function homeData() {
  const cfg = labConfig_();
  const today = today_(), next = nextSchoolDay_(today);
  const all = bookings_(today, next, cfg);
  const byDay = function (k) {
    return all.filter(function (b) { return b.date === k; }).sort(function (a, b) {
      return (a.pis[0] === undefined ? 99 : a.pis[0]) - (b.pis[0] === undefined ? 99 : b.pis[0]) || naturalCompare_(a.lab, b.lab);
    });
  };
  let todos = [];
  if (typeof todoData === 'function') {
    todos = todoData().list.filter(function (x) { return x.status === '待準備' && (!x.date || x.date <= next); });
  }
  let restock = null;
  try { if (typeof restockItems_ === 'function') restock = restockItems_().length; } catch (e) { restock = null; }
  const now = new Date();
  return {
    colors: cfg.colors, periods: cfg.periods, today: { key: today, roc: rocText_(today), wd: weekdayOf_(today), list: byDay(today) },
    next: { key: next, roc: rocText_(next), wd: weekdayOf_(next), list: byDay(next) },
    todos: todos, restock: restock, nowHm: Utilities.formatDate(now, SCHOOL_TZ, 'HH:mm'),
  };
}

/** 手機「今天」分頁的畫面程式：homeInit(容器id) */
function homeJs_() {
  return `
  (function () {
    const css = document.createElement('style');
    css.textContent = '.hm h3{margin:14px 2px 6px;font-size:16px;color:#1a73e8}.hm h3 small{color:#5f6368;font-weight:normal;font-size:13px}' +
      '.hm .ls{background:#fff;border-radius:10px;padding:8px 10px;margin-bottom:6px;border-left:6px solid #d2e3fc;display:flex;gap:10px;align-items:center}' +
      '.hm .ls.now{box-shadow:0 0 0 2px #d93025}.hm .ls .p{font-weight:700;min-width:44px;text-align:center;font-size:15px}.hm .ls .t{font-size:15px;font-weight:600}' +
      '.hm .ls .s{font-size:13px;color:#5f6368}.hm .none{color:#888;padding:6px 4px}.hm .td{background:#fff;border-radius:10px;padding:8px 10px;margin-bottom:6px;display:flex;gap:10px;align-items:flex-start;border-left:6px solid #fbbc04}' +
      '.hm .td.late{border-left-color:#d93025}.hm .td input{width:22px;height:22px;flex:none;margin-top:2px}.hm .td.done{opacity:.5}' +
      '.hm .rs{background:#fff;border-radius:10px;padding:10px;display:flex;justify-content:space-between;align-items:center;margin-top:6px}.hm .rs b{color:#d93025;font-size:20px}' +
      '.hm .rf{float:right;font-size:13px;font-weight:normal}';
    document.head.appendChild(css);
    function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
    function short(lab) { return String(lab).replace(/學?實驗室/, '').replace(/能?教室/, ''); }
    window.homeInit = function (id) {
      const box = document.getElementById(id);
      let H = null;
      function load() {
        box.innerHTML = '<p style="color:#666">讀取中…</p>';
        google.script.run.withSuccessHandler(function (d) { H = d; draw(); }).withFailureHandler(function (e) { box.textContent = '讀取失敗：' + (e.message || e); }).homeData();
      }
      function lessons(day, isToday) {
        if (!day.list.length) return '<div class="none">沒有登記的課</div>';
        return day.list.map(function (b, i) {
          let now = false;
          if (isToday) b.pis.forEach(function (pi) { const p = H.periods[pi]; if (p && H.nowHm >= p.start && H.nowHm < p.end) now = true; });
          return '<div class="ls' + (now ? ' now' : '') + '" data-d="' + (isToday ? 't' : 'n') + '" data-i="' + i + '" style="border-left-color:' + (H.colors[b.type] || '#ddd') + '">' +
            '<div class="p">' + esc(b.periodText === '午' ? '午' : b.periodText) + '</div><div><div class="t">' + esc(b.content || b.type) + '</div>' +
            '<div class="s">' + esc([short(b.lab), b.cls, b.teacher].filter(String).join('　')) + (now ? '　<b style="color:#d93025">● 現在</b>' : '') + '</div></div></div>';
        }).join('');
      }
      function draw() {
        let h = '<div class="hm"><h3>今天 <small>' + esc(H.today.roc + '（' + H.today.wd + '）') + '</small><a href="#" class="rf">重新整理</a></h3>' + lessons(H.today, true) +
          '<h3>下一個上課日 <small>' + esc(H.next.roc + '（' + H.next.wd + '）') + '</small></h3>' + lessons(H.next, false) +
          '<h3>待準備事項 <small>' + H.todos.length + ' 項（過期＋到下一個上課日）</small></h3>';
        if (!H.todos.length) h += '<div class="none">沒有 👍</div>';
        H.todos.forEach(function (x) {
          const late = x.date && x.date < H.today.key;
          h += '<div class="td' + (late ? ' late' : '') + '"><input type="checkbox" data-id="' + esc(x.id) + '"><div><div class="t" style="font-weight:600">' + esc(x.what) +
            (x.qty ? ' × ' + esc(x.qty) + ' ' + esc(x.unit) : '') + '</div><div class="s" style="font-size:13px;color:#5f6368">' +
            esc([x.date ? x.roc.slice(4) + '（' + x.wd + '）' + (late ? ' 已過期' : '') : '', short(x.lab), x.period ? (x.period === '午' ? '中午' : '第' + x.period + '節') : '', x.cls].filter(String).join('　')) +
            (x.place ? '<br>📍 ' + esc(x.place) : '') + (x.short ? '　<b style="color:#d93025">⚠ 可能不夠</b>' : '') + '</div></div></div>';
        });
        if (H.restock !== null) h += '<div class="rs"><span>需補充（低於安全存量）</span><span><b>' + H.restock + '</b> 項 ' +
          (window.tab ? '<a href="#" class="go">看清單</a>' : '') + '</span></div>';
        box.innerHTML = h + '</div>';
        box.querySelector('.rf').onclick = function (e) { e.preventDefault(); load(); };
        const go = box.querySelector('.go'); if (go) go.onclick = function (e) { e.preventDefault(); window.tab('need'); };
        box.querySelectorAll('.td input').forEach(function (c) {
          c.onchange = function () {
            c.parentNode.classList.toggle('done', c.checked);
            google.script.run.withFailureHandler(function (e) { alert('沒有存到：' + (e.message || e)); load(); })
              .todoSet([c.getAttribute('data-id')], c.checked ? '已準備' : '待準備');
          };
        });
        if (window.lessonPanel) box.querySelectorAll('.ls').forEach(function (e) {
          e.onclick = function () {
            const b = (e.getAttribute('data-d') === 't' ? H.today : H.next).list[+e.getAttribute('data-i')];
            window.lessonPanel(b, { colors: H.colors }, {}, load);
          };
        });
      }
      load();
    };
  })();`;
}
