/**
 * 實驗室藥品器材耗材管理系統 — 每天早上 email 提醒、實驗套組一鍵產生準備事項
 *
 * 📧 每天早上 email 提醒（🔧 維護 → 📧 每天早上 email 提醒（設定））
 *   ・每天早上（預設 7 點）寄一封信：今天的課、到下一個上課日要準備的東西（含過期沒準備的）、
 *     器材不夠的實驗、還沒填實驗名稱的課、借出逾期、請購待處理、需補充
 *   ・週末、沒有任何事情的日子不寄（可以在設定裡改）
 *   ・設定存在這份試算表的「文件屬性」，不會多一張工作表
 *
 * 📝 實驗套組一鍵產生準備事項
 *   ・網頁：點一堂課 → 器材下面「一鍵產生準備事項」→ 帶入套組的器材和數量（組數×每組數量）→ 看一眼按儲存
 *   ・試算表：🗓 實驗室使用 → 📝 依實驗套組產生準備事項（一段期間）→ 這段期間每堂有套組的實驗課一次產生
 *   ・已經在準備事項裡的（同一天、同一間、同一節、同樣東西）不會重複產生
 *
 * 需要「01」～「17」。
 */

// ---------------------------------------------------------------- 📝 實驗套組 → 準備事項

/**
 * 一堂課的套組器材，轉成準備事項的列。
 * lesson = {date, lab, period, cls, teacher, name, groups}
 * 回傳 {kit, groups, note, items: [{what, qty, unit, code, info, short}], dup: 已經有的項數}；沒有這個套組回傳 null。
 */
function kitTodoItems_(lesson, stock, existing) {
  const name = String(lesson.name || '').trim();
  const k = kits_().map[name];
  if (!k) return null;
  stock = stock || stockIndex_();
  const ck = experimentCheck_(name, Number(lesson.groups) || 0, stock);
  const have = existing || todoKeys_();
  const at = [dateKey_(lesson.date) || '', String(lesson.lab || '').trim(), periodVal_(lesson.period)].join('|');
  let dup = 0;
  const items = [];
  k.items.forEach(function (it, i) {
    const r = ck.rows[i], s = stock(it.code, it.name);
    const code = it.code || (s ? s.code : '');
    if (have[at + '|' + (code ? 'c:' + code : 'n:' + it.name)] || have[at + '|n:' + it.name]) { dup++; return; }
    const perUnit = String(it.perText).replace(/^[\d.\s]+/, '').trim();
    const noNum = r.need === '';
    items.push({
      what: it.name, code: code,
      qty: noNum ? (it.perText ? '每組' + it.perText : '') : r.need,
      unit: noNum ? '' : (perUnit || (s ? s.unit : '')),
      info: '📍 ' + (r.loc || '—') + '　庫存 ' + r.haveText + (r.unit ? ' ' + r.unit : '') +
        (r.state === 'bad' ? '　⚠ 不夠' : r.state === 'warn' && r.note ? '　⚠ ' + r.note : ''),
      short: r.state === 'bad',
    });
  });
  return { kit: name, groups: ck.groups, dup: dup, items: items,
    note: '實驗套組：' + name + (ck.groups ? '（' + ck.groups + ' 組）' : '') };
}

/** 已經有的準備事項（不含取消）：「日期|實驗室|節次|c:編號」「…|n:事項」→ true */
function todoKeys_() {
  const out = {};
  const sh = SpreadsheetApp.getActive().getSheetByName(TODO_SHEET);
  if (!sh || sh.getLastRow() < 2) return out;
  const t = getTable_(TODO_SHEET), c = t.col;
  t.rows.forEach(function (r) {
    if ((String(r[c['狀態']]).trim() || '待準備') === '取消') return;
    const at = [dateKey_(r[c['需要日期']]) || '', String(r[c['實驗室']]).trim(), periodVal_(r[c['節次']])].join('|');
    const code = String(r[c['編號']]).trim(), what = String(r[c['事項']]).trim();
    if (code) out[at + '|c:' + code] = true;
    if (what) out[at + '|n:' + what] = true;
  });
  return out;
}

/** 網頁：一堂課要帶入的準備事項（先給使用者看，按儲存才寫入）。 */
function kitTodoPreview(lesson) {
  return withReadMemo_(function () { return kitTodoItems_(lesson || {}); });
}

/** 某段期間有套組的實驗課（不含取消、已歸還）。 */
function kitLessons_(fromKey, toKey) {
  const cfg = labConfig_();
  const kits = kits_().map;
  return bookings_(fromKey, toKey, cfg).filter(function (b) {
    return (b.type === '實驗課' || b.type === '補做實驗') && b.content && kits[b.content] && b.status !== '已歸還';
  }).sort(function (a, b) {
    return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.pis[0] || 0) - (b.pis[0] || 0) || naturalCompare_(a.lab, b.lab);
  }).map(function (b) {
    return { date: b.date, lab: b.lab, period: b.periodText, cls: b.cls, teacher: b.teacher, name: b.content, groups: b.groups };
  });
}

/** 試算表：一段期間一次產生。apply＝false 只算要產生幾項（預覽）。 */
function kitTodoRange(from, to, apply) {
  const f = dateKey_(from), t = dateKey_(to);
  if (!f || !t || f > t) throw new Error('日期範圍不對。');
  if (addDays_(f, 62) < t) throw new Error('一次最多兩個月，請把範圍縮短。');
  const stock = stockIndex_(), have = todoKeys_();
  const plan = kitLessons_(f, t).map(function (l) {
    const x = kitTodoItems_(l, stock, have);
    return { lesson: l, x: x };
  }).filter(function (p) { return p.x; });
  const total = plan.reduce(function (s, p) { return s + p.x.items.length; }, 0);
  const dup = plan.reduce(function (s, p) { return s + p.x.dup; }, 0);
  const lines = plan.filter(function (p) { return p.x.items.length; }).map(function (p) {
    const l = p.lesson;
    const sh = p.x.items.filter(function (i) { return i.short; }).length;
    return rocText_(l.date) + '（' + weekdayOf_(l.date) + '）' + (l.period ? perText_(l.period) : '') + ' ' + l.lab + ' ' + (l.cls || '') +
      '　' + l.name + '：' + p.x.items.length + ' 項' + (sh ? '（' + sh + ' 項庫存不夠）' : '');
  });
  if (!apply) return { lessons: plan.length, total: total, dup: dup, lines: lines };
  if (!total) return '沒有要新增的：這段期間的實驗課' + (plan.length ? '，準備事項都已經產生過了。' : '都沒有對應的實驗套組。');
  let n = 0;
  plan.forEach(function (p) {
    if (!p.x.items.length) return;
    const l = p.lesson;
    todoAdd({ date: l.date, period: l.period, lab: l.lab, cls: l.cls, teacher: l.teacher, note: p.x.note,
      items: p.x.items.map(function (i) { return { what: i.what, qty: i.qty, unit: i.unit, code: i.code }; }) });
    n += p.x.items.length;
  });
  return '已產生 ' + n + ' 項準備事項（' + lines.length + ' 堂課）' + (dup ? '，' + dup + ' 項之前已經有了，沒有重複加。' : '。') +
    '\n\n' + lines.join('\n');
}

function kitTodoDialog() {
  const d = new Date();
  d.setDate(d.getDate() + 7);
  const html = DIALOG_STYLE + `
    <p>這段期間「實驗課」「補做實驗」中，<b>實驗名稱有對應實驗套組</b>的課，會把套組的器材（組數 × 每組數量）一次加進準備事項。</p>
    <div class="row"><div style="flex:1"><label>從</label><input type="date" id="from" onchange="pv()"></div>
      <div style="flex:1"><label>到</label><input type="date" id="to" onchange="pv()"></div></div>
    <div id="pv" class="hint" style="margin-top:12px;white-space:pre-line;max-height:170px;overflow:auto">計算中…</div>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()" disabled>產生</button></div>
    <script>
      const D = __DATA__;
      document.getElementById('from').value = D.from; document.getElementById('to').value = D.to;
      function pv() {
        const el = document.getElementById('pv'), b = document.getElementById('go');
        el.textContent = '計算中…'; b.disabled = true;
        google.script.run.withSuccessHandler(function (r) {
          el.textContent = r.lessons ? ('有套組的實驗課 ' + r.lessons + ' 堂，要新增 ' + r.total + ' 項' + (r.dup ? '（已經有的 ' + r.dup + ' 項會略過）' : '') + '：\\n' + r.lines.join('\\n')) :
            '這段期間沒有「實驗名稱」對應到實驗套組的實驗課。';
          b.disabled = !r.total;
        }).withFailureHandler(function (e) { el.textContent = '⚠ ' + e.message; })
          .kitTodoRange(document.getElementById('from').value, document.getElementById('to').value, false);
      }
      function go() {
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '產生中…';
        google.script.run.withSuccessHandler(showDone)
          .withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '產生'; })
          .kitTodoRange(document.getElementById('from').value, document.getElementById('to').value, true);
      }
      pv();
    </script>`;
  showDialog_(html, { from: today_(), to: Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') }, '📝 依實驗套組產生準備事項', 430);
}

// ---------------------------------------------------------------- 📧 每天早上 email 提醒

const MAIL_PROP = 'morningMail';

/** 設定：{on, to, hour, weekend, empty} */
function mailConfig_() {
  let c = {};
  try { c = JSON.parse(PropertiesService.getDocumentProperties().getProperty(MAIL_PROP) || '{}'); } catch (e) { c = {}; }
  return { on: !!c.on, to: c.to || '', hour: c.hour >= 5 && c.hour <= 9 ? c.hour : 7, weekend: !!c.weekend, empty: !!c.empty };
}

/** 收件人：用逗號、分號、空白、換行分開；最多 10 個，格式不對的會報錯。 */
function mailList_(s) {
  const list = String(s || '').split(/[,，;；\s]+/).map(function (x) { return x.trim(); }).filter(String);
  const bad = list.filter(function (x) { return !/^[^@\s<>"']+@[^@\s<>"']+\.[^@\s<>"']+$/.test(x); });
  if (bad.length) throw new Error('信箱格式不對：' + bad.join('、'));
  if (list.length > 10) throw new Error('收件人最多 10 個。');
  return list;
}

function mailMe_() {
  try { return Session.getEffectiveUser().getEmail() || ''; } catch (e) { return ''; }
}

function mailSetupDialog() {
  const c = mailConfig_();
  const html = DIALOG_STYLE + `
    <p>每天早上寄一封信，告訴你今天的課、要準備的東西、器材不夠的實驗等。</p>
    <label>寄給誰（多個用逗號分開）</label><input id="to">
    <label>幾點寄</label><select id="hour"></select>
    <label class="inline"><input type="checkbox" id="weekend">週六、週日也寄（只有當天有登記課時才會寄）</label>
    <label class="inline"><input type="checkbox" id="empty">沒有任何事情的上課日也寄（信裡會寫「今天沒有要準備的」）</label>
    <p class="hint" id="st"></p>
    <div class="btns" style="display:flex;flex-wrap:wrap;gap:8px;justify-content:flex-end">
      <button onclick="test(this)">先寄一封測試信</button>
      <button id="off" onclick="save(false, this)">關閉提醒</button>
      <button class="primary" onclick="save(true, this)">儲存並開啟</button></div>
    <script>
      const D = __DATA__;
      const h = document.getElementById('hour');
      [5, 6, 7, 8, 9].forEach(function (x) { h.add(new Option('早上 ' + x + ' 點左右', x)); });
      h.value = D.hour; document.getElementById('to').value = D.to || D.me;
      document.getElementById('weekend').checked = D.weekend; document.getElementById('empty').checked = D.empty;
      document.getElementById('st').textContent = D.on ? '目前：✔ 已開啟，每天早上 ' + D.hour + ' 點左右寄給 ' + D.to : '目前：還沒開啟';
      document.getElementById('off').style.display = D.on ? '' : 'none';
      function val() { return { to: document.getElementById('to').value, hour: Number(h.value),
        weekend: document.getElementById('weekend').checked, empty: document.getElementById('empty').checked }; }
      function busy(b, t) { b.disabled = true; b.dataset.t = b.textContent; b.textContent = t; }
      function free(b) { b.disabled = false; b.textContent = b.dataset.t; }
      function save(on, b) {
        busy(b, on ? '設定中…' : '關閉中…');
        google.script.run.withSuccessHandler(showDone).withFailureHandler(function (e) { alert(e.message); free(b); })
          .mailSave(Object.assign(val(), { on: on }));
      }
      function test(b) {
        busy(b, '寄送中…');
        google.script.run.withSuccessHandler(function (m) { free(b); alert(m); document.getElementById('__msg').className = 'msg'; })
          .withFailureHandler(function (e) { alert(e.message); free(b); }).mailTest(val().to);
      }
    </script>`;
  showDialog_(html, Object.assign(c, { me: mailMe_() }), '📧 每天早上 email 提醒', 470);
}

/** 儲存設定，並建立（或移除）每天的排程。 */
function mailSave(p) {
  const to = mailList_(p.to);
  if (p.on && !to.length) throw new Error('請填至少一個收件信箱。');
  const hour = Math.max(5, Math.min(9, Number(p.hour) || 7));
  const c = { on: !!p.on, to: to.join(', '), hour: hour, weekend: !!p.weekend, empty: !!p.empty };
  PropertiesService.getDocumentProperties().setProperty(MAIL_PROP, JSON.stringify(c));
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'morningMail') ScriptApp.deleteTrigger(t); });
  if (!c.on) return '已關閉每天早上的 email 提醒。';
  ScriptApp.newTrigger('morningMail').timeBased().everyDays(1).atHour(hour).nearMinute(0).inTimezone(SCHOOL_TZ).create();
  return '已開啟：每天早上 ' + hour + ' 點左右（前後約 15 分鐘）寄到\n' + c.to +
    '\n\n' + (c.weekend ? '週末有登記課的日子也會寄。' : '週六、週日不寄。') + (c.empty ? '' : '\n沒有任何事情的日子不寄。');
}

/** 測試：馬上寄一封（不管週末、有沒有事情）。 */
function mailTest(toText) {
  const to = mailList_(toText);
  if (!to.length) throw new Error('請填收件信箱。');
  const d = morningDigest_(today_());
  sendDigest_(d, to, true);
  return '已寄出測試信到 ' + to.join(', ') + '，請到信箱看看（可能在「促銷內容」或垃圾郵件匣）。\n今天剩下可寄 ' + MailApp.getRemainingDailyQuota() + ' 封。';
}

/** 排程每天呼叫。 */
function morningMail() {
  const c = mailConfig_();
  if (!c.on || !c.to) return;
  const props = PropertiesService.getDocumentProperties();
  const today = today_();
  if (props.getProperty('morningMailLast') === today) return;   // 同一天不重複寄
  const d = morningDigest_(today);
  const weekend = weekdayOf_(today) === '六' || weekdayOf_(today) === '日';
  if (weekend && !(c.weekend && d.today.length)) return;
  if (d.empty && !c.empty) return;
  if (d.holiday && !d.today.length && !d.urgent) return;   // 整天放假，又沒有過期／今天要準備的
  sendDigest_(d, mailList_(c.to), false);
  props.setProperty('morningMailLast', today);
}

/** 整理信件內容（只讀，不改資料）。 */
function morningDigest_(today) {
  return withReadMemo_(function () {
    const cfg = labConfig_();
    const next = nextSchoolDay_(today);
    const sortB = function (a, b) {
      return (a.pis[0] === undefined ? 99 : a.pis[0]) - (b.pis[0] === undefined ? 99 : b.pis[0]) || naturalCompare_(a.lab, b.lab);
    };
    const all = bookings_(today, next, cfg);
    const holiday = all.some(function (b) { return b.date === today && b.type === '放假'; });
    const real = all.filter(function (b) { return b.type !== '放假'; });
    const dayT = real.filter(function (b) { return b.date === today; }).sort(sortB);
    const dayN = real.filter(function (b) { return b.date === next; }).sort(sortB);

    // 器材不夠的實驗、沒填實驗名稱的課（今天＋下一個上課日）
    const stock = stockIndex_(), kits = kits_().map;
    const short = [], unnamed = [];
    dayT.concat(dayN).forEach(function (b) {
      if (b.type !== '實驗課' && b.type !== '補做實驗') return;
      if (!b.content) { unnamed.push(b); return; }
      if (!kits[b.content]) return;
      const ck = experimentCheck_(b.content, Number(b.groups) || 0, stock);
      const bad = ck.rows.filter(function (r) { return r.state === 'bad'; });
      if (bad.length) short.push({ b: b, rows: bad });
    });

    // 準備事項：過期沒準備、今天、下一個上課日、沒寫日期
    let todos = [];
    if (typeof todoData === 'function') {
      todos = todoData().list.filter(function (x) { return x.status === '待準備' && (!x.date || x.date <= next); });
    }
    const late = todos.filter(function (x) { return x.date && x.date < today; });
    const tdT = todos.filter(function (x) { return x.date === today; });
    const tdN = todos.filter(function (x) { return x.date && x.date > today && x.date <= next; });
    const tdNo = todos.filter(function (x) { return !x.date; });

    const kpi = {};
    try { if (typeof restockItems_ === 'function') kpi.restock = restockItems_().length; } catch (e) { }
    if (typeof appSheetCount_ === 'function') {
      try { kpi.loans = appSheetCount_(LOAN_SHEET, '狀態', ['借出中'], '預計歸還'); } catch (e) { }
      try { kpi.purchase = appSheetCount_(REQ_SHEET, '狀態', ['待處理']); } catch (e) { }
    }
    let url = '';
    try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }
    const s = getSettings_();
    const d = {
      school: String(s.params['學校名稱'] || ''), key: today, roc: rocText_(today), wd: weekdayOf_(today),
      next: next, nextRoc: rocText_(next), nextWd: weekdayOf_(next), holiday: holiday,
      today: dayT, nextList: dayN, short: short, unnamed: unnamed,
      late: late, tdT: tdT, tdN: tdN, tdNo: tdNo, kpi: kpi, periods: cfg.periods,
      url: url, sheetUrl: SpreadsheetApp.getActive().getUrl ? SpreadsheetApp.getActive().getUrl() : '',
    };
    d.urgent = late.length + tdT.length + short.length;
    d.empty = !dayT.length && !dayN.length && !todos.length && !short.length &&
      !(kpi.loans && kpi.loans.late);
    return d;
  });
}

function sendDigest_(d, to, test) {
  if (MailApp.getRemainingDailyQuota() < to.length) throw new Error('今天 Google 寄信的額度用完了，明天再試。');
  const n = d.today.length, m = d.late.length + d.tdT.length + d.tdN.length + d.tdNo.length;
  const subject = (test ? '【測試】' : '') + '🧪 實驗室提醒 ' + d.roc + '（' + d.wd + '）：' +
    (n ? n + ' 堂課' : '今天沒有登記課') + (m ? '・' + m + ' 項要準備' : '') +
    (d.late.length ? '・⚠ ' + d.late.length + ' 項過期' : '') + (d.short.length ? '・⚠ ' + d.short.length + ' 堂器材不夠' : '');
  MailApp.sendEmail({ to: to.join(','), subject: subject, htmlBody: digestHtml_(d), body: digestText_(d), name: '實驗室管理' });
}

// ---- 信件內容

function digestPer_(b, periods) {
  const t = b.periodText ? perText_(b.periodText) : '';
  const p0 = periods[b.pis[0]], p1 = periods[b.pis[b.pis.length - 1]];
  return t + (p0 && p1 ? ' ' + p0.start + '–' + p1.end : '');
}
function digestLesson_(b) {
  return [b.lab, b.cls, b.teacher ? b.teacher + '老師' : '', b.content || (b.type === '實驗課' || b.type === '補做實驗' ? '（實驗名稱未填）' : b.type),
    b.groups ? b.groups + ' 組' : ''].filter(String).join('　');
}
function digestTodo_(x) {
  return x.what + (x.qty ? ' × ' + x.qty + (x.unit ? ' ' + x.unit : '') : '');
}
function digestTodoWhere_(x, noDate) {
  return [x.date && !noDate ? md_(x.date) + '（' + x.wd + '）' : '', x.lab, x.period ? perText_(x.period) : '', x.cls].filter(String).join(' ');
}
function md_(k) { return Number(k.slice(5, 7)) + '/' + Number(k.slice(8, 10)); }

function digestText_(d) {
  const L = [];
  L.push(d.school + ' 實驗室提醒　' + d.roc + '（' + d.wd + '）', '');
  const sec = function (title, rows) { if (!rows.length) return; L.push('■ ' + title); rows.forEach(function (r) { L.push('・' + r); }); L.push(''); };
  sec('今天的課', d.today.map(function (b) { return digestPer_(b, d.periods) + '　' + digestLesson_(b); }));
  sec('⚠ 過期還沒準備', d.late.map(function (x) { return digestTodo_(x) + '（' + digestTodoWhere_(x) + '）'; }));
  sec('今天要準備', d.tdT.map(function (x) { return digestTodo_(x) + '（' + digestTodoWhere_(x) + '）'; }));
  sec(d.nextRoc + '（' + d.nextWd + '）要準備', d.tdN.map(function (x) { return digestTodo_(x) + '（' + digestTodoWhere_(x) + '）'; }));
  sec('沒寫日期的準備事項', d.tdNo.map(digestTodo_));
  sec('⚠ 器材不夠', d.short.map(function (s) {
    return md_(s.b.date) + ' ' + digestLesson_(s.b) + '：' + s.rows.map(function (r) { return r.name + '（要 ' + numText_(r.need) + '，有 ' + r.haveText + '）'; }).join('、');
  }));
  sec('還沒填實驗名稱', d.unnamed.map(function (b) { return md_(b.date) + ' ' + digestPer_(b, d.periods) + '　' + digestLesson_(b); }));
  sec(d.nextRoc + '（' + d.nextWd + '）的課', d.nextList.map(function (b) { return digestPer_(b, d.periods) + '　' + digestLesson_(b); }));
  const k = [];
  if (d.kpi.loans && d.kpi.loans.n) k.push('借出中 ' + d.kpi.loans.n + ' 筆' + (d.kpi.loans.late ? '（逾期 ' + d.kpi.loans.late + '）' : ''));
  if (d.kpi.purchase && d.kpi.purchase.n) k.push('請購待處理 ' + d.kpi.purchase.n + ' 筆');
  if (d.kpi.restock) k.push('需補充 ' + d.kpi.restock + ' 項');
  if (k.length) L.push('■ 其他：' + k.join('　'), '');
  if (d.empty) L.push('今天沒有要準備的事情 🎉', '');
  if (d.url) L.push('打開網頁工作台：' + d.url);
  L.push('', '（這封信由實驗室管理系統自動寄出；要停止或改時間：試算表 → 🧪 實驗室管理 → 🔧 維護 → 📧 每天早上 email 提醒）');
  return L.join('\n');
}

function digestHtml_(d) {
  const e = esc_;
  const box = 'margin:0 0 14px;border:1px solid #e3e6ea;border-radius:12px;overflow:hidden';
  const head = function (t, color) {
    return '<div style="background:' + (color || '#f1f5fb') + ';padding:8px 12px;font-weight:700;font-size:15px">' + t + '</div>';
  };
  const row = function (a, b, warn) {
    return '<tr><td style="padding:7px 12px;border-top:1px solid #eef0f3;width:34%;vertical-align:top;color:' + (warn ? '#b3261e' : '#5f6368') +
      ';font-size:13px">' + a + '</td><td style="padding:7px 12px;border-top:1px solid #eef0f3;font-size:14px">' + b + '</td></tr>';
  };
  const sec = function (title, rows, color) {
    if (!rows.length) return '';
    return '<div style="' + box + '">' + head(title, color) + '<table style="border-collapse:collapse;width:100%">' + rows.join('') + '</table></div>';
  };
  const lessonRows = function (list) {
    return list.map(function (b) { return row(e(digestPer_(b, d.periods)), e(digestLesson_(b)) + (b.note ? '<br><small style="color:#5f6368">📝 ' + e(b.note) + '</small>' : '')); });
  };
  const todoRows = function (list, warn, noDate) {
    return list.map(function (x) {
      const j = x.judge && x.judge.state === 'bad' ? '<br><small style="color:#b3261e">⚠ ' + e(x.judge.text) + '</small>' : '';
      return row(e(digestTodoWhere_(x, noDate) || '—'), '<b>' + e(digestTodo_(x)) + '</b>' + (x.place ? '　<small style="color:#5f6368">📍 ' + e(x.place) + '</small>' : '') +
        (x.buy ? '　<small>🛒 ' + e(x.buy) + '</small>' : '') + j, warn);
    });
  };
  let h = '<div style="font-family:\'Noto Sans TC\',\'Microsoft JhengHei\',\'PingFang TC\',sans-serif;max-width:640px;margin:0 auto;color:#1b2433;line-height:1.5">' +
    '<div style="padding:4px 2px 12px"><div style="font-size:13px;color:#5f6368">' + e(d.school) + '</div>' +
    '<div style="font-size:22px;font-weight:800">🧪 實驗室提醒　' + e(d.roc) + '（' + e(d.wd) + '）</div></div>';
  if (d.holiday) h += '<p style="margin:0 0 12px;color:#5f6368">今天有登記「放假」。</p>';
  h += sec('📅 今天的課（' + d.today.length + ' 堂）', lessonRows(d.today));
  h += sec('⚠ 過期還沒準備（' + d.late.length + '）', todoRows(d.late, true), '#fce8e6');
  h += sec('📝 今天要準備（' + d.tdT.length + '）', todoRows(d.tdT, false, true), '#fef3e2');
  h += sec('📝 ' + e(d.nextRoc) + '（' + e(d.nextWd) + '）要準備（' + d.tdN.length + '）', todoRows(d.tdN, false, d.tdN.every(function (x) { return x.date === d.next; })));
  h += sec('📝 沒寫日期的準備事項（' + d.tdNo.length + '）', todoRows(d.tdNo));
  h += sec('⚠ 器材不夠', d.short.map(function (s) {
    return row(e(md_(s.b.date) + ' ' + digestPer_(s.b, d.periods)), e(digestLesson_(s.b)) + '<br>' + s.rows.map(function (r) {
      return '<span style="color:#b3261e">' + e(r.name) + '：要 ' + e(numText_(r.need)) + '，有 ' + e(r.haveText) + (r.unit ? ' ' + e(r.unit) : '') + '</span>';
    }).join('<br>'), true);
  }), '#fce8e6');
  h += sec('✏️ 還沒填實驗名稱（' + d.unnamed.length + '）', lessonRows(d.unnamed));
  h += sec('🗓 ' + e(d.nextRoc) + '（' + e(d.nextWd) + '）的課（' + d.nextList.length + ' 堂）', lessonRows(d.nextList));
  const k = [];
  if (d.kpi.loans && d.kpi.loans.n) k.push('借出中 <b>' + d.kpi.loans.n + '</b> 筆' + (d.kpi.loans.late ? '（<span style="color:#b3261e">逾期 ' + d.kpi.loans.late + '</span>）' : ''));
  if (d.kpi.purchase && d.kpi.purchase.n) k.push('請購待處理 <b>' + d.kpi.purchase.n + '</b> 筆');
  if (d.kpi.restock) k.push('需補充 <b>' + d.kpi.restock + '</b> 項');
  if (k.length) h += '<p style="margin:4px 2px 14px;font-size:14px">' + k.join('　・　') + '</p>';
  if (d.empty) h += '<p style="font-size:16px;margin:10px 2px 16px">今天沒有要準備的事情 🎉</p>';
  if (d.url) h += '<p style="margin:6px 0 18px"><a href="' + e(d.url) + '" style="display:inline-block;background:#1F5EDB;color:#fff;text-decoration:none;padding:10px 18px;border-radius:10px;font-weight:700">打開網頁工作台</a></p>';
  h += '<p style="font-size:12px;color:#9aa0a6">這封信由實驗室管理系統自動寄出。要停止或改時間：試算表 → 🧪 實驗室管理 → 🔧 維護 → 📧 每天早上 email 提醒。</p></div>';
  return h;
}
