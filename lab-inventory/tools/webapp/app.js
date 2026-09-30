// 網頁工作台：今天、課表、待辦、盤點、更多（需補充、品項查詢、外觀）
var S = {
  D: null, route: 'today', wide: false, kits: null, restock: null,
  todayTab: 'today', heatTab: 'today',
  week: { mode: 'week', start: '', lab: '', kw: '', U: null, key: '', sel: null, loading: false },
  todoView: 'todo',
  count: { data: null, room: '', cab: '', filter: 'todo', idx: 0 },
  search: { kw: '', cat: '', rows: null },
  lastLoad: 0,
};

// ---------------------------------------------------------------- 小工具
function $(s, el) { return (el || document).querySelector(s); }
function $$(s, el) { return Array.prototype.slice.call((el || document).querySelectorAll(s)); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function ic(n, cls) { return '<svg class="i ' + (cls || '') + '"><use href="#' + n + '"/></svg>'; }
function call(fn) {
  var args = Array.prototype.slice.call(arguments, 1);
  return new Promise(function (ok, bad) {
    var r = google.script.run.withSuccessHandler(ok).withFailureHandler(function (e) { bad(e && e.message ? e.message : String(e)); });
    r[fn].apply(r, args);
  });
}
function pad(n) { return (n < 10 ? '0' : '') + n; }
function keyOf(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
function parseKey(k) { var a = k.split('-'); return new Date(+a[0], +a[1] - 1, +a[2]); }
function addDays(k, n) { var d = parseKey(k); d.setDate(d.getDate() + n); return keyOf(d); }
function monday(k) { var d = parseKey(k); d.setDate(d.getDate() - (d.getDay() + 6) % 7); return keyOf(d); }
function wdOf(k) { return '日一二三四五六'.charAt(parseKey(k).getDay()); }
function roc(k) { return (+k.slice(0, 4) - 1911) + '.' + k.slice(5, 7) + '.' + k.slice(8, 10); }
function md(k) { return (+k.slice(5, 7)) + '/' + (+k.slice(8, 10)); }
function short(lab) { return String(lab || '').replace(/學?實驗室/, '').replace(/能?教室/, ''); }
function perLabel(t) { t = String(t || '').trim(); return !t ? '' : t === '午' ? '中午' : '第' + t + '節'; }
function colorOf(type) { return (S.D.colors || {})[type] || '#E8EAED'; }
function evAttr(type) { return 'class="ev-c" style="--c:' + esc(colorOf(type)) + '"'; }
function isPrep(b) { return b.type === '實驗課' || b.type === '補做實驗'; }
function daysBetween(a, b) { return Math.round((parseKey(b) - parseKey(a)) / 86400000); }
function nowHm() { var d = new Date(); return pad(d.getHours()) + ':' + pad(d.getMinutes()); }
function nowIdx() {
  if (!S.D || keyOf(new Date()) !== S.D.today.key) return -1;
  var hm = nowHm();
  for (var i = 0; i < S.D.periods.length; i++) if (hm >= S.D.periods[i].start && hm < S.D.periods[i].end) return i;
  return -1;
}
// 節次文字 → 索引（和伺服器 parsePeriods_ 一樣：範圍中間跨過中午不算中午）
function parsePer(str) {
  var shorts = S.D.periods.map(function (p) { return p.short; }), out = {};
  var idx = function (t) { t = String(t).trim(); if (!t) return -1; if (/午/.test(t)) return shorts.indexOf('午'); var d = t.replace(/[^0-9]/g, ''); return d ? shorts.indexOf(d) : -1; };
  String(str || '').split(/[,，、\s]+/).forEach(function (tok) {
    if (!tok) return;
    var m = tok.split(/[-~～–]/);
    if (m.length === 2) {
      var a = idx(m[0]), b = idx(m[1]); if (a < 0 || b < 0) return;
      for (var i = Math.min(a, b); i <= Math.max(a, b); i++) { if (shorts[i] === '午' && i !== a && i !== b) continue; out[i] = 1; }
    } else { var i2 = idx(tok); if (i2 >= 0) out[i2] = 1; }
  });
  return Object.keys(out).map(Number);
}
function todosOf(b) {
  return (S.D.todos || []).filter(function (t) {
    if (t.date !== b.date || t.lab !== b.lab) return false;
    var p = parsePer(t.period);
    return !p.length || p.some(function (i) { return b.pis.indexOf(i) >= 0; });
  });
}
function pendingTodos() {
  return (S.D.todos || []).filter(function (t) { return t.status === '待準備'; });
}
function soonTodos() {
  return pendingTodos().filter(function (t) { return !t.date || t.date <= S.D.next.key; });
}
function toast(msg, err) {
  var t = $('#toast');
  t.className = 'toast on' + (err ? ' err' : '');
  t.innerHTML = (err ? ic('alert', 's') : ic('check', 's')) + '<span>' + esc(msg) + '</span>';
  clearTimeout(toast.tm);
  toast.tm = setTimeout(function () { t.className = 'toast' + (err ? ' err' : ''); }, err ? 4200 : 2200);
}
function greet() { var h = new Date().getHours(); return h < 11 ? '早安' : h < 14 ? '午安' : h < 18 ? '午安' : '晚安'; }
function sk(h, n) { var s = ''; for (var i = 0; i < (n || 1); i++) s += '<div class="sk" style="height:' + h + 'px;margin-bottom:10px"></div>'; return s; }

// ---------------------------------------------------------------- 外觀（淺色／深色）
function getTheme() { try { return localStorage.getItem('lm-theme') || 'auto'; } catch (e) { return 'auto'; } }
function setTheme(t) { try { localStorage.setItem('lm-theme', t); } catch (e) { } applyTheme(); }
function applyTheme() {
  var t = getTheme();
  var dark = t === 'dark' || (t === 'auto' && window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches);
  document.documentElement.classList.toggle('dark', dark);
}
applyTheme();
if (window.matchMedia) try { matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme); } catch (e) { }

// ---------------------------------------------------------------- 外框：側邊選單（電腦、平板）、底部導覽（手機）
var NAV = [['today', '今天', 'home'], ['week', '課表', 'cal'], ['todo', '待辦', 'todo'], ['count', '盤點', 'count'], ['more', '更多', 'grid']];
function shell() {
  var side = '<div class="brand"><div class="lg">' + ART.koala(40) + '</div><div><b>實驗室管理</b><small>' + esc(S.D ? S.D.school.replace('臺北市立', '') : '') + '</small></div></div>' +
    '<button class="nv" data-go="today">' + ic('home') + '今天</button>' +
    '<button class="nv" data-go="week">' + ic('cal') + '實驗室課表</button>' +
    '<button class="nv" data-go="todo">' + ic('todo') + '準備事項<span class="n w" id="sTodo"></span></button>' +
    '<button class="nv" data-go="count">' + ic('count') + '盤點</button>' +
    '<div class="grp">庫存</div>' +
    '<button class="nv" data-go="restock">' + ic('alert') + '需補充<span class="n" id="sRes"></span></button>' +
    '<button class="nv" data-go="search">' + ic('box') + '品項查詢</button>' +
    '<div class="grp">其他</div>' +
    '<button class="nv" data-act="theme">' + ic('moon') + '<span id="sTheme">外觀</span></button>' +
    '<button class="nv" data-act="refresh">' + ic('refresh') + '重新整理資料</button>' +
    '<div class="foot">' + esc(S.D ? S.D.term + ' 學期' : '') + '</div>';
  $('#side').innerHTML = side;
  $('#nav').innerHTML = NAV.map(function (n) {
    return '<button data-go="' + n[0] + '">' + ic(n[2]) + n[1] + (n[0] === 'todo' ? '<span class="dotn hide" id="nTodo"></span>' : '') + '</button>';
  }).join('');
  $$('[data-go]').forEach(function (b) { b.onclick = function () { go(b.getAttribute('data-go')); }; });
  $$('[data-act="theme"]').forEach(function (b) { b.onclick = cycleTheme; });
  $$('[data-act="refresh"]').forEach(function (b) { b.onclick = function () { refresh(true); }; });
  badges();
}
function cycleTheme() {
  var t = getTheme(), n = t === 'auto' ? 'light' : t === 'light' ? 'dark' : 'auto';
  setTheme(n); badges();
  toast('外觀：' + { auto: '跟著系統', light: '淺色', dark: '深色' }[n]);
  if (S.route === 'more') render();
}
function badges() {
  if (!S.D) return;
  var n = pendingTodos().length, r = S.D.kpi.restock;
  var a = $('#sTodo'); if (a) a.textContent = n || '';
  if (a) a.classList.toggle('hide', !n);
  var b = $('#nTodo'); if (b) { b.textContent = n; b.classList.toggle('hide', !n); }
  var c = $('#sRes'); if (c) { c.textContent = r || ''; c.classList.toggle('hide', !r); }
  var t = $('#sTheme'); if (t) t.textContent = '外觀：' + { auto: '跟著系統', light: '淺色', dark: '深色' }[getTheme()];
  var nav = S.route === 'restock' || S.route === 'search' ? 'more' : S.route;
  $$('#nav button').forEach(function (x) { x.classList.toggle('on', x.getAttribute('data-go') === nav); });
  $$('#side .nv[data-go]').forEach(function (x) { x.classList.toggle('on', x.getAttribute('data-go') === S.route); });
}
function go(r) {
  S.route = r; closeSheet(); badges(); render();
  window.scrollTo(0, 0);
}
function render() {
  var m = $('#main');
  var f = { today: vToday, week: vWeek, todo: vTodo, count: vCount, more: vMore, restock: vRestock, search: vSearch }[S.route] || vToday;
  f(m);
  $('#fab').classList.toggle('hide', !(S.route === 'today' || S.route === 'todo'));
}

// ---------------------------------------------------------------- 抽屜（手機從下面滑出、電腦在中間）
function openSheet(html, onOpen) {
  closeSheet(true);
  var sc = document.createElement('div'); sc.className = 'scrim'; sc.id = 'scrim';
  var sh = document.createElement('div'); sh.className = 'sheet'; sh.id = 'sheet';
  sh.innerHTML = '<div class="grab"></div>' + html;
  document.body.appendChild(sc); document.body.appendChild(sh);
  sc.onclick = function () { closeSheet(); };
  requestAnimationFrame(function () { sc.classList.add('on'); sh.classList.add('on'); });
  $$('[data-close]', sh).forEach(function (b) { b.onclick = function () { closeSheet(); }; });
  if (onOpen) onOpen(sh);
  return sh;
}
function closeSheet(now) {
  var sc = $('#scrim'), sh = $('#sheet');
  if (!sh) return;
  sc.id = ''; sh.id = '';
  if (now) { sc.remove(); sh.remove(); return; }
  sc.classList.remove('on'); sh.classList.remove('on');
  setTimeout(function () { sc.remove(); sh.remove(); }, 260);
}

// ---------------------------------------------------------------- 資料
function load() {
  $('#main').innerHTML = '<div class="empty" style="padding-top:18vh">' + ART.duckFlask(120) + '<b>準備中…</b>第一次開啟約需 3～6 秒</div>';
  call('appInit').then(function (d) {
    S.D = d; S.lastLoad = Date.now();
    S.week.start = d.today.key; S.week.lab = S.week.lab || d.rooms[0];
    shell(); render();
  }).catch(function (e) {
    $('#main').innerHTML = '<div class="empty" style="padding-top:18vh">' + ART.koalaSleep(120) + '<b>讀取失敗</b>' + esc(e) +
      '<div style="margin-top:14px"><button class="btn p" onclick="load()" style="margin:auto">再試一次</button></div></div>';
  });
}
function refresh(show) {
  return call('appInit').then(function (d) {
    S.D = d; S.lastLoad = Date.now(); S.week.U = null; S.week.key = ''; S.restock = null;
    shell(); render(); if (show) toast('已更新');
  }).catch(function (e) { if (show) toast('更新失敗：' + e, true); });
}
function reloadTodos() {
  return call('todoData').then(function (d) { S.D.todos = d.list; badges(); render(); });
}
document.addEventListener('visibilitychange', function () {
  if (!document.hidden && S.D && Date.now() - S.lastLoad > 3 * 60 * 1000) refresh(false);
});
setInterval(function () { if (S.D && S.route === 'today' && !$('#sheet')) render(); }, 60000);
window.addEventListener('resize', function () {
  var w = window.innerWidth >= 900;
  if (w !== S.wide) { S.wide = w; if (S.D) render(); }
});
S.wide = window.innerWidth >= 900;

// ---------------------------------------------------------------- 今天
function lessonRow(b, day) {
  var idx = nowIdx(), isToday = day === S.D.today.key;
  var cur = isToday && b.pis.indexOf(idx) >= 0;
  var past = isToday && idx >= 0 && b.pis.length && Math.max.apply(null, b.pis) < idx;
  var n = todosOf(b).filter(function (t) { return t.status === '待準備'; }).length;
  return '<button class="it' + (cur ? ' cur' : '') + (past ? ' past' : '') + '" data-b="' + esc(JSON.stringify([day, S._bi.push(b) - 1])) + '">' +
    '<div class="p">' + esc(b.periodText === '午' ? '午' : b.periodText) + '</div><div style="min-width:0"><div class="t">' + esc(b.content || b.type) + '</div>' +
    '<div class="s"><span class="tag" ' + evAttr(b.type) + '>' + esc(short(b.lab)) + '</span>' + esc([b.cls, b.teacher].filter(String).join('　')) + '</div></div>' +
    (n ? '<span class="badge">📝 ' + n + '</span>' : '') + '</button>';
}
function todoRow(t, opt) {
  opt = opt || {};
  var d = t.date, today = S.D.today.key;
  var when = d ? (d === today ? '今天' : d === addDays(today, 1) ? '明天' : md(d) + '（' + wdOf(d) + '）') : '';
  var late = d && d < today && t.status === '待準備';
  if (opt.noDate) when = '';
  var where = [when ? (late ? '<span class="late">' + esc(when) + ' 已過期</span>' : esc(when)) : '', esc(short(t.lab)), esc(perLabel(t.period)), esc(t.cls), t.teacher ? esc(t.teacher + '老師') : '']
    .filter(String).join('　');
  return '<div class="it' + (t.status !== '待準備' ? ' done' : '') + '"><button class="ck' + (t.status !== '待準備' ? ' on' : '') + '" data-ck="' + esc(t.id) + '">' +
    (t.status !== '待準備' ? ic('check', 's') : '') + '</button><div style="min-width:0;flex:1"><div class="t">' + esc(t.what) + (t.qty ? ' × ' + esc(t.qty) + ' ' + esc(t.unit) : '') + '</div>' +
    '<div class="s">' + where + (t.place || t.have !== '' ? '<br>' + ic('pin', 's') + ' ' + esc(t.place || '—') + (t.have !== '' && t.have != null ? '　庫存 ' + esc(t.have) + (t.sunit ? ' ' + esc(t.sunit) : '') : '') : '') +
    (t.judge && t.judge.text ? '<br><span class="jd ' + t.judge.state + '">' + (t.judge.state === 'bad' ? '✘ ' : t.judge.state === 'warn' ? '⚠ ' : '✔ ') + esc(t.judge.text) + '</span>' : '') +
    (t.note ? '<br>📝 ' + esc(t.note) : '') +
    (t.buy ? '<br><span class="buyb' + (t.buy === '已到貨' ? ' ok' : '') + '">🛒 請購：' + esc(t.buy) + '</span>' : '') + '</div></div>' +
    (opt.actions === false ? '' : t.status === '待準備' ? '<div style="display:flex;flex-direction:column;gap:4px;align-items:flex-end">' +
      (t.buy ? '' : '<button class="x" data-buy="' + esc(t.id) + '">🛒 要先買</button>') + '<button class="x" data-cancel="' + esc(t.id) + '">取消</button></div>' :
      t.status === '已準備' ? '<button class="x" data-back="' + esc(t.id) + '">已歸還</button>' : '<span class="x">' + esc(t.status) + '</span>') + '</div>';
}
function bindTodoRows(el) {
  $$('[data-ck]', el).forEach(function (b) {
    b.onclick = function () {
      var id = b.getAttribute('data-ck'), t = S.D.todos.filter(function (x) { return x.id === id; })[0];
      if (!t) return;
      var st = t.status === '待準備' ? '已準備' : '待準備';
      setTodo(t, st, st === '已準備' ? '✔ 已準備：' + t.what : '改回待準備');
    };
  });
  $$('[data-cancel]', el).forEach(function (b) {
    b.onclick = function () {
      if (!b.classList.contains('cf')) { b.classList.add('cf'); b.textContent = '確定取消？'; setTimeout(function () { b.classList.remove('cf'); b.textContent = '取消'; }, 3000); return; }
      var t = S.D.todos.filter(function (x) { return x.id === b.getAttribute('data-cancel'); })[0];
      if (t) setTodo(t, '取消', '已取消：' + t.what);
    };
  });
  $$('[data-buy]', el).forEach(function (b) {
    b.onclick = function () {
      b.disabled = true; b.textContent = '加入中…';
      call('todoToPurchase', b.getAttribute('data-buy')).then(function (m) { toast(m); return reloadTodos(); })
        .catch(function (e) { toast(e, true); b.disabled = false; b.textContent = '🛒 要先買'; });
    };
  });
  $$('[data-back]', el).forEach(function (b) {
    b.onclick = function () { var t = S.D.todos.filter(function (x) { return x.id === b.getAttribute('data-back'); })[0]; if (t) setTodo(t, '已歸還', '已歸還：' + t.what); };
  });
}
function setTodo(t, st, msg) {
  var old = t.status;
  t.status = st;
  if (st === '取消') S.D.todos = S.D.todos.filter(function (x) { return x !== t; });
  badges(); rerenderKeepSheet(); toast(msg);
  call('todoSet', [t.id], st).catch(function (e) { t.status = old; if (st === '取消') S.D.todos.push(t); badges(); rerenderKeepSheet(); toast('沒有存到：' + e, true); });
}
function rerenderKeepSheet() {
  render();
  var sh = $('#sheet');
  if (sh && sh._lesson) fillLesson(sh, sh._lesson);
}
function nowCard() {
  var idx = nowIdx(), p = S.D.periods[idx];
  if (idx < 0) {
    var hm = nowHm(), nextP = null;
    if (keyOf(new Date()) === S.D.today.key) S.D.periods.forEach(function (x, i) { if (!nextP && x.start > hm) nextP = x; });
    var nl = nextP ? S.D.today.list.filter(function (b) { return b.pis.indexOf(S.D.periods.indexOf(nextP)) >= 0; }) : [];
    return '<div class="now off"><div style="flex:1"><div class="h">' + ic('clock', 's') + (nextP ? '下課時間　下一節：' + esc(nextP.name) + ' ' + esc(nextP.start) : '現在沒有上課') + '</div>' +
      '<div class="big">' + (nextP ? (nl.length ? nl.length + ' 間實驗室下一節有課' : '下一節實驗室都空著') : S.D.today.list.length ? '今天的課都上完了' : '今天沒有登記的課') + '</div>' +
      (nl.length ? '<div class="labs">' + nl.slice(0, 4).map(labBtn).join('') + '</div>' : '') + '</div>' + (nl.length ? '' : ART.koalaSleep(70)) + '</div>';
  }
  var bs = S.D.today.list.filter(function (b) { return b.pis.indexOf(idx) >= 0; });
  return '<div class="now"><div style="flex:1;min-width:0"><div class="h"><span class="dot"></span>現在　' + esc(p.name) + '　' + esc(p.start + '–' + p.end) + '</div>' +
    '<div class="big">' + (bs.length ? bs.length + ' 間實驗室上課中' : '這一節實驗室都空著') + '</div>' +
    (bs.length ? '<div class="labs">' + bs.slice(0, 6).map(labBtn).join('') + '</div>' : '') + '</div>' + (bs.length ? '' : ART.koala(64)) + '</div>';
}
function labBtn(b) {
  return '<button class="lab" data-b="' + esc(JSON.stringify([S.D.today.key, S._bi.push(b) - 1])) + '"><small>' + esc(b.lab) + '</small><b>' + esc(b.content || b.type) + '</b><small>' +
    esc([b.cls, b.teacher ? b.teacher + '老師' : ''].filter(String).join('　')) + '</small></button>';
}
function bindLessons(el) {
  $$('[data-b]', el).forEach(function (x) {
    x.onclick = function () { var a = JSON.parse(x.getAttribute('data-b')); openLesson(S._bi[a[1]]); };
  });
}
function vToday(m) {
  S._bi = [];
  if (S.wide) return vDash(m);
  var D = S.D, soon = soonTodos();
  var cnt = D.today.list.filter(function (b) { return b.type !== '放假'; }).length;
  var todayDue = soon.filter(function (t) { return t.date && t.date <= D.today.key; }).length;
  var day = S.todayTab === 'today' ? D.today : D.next;
  var h = '<div class="top"><div><h1>今天</h1><div class="sub">' + esc(D.today.roc + '（' + D.today.wd + '）') + '</div></div><div class="sp"></div>' +
    '<button class="icbtn" data-go2="search" aria-label="查詢">' + ic('search') + '</button><button class="icbtn" data-act2="refresh" aria-label="重新整理">' + ic('refresh') + '</button></div>' +
    '<div class="chips"><div class="chip">' + ic('flask', 's') + '今天 <b>' + cnt + '</b> 堂</div>' +
    '<button class="chip' + (soon.length ? ' warn' : '') + '" data-go2="todo">' + ic('todo', 's') + '待準備 <b>' + soon.length + '</b></button>' +
    (D.kpi.restock !== null ? '<button class="chip' + (D.kpi.restock ? ' bad' : '') + '" data-go2="restock">' + ic('alert', 's') + '需補充 <b>' + D.kpi.restock + '</b></button>' : '') +
    (D.kpi.loans.n ? '<div class="chip">' + ic('swap', 's') + '借出中 <b>' + D.kpi.loans.n + '</b></div>' : '') +
    (D.kpi.count.total ? '<button class="chip" data-go2="count">' + ic('count', 's') + '盤點 <b>' + D.kpi.count.done + '/' + D.kpi.count.total + '</b></button>' : '') + '</div>' +
    '<div class="pad">' + nowCard() +
    '<div class="h2" style="margin-top:18px"><div class="seg" style="flex:1"><button data-tt="today" class="' + (S.todayTab === 'today' ? 'on' : '') + '">今天 ' + md(D.today.key) + '（' + D.today.wd + '）</button>' +
    '<button data-tt="next" class="' + (S.todayTab === 'next' ? 'on' : '') + '">下一個上課日 ' + md(D.next.key) + '（' + D.next.wd + '）</button></div></div>' +
    (day.list.length ? '<div class="tl">' + day.list.map(function (b) { return lessonRow(b, day.key); }).join('') + '</div>' :
      '<div class="tl"><div class="empty">' + ART.koalaSleep(90) + '<b>沒有登記的課</b>實驗室都空著</div></div>') +
    '<div class="h2">待準備' + (todayDue ? '　<span class="late">' + todayDue + ' 項今天要用或已過期</span>' : '') + '<button class="more" data-go2="todo">全部 ' + pendingTodos().length + ' 項</button></div>' +
    (soon.length ? '<div class="todo">' + soon.slice(0, 8).map(function (t) { return todoRow(t, { actions: false }); }).join('') + '</div>' :
      '<div class="todo"><div class="empty">' + ART.duck(96) + '<b>全部準備好了！</b>到下一個上課日都沒有待準備的東西</div></div>') + '</div>';
  m.innerHTML = h;
  bindCommon(m); bindLessons(m); bindTodoRows(m);
  $$('[data-tt]', m).forEach(function (b) { b.onclick = function () { S.todayTab = b.getAttribute('data-tt'); render(); }; });
}
function bindCommon(m) {
  $$('[data-go2]', m).forEach(function (b) { b.onclick = function () { go(b.getAttribute('data-go2')); }; });
  $$('[data-act2="refresh"]', m).forEach(function (b) { b.onclick = function () { refresh(true); }; });
}
function vDash(m) {
  var D = S.D, idx = nowIdx(), soon = soonTodos(), p = D.periods[idx];
  var cnt = D.today.list.filter(function (b) { return b.type !== '放假'; });
  var per = {}; cnt.forEach(function (b) { per[short(b.lab)] = (per[short(b.lab)] || 0) + 1; });
  var due = soon.filter(function (t) { return t.date && t.date <= D.today.key; }).length;
  var h = '<div class="hello">' + ART.koala(64) + '<div style="flex:1"><h1>' + greet() + '！今天是 ' + md(D.today.key) + '（' + D.today.wd + '）</h1>' +
    '<div class="sub">' + esc(D.term) + ' 學期　｜　' + (p ? '現在' + esc(p.name) + '（' + esc(p.start + '–' + p.end) + '）' : '現在是下課時間') + '</div></div>' +
    '<button class="icbtn" data-act2="refresh" title="重新整理">' + ic('refresh') + '</button></div>' +
    '<div class="kpis">' +
    '<button class="card kpi" data-go2="week"><small>' + ic('flask', 's') + '今天的課</small><b>' + cnt.length + '</b><span>' + (Object.keys(per).map(function (k) { return k + ' ' + per[k]; }).join('・') || '沒有登記') + '</span></button>' +
    '<button class="card kpi" data-go2="todo"><small>' + ic('todo', 's') + '待準備</small><b style="color:' + (soon.length ? 'var(--warn)' : 'var(--ok)') + '">' + pendingTodos().length + '</b><span>' + (due ? due + ' 項今天要用或已過期' : '到下一個上課日 ' + soon.length + ' 項') + '</span></button>' +
    '<button class="card kpi" data-go2="restock"><small>' + ic('alert', 's') + '需補充</small><b style="color:' + (D.kpi.restock ? 'var(--bad)' : 'var(--ok)') + '">' + (D.kpi.restock === null ? '—' : D.kpi.restock) + '</b><span>低於安全存量</span></button>' +
    '<div class="card kpi"><small>' + ic('cart', 's') + '請購中</small><b>' + D.kpi.purchase.n + '</b><span>待處理＋已請購</span></div>' +
    '<div class="card kpi"><small>' + ic('swap', 's') + '借出中</small><b>' + D.kpi.loans.n + '</b><span>' + (D.kpi.loans.late ? '<span class="late">逾期 ' + D.kpi.loans.late + ' 筆</span>' : '沒有逾期') + '</span></div></div>' +
    '<div class="dash"><div class="card"><div class="ch"><h3>各實驗室・現在</h3>' + (p ? '<span class="st b">' + esc(p.name) + '</span>' : '<span class="st f">下課時間</span>') + '<button class="more" data-go2="week">看課表 ›</button></div>' +
    '<div class="labs6">' + D.rooms.map(function (r) {
      var b = idx >= 0 ? D.today.list.filter(function (x) { return x.lab === r && x.pis.indexOf(idx) >= 0; })[0] : null;
      if (b) return '<button class="lt busy ev-c" style="--c:' + esc(colorOf(b.type)) + '" data-b="' + esc(JSON.stringify([D.today.key, S._bi.push(b) - 1])) + '"><div class="n">' + esc(r) + '<span class="st b">上課中</span></div><b>' + esc(b.content || b.type) + '</b><span>' + esc([b.cls, b.teacher ? b.teacher + '老師' : ''].filter(String).join('　')) + '</span></button>';
      var nx = D.today.list.filter(function (x) { return x.lab === r && x.pis.length && Math.min.apply(null, x.pis) > idx && x.type !== '放假'; })[0];
      return '<div class="lt free"><div class="n">' + esc(r) + '<span class="st f">空堂</span></div><b>—</b><span>' + (nx ? '下一堂：' + esc(perLabel(nx.periodText)) + ' ' + esc(nx.cls) : '今天沒有其他課') + '</span></div>';
    }).join('') + '</div>' +
    '<div class="ch" style="padding-top:0"><div class="seg"><button data-ht="today" class="' + (S.heatTab === 'today' ? 'on' : '') + '">今天各節</button><button data-ht="next" class="' + (S.heatTab === 'next' ? 'on' : '') + '">' + md(D.next.key) + '（' + D.next.wd + '）各節</button></div></div>' +
    heat(S.heatTab === 'today' ? D.today : D.next) + '</div>' +
    '<div style="display:flex;flex-direction:column;gap:16px"><div class="card"><div class="ch"><h3>待準備</h3><button class="more" data-go2="todo">全部 ›</button></div>' +
    (soon.length ? '<div class="todo" style="box-shadow:none;border-radius:0 0 16px 16px">' + soon.slice(0, 7).map(function (t) { return todoRow(t, { actions: false }); }).join('') + '</div>' :
      '<div class="empty">' + ART.duck(80) + '<b>全部準備好了！</b></div>') + '</div>' +
    '<div class="card" id="dRes"><div class="ch"><h3>需補充</h3><button class="more" data-go2="restock">全部 ›</button></div><div style="padding:0 16px 12px">' + sk(18, 3) + '</div></div>' +
    (D.kpi.count.total ? '<button class="card" data-go2="count" style="text-align:left"><div class="ch"><h3>盤點進度</h3><span class="more">繼續盤點 ›</span></div><div class="prog" style="padding-top:0"><div class="l"><span>' + esc(D.kpi.count.name || '進行中的盤點') + '</span><span>' + D.kpi.count.done + ' / ' + D.kpi.count.total + '</span></div><div class="bar"><i style="width:' + Math.round(D.kpi.count.done / D.kpi.count.total * 100) + '%"></i></div></div></button>' : '') +
    '</div></div>';
  m.innerHTML = h;
  bindCommon(m); bindLessons(m); bindTodoRows(m);
  $$('[data-ht]', m).forEach(function (b) { b.onclick = function () { S.heatTab = b.getAttribute('data-ht'); render(); }; });
  loadRestock().then(function (rows) {
    var el = $('#dRes'); if (!el) return;
    el.innerHTML = '<div class="ch"><h3>需補充</h3><button class="more" data-go2="restock">全部 ' + rows.length + ' ›</button></div>' +
      (rows.length ? '<div class="lst">' + rows.slice(0, 6).map(function (r) {
        return '<div class="it"><span class="dt" style="background:var(--bad)"></span>' + esc(r.name) + ' <small class="muted">' + esc(r.spec) + '</small><span class="r">剩 ' + esc(r.qty) + ' ' + esc(r.unit) + '，安全 ' + esc(r.safe) + '</span></div>';
      }).join('') + '</div>' : '<div class="empty" style="padding:14px">' + '<b>沒有需要補充的</b></div>');
    bindCommon(el);
  });
}
function heat(day) {
  var idx = day.key === S.D.today.key ? nowIdx() : -1;
  var h = '<div class="heat"><table><tr class="hd"><td class="h"></td>' + S.D.periods.map(function (p, i) { return '<td' + (i === idx ? ' style="color:var(--pri);font-weight:800"' : '') + '>' + esc(p.short) + '</td>'; }).join('') + '</tr>';
  S.D.rooms.forEach(function (r) {
    h += '<tr><td class="h">' + esc(short(r)) + '</td>';
    S.D.periods.forEach(function (p, i) {
      var b = day.list.filter(function (x) { return x.lab === r && x.pis.indexOf(i) >= 0; })[0];
      h += b ? '<td class="x ev-c' + (i === idx ? ' nw' : '') + '" style="--c:' + esc(colorOf(b.type)) + '" title="' + esc((b.content || b.type) + ' ' + b.cls + ' ' + b.teacher) + '" data-b="' + esc(JSON.stringify([day.key, S._bi.push(b) - 1])) + '">' +
        esc(b.cls || b.content || b.type) + '</td>' : '<td class="' + (i === idx ? 'nw' : '') + '"></td>';
    });
    h += '</tr>';
  });
  return h + '</table></div>';
}
function loadRestock() {
  if (S.restock) return Promise.resolve(S.restock);
  return call('mobileRestock').then(function (r) { S.restock = r.ok ? r.rows : []; return S.restock; }).catch(function () { return []; });
}

// ---------------------------------------------------------------- 一堂課（詳細）
function openLesson(b) {
  if (S.route === 'week' && S.wide && window.innerWidth >= 1100) { S.week.sel = b; render(); return; }
  var sh = openSheet('<div class="lsn det" style="padding:0"></div>');
  sh._lesson = b; fillLesson(sh, b);
}
function lessonHtml(b) {
  var p0 = S.D.periods[b.pis[0]], p1 = S.D.periods[b.pis[b.pis.length - 1]];
  var time = p0 && p1 ? p0.start + '–' + p1.end : '';
  var ts = todosOf(b);
  var h = '<div style="display:flex;align-items:center"><span class="tag" ' + evAttr(b.type) + '>' + esc(b.type) + '</span><div class="sp"></div>' +
    (S.wide && S.route === 'week' && window.innerWidth >= 1100 ? '<button class="x muted" data-unsel>' + ic('x') + '</button>' : '<button class="x muted" data-close>' + ic('x') + '</button>') + '</div>' +
    '<h3>' + esc(b.content || (isPrep(b) ? '實驗課（名稱未填）' : b.type)) + '</h3>' +
    '<div class="m">' + ic('clock', 's') + esc(md(b.date) + '（' + wdOf(b.date) + '）' + perLabel(b.periodText) + '　' + time) + '</div>' +
    '<div class="m">' + ic('pin', 's') + esc([b.lab, b.cls, b.teacher ? b.teacher + '老師' : '', b.groups ? b.groups + ' 組' : ''].filter(String).join('　｜　')) + '</div>' +
    (b.note ? '<div class="m">📝 ' + esc(b.note) + '</div>' : '');
  if (isPrep(b)) {
    h += '<div class="blk"><b>實驗名稱</b><div class="nmrow"><input class="inp" id="lName" list="kitList" placeholder="例：酸鹼滴定" value="' + esc(b.content) + '">' +
      '<input class="inp" id="lGroups" inputmode="numeric" placeholder="組數" value="' + esc(b.groups || '') + '"></div>' +
      '<label style="display:flex;gap:8px;align-items:center;font-size:13px;color:var(--sub);margin-top:8px"><input type="checkbox" id="lSame" checked> 同一週、同一間、' + esc(b.teacher || '同一位') + '老師還沒填的也一起填</label>' +
      '<button class="btn p s" id="lSave" style="margin-top:10px;width:100%">' + ic('check', 's') + '儲存實驗名稱</button><datalist id="kitList"></datalist></div>' +
      '<div class="blk"><b>器材' + (b.groups ? '（' + esc(b.groups) + ' 組）' : '') + '</b><div id="lKit">' + (b.content ? sk(16, 3) : '<span class="muted">填好實驗名稱後，會用「實驗套組」檢查器材夠不夠。</span>') + '</div></div>';
  }
  h += '<div class="blk"><b>準備事項</b>' + (ts.length ? ts.map(function (t) {
    return '<div class="mini' + (t.status !== '待準備' ? ' done' : '') + '"><button class="ck' + (t.status !== '待準備' ? ' on' : '') + '" data-ck="' + esc(t.id) + '">' + (t.status !== '待準備' ? ic('check', 's') : '') + '</button><span>' +
      esc(t.what + (t.qty ? ' × ' + t.qty + ' ' + t.unit : '')) + '</span>' + (t.buy ? '<span class="buyb">🛒 ' + esc(t.buy) + '</span>' : '') + '</div>';
  }).join('') : '<span class="muted">還沒有</span>') + '</div>' +
    '<div class="btns2"><button class="btn g s" id="lAdd">' + ic('plus', 's') + '加準備事項</button>' +
    (S.D.url && b.content && isPrep(b) ? '<button class="btn p s" id="lPrep">' + ic('print', 's') + '器材準備單</button>' : '<span></span>') + '</div>';
  return h;
}
function fillLesson(el, b) {
  var box = $('.lsn', el) || el;
  box.innerHTML = lessonHtml(b);
  $$('[data-close]', box).forEach(function (x) { x.onclick = function () { closeSheet(); }; });
  $$('[data-unsel]', box).forEach(function (x) { x.onclick = function () { S.week.sel = null; render(); }; });
  bindTodoRows(box);
  $('#lAdd', box).onclick = function () { openAdd({ date: b.date, lab: b.lab, period: b.periodText, cls: b.cls, teacher: b.teacher }); };
  var pr = $('#lPrep', box);
  if (pr) pr.onclick = function () { window.open(S.D.url + '?page=prep&from=' + b.date + '&to=' + b.date, '_blank'); };
  if (!isPrep(b)) return;
  var fillKits = function () { var dl = $('#kitList', box); if (dl) dl.innerHTML = S.kits.map(function (k) { return '<option value="' + esc(k) + '">'; }).join(''); };
  if (S.kits) fillKits(); else call('lessonKits').then(function (k) { S.kits = k; fillKits(); });
  if (b.content) call('appLessonKit', b.content, b.groups || 0).then(function (k) {
    var el2 = $('#lKit', box); if (!el2) return;
    if (!k) { el2.innerHTML = '<span class="muted">「實驗套組」裡沒有「' + esc(b.content) + '」，建好之後這裡會自動檢查器材。</span>'; return; }
    el2.innerHTML = k.rows.map(function (r) {
      var hv = esc(r.have) + (r.unit ? ' ' + esc(r.unit) : '');
      var st = r.state === 'bad' ? '<span class="ng">' + hv + ' ✘</span>' : r.state === 'warn' ? '<span class="soon" style="white-space:nowrap">' + hv + ' ⚠</span>' :
        '<span class="ok">' + hv + (r.bulk ? '（大包裝）' : '') + ' ✔</span>';
      return '<div class="kit"><span>' + esc(r.name) + (r.need !== '' ? ' × ' + esc(r.need) : '') + (r.note ? '<br><small class="muted">' + esc(r.note) + '</small>' : '') + '</span>' + st + '</div>';
    }).join('') || '<span class="muted">套組沒有器材</span>';
  }).catch(function () { });
  $('#lSave', box).onclick = function () {
    var name = $('#lName', box).value.trim(), g = $('#lGroups', box).value.trim(), btn = this;
    if (!name) { toast('請填實驗名稱', true); return; }
    btn.disabled = true; btn.textContent = '儲存中…';
    call('applyLessonName', { row: b.row, key: [b.date, b.lab, String(b.periodText).trim(), b.cls].join('|'), name: name, groups: g, same: $('#lSame', box).checked })
      .then(function (msg) {
        b.content = name; if (g) b.groups = g;
        toast(msg.split('\n')[0]); S.week.key = '';
        refresh(false).then(function () { if (el._lesson) fillLesson(el, b); });
      }).catch(function (e) { btn.disabled = false; btn.textContent = '儲存實驗名稱'; toast(e, true); });
  };
}

// ---------------------------------------------------------------- 課表
function weekRange() {
  var w = S.week;
  if (w.mode === 'day') return [w.start, 1];
  if (w.mode === 'week') return [monday(w.start), 7];
  if (w.mode === 'month') { var d = parseKey(w.start); return [keyOf(new Date(d.getFullYear(), d.getMonth(), 1)), new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate()]; }
  return [w.start, 30];
}
function vWeek(m) {
  var w = S.week, r = weekRange(), key = r.join('/');
  if (w.mode === 'week' && !w.lab) w.lab = S.D.rooms[0];
  var title = w.mode === 'day' ? md(r[0]) + '（' + wdOf(r[0]) + '）' : w.mode === 'month' ? (+r[0].slice(0, 4) - 1911) + ' 年 ' + (+r[0].slice(5, 7)) + ' 月' :
    md(r[0]) + '（' + wdOf(r[0]) + '）～' + md(addDays(r[0], w.mode === 'week' ? 4 : r[1] - 1));
  var labs = (w.mode === 'week' ? [] : [['', '全部']]).concat(S.D.rooms.map(function (x) { return [x, S.wide ? x : short(x)]; }));
  var h = '<div class="top"><div style="min-width:0"><h1>實驗室課表</h1><div class="sub">' + esc(title) + '</div></div><div class="sp"></div>' +
    '<input class="inp" id="wKw" placeholder="找班級、老師、實驗…" value="' + esc(w.kw) + '" style="max-width:' + (S.wide ? '280px' : '150px') + ';padding:9px 11px"></div>' +
    '<div class="pad"><div class="tbar"><button class="pill" data-mv="-1">' + ic('left', 's') + '</button><button class="pill" data-mv="0">' + (w.mode === 'month' ? '這個月' : w.mode === 'week' ? '本週' : '今天') + '</button>' +
    '<button class="pill" data-mv="1">' + ic('right', 's') + '</button><input type="date" class="pill" id="wDate" value="' + esc(w.start) + '" style="padding:5px 8px">' +
    '<div class="seg">' + [['day', '單日'], ['week', '整週'], ['month', '月曆'], ['list', '清單']].map(function (x) { return '<button data-md="' + x[0] + '" class="' + (w.mode === x[0] ? 'on' : '') + '">' + x[1] + '</button>'; }).join('') + '</div></div>' +
    '<div class="labsel" style="margin-bottom:12px">' + labs.map(function (x) { return '<button class="pill' + (w.lab === x[0] || (!w.lab && !x[0]) ? ' on' : '') + '" data-lab="' + esc(x[0]) + '">' + esc(x[1]) + '</button>'; }).join('') + '</div>' +
    '<div class="split' + (w.sel && S.wide && window.innerWidth >= 1100 ? ' has' : '') + '"><div id="wBody">' + (w.U && w.key === key ? '' : sk(60, 6)) + '</div>' +
    (w.sel && S.wide && window.innerWidth >= 1100 ? '<div class="card det" style="position:sticky;top:16px"><div class="lsn"></div></div>' : '') + '</div></div>';
  m.innerHTML = h;
  $$('[data-mv]', m).forEach(function (b) {
    b.onclick = function () {
      var n = +b.getAttribute('data-mv');
      if (!n) w.start = S.D.today.key;
      else if (w.mode === 'month') { var d = parseKey(w.start); w.start = keyOf(new Date(d.getFullYear(), d.getMonth() + n, 1)); }
      else w.start = addDays(w.start, n * (w.mode === 'week' ? 7 : w.mode === 'list' ? 30 : 1));
      w.sel = null; render();
    };
  });
  $('#wDate', m).onchange = function () { if (this.value) { w.start = this.value; w.sel = null; render(); } };
  $$('[data-md]', m).forEach(function (b) { b.onclick = function () { w.mode = b.getAttribute('data-md'); if (w.mode === 'week' && !w.lab) w.lab = S.D.rooms[0]; w.sel = null; render(); }; });
  $$('[data-lab]', m).forEach(function (b) { b.onclick = function () { w.lab = b.getAttribute('data-lab'); w.sel = null; render(); }; });
  var tm;
  $('#wKw', m).oninput = function () { var v = this.value; clearTimeout(tm); tm = setTimeout(function () { w.kw = v.trim(); drawWeek(); }, 250); };
  if (w.sel && $('.det', m)) fillLesson($('.det', m), w.sel);
  if (w.U && w.key === key) drawWeek();
  else call('usageData', r[0], r[1]).then(function (U) {
    w.U = U; w.key = key;
    // 待辦用最新的（一覽裡的是那時候的）
    if (S.route === 'week') drawWeek();
  }).catch(function (e) { var b = $('#wBody'); if (b) b.innerHTML = '<div class="empty">讀取失敗：' + esc(e) + '</div>'; });
}
function drawWeek() {
  var w = S.week, U = w.U, el = $('#wBody');
  if (!el || !U) return;
  S._bi = [];
  var kw = w.kw.toLowerCase();
  var list = U.bookings.filter(function (b) {
    if (w.lab && b.lab !== w.lab) return false;
    return !kw || [b.content, b.type, b.cls, b.teacher, b.lab, b.note].join(' ').toLowerCase().indexOf(kw) >= 0;
  });
  var idx = nowIdx(), today = S.D.today.key;
  var ev = function (b, withLab) {
    var n = todosOf(b).filter(function (t) { return t.status === '待準備'; }).length;
    return '<button class="ev ev-c' + (w.sel && w.sel.row === b.row && w.sel.date === b.date ? ' sel' : '') + '" style="--c:' + esc(colorOf(b.type)) + '" data-b="' + esc(JSON.stringify([b.date, S._bi.push(b) - 1])) + '">' +
      '<b>' + esc(b.content || b.type) + '</b>' + esc([withLab ? short(b.lab) : '', b.cls, b.teacher].filter(String).join(' ')) + (n ? '<span class="td">📝' + n + '</span>' : '') + '</button>';
  };
  var h = '';
  if (w.mode === 'day') {
    var rooms = w.lab ? [w.lab] : S.D.rooms, d0 = U.days[0].key;
    h = '<div class="card wkwrap"><table class="wk" style="min-width:' + (rooms.length > 3 ? 760 : 360) + 'px"><tr><th style="width:62px"></th>' + rooms.map(function (r) { return '<th>' + esc(S.wide ? r : short(r)) + '</th>'; }).join('') + '</tr>';
    S.D.periods.forEach(function (p, pi) {
      h += '<tr class="' + (d0 === today && pi === idx ? 'nowr' : '') + '"><td class="pr"><b>' + esc(p.short) + '</b>' + esc(p.start) + '</td>' + rooms.map(function (r) {
        return '<td>' + list.filter(function (b) { return b.lab === r && b.date === d0 && b.pis.indexOf(pi) >= 0; }).map(function (b) { return ev(b); }).join('') + '</td>';
      }).join('') + '</tr>';
    });
    h += '</table></div>';
  } else if (w.mode === 'week') {
    var days = U.days.filter(function (d, i) { return i < 5 || list.some(function (b) { return b.date === d.key; }); });
    h = '<div class="card wkwrap"><table class="wk"><tr><th style="width:56px"></th>' + days.map(function (d) {
      return d.key === today ? '<th class="today">' + d.wd + '<div>' + md(d.key) + '</div></th>' : '<th>' + d.wd + ' ' + md(d.key) + '</th>';
    }).join('') + '</tr>';
    S.D.periods.forEach(function (p, pi) {
      h += '<tr class="' + (pi === idx && days.some(function (d) { return d.key === today; }) ? 'nowr' : '') + '"><td class="pr"><b>' + esc(p.short) + '</b>' + esc(p.start) + '</td>' + days.map(function (d) {
        return '<td>' + list.filter(function (b) { return b.date === d.key && b.pis.indexOf(pi) >= 0; }).map(function (b) { return ev(b); }).join('') + '</td>';
      }).join('') + '</tr>';
    });
    h += '</table></div>';
  } else if (w.mode === 'month') {
    var first = parseKey(U.days[0].key), last = U.days[U.days.length - 1].key;
    var weekend = list.some(function (b) { var x = parseKey(b.date).getDay(); return x === 0 || x === 6; }), cols = weekend ? 7 : 5;
    h = '<div class="wkwrap"><table class="mon"><tr>' + '一二三四五六日'.slice(0, cols).split('').map(function (x) { return '<th>' + x + '</th>'; }).join('') + '</tr>';
    for (var k = monday(U.days[0].key); k <= last;) {
      h += '<tr>';
      for (var i = 0; i < 7; i++, k = addDays(k, 1)) {
        if (i >= cols) continue;
        var inM = parseKey(k).getMonth() === first.getMonth();
        var seen = {}, es = list.filter(function (b) { return b.date === k; }).sort(function (a, b) { return (a.pis[0] || 0) - (b.pis[0] || 0); }).filter(function (b) {
          if (b.type !== '放假' && b.type !== '考試') return true;
          var g = b.type + b.content + b.periodText; if (seen[g]) return false; seen[g] = 1; return true;
        });
        h += '<td class="' + (inM ? '' : 'out') + (k === today ? ' today' : '') + '"><div class="dn">' + (+k.slice(8)) + '</div>' + (inM ? es.slice(0, 5).map(function (b) {
          return '<button class="e ev-c" style="--c:' + esc(colorOf(b.type)) + '" data-b="' + esc(JSON.stringify([b.date, S._bi.push(b) - 1])) + '">' +
            esc((b.periodText === '午' ? '午' : b.periodText) + ' ' + (b.type === '放假' || b.type === '考試' ? b.content : [w.lab ? '' : short(b.lab), b.cls, b.content].filter(String).join(' '))) + '</button>';
        }).join('') + (es.length > 5 ? '<button class="e muted" data-day="' + k + '">…還有 ' + (es.length - 5) + ' 筆</button>' : '') : '') + '</td>';
      }
      h += '</tr>';
    }
    h += '</table></div>';
  } else {
    var ls = list.slice().sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.pis[0] || 0) - (b.pis[0] || 0); });
    h = ls.length ? '<div class="card" style="overflow:hidden">' + ls.map(function (b) {
      return '<button class="lrow" data-b="' + esc(JSON.stringify([b.date, S._bi.push(b) - 1])) + '"><div class="d">' + md(b.date) + '（' + wdOf(b.date) + '）<br>' + esc(perLabel(b.periodText)) + '</div>' +
        '<span class="tag ev-c" style="--c:' + esc(colorOf(b.type)) + '">' + esc(short(b.lab)) + '</span><div style="min-width:0"><div class="t">' + esc(b.content || b.type) + '</div><div class="s">' +
        esc([b.cls, b.teacher].filter(String).join('　')) + '</div></div></button>';
    }).join('') + '</div>' : '';
  }
  if (!list.length && (w.mode === 'list' || w.kw)) h += '<div class="card"><div class="empty">' + ART.koalaSleep(90) + '<b>沒有符合的登記</b>換個關鍵字或日期看看</div></div>';
  el.innerHTML = h;
  bindLessons(el);
  $$('[data-day]', el).forEach(function (x) { x.onclick = function () { w.mode = 'day'; w.start = x.getAttribute('data-day'); render(); }; });
}

// ---------------------------------------------------------------- 待辦
function vTodo(m) {
  var all = S.D.todos || [], today = S.D.today.key;
  var pend = all.filter(function (t) { return t.status === '待準備'; });
  var done = all.filter(function (t) { return t.status !== '待準備'; }).reverse();
  var list = S.todoView === 'todo' ? pend : done;
  var h = '<div class="top"><div><h1>準備事項</h1><div class="sub">老師交代要準備的東西</div></div><div class="sp"></div>' +
    '<button class="icbtn" data-act2="refresh">' + ic('refresh') + '</button></div><div class="pad" style="max-width:820px">' +
    '<div class="seg" style="margin-bottom:6px"><button data-tv="todo" class="' + (S.todoView === 'todo' ? 'on' : '') + '">待準備 ' + pend.length + '</button><button data-tv="done" class="' + (S.todoView === 'done' ? 'on' : '') + '">最近完成 ' + done.length + '</button></div>';
  if (!list.length) h += '<div class="card" style="margin-top:12px"><div class="empty">' + ART.duck(110) + '<b>' + (S.todoView === 'todo' ? '全部準備好了！' : '最近 7 天沒有完成的事項') + '</b>' +
    (S.todoView === 'todo' ? '老師交代新的東西，按右下角 ＋ 記下來' : '') + '</div></div>';
  var last = null, grp = '';
  list.forEach(function (t) {
    var d = t.date, lbl, cls = '';
    if (!d) lbl = '沒有指定日期';
    else {
      var n = daysBetween(today, d);
      lbl = md(d) + '（' + wdOf(d) + '）' + (n === 0 ? '　今天' : n === 1 ? '　明天' : n === 2 ? '　後天' : n < 0 ? '　已過 ' + (-n) + ' 天' : '');
      cls = n < 0 && t.status === '待準備' ? 'late' : n <= 1 ? 'soon' : '';
    }
    if (lbl !== last) { if (last !== null) grp += '</div>'; grp += '<div class="day ' + cls + '">' + ic('cal', 's') + esc(lbl) + '</div><div class="todo">'; last = lbl; }
    grp += todoRow(t, { noDate: true });
  });
  if (last !== null) grp += '</div>';
  m.innerHTML = h + grp + '</div>';
  bindCommon(m); bindTodoRows(m);
  $$('[data-tv]', m).forEach(function (b) { b.onclick = function () { S.todoView = b.getAttribute('data-tv'); render(); }; });
}

// ---------------------------------------------------------------- 新增準備事項（抽屜）
function openAdd(pre) {
  pre = pre || {};
  var st = { date: pre.date || S.D.next.key, lab: pre.lab || '', period: pre.period || '', cls: pre.cls || '', teacher: pre.teacher || '', note: '', rows: [{}], lessons: [], manual: !!(pre.lab && !pre.cls) };
  var sh = openSheet('<h3>新增準備事項<button class="x" data-close>' + ic('x') + '</button></h3><div id="aBody"></div>');
  var body = $('#aBody', sh);
  function keep() {
    var g = function (id) { var e = $('#' + id, sh); return e ? e.value : undefined; };
    ['date', 'lab', 'period', 'cls', 'teacher', 'note'].forEach(function (k) { var v = g('a_' + k); if (v !== undefined) st[k] = v; });
    $$('.aWhat', sh).forEach(function (e) { st.rows[+e.getAttribute('data-i')].what = e.value; });
    $$('.aQty', sh).forEach(function (e) { st.rows[+e.getAttribute('data-i')].qty = e.value; });
    $$('.aUnit', sh).forEach(function (e) { st.rows[+e.getAttribute('data-i')].unit = e.value; });
    $$('.aBuy', sh).forEach(function (e) { st.rows[+e.getAttribute('data-i')].buy = e.checked; });
  }
  function isSel(l) { return l.lab === st.lab && String(l.periodText).trim() === String(st.period).trim() && l.cls === st.cls; }
  function draw() {
    var h = '<label class="lbl">哪天要用</label><input type="date" class="inp" id="a_date" value="' + esc(st.date) + '">' +
      '<label class="lbl">哪一堂課（' + md(st.date) + ' ' + wdOf(st.date) + '）</label><div class="lesson">' +
      (st.lessons === null ? '<span class="muted" style="padding:8px 0">讀取中…</span>' : st.lessons.map(function (l, i) {
        return '<button data-l="' + i + '" class="' + (isSel(l) ? 'on' : '') + '">' + esc(perLabel(l.periodText) + ' ' + short(l.lab) + ' ' + (l.cls || l.content || l.type)) + '</button>';
      }).join('') + '<button data-l="-1" class="' + (st.manual ? 'on' : '') + '">不對應／自己填</button>') + '</div>' +
      (st.manual || (!st.lessons || !st.lessons.some(isSel)) && (st.lab || st.cls) ? '<div class="g2">' +
        '<div><label class="lbl">實驗室</label><select class="inp" id="a_lab"><option value=""></option>' + S.D.rooms.concat(['準備室']).map(function (r) { return '<option' + (r === st.lab ? ' selected' : '') + '>' + esc(r) + '</option>'; }).join('') + '</select></div>' +
        '<div><label class="lbl">節次（3、3-4、午）</label><input class="inp" id="a_period" value="' + esc(st.period) + '"></div>' +
        '<div><label class="lbl">班級</label><input class="inp" id="a_cls" value="' + esc(st.cls) + '"></div>' +
        '<div><label class="lbl">老師</label><input class="inp" id="a_teacher" value="' + esc(st.teacher) + '"></div></div>' : '') +
      '<label class="lbl">要準備的東西（打字會跳出品項，找不到直接打）</label>' +
      st.rows.map(function (r, i) {
        return '<div class="row3"><input class="inp aWhat" data-i="' + i + '" placeholder="例：燒杯、冰塊、報紙" value="' + esc(r.what || '') + '" autocomplete="off">' +
          '<input class="inp aQty" data-i="' + i + '" inputmode="decimal" placeholder="數量" value="' + esc(r.qty || '') + '"><input class="inp aUnit" data-i="' + i + '" placeholder="單位" value="' + esc(r.unit || '') + '">' +
          '<button class="muted" data-rm="' + i + '">' + ic('x') + '</button><div class="info"><label class="buyck"><input type="checkbox" class="aBuy" data-i="' + i + '"' + (r.buy ? ' checked' : '') +
          '>🛒 要先買（也加進請購清單）</label>' + (r.info ? '　' + esc(r.info) : '') + '</div></div>';
      }).join('') +
      '<button class="btn g s" id="aMore" style="width:100%;margin-top:10px;color:var(--pri)">' + ic('plus', 's') + '再加一樣</button>' +
      '<label class="lbl">備註</label><input class="inp" id="a_note" placeholder="例：便條內容、老師自己來拿" value="' + esc(st.note) + '">' +
      '<button class="btn p" id="aSave" style="width:100%;margin-top:16px">儲存</button>';
    body.innerHTML = h;
    $('#a_date', sh).onchange = function () { keep(); st.lab = st.period = st.cls = st.teacher = ''; st.manual = false; fetchLessons(); };
    $$('[data-l]', sh).forEach(function (b) {
      b.onclick = function () {
        keep(); var i = +b.getAttribute('data-l');
        if (i < 0) { st.manual = true; }
        else { var l = st.lessons[i]; st.manual = false; st.lab = l.lab; st.period = String(l.periodText).trim(); st.cls = l.cls; st.teacher = l.teacher; }
        draw();
      };
    });
    $('#aMore', sh).onclick = function () { keep(); st.rows.push({}); draw(); var w = $$('.aWhat', sh); w[w.length - 1].focus(); };
    $$('[data-rm]', sh).forEach(function (b) { b.onclick = function () { keep(); st.rows.splice(+b.getAttribute('data-rm'), 1); if (!st.rows.length) st.rows.push({}); draw(); }; });
    $$('.aWhat', sh).forEach(function (inp) {
      inp.oninput = function () { var i = +inp.getAttribute('data-i'); st.rows[i].code = ''; st.rows[i].info = ''; suggest(inp); };
      inp.onblur = function () { setTimeout(function () { var s = inp.parentNode.querySelector('.sug'); if (s) s.remove(); }, 200); };
    });
    $('#aSave', sh).onclick = save;
  }
  function suggest(inp) {
    var old = inp.parentNode.querySelector('.sug'); if (old) old.remove();
    var k = inp.value.trim().toLowerCase(); if (!k) return;
    var hit = (S.D.items || []).filter(function (x) { return (x[1] + ' ' + x[2] + ' ' + x[0]).toLowerCase().indexOf(k) >= 0; }).slice(0, 8);
    if (!hit.length) return;
    var s = document.createElement('div'); s.className = 'sug';
    s.innerHTML = hit.map(function (x, i) { return '<button data-h="' + i + '">' + esc(x[1]) + ' <small>' + esc(x[2]) + '　' + esc(x[5]) + '・庫存 ' + esc(x[4] || '—') + ' ' + esc(x[3]) + (x[6] ? '（大包裝）' : '') + '</small></button>'; }).join('');
    $$('button', s).forEach(function (b) {
      b.onmousedown = function (e) {
        e.preventDefault(); keep();
        var x = hit[+b.getAttribute('data-h')], i = +inp.getAttribute('data-i');
        st.rows[i].what = x[1] + (x[2] ? ' ' + x[2] : ''); st.rows[i].code = x[0]; st.rows[i].unit = st.rows[i].unit || x[3];
        st.rows[i].info = '📍 ' + (x[5] || '—') + '　庫存 ' + (x[4] || '—') + ' ' + x[3];
        draw(); var q = $$('.aQty', sh)[i]; if (q) q.focus();
      };
    });
    inp.parentNode.appendChild(s);
  }
  function fetchLessons() {
    st.lessons = null; draw();
    call('usageData', st.date, 1).then(function (u) {
      st.lessons = u.bookings.filter(function (b) { return b.type !== '放假' && b.type !== '考試'; })
        .sort(function (a, b) { return (a.pis[0] || 0) - (b.pis[0] || 0) || (a.lab < b.lab ? -1 : 1); });
      if (!st.manual && st.lab && !st.lessons.some(isSel)) st.manual = true;
      draw();
    }).catch(function () { st.lessons = []; draw(); });
  }
  function save() {
    keep();
    var items = st.rows.filter(function (r) { return (r.what || '').trim(); });
    if (!items.length) { toast('請至少填一樣要準備的東西', true); return; }
    var b = $('#aSave', sh); b.disabled = true; b.textContent = '儲存中…';
    call('todoAdd', { date: st.date, lab: st.lab, period: st.period, cls: st.cls, teacher: st.teacher, note: st.note,
      items: items.map(function (r) { return { what: r.what, qty: r.qty, unit: r.unit, code: r.code || '', buy: !!r.buy }; }) })
      .then(function () {
        var nb = items.filter(function (r) { return r.buy; }).length;
        closeSheet(); toast('已新增 ' + items.length + ' 項準備事項' + (nb ? '，' + nb + ' 項加進請購清單' : ''));
        return reloadTodos();
      }).catch(function (e) { b.disabled = false; b.textContent = '儲存'; toast(e, true); });
  }
  fetchLessons();
}

// ---------------------------------------------------------------- 盤點
function cntRows() { return (S.count.data && S.count.data.rows) || []; }
function isDone(r) { return String(r.qty).trim() !== '' || String(r.note).trim() !== ''; }
function isDiff(r) { var q = String(r.qty).trim(), l = String(r.lastQty).trim(); return q !== '' && l !== '' && !isNaN(q) && !isNaN(l) && Number(q) !== Number(l); }
function natural(a, b) {
  var pa = String(a).match(/\d+|\D+/g) || [], pb = String(b).match(/\d+|\D+/g) || [];
  for (var i = 0; i < Math.max(pa.length, pb.length); i++) {
    var x = pa[i], y = pb[i]; if (x === undefined) return -1; if (y === undefined) return 1;
    if (!isNaN(x) && !isNaN(y)) { if (+x !== +y) return +x - +y; } else if (x !== y) return x < y ? -1 : 1;
  }
  return 0;
}
function vCount(m) {
  var C = S.count;
  if (!C.data) {
    m.innerHTML = '<div class="top"><div><h1>盤點</h1></div></div><div class="pad">' + sk(90) + sk(300) + '</div>';
    call('mobileData').then(function (d) { C.data = d; if (S.route === 'count') render(); })
      .catch(function (e) { m.innerHTML = '<div class="empty">讀取失敗：' + esc(e) + '</div>'; });
    return;
  }
  var rows = cntRows();
  if (!rows.length) {
    m.innerHTML = '<div class="top"><div><h1>盤點</h1></div></div><div class="pad"><div class="card"><div class="empty">' + ART.koalaSleep(110) +
      '<b>目前沒有進行中的盤點</b>要盤點時，在電腦的試算表按「📋 盤點 → 產生盤點表」，這裡就會出現要盤的品項。</div></div></div>';
    return;
  }
  var rooms = []; rows.forEach(function (r) { if (rooms.indexOf(r.room) < 0) rooms.push(r.room); });
  if (rooms.indexOf(C.room) < 0) C.room = rooms[0];
  var cabs = []; rows.forEach(function (r) { if (r.room === C.room && cabs.indexOf(r.cab) < 0) cabs.push(r.cab); });
  cabs.sort(natural);
  if (cabs.indexOf(C.cab) < 0) C.cab = cabs[0];
  var inCab = rows.filter(function (r) { return r.room === C.room && r.cab === C.cab; }).sort(function (a, b) { return natural(a.pos, b.pos) || natural(a.name, b.name); });
  var list = inCab.filter(function (r) { return C.filter === 'todo' ? !isDone(r) : C.filter === 'diff' ? isDiff(r) : true; });
  if (C.idx >= list.length) C.idx = Math.max(0, list.length - 1);
  var cur = list[C.idx];
  var doneCab = inCab.filter(isDone).length, doneAll = rows.filter(isDone).length;
  var h = '<div class="top"><div style="min-width:0"><h1 style="font-size:22px">' + esc(C.room + ' ' + C.cab) + '</h1><div class="sub">' + esc(C.data.name || '盤點') + (C.data.date ? '　' + esc(C.data.date) : '') + '</div></div><div class="sp"></div>' +
    '<button class="icbtn" id="cReload" title="重新讀取">' + ic('refresh') + '</button><button class="icbtn" id="cFinish" title="完成盤點">' + ic('check') + '</button></div>' +
    '<div class="pad"><div class="cgrid"><div class="cwrap">' +
    '<div class="g2" style="margin-bottom:10px"><select class="inp" id="cRoom">' + rooms.map(function (r) { return '<option' + (r === C.room ? ' selected' : '') + '>' + esc(r) + '</option>'; }).join('') + '</select>' +
    '<select class="inp" id="cCab">' + cabs.map(function (c) { var n = rows.filter(function (r) { return r.room === C.room && r.cab === c && !isDone(r); }).length; return '<option value="' + esc(c) + '"' + (c === C.cab ? ' selected' : '') + '>' + esc(c) + (n ? '（剩 ' + n + '）' : ' ✔') + '</option>'; }).join('') + '</select></div>' +
    '<div class="card prog"><div class="l"><span>這櫃進度</span><span><b style="color:var(--text)">' + doneCab + '</b> / ' + inCab.length + '　全部 ' + doneAll + ' / ' + rows.length + '</span></div>' +
    '<div class="bar"><i style="width:' + Math.round(doneCab / Math.max(1, inCab.length) * 100) + '%"></i></div>' +
    '<div class="seg" style="margin-top:10px;box-shadow:none">' + [['all', '全部 ' + inCab.length], ['todo', '還沒盤 ' + (inCab.length - doneCab)], ['diff', '有差異 ' + inCab.filter(isDiff).length]].map(function (x) {
      return '<button data-cf="' + x[0] + '" class="' + (C.filter === x[0] ? 'on' : '') + '">' + x[1] + '</button>';
    }).join('') + '</div></div><div style="height:12px"></div>';
  if (!cur) {
    var nextCab = cabs.filter(function (c) { return rows.some(function (r) { return r.room === C.room && r.cab === c && !isDone(r); }); })[0];
    h += '<div class="card"><div class="empty">' + ART.duck(110) + '<b>' + (C.filter === 'diff' ? '這櫃沒有差異' : '這櫃都盤完了！') + '</b>' +
      (nextCab && C.filter === 'todo' ? '<button class="btn p s" id="cNext" style="margin:12px auto 0">下一櫃：' + esc(nextCab) + ' ›</button>' : '') + '</div></div>';
  } else {
    var val = String(cur.qty).trim();
    h += '<div class="cur"><div class="loc">' + ic('pin', 's') + esc([cur.room, cur.cab, cur.pos].filter(String).join('　')) + '<span class="sp"></span>' + (C.idx + 1) + ' / ' + list.length + '</div>' +
      '<div class="nm">' + esc(cur.name) + '</div><div class="fx">' + esc(cur.spec) + (cur.unit ? '　單位：' + esc(cur.unit) : '') + '</div>' +
      '<div class="lastq"><div>上次 ' + esc(cur.lastDate || '') + '<b>' + esc(cur.lastQty === '' ? '—' : cur.lastQty) + ' ' + esc(cur.unit) + '</b></div>' +
      '<div>上次說明<b style="font-size:14px">' + esc(cur.lastNote || '—') + '</b></div></div>' +
      '<div class="stepper"><button class="b" id="cMinus">−</button><div class="v"><input id="cQty" inputmode="decimal" value="' + esc(val) + '" placeholder="' + esc(cur.lastQty) + '"><small>' + esc(cur.unit) + '</small></div><button class="b" id="cPlus">＋</button></div>' +
      '<input class="inp" id="cNote" style="margin-top:10px" placeholder="說明（可空白，例：1 瓶剩半、找不到）" value="' + esc(cur.note) + '">' +
      '<div class="acts"><button class="btn g" id="cPrev">' + ic('left', 's') + '上一項</button><button class="btn p" id="cSave">存好，下一項' + ic('right', 's') + '</button></div></div>';
  }
  h += '</div><div>';
  var rest = cur ? inCab.slice(inCab.indexOf(cur) + 1).concat(inCab.slice(0, inCab.indexOf(cur))) : inCab;
  h += '<div class="h2">' + (S.wide ? '這櫃全部' : '接下來') + '</div><div class="card nx" style="overflow:hidden">' + (S.wide ? inCab : rest.slice(0, 6)).map(function (r) {
    var c = isDiff(r) ? 'var(--warn)' : isDone(r) ? 'var(--ok)' : '#C3CBD8';
    var q = isDiff(r) ? r.lastQty + ' → ' + r.qty + '（差 ' + Math.round((r.qty - r.lastQty) * 1000) / 1000 + '）' : isDone(r) ? '✔ ' + (r.qty || r.note) + ' ' + r.unit : '上次 ' + (r.lastQty === '' ? '—' : r.lastQty) + ' ' + r.unit;
    return '<button class="it" data-jump="' + esc(r.code + '|' + r.place) + '"' + (r === cur ? ' style="background:var(--pri-weak)"' : '') + '><span class="dt" style="background:' + c + '"></span>' + esc(r.name) +
      ' <small class="muted">' + esc(r.spec) + '</small><span class="q" style="color:' + (isDiff(r) ? 'var(--warn)' : isDone(r) ? 'var(--ok)' : '') + '">' + esc(q) + '</span></button>';
  }).join('') + '</div></div></div></div>';
  m.innerHTML = h;
  $('#cRoom', m).onchange = function () { C.room = this.value; C.cab = ''; C.idx = 0; render(); };
  $('#cCab', m).onchange = function () { C.cab = this.value; C.idx = 0; render(); };
  $$('[data-cf]', m).forEach(function (b) { b.onclick = function () { C.filter = b.getAttribute('data-cf'); C.idx = 0; render(); }; });
  $('#cReload', m).onclick = function () { C.data = null; render(); };
  $('#cFinish', m).onclick = finishCountSheet;
  var nx = $('#cNext', m); if (nx) nx.onclick = function () { C.cab = nx.textContent.replace(/^下一櫃：| ›$/g, ''); C.idx = 0; render(); };
  $$('[data-jump]', m).forEach(function (b) {
    b.onclick = function () {
      var k = b.getAttribute('data-jump');
      C.filter = 'all'; render();
      var l2 = inCab; C.idx = l2.map(function (r) { return r.code + '|' + r.place; }).indexOf(k); render();
    };
  });
  if (!cur) return;
  var q = $('#cQty', m);
  var step = function (d) { var v = q.value.trim() === '' ? (isNaN(cur.lastQty) || cur.lastQty === '' ? 0 : Number(cur.lastQty)) : Number(q.value); if (isNaN(v)) v = 0; q.value = Math.max(0, Math.round((v + d) * 1000) / 1000); };
  $('#cMinus', m).onclick = function () { step(-1); };
  $('#cPlus', m).onclick = function () { step(1); };
  $('#cPrev', m).onclick = function () { if (C.idx > 0) { C.idx--; render(); } };
  q.onkeydown = function (e) { if (e.key === 'Enter') $('#cSave', m).click(); };
  $('#cSave', m).onclick = function () {
    var qty = q.value.trim(), note = $('#cNote', m).value.trim();
    if (qty === '' && note === '') { qty = String(cur.lastQty); if (qty === '') { toast('請填數量或說明', true); return; } }
    var old = [cur.qty, cur.note];
    cur.qty = qty; cur.note = note;
    if (C.filter !== 'todo') C.idx++;
    render();
    toast('已存：' + cur.name + ' ' + (qty || note) + ' ' + cur.unit);
    call('mobileSave', { code: cur.code, place: cur.place, qty: qty, note: note }).catch(function (e) {
      cur.qty = old[0]; cur.note = old[1]; render(); toast('沒有存到「' + cur.name + '」：' + e, true);
    });
  };
}
function finishCountSheet() {
  var rows = cntRows(), done = rows.filter(isDone).length;
  openSheet('<h3>完成盤點<button class="x" data-close>' + ic('x') + '</button></h3>' +
    '<p style="color:var(--sub);margin:4px 0 10px">已填 <b style="color:var(--text)">' + done + '</b> / ' + rows.length + ' 項。存進「盤點紀錄」後，品項的最新數量會更新，並自動備份。</p>' +
    '<label style="display:flex;gap:10px;align-items:flex-start;padding:10px 0"><input type="radio" name="fm" value="keep" checked style="margin-top:4px"><span><b>沒填的留著</b><br><small class="muted">下次繼續盤</small></span></label>' +
    '<label style="display:flex;gap:10px;align-items:flex-start;padding:10px 0"><input type="radio" name="fm" value="carry" style="margin-top:4px"><span><b>沒填的沿用上次數量</b><br><small class="muted">一起存進盤點紀錄</small></span></label>' +
    '<button class="btn p" id="fGo" style="width:100%;margin-top:10px">存入盤點紀錄</button>', function (sh) {
    $('#fGo', sh).onclick = function () {
      var b = this, mode = $('input[name=fm]:checked', sh).value;
      b.disabled = true; b.textContent = '儲存中，約 10～20 秒…';
      call('mobileFinish', mode).then(function (msg) {
        closeSheet(); toast(String(msg).split('\n')[0] || '完成'); S.count.data = null; refresh(false);
      }).catch(function (e) { b.disabled = false; b.textContent = '存入盤點紀錄'; toast(e, true); });
    };
  });
}

// ---------------------------------------------------------------- 更多、需補充、品項查詢
function vMore(m) {
  var t = getTheme();
  m.innerHTML = '<div class="top"><div><h1>更多</h1></div></div><div class="pad">' +
    '<div class="menu"><button data-go2="restock"><span class="ic" style="background:var(--bad-weak);color:var(--bad)">' + ic('alert') + '</span>需補充<span class="r">' + (S.D.kpi.restock === null ? '' : S.D.kpi.restock + ' 項') + ' ›</span></button>' +
    '<button data-go2="search"><span class="ic">' + ic('box') + '</span>品項查詢<span class="r">›</span></button>' +
    '<button data-go2="count"><span class="ic" style="background:var(--ok-weak);color:var(--ok)">' + ic('count') + '</span>盤點<span class="r">' + (S.D.kpi.count.total ? S.D.kpi.count.done + ' / ' + S.D.kpi.count.total : '') + ' ›</span></button></div>' +
    '<div class="h2">外觀</div><div class="seg">' + [['auto', '跟著系統'], ['light', '淺色'], ['dark', '深色']].map(function (x) { return '<button data-th="' + x[0] + '" class="' + (t === x[0] ? 'on' : '') + '">' + x[1] + '</button>'; }).join('') + '</div>' +
    '<div class="h2">其他</div><div class="menu"><button data-act2="refresh"><span class="ic">' + ic('refresh') + '</span>重新整理資料<span class="r">›</span></button>' +
    (S.D.url ? '<button id="mOld"><span class="ic">' + ic('grid') + '</span>舊版網頁<span class="r">›</span></button>' : '') + '</div>' +
    '<div class="empty" style="margin-top:10px">' + ART.koala(70) + '<b>實驗室管理</b>' + esc(S.D.school) + '　設備組<br><small class="muted">把網頁「加到主畫面」，用起來就像 App</small></div></div>';
  bindCommon(m);
  $$('[data-th]', m).forEach(function (b) { b.onclick = function () { setTheme(b.getAttribute('data-th')); badges(); render(); }; });
  var o = $('#mOld', m); if (o) o.onclick = function () { window.open(S.D.url + '?page=old', '_blank'); };
}
function vRestock(m) {
  m.innerHTML = '<div class="top">' + (S.wide ? '' : '<button class="icbtn" data-go2="more">' + ic('left') + '</button>') + '<div><h1>需補充</h1><div class="sub">低於安全存量的品項</div></div></div><div class="pad" id="rBody">' + sk(70, 5) + '</div>';
  bindCommon(m);
  loadRestock().then(function (rows) {
    var el = $('#rBody'); if (!el) return;
    el.innerHTML = rows.length ? '<div class="rgrid">' + rows.map(function (r) {
      return '<div class="res"><div class="t">' + esc(r.name) + ' <small class="muted">' + esc(r.spec) + '</small></div><div class="s">' + ic('pin', 's') + ' ' + esc(r.loc || '—') +
        '<br>目前 <b class="late">' + esc(r.qty) + '</b> ' + esc(r.unit) + '　安全存量 ' + esc(r.safe) + '　建議補 <b>' + esc(r.need) + '</b> ' + esc(r.unit) + (r.date ? '<br><small class="muted">盤點 ' + esc(r.date) + '</small>' : '') + '</div></div>';
    }).join('') + '</div>' : '<div class="card"><div class="empty">' + ART.duck(110) + '<b>沒有需要補充的</b>庫存都在安全存量以上</div></div>';
  });
}
function vSearch(m) {
  var Q = S.search;
  m.innerHTML = '<div class="top">' + (S.wide ? '' : '<button class="icbtn" data-go2="more">' + ic('left') + '</button>') + '<div><h1>品項查詢</h1></div></div><div class="pad" style="max-width:980px">' +
    '<input class="inp" id="qKw" placeholder="品名、化學式、編號、櫃別" value="' + esc(Q.kw) + '" style="font-size:16px">' +
    '<div class="labsel" style="margin:10px 0">' + [['', '全部'], ['藥品', '藥品'], ['器材', '器材'], ['耗材', '耗材']].map(function (x) { return '<button class="pill' + (Q.cat === x[0] ? ' on' : '') + '" data-cat="' + x[0] + '">' + x[1] + '</button>'; }).join('') + '</div>' +
    '<div id="qBody"></div></div>';
  bindCommon(m);
  var tm, q = $('#qKw', m);
  var run = function () {
    var el = $('#qBody'); el.innerHTML = sk(64, 4);
    call('mobileSearch', { keyword: Q.kw, category: Q.cat, room: '' }).then(function (r) { Q.rows = r; draw(); }).catch(function (e) { el.innerHTML = '<div class="empty">查詢失敗：' + esc(e) + '</div>'; });
  };
  var draw = function () {
    var el = $('#qBody'), r = Q.rows; if (!el || !r) return;
    el.innerHTML = r.rows.length ? '<div class="rgrid">' + r.rows.map(function (x) {
      return '<div class="res"><div class="t">' + esc(x.name) + ' <small class="muted">' + esc(x.spec) + '</small>' + (x.status === '已淘汰' ? ' <span class="tag" style="background:var(--bad-weak)">已淘汰</span>' : '') + '</div>' +
        '<div class="s">' + esc(x.code) + '　' + ic('pin', 's') + ' ' + esc([x.room, x.cab, x.pos].filter(String).join(' ')) + (x.places ? '（另放：' + esc(x.places) + '）' : '') +
        '<br>最新 <b style="color:var(--text)">' + esc(x.note || x.qty || '—') + '</b> ' + esc(x.unit) + (x.date ? '<small class="muted">（' + esc(x.date) + '）</small>' : '') + '</div></div>';
    }).join('') + '</div>' + (r.more ? '<p class="muted" style="text-align:center">只顯示前 60 筆，請打更精確的關鍵字</p>' : '') :
      '<div class="card"><div class="empty">' + ART.koalaSleep(90) + '<b>找不到</b>換個關鍵字試試</div></div>';
  };
  q.oninput = function () { Q.kw = q.value.trim(); clearTimeout(tm); tm = setTimeout(run, 400); };
  $$('[data-cat]', m).forEach(function (b) { b.onclick = function () { Q.cat = b.getAttribute('data-cat'); $$('[data-cat]', m).forEach(function (x) { x.classList.toggle('on', x === b); }); run(); }; });
  if (Q.rows) draw(); else if (Q.kw || true) run();
  if (!S.wide) setTimeout(function () { q.focus(); }, 50);
}

// ---------------------------------------------------------------- 開始
document.getElementById('fab').onclick = function () { openAdd({}); };
load();
