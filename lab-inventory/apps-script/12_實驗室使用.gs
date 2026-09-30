/**
 * 實驗室藥品器材耗材管理系統 — 實驗室使用登記、門口海報、本週課表、六間使用一覽（第四批）
 *
 * 🗓 實驗室使用
 *   ・登記使用（一週課表）：選實驗室、週次、用途類型（實驗課、社團、自主學習…），在課表格子上點選時段、填班級，一次存檔
 *       可每週重複到某日；同一間、同一天、同一節已有人登記會提醒（撞堂）
 *   ・六間實驗室使用一覽：今天／整週；現在這一節會標亮；手機網頁也有「使用」分頁
 *   ・列印門口海報（A4 橫式）：實驗名稱、年級、實驗室、本週日期、授課教師、班級、安全注意事項、講義 QR code
 *   ・列印本週課表（A4 直式）：該實驗室這週每一節誰在用
 * 資料存在「實驗排程」工作表（一列＝某間實驗室、某天、某幾節的一次使用）。
 * 節次時間、用途類型與顏色、上課教室、安全注意事項都在「設定」可以改。
 * 需要「01_基礎」「04_列印」「09_請購借用」「10_實驗準備」。
 */

const LAB_DEFAULTS = {
  rooms: ['化學實驗室一', '化學實驗室二', '化學實驗室三', '生物實驗室一', '生物實驗室二', '生物實驗室三'],
  periods: [['第1節', '08:10', '09:00'], ['第2節', '09:10', '10:00'], ['第3節', '10:10', '11:00'], ['第4節', '11:10', '12:00'],
    ['中午', '12:00', '13:10'], ['第5節', '13:10', '14:00'], ['第6節', '14:10', '15:00'], ['第7節', '15:20', '16:10']],
  types: [['實驗課', '#D2E3FC'], ['補做實驗', '#C6DAFC'], ['社團', '#CEEAD6'], ['自主學習', '#FEEFC3'], ['專題研究', '#FAD2CF'],
    ['多元選修', '#FDE2F3'], ['老師借用', '#E9D2FD'], ['借用教室', '#EDE7F6'], ['考試', '#FFE0B2'], ['研習', '#D7F3F5'],
    ['放假', '#DADCE0'], ['其他', '#E8EAED']],
  // 班名（照班序：仁＝01、義＝02…廉＝21）；課表顯示「二敬」，打「213」也會顯示成「二敬」
  classNames: ['仁', '義', '禮', '智', '忠', '孝', '博', '愛', '和', '平', '誠', '信', '敬', '業', '樂', '群', '簡', '捷', '敏', '慧', '廉'],
  safety: '進入實驗室請穿實驗衣、戴護目鏡，長髮請綁好｜實驗室內禁止飲食｜依老師指示操作，不可擅自取用藥品或器材｜' +
    '廢液、廢棄物依規定分類回收，不可倒入水槽｜發生意外或受傷，立即報告老師',
};
const WEEKDAYS = ['一', '二', '三', '四', '五', '六', '日'];
const PREP_TYPES = ['實驗課', '補做實驗'];

// ---------------------------------------------------------------- 設定

/** 「設定」沒有上課教室、節次、用途類型、安全注意事項時自動補上（之後可在設定直接改）。 */
function ensureLabSettings_() {
  const s = getSettings_();
  const sh = s.sheet;
  let lastCol = sh.getLastColumn();
  const addTable = function (headers, rows, colorCol) {
    if (headers.every(function (h) { return s.header.indexOf(h) >= 0; })) return;
    const c = lastCol + 2;
    sh.getRange(1, c, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#DDEBF7');
    sh.getRange(2, c, rows.length, headers.length).setNumberFormat('@').setValues(rows);
    if (colorCol !== undefined) rows.forEach(function (r, i) { sh.getRange(2 + i, c + colorCol).setBackground(r[colorCol]); });
    lastCol = c + headers.length - 1;
  };
  addTable(['上課教室'], LAB_DEFAULTS.rooms.map(function (r) { return [r]; }));
  addTable(['節次', '開始', '結束'], LAB_DEFAULTS.periods);
  addTable(['用途類型', '顏色'], LAB_DEFAULTS.types, 1);
  addTable(['班名'], LAB_DEFAULTS.classNames.map(function (r) { return [r]; }));
  if (!('安全注意事項' in s.params)) {
    const colA = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
    let last = 0;
    colA.forEach(function (v, i) { if (String(v[0]).trim()) last = i; });
    sh.getRange(last + 2, 1, 1, 2).setValues([['安全注意事項', LAB_DEFAULTS.safety]]);
    sh.getRange(last + 2, 2).setWrap(true).setNote('用「｜」分隔每一條，會印在門口海報上。');
  }
}

/** 實驗室使用相關設定。 */
function labConfig_() {
  ensureLabSettings_();
  const s = getSettings_();
  // 時間用「顯示的文字」讀，避免被當成日期時間；8:10 → 08:10
  const t = function (v) {
    const m = String(v || '').trim().match(/^(\d{1,2}):(\d{2})/);
    return m ? pad2_(Number(m[1])) + ':' + m[2] : String(v || '').trim();
  };
  const col = function (name) { return s.header.indexOf(name); };
  const rows = s.sheet.getDataRange().getDisplayValues().slice(1);
  const periods = [];
  const pc = col('節次');
  rows.forEach(function (r) {
    const n = String(r[pc]).trim();
    if (!n) return;
    periods.push({ name: n, short: periodShort_(n), start: t(r[pc + 1]), end: t(r[pc + 2]) });
  });
  const types = [], colors = {};
  const tc = col('用途類型');
  rows.forEach(function (r) {
    const n = String(r[tc]).trim();
    if (!n) return;
    types.push(n);
    colors[n] = String(r[tc + 1]).trim() || '#E8EAED';
  });
  return {
    rooms: s.lists['上課教室'] || LAB_DEFAULTS.rooms, periods: periods, types: types, colors: colors,
    classNames: s.lists['班名'] || LAB_DEFAULTS.classNames,
    safety: String(s.params['安全注意事項'] || '').split(/[｜|\n]/).map(function (x) { return x.trim(); }).filter(String),
    school: String(s.params['學校名稱'] || ''),
  };
}

function periodShort_(name) {
  return /午/.test(name) ? '午' : (String(name).replace(/[^0-9]/g, '') || String(name));
}

/** 「3-4」「午」「1,2」「第3節」→ 節次索引（依設定的節次順序）。範圍中間跨過中午時不含中午。 */
function parsePeriods_(str, periods) {
  const shorts = periods.map(function (p) { return p.short; });
  const idx = function (tok) {
    tok = String(tok).trim();
    if (!tok) return -1;
    if (/午/.test(tok)) return shorts.indexOf('午');
    const d = tok.replace(/[^0-9]/g, '');
    let i = d ? shorts.indexOf(d) : -1;
    if (i < 0) i = periods.map(function (p) { return p.name; }).indexOf(tok);
    return i;
  };
  const out = {};
  String(str || '').split(/[,，、\s]+/).forEach(function (tok) {
    if (!tok) return;
    const m = tok.split(/[-~～–]/);
    if (m.length === 2) {
      const a = idx(m[0]), b = idx(m[1]);
      if (a < 0 || b < 0) return;
      for (let i = Math.min(a, b); i <= Math.max(a, b); i++) {
        if (shorts[i] === '午' && i !== a && i !== b) continue;
        out[i] = true;
      }
    } else {
      const i = idx(tok);
      if (i >= 0) out[i] = true;
    }
  });
  return Object.keys(out).map(Number).sort(function (a, b) { return a - b; });
}

/** 節次索引 → 連續的分段（中午自成一段）：[[2,3],[4]] */
function periodRuns_(idxs, periods) {
  const runs = [];
  idxs.slice().sort(function (a, b) { return a - b; }).forEach(function (i) {
    const last = runs[runs.length - 1];
    const noon = periods[i].short === '午';
    if (last && !noon && last[last.length - 1] === i - 1 && periods[i - 1].short !== '午') last.push(i);
    else runs.push([i]);
  });
  return runs;
}

/** 節次文字 → 給人看的：「3-4」→「第 3-4 節」、「午」→「中午」 */
function periodLabel_(text) {
  text = String(text).trim();
  return text === '午' ? '中午' : '第 ' + text + ' 節';
}

function runText_(run, periods) {
  return run.length === 1 ? periods[run[0]].short : periods[run[0]].short + '-' + periods[run[run.length - 1]].short;
}

/** 某日期所在那週的星期一（yyyy-MM-dd）。 */
function mondayOf_(key) {
  const d = Utilities.parseDate(key, tz_(), 'yyyy-MM-dd');
  const wd = Number(Utilities.formatDate(d, tz_(), 'u'));   // 1＝星期一
  return addDays_(key, 1 - wd);
}

function addDays_(key, n) {
  const d = Utilities.parseDate(key, tz_(), 'yyyy-MM-dd');
  return Utilities.formatDate(new Date(d.getTime() + n * 86400000 + 3600000), tz_(), 'yyyy-MM-dd');
}

function weekdayOf_(key) {
  return WEEKDAYS[Number(Utilities.formatDate(Utilities.parseDate(key, tz_(), 'yyyy-MM-dd'), tz_(), 'u')) - 1];
}

// ---------------------------------------------------------------- 讀登記資料

/** 讀「實驗排程」；工作表或「用途類型」欄還沒有時才建立（建立時會套下拉選單，比較慢）。 */
function expTable_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(EXP_SHEET);
  if (sh) {
    const t = getTable_(EXP_SHEET);
    if ('用途類型' in t.col) return t;
  }
  return setupExperimentSheet_();
}

/** 期間內的登記（不含取消）：[{row, date, lab, pis, type, content, cls, teacher, groups, status, note}] */
function bookings_(fromKey, toKey, cfg) {
  const t = expTable_();
  const c = t.col;
  const out = [];
  t.rows.forEach(function (r, i) {
    const k = dateKey_(r[c['日期']]);
    const st = String(r[c['狀態']]).trim();
    if (!k || k < fromKey || k > toKey || st === '取消') return;
    out.push({
      row: i + 2, date: k, lab: String(r[c['教室']]).trim(), pis: parsePeriods_(r[c['節次']], cfg.periods),
      periodText: String(r[c['節次']]), type: String(r[c['用途類型']] || '').trim() || '實驗課',
      content: String(r[c['實驗名稱']]).trim(), cls: classLabel_(r[c['班級']], cfg.classNames), teacher: String(r[c['教師']]).trim(),
      groups: r[c['組數']], status: st, note: String(r[c['備註']]).trim(),
    });
  });
  return out;
}

/** 給對話框、手機的使用一覽資料。 */
function usageData(fromKey, days) {
  const cfg = labConfig_();
  fromKey = dateKey_(fromKey) || today_();
  days = Math.max(1, Math.min(120, Number(days) || 1));
  const toKey = addDays_(fromKey, days - 1);
  const list = [];
  for (let i = 0; i < days; i++) {
    const k = addDays_(fromKey, i);
    list.push({ key: k, roc: rocText_(k), wd: weekdayOf_(k) });
  }
  const now = new Date();
  return {
    rooms: cfg.rooms, periods: cfg.periods, colors: cfg.colors, days: list,
    bookings: bookings_(fromKey, toKey, cfg),
    // 有安裝「14_準備事項」時，附上這段期間待準備的事項（格子上顯示 📝）
    todos: typeof globalThis.todoPending_ === 'function' ? globalThis.todoPending_(fromKey, toKey) : [],
    today: today_(), nowHm: Utilities.formatDate(now, SCHOOL_TZ, 'HH:mm'),
  };
}

// ---------------------------------------------------------------- 登記使用（一週課表）

function bookDialog() {
  const cfg = labConfig_();
  const k = kits_();
  const data = {
    rooms: cfg.rooms, periods: cfg.periods, types: cfg.types, colors: cfg.colors,
    kits: k.order.map(function (n) { return { name: n, groups: k.map[n].groups }; }),
    monday: mondayOf_(today_()),
  };
  const html = DIALOG_STYLE + `
    <style>
      body { font-size: 14px; } .top { display: grid; grid-template-columns: repeat(4, 1fr); gap: 0 10px; }
      .top label { margin-top: 8px; } .span2 { grid-column: span 2; }
      table.g { border-collapse: collapse; width: 100%; margin-top: 12px; table-layout: fixed; }
      table.g th, table.g td { border: 1px solid #dadce0; padding: 3px; text-align: center; font-size: 12px; height: 44px; vertical-align: middle; }
      table.g th { background: #f1f3f4; } table.g th small { color: #5f6368; font-weight: normal; }
      td.free { cursor: pointer; } td.free:hover { background: #e8f0fe; }
      td.sel { background: #1a73e8 !important; color: #fff; cursor: pointer; } td.sel input { width: 92%; font-size: 12px; padding: 2px; border: none; border-radius: 3px; text-align: center; }
      td.busy { color: #3c4043; font-size: 11px; line-height: 1.25; } td.busy.sel { outline: 3px solid #d93025; }
      .legend span { display: inline-block; padding: 2px 8px; border-radius: 10px; margin: 2px; font-size: 12px; }
      .nav button { padding: 4px 10px; margin: 0 2px; } .conf { background: #fce8e6; color: #b3261e; padding: 8px 10px; border-radius: 6px; margin-top: 10px; white-space: pre-line; }
    </style>
    <div class="top">
      <div><label>實驗室</label><select id="lab" onchange="load()"></select></div>
      <div class="span2"><label>週次（選該週任一天）</label>
        <div class="row nav"><button onclick="shift(-7)">◀</button><input type="date" id="week" onchange="load()"><button onclick="shift(7)">▶</button></div></div>
      <div><label>&nbsp;</label><label class="inline"><input type="checkbox" id="wk" onchange="draw()"> 顯示週末</label></div>
      <div><label>用途類型</label><select id="type" onchange="typeChange()"></select></div>
      <div class="span2"><label>內容（實驗名稱／社團名稱…）</label>
        <select id="kit" onchange="kitChange()"></select><input id="content" placeholder="內容" style="display:none;margin-top:4px"></div>
      <div><label>組數</label><input id="groups" inputmode="numeric"></div>
      <div><label>教師／負責人</label><input id="teacher"></div>
      <div><label>預設班級／對象</label><input id="cls" placeholder="點格子時自動帶入"></div>
      <div><label>每週重複到（可空白）</label><input type="date" id="repeat"></div>
      <div><label>備註</label><input id="note"></div>
    </div>
    <p class="hint">點課表上的空格選時段（藍色），可以直接改每一格的班級；有顏色的格子是已經有人登記的（點了會提醒撞堂）。連續的節次會自動合併成一筆（例：3-4）。</p>
    <div id="grid"></div><div class="legend" id="legend"></div><div id="conf"></div>
    <div class="btns"><button onclick="google.script.host.close()">關閉</button>
      <button class="primary" id="go" onclick="save(false)">登記</button></div>
    <script>
      const D = __DATA__; let U = null; const sel = {};
      const $ = function (id) { return document.getElementById(id); };
      D.rooms.forEach(function (x) { $('lab').add(new Option(x, x)); });
      D.types.forEach(function (x) { $('type').add(new Option(x, x)); });
      D.kits.forEach(function (x) { $('kit').add(new Option(x.name, x.name)); }); $('kit').add(new Option('其他（自行輸入）', '__other'));
      $('week').value = D.monday;
      D.types.forEach(function (t) { const s = document.createElement('span'); s.textContent = t; s.style.background = D.colors[t]; $('legend').appendChild(s); });
      function isPrep() { return $('type').value === '實驗課' || $('type').value === '補做實驗'; }
      function typeChange() { const p = isPrep(); $('kit').style.display = p ? '' : 'none'; $('content').style.display = p && $('kit').value !== '__other' ? 'none' : ''; $('groups').disabled = !p; }
      function kitChange() { const k = D.kits.filter(function (x) { return x.name === $('kit').value; })[0]; if (k && k.groups) $('groups').value = k.groups; typeChange(); }
      typeChange(); kitChange();
      function key(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
      function parse(k) { const a = k.split('-'); return new Date(+a[0], +a[1] - 1, +a[2]); }
      function monday(k) { const d = parse(k); const wd = (d.getDay() + 6) % 7; d.setDate(d.getDate() - wd); return key(d); }
      function shift(n) { const d = parse(monday($('week').value || D.monday)); d.setDate(d.getDate() + n); $('week').value = key(d); load(); }
      function load() {
        const m = monday($('week').value || D.monday); $('week').value = m;
        $('grid').innerHTML = '<p class="hint">讀取中…</p>';
        google.script.run.withSuccessHandler(function (u) { U = u; draw(); }).withFailureHandler(function (e) { alert(e.message); }).usageData(m, 7);
      }
      function draw() {
        if (!U) return;
        const lab = $('lab').value, days = U.days.slice(0, $('wk').checked ? 7 : 5);
        let h = '<table class="g"><tr><th style="width:74px">節次</th>';
        days.forEach(function (d) { h += '<th>' + d.roc.slice(4) + '<br><small>（' + d.wd + '）</small></th>'; });
        h += '</tr>';
        U.periods.forEach(function (p, pi) {
          h += '<tr><th>' + p.name + '<br><small>' + p.start + '–' + p.end + '</small></th>';
          days.forEach(function (d) {
            const id = d.key + '|' + pi;
            const busy = U.bookings.filter(function (b) { return b.lab === lab && b.date === d.key && b.pis.indexOf(pi) >= 0; });
            const cls = (busy.length ? 'busy' : 'free') + (id in sel ? ' sel' : '');
            const bg = busy.length ? ' style="background:' + (U.colors[busy[0].type] || '#eee') + '"' : '';
            h += '<td class="' + cls + '"' + bg + ' data-id="' + id + '" onclick="tog(this)">';
            if (id in sel) h += '<input value="' + (sel[id] || '').replace(/"/g, '&quot;') + '" onclick="event.stopPropagation()" oninput="sel[\\'' + id + '\\']=this.value">';
            else if (busy.length) h += busy.map(function (b) { return '<b>' + esc(b.content || b.type) + '</b><br>' + esc(b.cls) + ' ' + esc(b.teacher); }).join('<hr>');
            h += '</td>';
          });
          h += '</tr>';
        });
        $('grid').innerHTML = h + '</table>';
      }
      function esc(s) { return String(s || '').replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
      function tog(td) { const id = td.getAttribute('data-id'); if (id in sel) delete sel[id]; else sel[id] = $('cls').value.trim(); $('conf').innerHTML = ''; draw(); }
      function save(force) {
        const slots = Object.keys(sel).map(function (id) { const a = id.split('|'); return { date: a[0], pi: +a[1], cls: sel[id] }; });
        if (!slots.length) { alert('請在課表上點選要登記的時段。'); return; }
        const content = isPrep() && $('kit').value !== '__other' ? $('kit').value : $('content').value.trim();
        if (!content) { alert('請填內容（實驗名稱、社團名稱…）。'); return; }
        const b = $('go'); b.disabled = true; b.textContent = '登記中…';
        google.script.run.withSuccessHandler(function (res) {
          b.disabled = false; b.textContent = '登記';
          if (res.conflicts) {
            $('conf').innerHTML = '<div class="conf"></div>';
            $('conf').firstChild.textContent = '⚠ 這些時段已經有人登記：\\n' + res.conflicts.join('\\n') + '\\n\\n確定仍要登記（同時段兩筆）嗎？';
            const y = document.createElement('button'); y.textContent = '仍要登記'; y.onclick = function () { save(true); };
            $('conf').firstChild.appendChild(document.createElement('br')); $('conf').firstChild.appendChild(y);
            return;
          }
          showDone(res.msg);
        }).withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '登記'; })
          .saveBooking({ lab: $('lab').value, type: $('type').value, content: content, teacher: $('teacher').value.trim(),
            groups: $('groups').value.trim(), note: $('note').value.trim(), repeat: $('repeat').value, slots: slots, force: force });
      }
      load();
    </script>`;
  const json = JSON.stringify(data).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html.replace('__DATA__', function () { return json; }))
    .setWidth(960).setHeight(760), '🗓 登記實驗室使用');
}

/** 存登記：同一天、同班級的連續節次合併成一筆；可每週重複；撞堂時先回報（force 才寫入）。 */
function saveBooking(p) {
  const cfg = labConfig_();
  if (!p.slots || !p.slots.length) throw new Error('沒有選時段。');
  // 每週重複
  let slots = p.slots.slice();
  const until = dateKey_(p.repeat);
  if (until) {
    p.slots.forEach(function (s) {
      for (let k = addDays_(s.date, 7); k <= until; k = addDays_(k, 7)) slots.push({ date: k, pi: s.pi, cls: s.cls });
    });
  }
  // 依 日期＋班級 分組 → 連續節次合併
  const groups = {};
  slots.forEach(function (s) {
    const g = s.date + '\u0001' + (s.cls || '');
    (groups[g] = groups[g] || []).push(s.pi);
  });
  const rows = [];
  Object.keys(groups).sort().forEach(function (g) {
    const a = g.split('\u0001');
    periodRuns_(groups[g], cfg.periods).forEach(function (run) {
      rows.push({ date: a[0], cls: a[1], pis: run, text: runText_(run, cfg.periods) });
    });
  });
  // 撞堂檢查
  const keys = rows.map(function (r) { return r.date; }).sort();
  const existing = bookings_(keys[0], keys[keys.length - 1], cfg).filter(function (b) { return b.lab === p.lab; });
  const conflicts = [];
  rows.forEach(function (r) {
    existing.forEach(function (b) {
      if (b.date === r.date && b.pis.some(function (x) { return r.pis.indexOf(x) >= 0; })) {
        conflicts.push(rocText_(r.date) + '（' + weekdayOf_(r.date) + '）' + periodLabel_(b.periodText) + '：' +
          [b.content || b.type, b.cls, b.teacher].filter(String).join(' '));
      }
    });
  });
  if (conflicts.length && !p.force) return { conflicts: conflicts.filter(function (x, i, a) { return a.indexOf(x) === i; }) };

  const t = expTable_();
  const prep = PREP_TYPES.indexOf(p.type) >= 0;
  rows.forEach(function (r) {
    appendRow_(t, { '日期': dateValue_(r.date), '節次': r.text, '班級': r.cls, '教師': p.teacher, '教室': p.lab,
      '實驗名稱': p.content, '組數': prep && p.groups ? Number(p.groups) : '', '狀態': prep ? '待準備' : '',
      '備註': p.note, '用途類型': p.type });
  });
  let msg = '已登記 ' + rows.length + ' 筆：' + p.lab + '　' + p.content + '（' + p.type + '）\n' +
    rows.slice(0, 12).map(function (r) { return rocText_(r.date) + '（' + weekdayOf_(r.date) + '）' + periodLabel_(r.text) + ' ' + r.cls; }).join('\n') +
    (rows.length > 12 ? '\n…共 ' + rows.length + ' 筆' : '');
  if (conflicts.length) msg += '\n\n⚠ 有 ' + conflicts.length + ' 個時段與其他登記重疊（已照您的選擇登記）。';
  if (prep) {
    const check = experimentCheck_(p.content, Number(p.groups) || 0);
    if (check) {
      const short = check.rows.filter(function (x) { return x.short; });
      msg += short.length ? '\n\n⚠ 器材不足：' + short.map(function (x) { return x.name + '（需 ' + x.need + '，有 ' + x.have + '）'; }).join('、') :
        '\n\n✔ 器材數量足夠';
    }
  }
  return { msg: msg };
}

// ---------------------------------------------------------------- 六間實驗室使用一覽

function usageBoard() {
  const html = DIALOG_STYLE + '<div id="board"></div><script>' + usageBoardJs_() + extraJs_(['todoJs_', 'lessonJs_']) +
    'boardInit("board", { print: true });</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(1100).setHeight(780), '🗓 六間實驗室使用一覽');
}

/**
 * 使用一覽的畫面程式（電腦對話框、手機網頁共用）。呼叫 boardInit(容器id, {print}) 開始。
 * 檢視：單日（六間 × 各節）、整週、月曆、清單（任選期間）；可跳到任一天、只看某一間、用關鍵字找（例：二敬、仲文、多元選修）。
 */
/** 其他檔案的畫面程式（有安裝才加進來） */
function extraJs_(names) {
  return names.map(function (n) { return typeof globalThis[n] === 'function' ? globalThis[n]() : ''; }).join('');
}

function usageBoardJs_() {
  return `
  (function () {
    const css = document.createElement('style');
    css.textContent = '.ub-bar{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:4px 0 6px}.ub-bar button{padding:6px 11px;font-size:14px;border:1px solid #dadce0;border-radius:6px;background:#fff;cursor:pointer}' +
      '.ub-bar button.on{background:#1a73e8;color:#fff;border-color:#1a73e8}.ub-bar input,.ub-bar select{width:auto;margin:0;font-size:14px;padding:5px 6px;border:1px solid #dadce0;border-radius:6px;background:#fff;max-width:100%}' +
      '.ub-bar .ub-date,.ub-bar .ub-end,.ub-bar .ub-lab{width:auto!important}.ub-bar .ub-kw{width:auto!important;flex:1;min-width:140px}.ub-to{display:inline-flex;align-items:center;gap:4px}' +
      '.ub-title{font-weight:bold;font-size:16px;margin:2px 0 8px}.ub-sp{flex:1}' +
      '.ub-wrap{overflow-x:auto;margin-bottom:14px}table.ub{border-collapse:collapse;width:100%;min-width:560px;table-layout:fixed;background:#fff}table.ub th,table.ub td{border:1px solid #dadce0;padding:3px;font-size:12px;text-align:center;vertical-align:middle;height:40px}' +
      'table.ub th{background:#f1f3f4}table.ub tr>th:first-child{position:sticky;left:0;z-index:1}table.ub th small{color:#5f6368;font-weight:normal}table.ub tr.now th,table.ub tr.now td{box-shadow:inset 0 0 0 2px #d93025}' +
      'table.ub tr.now th{background:#fce8e6;color:#b3261e}table.ub th.today{background:#d2e3fc}.ub-day{font-weight:bold;margin:6px 0 4px}.ub-b{border-radius:4px;padding:2px;line-height:1.25;margin:1px 0}' +
      'table.um{border-collapse:collapse;width:100%;min-width:620px;table-layout:fixed;background:#fff}table.um th{background:#f1f3f4;border:1px solid #dadce0;padding:4px;font-size:13px}' +
      'table.um td{border:1px solid #dadce0;vertical-align:top;padding:2px 3px;height:74px;font-size:11px;text-align:left}table.um td.out{background:#f8f9fa;color:#bbb}table.um td.today{box-shadow:inset 0 0 0 2px #1a73e8}' +
      'table.um .dn{font-weight:bold;font-size:12px;cursor:pointer;color:#1a73e8}table.um .e{border-radius:3px;padding:1px 3px;margin:1px 0;line-height:1.3;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}' +
      'table.ul{border-collapse:collapse;width:100%;background:#fff}table.ul th,table.ul td{border:1px solid #dadce0;padding:4px 6px;font-size:13px;text-align:left}table.ul th{background:#f1f3f4}' +
      '.ub-td{font-size:11px;background:#fff;border-radius:8px;padding:0 4px;margin-left:2px}.ub-cnt{color:#5f6368;font-size:13px;margin:4px 0}.ub-legend span{display:inline-block;padding:2px 8px;border-radius:10px;margin:2px;font-size:12px}' +
      '@media print{.ub-bar{display:none}.ub-wrap{overflow:visible}}';
    document.head.appendChild(css);
    function key(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
    function parse(k) { const a = k.split('-'); return new Date(+a[0], +a[1] - 1, +a[2]); }
    function add(k, n) { const d = parse(k); d.setDate(d.getDate() + n); return key(d); }
    function monday(k) { const d = parse(k); d.setDate(d.getDate() - (d.getDay() + 6) % 7); return key(d); }
    function roc(k) { return (+k.slice(0, 4) - 1911) + '.' + k.slice(5, 7) + '.' + k.slice(8, 10); }
    function esc(s) { return String(s || '').replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
    function short(lab) { return String(lab).replace(/學?實驗室/, ''); }
    function per(t) { return t === '午' ? '午' : t; }
    window.boardInit = function (id, opt) {
      const box = document.getElementById(id);
      const st = { mode: 'day', start: key(new Date()), end: '', lab: '', kw: '' };
      let U = null;
      box.innerHTML = '<div class="ub-bar">' +
        '<button data-a="prev">◀</button><button data-a="today">今天</button><button data-a="next">▶</button>' +
        '<input type="date" class="ub-date" title="跳到這一天">' +
        '<span class="ub-to" style="display:none">到 <input type="date" class="ub-end"></span>' +
        '<span class="ub-sp"></span>' +
        '<button data-m="day">單日</button><button data-m="week">整週</button><button data-m="month">月曆</button><button data-m="list">清單</button>' +
        '</div><div class="ub-bar"><select class="ub-lab"><option value="">全部實驗室</option></select>' +
        '<input class="ub-kw" placeholder="找：班級、老師、實驗…">' +
        (opt && opt.print ? '<button data-a="print">🖨 列印</button>' : '') + '</div><div class="ub-title"></div><div class="ub-body"></div>';
      const $ = function (c) { return box.querySelector(c); };
      function range() {
        if (st.mode === 'day') return [st.start, 1];
        if (st.mode === 'week') return [monday(st.start), 7];
        if (st.mode === 'month') {
          const d = parse(st.start), first = key(new Date(d.getFullYear(), d.getMonth(), 1));
          return [first, new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate()];
        }
        if (!st.end || st.end < st.start) st.end = add(st.start, 29);
        const n = Math.round((parse(st.end) - parse(st.start)) / 86400000) + 1;
        if (n > 120) { st.end = add(st.start, 119); return [st.start, 120]; }
        return [st.start, n];
      }
      function load() {
        const r = range();
        box.querySelectorAll('[data-m]').forEach(function (b) { b.className = b.getAttribute('data-m') === st.mode ? 'on' : ''; });
        $('.ub-date').value = st.start; $('.ub-to').style.display = st.mode === 'list' ? '' : 'none'; $('.ub-end').value = st.end || '';
        $('.ub-body').innerHTML = '<p style="color:#666">讀取中…</p>';
        google.script.run.withSuccessHandler(function (u) {
          U = u;
          const sel = $('.ub-lab');
          if (sel.options.length === 1) U.rooms.forEach(function (x) { sel.add(new Option(x, x)); });
          sel.value = st.lab;
          draw();
        }).withFailureHandler(function (e) { $('.ub-body').textContent = '讀取失敗：' + (e.message || e); }).usageData(r[0], r[1]);
      }
      function hit(b) {
        if (st.lab && b.lab !== st.lab) return false;
        if (!st.kw) return true;
        return [b.content, b.type, b.cls, b.teacher, b.lab, b.note].join(' ').toLowerCase().indexOf(st.kw.toLowerCase()) >= 0;
      }
      // 這堂課有幾項待準備事項（同一天、同實驗室、節次重疊或沒寫節次）
      function todoN(b) {
        return (U.todos || []).filter(function (t) {
          return t.date === b.date && t.lab === b.lab && (!t.pis.length || t.pis.some(function (i) { return b.pis.indexOf(i) >= 0; }));
        }).length;
      }
      function badge(b) { const n = todoN(b); return n ? ' <span class="ub-td">📝' + n + '</span>' : ''; }
      function cell(b, withLab) {
        return '<div class="ub-b" data-bi="' + U.bookings.indexOf(b) + '" style="background:' + (U.colors[b.type] || '#eee') + '" title="' + esc(b.type + ' ' + b.note) + '"><b>' +
          esc(b.content || b.type) + '</b>' + badge(b) + '<br>' + esc([withLab ? short(b.lab) : '', b.cls, b.teacher].filter(String).join(' ')) + '</div>';
      }
      function isNow(dk, p) { return dk === U.today && U.nowHm >= p.start && U.nowHm < p.end; }
      function dayTable(d, list) {
        const rooms = st.lab ? [st.lab] : U.rooms;
        let h = '<div class="ub-wrap"><table class="ub"><tr><th style="width:70px">節次</th>' + rooms.map(function (r) { return '<th>' + esc(r) + '</th>'; }).join('') + '</tr>';
        U.periods.forEach(function (p, pi) {
          const now = isNow(d.key, p);
          h += '<tr' + (now ? ' class="now"' : '') + '><th>' + esc(p.name) + (now ? ' ●' : '') + '<br><small>' + p.start + '–' + p.end + '</small></th>';
          rooms.forEach(function (r) {
            h += '<td>' + list.filter(function (b) { return b.lab === r && b.date === d.key && b.pis.indexOf(pi) >= 0; }).map(function (b) { return cell(b); }).join('') + '</td>';
          });
          h += '</tr>';
        });
        return h + '</table></div>';
      }
      // 一間實驗室的一週：節次 × 天（和門口課表一樣）
      function weekGrid(days, list) {
        let h = '<div class="ub-wrap"><table class="ub"><tr><th style="width:70px">節次</th>' + days.map(function (d) {
          return '<th class="' + (d.key === U.today ? 'today' : '') + '">' + d.roc.slice(4) + '<br><small>（' + d.wd + '）</small></th>';
        }).join('') + '</tr>';
        U.periods.forEach(function (p, pi) {
          h += '<tr><th>' + esc(p.name) + '<br><small>' + p.start + '–' + p.end + '</small></th>';
          days.forEach(function (d) {
            h += '<td style="' + (isNow(d.key, p) ? 'box-shadow:inset 0 0 0 2px #d93025' : '') + '">' +
              list.filter(function (b) { return b.date === d.key && b.pis.indexOf(pi) >= 0; }).map(function (b) { return cell(b, !st.lab); }).join('') + '</td>';
          });
          h += '</tr>';
        });
        return h + '</table></div>';
      }
      function month(list) {
        const days = U.days, first = parse(days[0].key);
        const weekend = list.some(function (b) { const w = parse(b.date).getDay(); return w === 0 || w === 6; });
        const cols = weekend ? 7 : 5;
        let h = '<div class="ub-wrap"><table class="um"><tr>' + ['一', '二', '三', '四', '五', '六', '日'].slice(0, cols).map(function (w) { return '<th>' + w + '</th>'; }).join('') + '</tr>';
        let k = monday(days[0].key);
        const last = days[days.length - 1].key;
        while (k <= last) {
          h += '<tr>';
          for (let i = 0; i < 7; i++, k = add(k, 1)) {
            if (i >= cols) continue;
            const inMonth = parse(k).getMonth() === first.getMonth();
            if (!inMonth) { h += '<td class="out">' + (+k.slice(8)) + '</td>'; continue; }
            // 放假、考試等同一天多間一樣的合併成一條
            const seen = {};
            const es = list.filter(function (b) { return b.date === k; }).sort(function (a, b) { return (a.pis[0] || 0) - (b.pis[0] || 0) || (a.lab < b.lab ? -1 : 1); })
              .filter(function (b) {
                if (b.type !== '放假' && b.type !== '考試') return true;
                const g = b.type + b.content + b.periodText;
                if (seen[g]) { seen[g].labs.push(short(b.lab)); return false; }
                seen[g] = b; b.labs = [short(b.lab)]; return true;
              });
            h += '<td class="' + (k === U.today ? 'today' : '') + '"><div class="dn" data-k="' + k + '">' + (+k.slice(8)) + '</div>' + es.slice(0, 7).map(function (b) {
              const who = b.labs ? b.labs.join('、') : (st.lab ? '' : short(b.lab));
              return '<div class="e" data-bi="' + U.bookings.indexOf(b) + '" style="background:' + (U.colors[b.type] || '#eee') + '" title="' + esc([per(b.periodText), b.lab, b.content || b.type, b.cls, b.teacher].join(' ')) + '">' +
                esc(per(b.periodText) + ' ' + [who, b.cls || '', b.content || b.type].filter(String).join(' ')) + (b.labs ? '' : badge(b)) + '</div>';
            }).join('') + (es.length > 7 ? '<div class="dn" data-k="' + k + '">…還有 ' + (es.length - 7) + ' 筆</div>' : '') + '</td>';
          }
          h += '</tr>';
        }
        return h + '</table></div><div class="ub-cnt">點日期可以看那一天的詳細課表。</div>';
      }
      function listView(list) {
        if (!list.length) return '<p class="ub-cnt">這段期間沒有符合的登記。</p>';
        let h = '<div class="ub-cnt">共 ' + list.length + ' 筆</div><div class="ub-wrap"><table class="ul"><tr><th>日期</th><th>節次</th><th>實驗室</th><th>類型</th><th>內容</th><th>班級</th><th>老師</th></tr>';
        list.slice().sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.pis[0] || 0) - (b.pis[0] || 0) || (a.lab < b.lab ? -1 : 1); })
          .forEach(function (b) {
            const w = '日一二三四五六'.charAt(parse(b.date).getDay());
            h += '<tr data-bi="' + U.bookings.indexOf(b) + '"><td style="white-space:nowrap">' + roc(b.date).slice(4) + '（' + w + '）</td><td>' + esc(b.periodText === '午' ? '中午' : b.periodText) + '</td><td>' + esc(short(b.lab)) +
              '</td><td><span style="background:' + (U.colors[b.type] || '#eee') + ';padding:1px 6px;border-radius:8px">' + esc(b.type) + '</span></td><td>' + esc(b.content) +
              badge(b) + '</td><td>' + esc(b.cls) + '</td><td>' + esc(b.teacher) + '</td></tr>';
          });
        return h + '</table></div>';
      }
      function draw() {
        if (!U) return;
        const list = U.bookings.filter(hit);
        const d0 = U.days[0], dn = U.days[U.days.length - 1];
        $('.ub-title').textContent = st.mode === 'day' ? d0.roc + '（' + d0.wd + '）' + (d0.key === U.today ? '　今天' : '') :
          st.mode === 'month' ? (+d0.key.slice(0, 4) - 1911) + ' 年 ' + (+d0.key.slice(5, 7)) + ' 月' :
          d0.roc + '（' + d0.wd + '）～' + dn.roc.slice(4) + '（' + dn.wd + '）';
        let h = '';
        if (st.mode === 'day') h = dayTable(d0, list);
        else if (st.mode === 'week') {
          const days = U.days.filter(function (d, i) { return i < 5 || list.some(function (b) { return b.date === d.key; }); });
          if (st.lab || st.kw) h = weekGrid(days, list);
          else days.forEach(function (d) { h += '<div class="ub-day">' + d.roc + '（' + d.wd + '）' + (d.key === U.today ? '　今天' : '') + '</div>' + dayTable(d, list); });
        } else if (st.mode === 'month') h = month(list);
        else h = listView(list);
        h += '<div class="ub-legend">' + Object.keys(U.colors).map(function (t) { return '<span style="background:' + U.colors[t] + '">' + esc(t) + '</span>'; }).join('') + '</div>';
        $('.ub-body').innerHTML = h;
        box.querySelectorAll('.ub-body .dn').forEach(function (e) {
          e.onclick = function () { st.mode = 'day'; st.start = e.getAttribute('data-k'); load(); };
        });
        // 點一堂課：看詳細、新增準備事項、補實驗名稱（有安裝「15_今日與便利」才有）
        if (window.lessonPanel) box.querySelectorAll('.ub-body [data-bi]').forEach(function (e) {
          e.style.cursor = 'pointer';
          e.onclick = function (ev) { ev.stopPropagation(); window.lessonPanel(U.bookings[+e.getAttribute('data-bi')], U, opt || {}, load); };
        });
      }
      function move(n) {
        if (st.mode === 'month') { const d = parse(st.start); st.start = key(new Date(d.getFullYear(), d.getMonth() + n, 1)); }
        else if (st.mode === 'list') { const r = range(); st.start = add(st.start, n * r[1]); st.end = add(st.end, n * r[1]); }
        else st.start = add(st.start, n * (st.mode === 'week' ? 7 : 1));
        load();
      }
      box.querySelectorAll('.ub-bar button').forEach(function (b) {
        b.onclick = function () {
          const a = b.getAttribute('data-a'), m = b.getAttribute('data-m');
          if (m) { st.mode = m; if (m === 'list' && !st.end) st.end = add(st.start, 29); load(); }
          else if (a === 'prev') move(-1);
          else if (a === 'next') move(1);
          else if (a === 'today') { st.start = key(new Date()); if (st.mode === 'list') st.end = add(st.start, 29); load(); }
          else if (a === 'print') window.print();
        };
      });
      $('.ub-date').onchange = function () { if (this.value) { st.start = this.value; if (st.mode === 'list' && st.end < st.start) st.end = add(st.start, 29); load(); } };
      $('.ub-end').onchange = function () { if (this.value) { st.end = this.value; load(); } };
      $('.ub-lab').onchange = function () { st.lab = this.value; draw(); };
      let tm = null;
      $('.ub-kw').oninput = function () { const v = this.value.trim(); clearTimeout(tm); tm = setTimeout(function () { st.kw = v; draw(); }, 250); };
      load();
    };
  })();`;
}

// ---------------------------------------------------------------- 列印：門口海報、本週課表

function printPosterDialog() { labPrintDialog_('poster', '🖨 列印門口海報（A4 橫式）', '每個實驗（實驗課、補做實驗）印一張，放在實驗室門口。'); }
function printLabWeekDialog() { labPrintDialog_('labweek', '🖨 列印本週課表（A4 直式）', '該實驗室這一週每一節的使用情形。'); }

function labPrintDialog_(page, title, hint) {
  const cfg = labConfig_();
  const html = DIALOG_STYLE + `
    <p class="hint">${hint}</p>
    <label>實驗室</label><select id="lab"><option value="">全部六間</option></select>
    <label>週次（選該週任一天）</label><input type="date" id="week">
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" onclick="openPrint({ page: '${page}', lab: document.getElementById('lab').value, week: document.getElementById('week').value }, this)">開啟列印頁</button></div>
    <script>${OPEN_PRINT_JS}
      const D = __DATA__; D.rooms.forEach(function (x) { document.getElementById('lab').add(new Option(x, x)); });
      document.getElementById('week').value = D.today;</script>`;
  showDialog_(html, { rooms: cfg.rooms, today: today_() }, title, 280);
}

function gradeOf_(kitGrade, classes) {
  if (kitGrade) return kitGrade;
  const g = {};
  classes.forEach(function (c) {
    const m = String(c).match(/^(?:高)?([一二三123])/);
    if (m) g['高' + ({ '1': '一', '2': '二', '3': '三' }[m[1]] || m[1])] = true;
  });
  return Object.keys(g).join('、');
}

/** 班級顯示用學校習慣的班名：「213」→「二敬」；已經是「二敬」或認不出來的照原樣。 */
function classLabel_(cls, names) {
  const s = String(cls || '').trim();
  const m = s.match(/^([123])(\d{2})$/);
  if (!m) return s;
  const n = names[Number(m[2]) - 1];
  return n ? '一二三'.charAt(Number(m[1]) - 1) + n : s;
}

/** 班級排序：依年級、班序（仁、義、禮…），認不出來的排後面。 */
function classCompare_(names) {
  const key = function (c) {
    const m = String(c).match(/^(?:高)?([一二三])(.+)$/);
    if (m && names.indexOf(m[2]) >= 0) return ('一二三'.indexOf(m[1]) + 1) * 100 + names.indexOf(m[2]) + 1;
    return /^\d+$/.test(c) ? Number(c) : 99999;
  };
  return function (a, b) { return key(a) - key(b) || naturalCompare_(a, b); };
}

function page_poster(p) {
  const cfg = labConfig_();
  const mon = mondayOf_(dateKey_(p.week) || today_()), sun = addDays_(mon, 6);
  const labs = p.lab ? [p.lab] : cfg.rooms;
  const kits = kits_().map;
  const list = bookings_(mon, sun, cfg).filter(function (b) { return PREP_TYPES.indexOf(b.type) >= 0 && labs.indexOf(b.lab) >= 0; });
  const groups = {}, order = [];
  list.forEach(function (b) {
    const g = b.lab + '\u0001' + b.content;
    if (!groups[g]) { groups[g] = []; order.push(g); }
    groups[g].push(b);
  });
  order.sort(function (a, b) { return labs.indexOf(a.split('\u0001')[0]) - labs.indexOf(b.split('\u0001')[0]) || naturalCompare_(a, b); });
  const uniq = function (a) { return a.filter(function (x, i) { return x && a.indexOf(x) === i; }); };
  let qrs = [];
  const pages = order.map(function (g, i) {
    const bs = groups[g], lab = bs[0].lab, name = bs[0].content, kit = kits[name] || {};
    const dates = uniq(bs.map(function (b) { return b.date; }).sort());
    const classes = uniq(bs.map(function (b) { return b.cls; })).sort(classCompare_(cfg.classNames));
    const teachers = uniq(bs.map(function (b) { return b.teacher; }));
    const grade = gradeOf_(kit.grade, classes);
    if (kit.link) qrs.push({ id: 'qr' + i, text: kit.link });
    return '<section class="poster' + (bs.length > 8 ? ' many' : '') + '"><div class="p-top"><span>' + esc_(cfg.school) + '</span><span>' +
      esc_(rocText_(dates[0]) + (dates.length > 1 ? '～' + rocText_(dates[dates.length - 1]).slice(4) : '')) + '</span></div>' +
      (grade ? '<div class="p-grade">' + esc_(grade) + '</div>' : '') +
      '<div class="p-title">' + esc_(name || '（實驗名稱未填）') + '</div><div class="p-room">📍 ' + esc_(lab) + '</div>' +
      '<div class="p-meta">' + (teachers.length ? '授課教師：' + esc_(teachers.join('、')) : '') +
      (classes.length ? '　｜　班級：' + esc_(classes.join('、')) : '') + '</div>' +
      '<div class="p-sched">' + bs.slice().sort(function (a, b) {
        return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.pis[0] || 0) - (b.pis[0] || 0);
      }).map(function (b) {
        return '<span>' + esc_(rocText_(b.date).slice(4).replace('.', '/') + '（' + weekdayOf_(b.date) + '）' +
          periodLabel_(b.periodText).replace(/ /g, '') + (b.cls ? ' ' + b.cls : '')) + '</span>';
      }).join('') + '</div>' +
      '<div class="p-bottom"><div class="p-safety"><b>⚠ 實驗室安全注意事項</b><ol>' +
      cfg.safety.map(function (x) { return '<li>' + esc_(x) + '</li>'; }).join('') + '</ol></div>' +
      (kit.link ? '<div class="p-qr"><div id="qr' + i + '"></div><small>掃描看實驗講義</small></div>' : '') + '</div></section>';
  });
  const body = pages.join('') || '<p style="padding:20px">這一週（' + rocText_(mon) + ' 起）' + (p.lab ? p.lab : '六間實驗室') +
    '沒有登記「實驗課」或「補做實驗」。</p>';
  const qrJs = qrs.length ? '<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script><script>' +
    JSON.stringify(qrs).replace(/</g, '\\u003c') + '.forEach(function (q) { const el = document.getElementById(q.id);' +
    ' if (window.QRCode) new QRCode(el, { text: q.text, width: 130, height: 130 });' +
    ' else { el.className = "qr-fail"; el.textContent = q.text; } });</script>' : '';   // 連不上產生 QR 的程式時改印網址
  return HtmlService.createHtmlOutput(labShell_('門口海報', body + qrJs, true, POSTER_CSS)).setTitle('門口海報（列印）');
}

const POSTER_CSS = `
  section.poster { height: 186mm; display: flex; flex-direction: column; border: 2.5pt solid #1a73e8; border-radius: 6mm;
    padding: 9mm 12mm; break-after: page; page-break-after: always; margin-bottom: 8mm; background: #fff; }
  section.poster:last-of-type { break-after: auto; page-break-after: auto; }
  .p-top { display: flex; justify-content: space-between; font-size: 13pt; color: #5f6368; }
  .p-grade { margin-top: 6mm; font-size: 26pt; font-weight: bold; color: #1a73e8; }
  .p-title { font-size: 60pt; font-weight: 900; line-height: 1.15; margin-top: 2mm; }
  .p-room { font-size: 30pt; font-weight: bold; margin-top: 4mm; }
  .p-meta { font-size: 15pt; margin-top: 4mm; color: #3c4043; }
  .p-sched { margin-top: 4mm; display: flex; flex-wrap: wrap; gap: 2mm; }
  .p-sched span { border: 1pt solid #1a73e8; color: #174ea6; border-radius: 3mm; padding: 1mm 3mm; font-size: 13pt; }
  .many .p-title { font-size: 46pt; } .many .p-room { font-size: 24pt; margin-top: 2mm; } .many .p-grade { margin-top: 3mm; }
  .many .p-sched { gap: 1.2mm; margin-top: 3mm; } .many .p-sched span { font-size: 10.5pt; padding: 0.4mm 2mm; }
  .many .p-safety { font-size: 10.5pt; padding: 2mm 4mm; }
  .p-bottom { margin-top: auto; display: flex; gap: 8mm; align-items: flex-end; }
  .p-safety { flex: 1; background: #fef7e0; border-left: 5pt solid #f9ab00; padding: 3mm 5mm; font-size: 12pt; }
  .p-safety ol { margin: 2mm 0 0; padding-left: 6mm; } .p-safety li { margin: 0.6mm 0; }
  .p-qr { text-align: center; } .qr-fail { width: 45mm; font-size: 9pt; word-break: break-all; } .p-qr small { display: block; font-size: 10pt; color: #5f6368; margin-top: 1mm; }
`;

function page_labweek(p) {
  const cfg = labConfig_();
  const mon = mondayOf_(dateKey_(p.week) || today_()), sun = addDays_(mon, 6);
  const labs = p.lab ? [p.lab] : cfg.rooms;
  const all = bookings_(mon, sun, cfg);
  const pages = labs.map(function (lab) {
    const list = all.filter(function (b) { return b.lab === lab; });
    let days = [];
    for (let i = 0; i < 7; i++) days.push(addDays_(mon, i));
    days = days.filter(function (k, i) { return i < 5 || list.some(function (b) { return b.date === k; }); });
    const w = Math.floor(84 / days.length);
    let h = '<table class="wk"><colgroup><col style="width:16%">' + days.map(function () { return '<col style="width:' + w + '%">'; }).join('') +
      '</colgroup><thead><tr><th>節次</th>' + days.map(function (k) {
        return '<th>' + esc_(rocText_(k).slice(4)) + '<br>（' + weekdayOf_(k) + '）</th>';
      }).join('') + '</tr></thead><tbody>';
    cfg.periods.forEach(function (per, pi) {
      h += '<tr' + (per.short === '午' ? ' class="noon"' : '') + '><th>' + esc_(per.name) + '<br><small>' + esc_(per.start + '–' + per.end) + '</small></th>';
      days.forEach(function (k) {
        const bs = list.filter(function (b) { return b.date === k && b.pis.indexOf(pi) >= 0; });
        h += '<td' + (bs.length ? ' style="background:' + (cfg.colors[bs[0].type] || '#eee') + '"' : '') + '>' +
          bs.map(function (b) {
            return '<div class="c1">' + esc_(b.content || b.type) + '</div><div class="c2">' + esc_(b.cls) + '</div><div class="c3">' +
              esc_(b.teacher) + '</div>' + (b.content && b.type !== '實驗課' ? '<div class="c4">' + esc_(b.type) + '</div>' : '');
          }).join('<hr>') + '</td>';
      });
      h += '</tr>';
    });
    h += '</tbody></table>';
    const used = Object.keys(cfg.colors).filter(function (t) { return list.some(function (b) { return b.type === t; }); });
    return '<section class="lw"><div class="lw-head"><div class="lw-school">' + esc_(cfg.school) + '</div><div class="lw-title">' + esc_(lab) +
      '　本週使用課表</div><div class="lw-date">' + esc_(rocText_(mon) + '（一）～' + rocText_(addDays_(mon, 4)).slice(4) + '（五）') + '</div></div>' + h +
      '<div class="lw-foot">' + used.map(function (t) { return '<span style="background:' + cfg.colors[t] + '">' + esc_(t) + '</span>'; }).join('') +
      '<span class="upd">更新：' + esc_(rocText_(today_())) + '</span></div></section>';
  });
  return HtmlService.createHtmlOutput(labShell_('本週課表', pages.join(''), false, LABWEEK_CSS)).setTitle('本週課表（列印）');
}

const LABWEEK_CSS = `
  section.lw { break-after: page; page-break-after: always; margin-bottom: 8mm; }
  section.lw:last-of-type { break-after: auto; page-break-after: auto; }
  .lw-head { text-align: center; margin-bottom: 4mm; } .lw-school { font-size: 12pt; color: #5f6368; }
  .lw-title { font-size: 26pt; font-weight: 900; margin: 1mm 0; } .lw-date { font-size: 14pt; }
  table.wk { width: 100%; border-collapse: collapse; table-layout: fixed; }
  table.wk th, table.wk td { border: 1pt solid #000; text-align: center; vertical-align: middle; padding: 1.5mm 1mm; }
  table.wk thead th { background: #e8eaed; font-size: 12pt; }
  table.wk tbody th { background: #f1f3f4; font-size: 12pt; } table.wk tbody th small { font-weight: normal; font-size: 8.5pt; }
  table.wk tbody tr { height: 27mm; } table.wk tbody tr.noon { height: 17mm; }
  .c1 { font-weight: bold; font-size: 12pt; line-height: 1.2; } .c2 { font-size: 11pt; } .c3 { font-size: 10pt; color: #3c4043; }
  .c4 { font-size: 8.5pt; color: #5f6368; } table.wk hr { border: none; border-top: 0.6pt dashed #666; margin: 1mm 0; }
  .lw-foot { margin-top: 3mm; font-size: 10pt; } .lw-foot span { display: inline-block; padding: 1px 8px; border-radius: 8px; margin-right: 4px; }
  .lw-foot .upd { float: right; color: #5f6368; }
`;

function labShell_(title, body, landscape, css) {
  return `<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><title>${esc_(title)}</title>
<style>
  @page { size: A4 ${landscape ? 'landscape' : 'portrait'}; margin: 10mm; }
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; box-sizing: border-box; }
  body { margin: 0; color: #000; background: #e9ecef; font-family: "Noto Sans TC", "Microsoft JhengHei", "PingFang TC", sans-serif; }
  .bar { position: sticky; top: 0; z-index: 9; background: #fff8e1; border-bottom: 1px solid #e0c97a; padding: 10px 16px; font-size: 14px;
    display: flex; gap: 16px; align-items: center; flex-wrap: wrap; }
  .bar button { font-size: 16px; padding: 8px 20px; background: #1a73e8; color: #fff; border: none; border-radius: 6px; cursor: pointer; }
  .paper { width: ${landscape ? '277mm' : '190mm'}; margin: 16px auto; }
  ${css}
  @media print { body { background: #fff; } .bar { display: none; } .paper { width: auto; margin: 0; } section { margin-bottom: 0 !important; } }
</style></head><body>
<div class="bar"><button onclick="window.print()">🖨 列印</button>
  <span>列印設定：紙張 A4、版面配置「<b>${landscape ? '橫向' : '直向'}</b>」；「更多設定」取消「頁首和頁尾」、勾選「背景圖形」。</span></div>
<div class="paper">${body}</div></body></html>`;
}
