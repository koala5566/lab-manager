/**
 * 實驗室藥品器材耗材管理系統 — 匯入實驗室課表（老師的 Excel）
 *
 * 老師的課表格式：每一週一塊，上面一列是週一～週五的日期，下面第 1～7 節；
 * 格子寫「二敬-仲文」（高二敬班、仲文老師）、「多元選修」「第一次期中考」「教師節放假」…，
 * 合併儲存格代表連續好幾節。
 *
 * 用法：
 *   1. 「檔案 → 匯入 → 上傳」老師的 Excel →「插入新工作表」（工作表名稱要有「化一」「化二」「生一」…）
 *   2. 「🗓 實驗室使用 → 匯入實驗室課表」→ 勾選工作表、確認對應的實驗室 → 預覽 → 匯入
 * 重新匯入同一張工作表時，會先刪掉上次從它匯入、且在「從哪天開始」之後的登記，再寫入新的；
 * 自己用「登記使用」登記的不會被動到。已經補上的實驗名稱、組數、狀態、備註會保留。
 * 需要「12_實驗室使用」。
 */

const IMPORT_COL = '匯入來源';
const IMPORT_NEW_TYPES = [['多元選修', '#FDE2F3'], ['放假', '#DADCE0']];

// ---------------------------------------------------------------- 對話框

function importScheduleDialog() {
  const cfg = labConfig_();
  const ss = SpreadsheetApp.getActive();
  const sheets = [];
  ss.getSheets().forEach(function (sh) {
    const a = sh.getRange(1, 1, Math.min(30, sh.getMaxRows()), 1).getValues();
    if (!a.some(function (r) { return String(r[0]).trim() === '週次'; })) return;
    sheets.push({ name: sh.getName(), lab: guessLab_(sh.getName(), cfg.rooms) });
  });
  if (!sheets.length) {
    SpreadsheetApp.getUi().alert('找不到老師的課表工作表',
      '請先「檔案 → 匯入 → 上傳」老師的 Excel，選「插入新工作表」。\n（工作表的 A 欄要有「週次」，名稱最好有「化一」「化二」這類簡稱。）',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const html = DIALOG_STYLE + `
    <style>
      body { font-size: 14px; } table.s { border-collapse: collapse; width: 100%; margin: 6px 0; }
      table.s td, table.s th { border: 1px solid #dadce0; padding: 4px 6px; font-size: 13px; text-align: left; vertical-align: top; }
      table.s th { background: #f1f3f4; } .warn { background: #fef7e0; padding: 8px 10px; border-radius: 6px; margin: 8px 0; white-space: pre-line; font-size: 13px; }
      .cnt span { display: inline-block; padding: 2px 8px; border-radius: 10px; margin: 2px; font-size: 12px; }
      #prev input { font-size: 13px; padding: 3px 5px; width: 100%; }
    </style>
    <p class="hint">勾選要匯入的工作表，確認對應的實驗室。只匯入「從哪天開始」以後的日期（之前的保留不動）。</p>
    <table class="s"><tr><th style="width:40px">匯入</th><th>工作表</th><th>實驗室</th></tr><tbody id="list"></tbody></table>
    <label>從哪天開始</label><input type="date" id="from">
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="pv" onclick="preview()">預覽</button></div>
    <div id="prev"></div>
    <script>
      const D = __DATA__; let P = null;
      const $ = function (id) { return document.getElementById(id); };
      function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
      D.sheets.forEach(function (s, i) {
        const tr = document.createElement('tr');
        tr.innerHTML = '<td><input type="checkbox" id="c' + i + '"' + (s.lab ? ' checked' : '') + '></td><td>' + esc(s.name) +
          '</td><td><select id="l' + i + '"><option value="">（選實驗室）</option>' +
          D.rooms.map(function (r) { return '<option' + (r === s.lab ? ' selected' : '') + '>' + esc(r) + '</option>'; }).join('') + '</select></td>';
        $('list').appendChild(tr);
      });
      $('from').value = D.from;
      function opts() {
        const pick = [];
        D.sheets.forEach(function (s, i) { if ($('c' + i).checked) pick.push({ sheet: s.name, lab: $('l' + i).value }); });
        return { sheets: pick, from: $('from').value };
      }
      function preview() {
        const o = opts();
        if (!o.sheets.length) { alert('請至少勾選一張工作表。'); return; }
        if (o.sheets.some(function (s) { return !s.lab; })) { alert('請幫勾選的工作表選對應的實驗室。'); return; }
        $('pv').disabled = true; $('pv').textContent = '讀取中…'; $('prev').innerHTML = '';
        google.script.run.withSuccessHandler(show).withFailureHandler(function (e) { alert(e.message); $('pv').disabled = false; $('pv').textContent = '預覽'; })
          .importPreview(o);
      }
      function show(res) {
        P = res; $('pv').disabled = false; $('pv').textContent = '重新預覽';
        let h = '<h3 style="margin:14px 0 4px">預覽：共 ' + res.total + ' 筆</h3><div class="cnt">' +
          Object.keys(res.byType).map(function (t) { return '<span style="background:' + (res.colors[t] || '#eee') + '">' + esc(t) + ' ' + res.byType[t] + '</span>'; }).join('') + '</div>';
        if (res.replace) h += '<p class="hint">會先刪掉上次從這些工作表匯入的 ' + res.replace + ' 筆（' + esc(res.fromRoc) + ' 以後），再寫入新的。</p>';
        if (res.warnings.length) h += '<div class="warn">⚠ 請確認：\\n' + res.warnings.map(esc).join('\\n') + '</div>';
        if (res.weeks.length) {
          h += '<p style="margin:10px 0 4px"><b>實驗名稱</b>（可以先空白，之後在「實驗排程」補，或下次匯入時再填；已經填過的會保留）</p>' +
            '<table class="s"><tr><th>週次</th><th>實驗室</th><th>日期</th><th>班級</th><th style="width:34%">實驗名稱</th></tr>';
          res.weeks.forEach(function (w, i) {
            h += '<tr><td>' + esc(w.week) + '</td><td>' + esc(w.lab) + '</td><td>' + esc(w.range) + '</td><td>' + esc(w.classes) +
              '</td><td><input id="w' + i + '" list="kits" value="' + esc(w.name) + '"></td></tr>';
          });
          h += '</table><datalist id="kits">' + res.kits.map(function (k) { return '<option value="' + esc(k) + '">'; }).join('') + '</datalist>';
        }
        h += '<div class="btns"><button class="primary" id="go" onclick="apply()">確定匯入 ' + res.total + ' 筆</button></div>';
        $('prev').innerHTML = h;
        $('prev').scrollIntoView();
      }
      function apply() {
        const o = opts(); o.names = {};
        P.weeks.forEach(function (w, i) { const v = $('w' + i).value.trim(); if (v) o.names[w.key] = v; });
        const b = $('go'); b.disabled = true; b.textContent = '匯入中…';
        google.script.run.withSuccessHandler(showDone).withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '確定匯入'; })
          .importApply(o);
      }
    </script>`;
  const json = JSON.stringify({ sheets: sheets, rooms: cfg.rooms, from: mondayOf_(today_()) }).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html.replace('__DATA__', function () { return json; }))
    .setWidth(820).setHeight(720), '📥 匯入實驗室課表');
}

/** 工作表名稱「115-1(化一)」→ 化學實驗室一 */
function guessLab_(name, rooms) {
  const m = String(name).match(/([化生])(?:學)?(?:實驗室)?([一二三123])/);
  if (!m) return '';
  const n = { '1': '一', '2': '二', '3': '三' }[m[2]] || m[2];
  const lab = (m[1] === '化' ? '化學' : '生物') + '實驗室' + n;
  return rooms.indexOf(lab) >= 0 ? lab : '';
}

// ---------------------------------------------------------------- 讀老師的課表


/** 格子文字 → { type, content, cls, teacher }，班名認不出來時 warn 有內容 */
function classifyCell_(text, cfg) {
  const t = String(text).replace(/\s+/g, '');
  const m = t.match(/^([一二三123])([^\-－—–~～]{1,3})[\-－—–~～](.+)$/);
  if (m) {
    // 班級照學校習慣存「二敬」
    const g = { '1': '一', '2': '二', '3': '三' }[m[1]] || m[1];
    const ok = cfg.classNames.indexOf(m[2]) >= 0;
    return { type: '實驗課', content: '', cls: g + m[2], teacher: m[3], warn: ok ? '' : '班名「' + m[2] + '」不在設定的班名裡，請確認' };
  }
  if (/多元選修/.test(t)) return { type: '多元選修', content: t };
  if (/研究方法|專題/.test(t)) return { type: '專題研究', content: t };
  if (/補做/.test(t)) return { type: '補做實驗', content: t };
  if (/考|學測|統測|會考/.test(t)) return { type: '考試', content: t };
  if (/放假|補假|停課|颱風/.test(t)) return { type: '放假', content: t };
  if (/研習/.test(t)) return { type: '研習', content: t };
  if (/社/.test(t)) return { type: '社團', content: t };
  return { type: '其他', content: t };
}

/**
 * 讀一張老師課表工作表 → [{date, periodText, type, content, cls, teacher, week, cell}]，並把疑點放進 warnings。
 * 找「日期」格當每一天的欄標題，往下讀第 1～7 節（遇到下一列日期就停），合併儲存格＝連續好幾節。
 */
function parseTeacherSheet_(sh, cfg, warnings) {
  const range = sh.getDataRange();
  const v = range.getValues();
  const span = {};
  range.getMergedRanges().forEach(function (m) { span[(m.getRow() - 1) + ',' + (m.getColumn() - 1)] = m.getNumRows(); });
  const shorts = cfg.periods.map(function (p) { return p.short; });
  const isDate = function (x) { return x instanceof Date && !isNaN(x); };
  const out = [];
  let prevKey = '', fixedYear = '';
  const name = sh.getName();

  for (let r = 0; r < v.length; r++) {
    for (let c = 0; c < v[r].length; c++) {
      if (!isDate(v[r][c])) continue;
      // 日期（跨年打錯年份時自動修正：比前一個日期早很多就 +1 年）
      let key = dateKey_(v[r][c]);
      if (prevKey && key < addDays_(prevKey, -120)) {
        const orig = key;
        key = (Number(key.slice(0, 4)) + 1) + key.slice(4);
        if (!fixedYear) fixedYear = orig.replace(/-/g, '/') + ' → ' + key.replace(/-/g, '/');
      }
      prevKey = key;
      // 這一天的節次欄：往左找第一個不是日期的欄
      let L = c - 1;
      while (L >= 0 && isDate(v[r][L])) L--;
      if (L < 0) continue;
      let week = '';
      for (let rr = r + 1; rr < Math.min(v.length, r + 14); rr++) {
        if (isDate(v[rr][c])) break;
        const cell = String(v[rr][c]).trim();
        if (/^星期/.test(cell)) { week = String(v[rr][L]).trim(); continue; }
        const label = String(v[rr][L]).trim().replace(/\.0$/, '');
        const pi = /午/.test(label) ? shorts.indexOf('午') : shorts.indexOf(label.replace(/[^0-9]/g, ''));
        if (pi < 0) { if (label === '' && cell === '') break; continue; }
        if (!cell) continue;
        const n = span[rr + ',' + c] || 1;
        const lastLabel = String((v[rr + n - 1] || [])[L]).trim().replace(/\.0$/, '');
        const periodText = n > 1 && lastLabel ? shorts[pi] + '-' + (/午/.test(lastLabel) ? '午' : lastLabel.replace(/[^0-9]/g, '')) : shorts[pi];
        const x = classifyCell_(cell, cfg);
        if (x.warn) warnings.push(name + ' ' + rocText_(key) + ' 第 ' + periodText + ' 節「' + cell + '」：' + x.warn);
        out.push({ date: key, periodText: periodText, type: x.type, content: x.content, cls: x.cls || '', teacher: x.teacher || '',
          week: week, cell: cell });
      }
    }
  }
  if (fixedYear) warnings.push(name + '：後面幾週的日期年份比前面早，已當成隔年（' + fixedYear + ' 起）。');
  return out;
}

/** 依選項讀出要匯入的列，並算出會被取代的舊列 */
function importPlan_(o) {
  const cfg = labConfig_();
  const from = dateKey_(o.from) || '0000-00-00';
  const ss = SpreadsheetApp.getActive();
  const warnings = [];
  const rows = [];
  (o.sheets || []).forEach(function (s) {
    const sh = ss.getSheetByName(s.sheet);
    if (!sh) throw new Error('找不到工作表「' + s.sheet + '」。');
    if (cfg.rooms.indexOf(s.lab) < 0) throw new Error('「' + s.lab + '」不在設定的上課教室裡。');
    const list = parseTeacherSheet_(sh, cfg, warnings);
    if (!list.length) warnings.push(s.sheet + '：沒有讀到任何登記，請確認格式（A 欄有「週次」、上面一列是日期）。');
    list.forEach(function (x) {
      if (x.date < from) return;
      x.lab = s.lab; x.source = s.sheet;
      rows.push(x);
    });
  });
  // 舊的：同一個來源、日期在 from 以後
  const t = importTable_();
  const c = t.col;
  const srcs = (o.sheets || []).map(function (s) { return s.sheet; });
  const old = [], keep = [];
  t.rows.forEach(function (r) {
    if (r.every(function (x) { return x === '' || x === null; })) return;
    const k = dateKey_(r[c['日期']]);
    if (srcs.indexOf(String(r[c[IMPORT_COL]]).trim()) >= 0 && k >= from) old.push(r); else keep.push(r);
  });
  // 保留已經填的實驗名稱、組數、狀態、備註（同實驗室、日期、節次、班級）
  const kept = {};
  old.forEach(function (r) {
    const key = [r[c['教室']], dateKey_(r[c['日期']]), String(r[c['節次']]).trim(), String(r[c['班級']]).trim()].join('|');
    kept[key] = { name: String(r[c['實驗名稱']]).trim(), groups: r[c['組數']], status: String(r[c['狀態']]).trim(), note: String(r[c['備註']]).trim() };
  });
  rows.forEach(function (x) {
    const k = kept[[x.lab, x.date, x.periodText, x.cls].join('|')];
    x.kept = k || null;
    if (k && k.name && x.type === '實驗課') x.content = k.name;
  });
  // 和自己手動登記的撞堂
  keep.forEach(function (r) {
    const st = String(r[c['狀態']]).trim();
    if (st === '取消' || String(r[c[IMPORT_COL]]).trim()) return;
    const k = dateKey_(r[c['日期']]), lab = String(r[c['教室']]).trim();
    const pis = parsePeriods_(r[c['節次']], cfg.periods);
    rows.forEach(function (x) {
      if (x.lab !== lab || x.date !== k) return;
      const xp = parsePeriods_(x.periodText, cfg.periods);
      if (xp.some(function (i) { return pis.indexOf(i) >= 0; })) {
        warnings.push('撞堂：' + lab + ' ' + rocText_(k) + ' 第 ' + x.periodText + ' 節「' + x.cell + '」和已登記的「' +
          [r[c['實驗名稱']], r[c['班級']], r[c['教師']]].filter(String).join(' ') + '」重疊');
      }
    });
  });
  return { cfg: cfg, rows: rows, old: old, keep: keep, t: t, warnings: warnings, from: from };
}

function importTable_() {
  if (!SpreadsheetApp.getActive().getSheetByName(EXP_SHEET)) setupExperimentSheet_();
  return ensureSheet_(EXP_SHEET, EXP_COLS.concat([IMPORT_COL]));
}

/** 預覽（不寫入） */
function importPreview(o) {
  const p = importPlan_(o);
  const byType = {};
  p.rows.forEach(function (x) { byType[x.type] = (byType[x.type] || 0) + 1; });
  // 實驗課依 實驗室＋週 分組，讓使用者填實驗名稱
  const weeks = {}, order = [];
  p.rows.forEach(function (x) {
    if (x.type !== '實驗課') return;
    const key = x.lab + '|' + mondayOf_(x.date);
    if (!weeks[key]) { weeks[key] = { key: key, lab: x.lab, week: x.week ? '第 ' + x.week + ' 週' : '', dates: [], classes: [], names: [] }; order.push(key); }
    const w = weeks[key];
    w.dates.push(x.date);
    if (w.classes.indexOf(x.cls) < 0) w.classes.push(x.cls);
    if (x.content && w.names.indexOf(x.content) < 0) w.names.push(x.content);
  });
  order.sort();
  const colors = {};
  Object.keys(byType).forEach(function (t) { colors[t] = p.cfg.colors[t] || (IMPORT_NEW_TYPES.filter(function (n) { return n[0] === t; })[0] || [])[1]; });
  return {
    total: p.rows.length, byType: byType, colors: colors, replace: p.old.length, fromRoc: rocText_(p.from),
    warnings: p.warnings.slice(0, 40).concat(p.warnings.length > 40 ? ['…還有 ' + (p.warnings.length - 40) + ' 項'] : []),
    kits: kits_().order,
    weeks: order.map(function (k) {
      const w = weeks[k];
      w.dates.sort();
      const classes = w.classes.slice().sort(classCompare_(p.cfg.classNames));
      return { key: k, lab: w.lab, week: w.week,
        range: rocText_(w.dates[0]).slice(4) + (w.dates[w.dates.length - 1] !== w.dates[0] ? '～' + rocText_(w.dates[w.dates.length - 1]).slice(4) : ''),
        classes: classes.length > 8 ? classes.slice(0, 8).join('、') + '…共 ' + classes.length + ' 班' : classes.join('、'),
        name: w.names.length === 1 ? w.names[0] : '' };
    }),
  };
}

/** 寫入「實驗排程」 */
function importApply(o) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const p = importPlan_(o);
    ensureLabTypes_(p.rows.map(function (x) { return x.type; }));
    const t = p.t, c = t.col, w = t.header.length;
    const names = o.names || {};
    const fresh = p.rows.map(function (x) {
      const row = new Array(w).fill('');
      const k = x.kept || {};
      let content = x.content;
      if (x.type === '實驗課') content = names[x.lab + '|' + mondayOf_(x.date)] || content;
      const prep = PREP_TYPES.indexOf(x.type) >= 0;
      const set = function (col, val) { if (col in c) row[c[col]] = val; };
      set('日期', dateValue_(x.date)); set('節次', x.periodText); set('班級', x.cls); set('教師', x.teacher);
      set('教室', x.lab); set('實驗名稱', content); set('組數', k.groups || '');
      set('狀態', k.status || (prep ? '待準備' : '')); set('備註', k.note || ''); set('用途類型', x.type); set(IMPORT_COL, x.source);
      return row;
    });
    const all = p.keep.concat(fresh);
    all.sort(function (a, b) {
      const x = dateKey_(a[c['日期']]) || '9999', y = dateKey_(b[c['日期']]) || '9999';
      return x < y ? -1 : x > y ? 1 : naturalCompare_(a[c['教室']], b[c['教室']]) || naturalCompare_(a[c['節次']], b[c['節次']]);
    });
    const sh = t.sheet;
    if (all.length + 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), all.length + 51 - sh.getMaxRows());
    sh.getRange(2, 1, sh.getMaxRows() - 1, w).clearContent();
    if (all.length) sh.getRange(2, 1, all.length, w).setValues(all);
    const noName = fresh.filter(function (r) { return r[c['用途類型']] === '實驗課' && !String(r[c['實驗名稱']]).trim(); }).length;
    return '已匯入 ' + fresh.length + ' 筆到「實驗排程」' + (p.old.length ? '（取代上次匯入的 ' + p.old.length + ' 筆）' : '') + '。' +
      (noName ? '\n\n還有 ' + noName + ' 筆實驗課沒有實驗名稱：問到老師後，在「實驗排程」的「實驗名稱」欄補上，或再匯入一次時填。' : '') +
      (p.warnings.length ? '\n\n⚠ 有 ' + p.warnings.length + ' 個疑點（預覽時列出的），請再確認。' : '');
  } finally {
    lock.releaseLock();
  }
}

/** 設定的用途類型缺「多元選修」「放假」等匯入會用到的類型時補上。 */
function ensureLabTypes_(types) {
  const s = getSettings_();
  const tc = s.header.indexOf('用途類型');
  if (tc < 0) return;
  const have = s.lists['用途類型'] || [];
  const need = IMPORT_NEW_TYPES.filter(function (n) { return types.indexOf(n[0]) >= 0 && have.indexOf(n[0]) < 0; });
  if (!need.length) return;
  const col = s.sheet.getRange(1, tc + 1, s.sheet.getMaxRows(), 1).getValues();
  let last = 0;
  col.forEach(function (v, i) { if (String(v[0]).trim()) last = i; });
  need.forEach(function (n, i) {
    const r = last + 2 + i;
    s.sheet.getRange(r, tc + 1, 1, 2).setValues([n]);
    s.sheet.getRange(r, tc + 2).setBackground(n[1]);
  });
  setupExperimentSheet_();   // 「用途類型」下拉選單也一起更新
}
