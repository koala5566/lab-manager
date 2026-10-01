/**
 * 實驗室藥品器材耗材管理系統 — 進階管理（資料變多、開放同仁時用）
 *
 * 🔧 維護
 *   ・📦 封存舊資料：把某天以前的實驗排程、已完成的準備事項、已結案的請購／借用，搬到「…_封存」工作表，
 *     平常讀的表變小、網頁和提醒信變快。搬之前會自動備份。
 *   ・🏷 教室改名（舊資料一起改）：設定、實驗排程、準備事項（可選：品項的存放教室）一次改好，舊登記不會對不起來。
 *   ・🔒 嚴格保護／🔓 警告模式：開放同仁編輯時，系統工作表改成「只有我能改」；平常用警告模式（會先問、可以改）。
 * 🖨 列印
 *   ・櫃子 QR Code 貼紙：手機相機掃描 → 打開網頁、列出這一櫃的品項，可以直接盤點這一櫃。
 * 網頁
 *   ・moveBooking：電腦／平板課表拖曳換時段（同一堂課的準備事項跟著移）。
 *   ・appCabItems：掃 QR Code 看一櫃的品項。
 * 需要「01」～「18」，QR Code 貼紙另外需要「20_QR產生器」。
 */

// ---------------------------------------------------------------- 📦 封存舊資料

const ARCHIVE_SUFFIX = '_封存';

/** 各表的封存條件：r＝一列，c＝欄位，cut＝這天以前（不含） */
function archiveRules_() {
  const before = function (v, cut) { const k = dateKey_(v); return !!k && k < cut; };
  const st = function (r, c) { return String(r[c['狀態']] == null ? '' : r[c['狀態']]).trim(); };
  return [
    { name: typeof EXP_SHEET === 'string' ? EXP_SHEET : '實驗排程', label: '實驗排程（這天以前的所有登記）',
      test: function (r, c, cut) { return before(r[c['日期']], cut); } },
    { name: typeof TODO_SHEET === 'string' ? TODO_SHEET : '準備事項', label: '準備事項（已準備／已歸還／取消；還沒準備的不搬）',
      test: function (r, c, cut) {
        const s = st(r, c) || '待準備';
        if (s === '待準備') return false;
        return '需要日期' in c && dateKey_(r[c['需要日期']]) ? before(r[c['需要日期']], cut) : before(String(r[c['完成時間']] || '').slice(0, 10), cut);
      } },
    { name: typeof REQ_SHEET === 'string' ? REQ_SHEET : '請購清單', label: '請購清單（已到貨／取消）',
      test: function (r, c, cut) { return (st(r, c) === '已到貨' || st(r, c) === '取消') && before(r[c['登記日期']], cut); } },
    { name: typeof LOAN_SHEET === 'string' ? LOAN_SHEET : '借用紀錄', label: '借用紀錄（已歸還）',
      test: function (r, c, cut) { return st(r, c) === '已歸還' && before(r[c['借出日期']], cut); } },
  ];
}

function archivePlan_(cut) {
  const ss = SpreadsheetApp.getActive();
  return archiveRules_().map(function (rule) {
    const sh = ss.getSheetByName(rule.name);
    if (!sh || sh.getLastRow() < 2) return { rule: rule, t: null, move: [], keep: [] };
    const t = getTable_(rule.name), move = [], keep = [];
    t.rows.forEach(function (r) {
      if (r.every(function (x) { return x === '' || x === null; })) return;
      (rule.test(r, t.col, cut) ? move : keep).push(r);
    });
    return { rule: rule, t: t, move: move, keep: keep };
  });
}

function archiveCheckCut_(cut) {
  const k = dateKey_(cut);
  if (!k) throw new Error('請選日期。');
  if (k > addDays_(today_(), -30)) throw new Error('只能封存 30 天以前的資料（避免把最近還會用到的搬走）。');
  return k;
}

function archiveDialog() {
  const d = new Date();
  d.setMonth(d.getMonth() - 6);
  const html = DIALOG_STYLE + `
    <p>把<b>這一天以前</b>的舊資料搬到「…${ARCHIVE_SUFFIX}」工作表（同一個檔案裡），平常讀的表變小，網頁和提醒信會比較快。</p>
    <label>這天以前的搬走（建議：上學期開學日）</label><input type="date" id="cut" onchange="pv()">
    <div id="pv" class="hint" style="margin-top:12px;white-space:pre-line">計算中…</div>
    <p class="hint">・搬之前會自動做一份永久備份。<br>・封存的資料還在，打開「…${ARCHIVE_SUFFIX}」工作表就看得到；網頁課表往前翻就看不到了。<br>・盤點紀錄是統計耗用量要用的，不會搬。</p>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()" disabled>封存</button></div>
    <script>
      const D = __DATA__; document.getElementById('cut').value = D.cut;
      function pv() {
        const el = document.getElementById('pv'), b = document.getElementById('go');
        el.textContent = '計算中…'; b.disabled = true;
        google.script.run.withSuccessHandler(function (r) {
          el.textContent = r.lines.join('\\n'); b.disabled = !r.total;
        }).withFailureHandler(function (e) { el.textContent = '⚠ ' + e.message; }).archivePreview(document.getElementById('cut').value);
      }
      function go() {
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '封存中…';
        google.script.run.withSuccessHandler(showDone).withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '封存'; })
          .archiveApply(document.getElementById('cut').value);
      }
      pv();
    </script>`;
  showDialog_(html, { cut: Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') }, '📦 封存舊資料', 470);
}

function archivePreview(cut) {
  const k = archiveCheckCut_(cut);
  const plan = archivePlan_(k);
  const total = plan.reduce(function (s, p) { return s + p.move.length; }, 0);
  return { total: total, lines: [rocText_(k) + ' 以前，要搬走：'].concat(plan.map(function (p) {
    return '・' + p.rule.label + '：' + p.move.length + ' 筆（留下 ' + p.keep.length + ' 筆）';
  })).concat(total ? [] : ['', '沒有要搬的資料。']) };
}

function archiveApply(cut) {
  const k = archiveCheckCut_(cut);
  try { backupNow('封存前', true, true); } catch (e) { throw new Error('備份沒有成功，先不封存：' + e.message); }
  return withLock_(function () {
    const ss = SpreadsheetApp.getActive();
    const plan = archivePlan_(k);
    const done = [];
    plan.forEach(function (p) {
      if (!p.move.length) return;
      const t = p.t, name = p.rule.name + ARCHIVE_SUFFIX;
      // 封存表：欄位和原表一樣（原表後來多的欄位也補上）
      let dst = ss.getSheetByName(name);
      if (!dst) {
        dst = ss.insertSheet(name);
        dst.getRange(1, 1, 1, t.header.length).setValues([t.header]).setFontWeight('bold').setBackground('#E8EAED');
        dst.setFrozenRows(1); dst.setTabColor('#9AA0A6');
      }
      const dh = dst.getRange(1, 1, 1, Math.max(1, dst.getLastColumn())).getValues()[0].map(function (h) { return String(h).trim(); });
      t.header.forEach(function (h) { if (h && dh.indexOf(h) < 0) { dst.getRange(1, dh.length + 1).setValue(h); dh.push(h); } });
      const rows = p.move.map(function (r) {
        const out = new Array(dh.length).fill('');
        t.header.forEach(function (h, j) { const x = dh.indexOf(h); if (h && x >= 0) out[x] = r[j]; });
        return out.map(safeCell_);
      });
      const start = Math.max(2, dst.getLastRow() + 1);
      if (start + rows.length - 1 > dst.getMaxRows()) dst.insertRowsAfter(dst.getMaxRows(), start + rows.length - 1 - dst.getMaxRows() + 10);
      if ('節次' in t.col && typeof fixPeriodCol_ === 'function') fixPeriodCol_(dst, dh.indexOf('節次') + 1);
      dst.getRange(start, 1, rows.length, dh.length).setValues(rows);
      // 原表：留下的寫回去（格式、下拉選單都還在）
      const sh = t.sheet, w = t.header.length;
      sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), w).clearContent();
      if (p.keep.length) sh.getRange(2, 1, p.keep.length, w).setValues(p.keep.map(function (r) { return r.slice(0, w).map(safeCell_); }));
      done.push('・' + p.rule.name + '：搬了 ' + p.move.length + ' 筆 → 「' + name + '」，留下 ' + p.keep.length + ' 筆');
    });
    if (readMemo_) readMemo_ = {};
    if (typeof recordBaseline_ === 'function') recordBaseline_();   // 不要被刪除偵測當成誤刪
    return done.length ? '封存完成（' + rocText_(k) + ' 以前）：\n' + done.join('\n') + '\n\n已先自動做一份永久備份。' : '沒有要搬的資料。';
  });
}

// ---------------------------------------------------------------- 🏷 教室改名（舊資料一起改）

/** 設定的上課教室，加上資料裡出現、但不在設定裡的教室（改名後對不上的舊資料） */
function roomInfo_() {
  const cfg = labConfig_();
  const used = {};
  [[typeof EXP_SHEET === 'string' ? EXP_SHEET : '實驗排程', '教室'], [typeof TODO_SHEET === 'string' ? TODO_SHEET : '準備事項', '實驗室']].forEach(function (x) {
    const sh = SpreadsheetApp.getActive().getSheetByName(x[0]);
    if (!sh || sh.getLastRow() < 2) return;
    const t = getTable_(x[0]);
    if (!(x[1] in t.col)) return;
    t.rows.forEach(function (r) { const v = String(r[t.col[x[1]]]).trim(); if (v) used[v] = (used[v] || 0) + 1; });
  });
  const orphans = Object.keys(used).filter(function (r) { return cfg.rooms.indexOf(r) < 0 && r !== '準備室'; }).sort(naturalCompare_);
  return { rooms: cfg.rooms, orphans: orphans.map(function (r) { return { name: r, n: used[r] }; }), used: used };
}

function roomRenameDialog() {
  const info = roomInfo_();
  const html = DIALOG_STYLE + `
    <p>教室改名時，用這裡改：<b>設定、實驗排程、準備事項</b>裡的舊名稱會一起換掉，舊的登記才不會對不上。</p>
    <label>要改的教室</label><select id="old"></select>
    <label>改成</label><input id="neu" placeholder="新的名稱">
    <label class="inline"><input type="checkbox" id="items">「品項」的存放教室也一起改（如果這間也放東西）</label>
    <p class="hint" id="orph"></p>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" id="go" onclick="go()">改名</button></div>
    <script>
      const D = __DATA__, sel = document.getElementById('old');
      D.rooms.forEach(function (r) { sel.add(new Option(r + (D.used[r] ? '（' + D.used[r] + ' 筆）' : ''), r)); });
      D.orphans.forEach(function (r) { sel.add(new Option('⚠ ' + r.name + '（不在設定裡，' + r.n + ' 筆）', r.name)); });
      if (D.orphans.length) document.getElementById('orph').textContent = '⚠ 有 ' + D.orphans.length + ' 個教室名稱在資料裡、但不在設定的上課教室裡（可能是以前改過名）：' +
        D.orphans.map(function (r) { return r.name; }).join('、') + '。可以選它，改成現在的名稱。';
      function go() {
        const o = sel.value, n = document.getElementById('neu').value.trim();
        if (!n) { alert('請填新的名稱。'); return; }
        const b = document.getElementById('go'); b.disabled = true; b.textContent = '改名中…';
        google.script.run.withSuccessHandler(showDone).withFailureHandler(function (e) { alert(e.message); b.disabled = false; b.textContent = '改名'; })
          .roomRename(o, n, document.getElementById('items').checked);
      }
    </script>`;
  showDialog_(html, info, '🏷 教室改名（舊資料一起改）', 440);
}

function roomRename(oldName, newName, alsoItems) {
  oldName = String(oldName || '').trim(); newName = String(newName || '').trim();
  if (!oldName || !newName) throw new Error('請選要改的教室，並填新的名稱。');
  if (oldName === newName) throw new Error('新舊名稱一樣。');
  if (/^[=+\-@]/.test(newName) || newName.length > 30) throw new Error('名稱不能用 = + - @ 開頭，最多 30 個字。');
  return withLock_(function () {
    const ss = SpreadsheetApp.getActive();
    const out = [];
    // 一欄裡等於舊名稱的換成新名稱（只寫一次）
    const swapCol = function (sh, col, label, row0) {
      if (!sh || col < 1 || sh.getLastRow() < (row0 || 2)) return 0;
      const r0 = row0 || 2, rg = sh.getRange(r0, col, sh.getLastRow() - r0 + 1, 1), v = rg.getValues();
      let n = 0;
      v.forEach(function (x) { if (String(x[0]).trim() === oldName) { x[0] = newName; n++; } });
      if (n) { rg.setValues(v); out.push('・' + label + '：' + n + ' 格'); }
      return n;
    };
    const s = getSettings_();
    const merged = (s.lists['上課教室'] || []).indexOf(newName) >= 0;
    if (merged) {
      // 新名稱已經在清單裡：清掉舊的那格（合併），其他格往上補
      const c = s.header.indexOf('上課教室') + 1, rg = s.sheet.getRange(2, c, s.sheet.getLastRow() - 1, 1);
      const list = rg.getValues().map(function (x) { return String(x[0]).trim(); }).filter(function (x) { return x && x !== oldName; });
      rg.clearContent();
      if (list.length) s.sheet.getRange(2, c, list.length, 1).setValues(list.map(function (x) { return [x]; }));
      out.push('・設定的上課教室：「' + oldName + '」併入「' + newName + '」');
    } else {
      swapCol(s.sheet, s.header.indexOf('上課教室') + 1, '設定的上課教室');
    }
    [[typeof EXP_SHEET === 'string' ? EXP_SHEET : '實驗排程', '教室'], [typeof TODO_SHEET === 'string' ? TODO_SHEET : '準備事項', '實驗室']].forEach(function (x) {
      [x[0], x[0] + ARCHIVE_SUFFIX].forEach(function (name) {
        const sh = ss.getSheetByName(name);
        if (!sh || sh.getLastRow() < 2) return;
        const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
        swapCol(sh, head.indexOf(x[1]) + 1, name);
      });
    });
    if (alsoItems) {
      const it = getTable_('品項');
      if ('教室' in it.col) swapCol(it.sheet, it.col['教室'] + 1, '品項的存放教室');
    }
    if (readMemo_) readMemo_ = {};
    if (typeof setupExperimentSheet_ === 'function') try { setupExperimentSheet_(); } catch (e) { }   // 教室下拉選單更新
    return '已把「' + oldName + '」改成「' + newName + '」：\n' + (out.length ? out.join('\n') : '（資料裡沒有用到這個名稱）') +
      '\n\n門口海報、課表、網頁都會用新名稱。之後匯入老師課表時，選實驗室也請選新名稱。';
  });
}

// ---------------------------------------------------------------- 🔒 嚴格保護／🔓 警告模式

/** 依目前模式，套用到所有本系統建立的保護（07_安全美化 的 protectSheets_ 跑完也會呼叫） */
function applyProtectMode_() {
  const strict = PropertiesService.getDocumentProperties().getProperty('protectMode') === 'strict';
  const ss = SpreadsheetApp.getActive();
  const tag = typeof PROTECT_TAG === 'string' ? PROTECT_TAG : '實驗室系統保護：';
  let me = null;
  try { me = Session.getEffectiveUser(); } catch (e) { me = null; }
  let n = 0;
  ss.getSheets().forEach(function (sh) {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).concat(sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)).forEach(function (p) {
      if (String(p.getDescription()).indexOf(tag) !== 0) return;
      if (strict) {
        p.setWarningOnly(false);
        if (me) p.addEditor(me);
        const others = p.getEditors().filter(function (u) { return !me || u.getEmail() !== me.getEmail(); });
        if (others.length) p.removeEditors(others);
        if (p.canDomainEdit()) p.setDomainEdit(false);
      } else {
        p.setWarningOnly(true);
      }
      n++;
    });
  });
  return n;
}

function protectStrict() { return protectModeSet_('strict'); }
function protectWarn() { return protectModeSet_('warn'); }
function protectModeSet_(mode) {
  PropertiesService.getDocumentProperties().setProperty('protectMode', mode);
  const n = applyProtectMode_();
  const ui = SpreadsheetApp.getUi();
  ui.alert(mode === 'strict' ? '🔒 已改成嚴格保護' : '🔓 已改回警告模式',
    (n ? '套用到 ' + n + ' 個保護範圍。' : '還沒有保護範圍：請先執行「🔧 維護 → 首頁、美化與資料保護（一次設定）」。') + '\n\n' +
    (mode === 'strict' ? '盤點紀錄、設定、首頁、需補充清單、品項的自動欄位：只有你能改，其他共用編輯的人改不了。\n（你自己、網頁和選單功能都照常。）' :
      '改到系統工作表時會先跳出「確定要編輯嗎？」，按確定就能改。只有你一個人用時建議用這個。'), ui.ButtonSet.OK);
  return n;
}

// ---------------------------------------------------------------- 🖨 櫃子 QR Code 貼紙

/** 品項裡所有的「教室＋櫃別」（不含已淘汰），依教室、櫃別排 */
function cabinetList_() {
  const it = getTable_('品項'), c = it.col, seen = {}, out = [];
  it.rows.forEach(function (r) {
    if (String(r[c['狀態']]).trim() === '已淘汰' || !String(r[c['編號']]).trim()) return;
    const room = String(r[c['教室']]).split(/[\/／]/)[0].trim(), cab = String(r[c['櫃別']]).trim();
    if (!room || !cab) return;
    const k = room + '|' + cab;
    if (!seen[k]) { seen[k] = { room: room, cab: cab, n: 0 }; out.push(seen[k]); }
    seen[k].n++;
  });
  return out.sort(function (a, b) { return naturalCompare_(a.room, b.room) || naturalCompare_(a.cab, b.cab); });
}

function printQrLabels() {
  const list = cabinetList_();
  const rooms = list.map(function (x) { return x.room; }).filter(function (x, i, a) { return a.indexOf(x) === i; });
  const html = DIALOG_STYLE + `
    <p>每個櫃子一張貼紙（約 4.5 × 5.5 公分）。手機相機掃描 → 打開網頁、列出這一櫃的品項，也可以直接盤點這一櫃。</p>
    <label>教室</label><select id="room"><option value="">全部（${list.length} 個櫃子）</option></select>
    <p class="hint">需要先部署成網頁應用程式（網址存在「設定」的「手機網頁網址」或自動抓）。掃描時要登入有權限的 Google 帳號。</p>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" onclick="openPrint({ page: 'qr', room: document.getElementById('room').value }, this)">開啟列印頁</button></div>
    <script>${OPEN_PRINT_JS}
      const D = __DATA__, s = document.getElementById('room');
      D.rooms.forEach(function (r) { s.add(new Option(r + '（' + D.count[r] + ' 個櫃子）', r)); });</script>`;
  const count = {};
  list.forEach(function (x) { count[x.room] = (count[x.room] || 0) + 1; });
  showDialog_(html, { rooms: rooms, count: count }, '🖨 櫃子 QR Code 貼紙', 300);
}

function page_qr(p) {
  const url = typeof webAppUrl_ === 'function' ? webAppUrl_() : (ScriptApp.getService().getUrl() || '');
  if (!url) return HtmlService.createHtmlOutput('<p style="padding:20px">還沒有網頁網址：請先「部署 → 新增部署作業 → 網頁應用程式」。</p>');
  const list = cabinetList_().filter(function (x) { return !p.room || x.room === p.room; });
  const items = list.map(function (x) {
    return { room: x.room, cab: x.cab, n: x.n, url: url + '?room=' + encodeURIComponent(x.room) + '&cab=' + encodeURIComponent(x.cab) };
  });
  const s = getSettings_();
  const body = '<div class="qrs">' + items.map(function (x, i) {
    return '<div class="q"><div class="qr" id="q' + i + '"></div><div class="r">' + esc_(x.room) + '</div><div class="c">' + esc_(x.cab) + (/^[A-Za-z0-9]{1,3}$/.test(x.cab) ? ' 櫃' : '') +
      '</div><div class="h">掃描看櫃內品項・' + x.n + ' 項</div></div>';
  }).join('') + '</div>' + (items.length ? '' : '<p style="padding:20px">沒有櫃子：「品項」的教室、櫃別要有填。</p>');
  const css = '.qrs{display:flex;flex-wrap:wrap;gap:4mm}.q{width:45mm;height:55mm;border:0.6pt dashed #999;border-radius:3mm;box-sizing:border-box;padding:3mm;text-align:center;break-inside:avoid;page-break-inside:avoid}' +
    '.qr{width:32mm;height:32mm;margin:0 auto 2mm}.qr svg{width:32mm;height:32mm;display:block}.r{font-size:9pt;color:#444}.c{font-size:15pt;font-weight:800;line-height:1.2}.h{font-size:7pt;color:#666;margin-top:1mm}';
  // QR Code 在伺服器算好（20_QR產生器），每個只傳黑白格子，頁面上畫成 SVG；不用連外部網站
  const mats = items.map(function (x) { return qrMatrix_(x.url); });
  const js = '<script>var M=' + JSON.stringify(mats) + ';' +
    'M.forEach(function(m,i){var n=m.n,d="";m.rows.forEach(function(h,y){var bits="";for(var k=0;k<h.length;k++)bits+=("000"+parseInt(h[k],16).toString(2)).slice(-4);' +
    'for(var x=0;x<n;){if(bits[x]==="1"){var s=x;while(x<n&&bits[x]==="1")x++;d+="M"+s+" "+y+"h"+(x-s)+"v1h-"+(x-s)+"z";}else x++;}});' +
    'document.getElementById("q"+i).innerHTML=\'<svg xmlns="http://www.w3.org/2000/svg" viewBox="-4 -4 \'+(n+8)+" "+(n+8)+\'" shape-rendering="crispEdges"><rect x="-4" y="-4" width="\'+(n+8)+\'" height="\'+(n+8)+\'" fill="#fff"/><path d="\'+d+\'" fill="#000"/></svg>\';});</script>';
  const html = printShell_('櫃子 QR Code 貼紙', '<div class="dochead"><div class="school">' + esc_(s.params['學校名稱']) + '</div><div class="ttl">櫃子 QR Code 貼紙</div></div>' + body)
    .replace('</style>', css + '</style>').replace('</body>', js + '</body>');
  return HtmlService.createHtmlOutput(html).setTitle('櫃子 QR Code 貼紙（列印）');
}

/** 文字 → QR Code 黑白格子：{n: 邊長, rows: [每列 16 進位字串]}（需要 20_QR產生器） */
function qrMatrix_(text) {
  if (typeof qrcode !== 'function') throw new Error('缺少「20_QR產生器」程式檔。');
  const q = qrcode(0, 'M');
  q.addData(String(text));
  q.make();
  const n = q.getModuleCount(), rows = [];
  for (let y = 0; y < n; y++) {
    let bits = '';
    for (let x = 0; x < n; x++) bits += q.isDark(y, x) ? '1' : '0';
    while (bits.length % 4) bits += '0';
    let hex = '';
    for (let i = 0; i < bits.length; i += 4) hex += parseInt(bits.substr(i, 4), 2).toString(16);
    rows.push(hex);
  }
  return { n: n, rows: rows };
}

/** 網頁：一櫃的品項（掃 QR Code 開的），依排序位置排 */
function appCabItems(room, cab) {
  room = String(room || '').trim(); cab = String(cab || '').trim();
  const it = getTable_('品項'), c = it.col, out = [];
  it.rows.forEach(function (r) {
    if (!String(r[c['編號']]).trim() || String(r[c['櫃別']]).trim() !== cab) return;
    if (String(r[c['教室']]).split(/[\/／]/)[0].trim() !== room) return;
    out.push({ code: String(r[c['編號']]), name: String(r[c['品名']]), spec: String(r[c['化學式或規格']] || ''), room: room, cab: cab,
      pos: String(r[c['排序位置']] || ''), places: String(r[c['分處存放']] || ''), unit: String(r[c['單位']] || ''), status: String(r[c['狀態']]),
      date: String(r[c['最新盤點日期']] || ''), qty: String(r[c['最新數量']]), note: String(r[c['最新數量說明']] || '') });
  });
  out.sort(function (a, b) { return naturalCompare_(a.pos, b.pos) || naturalCompare_(a.code, b.code); });
  return { rows: out, more: false };
}

// ---------------------------------------------------------------- 網頁：拖曳換時段

/**
 * p = {row, key（日期|教室|節次|班級）, date, period, lab, force}
 * 撞堂時回傳 {conflicts: [...]}（force 才硬移）；同一堂課待準備／已準備的準備事項跟著移。
 */
function moveBooking(p) {
  return withLock_(function () {
    const cfg = labConfig_(), t = expTable_(), c = t.col;
    const row = findLessonRow_(t, cfg, p.row, p.key);
    if (row < 0) throw new Error('找不到這堂課，可能已被改過。請重新整理。');
    const r = t.rows[row - 2];
    const date = dateKey_(p.date), lab = String(p.lab || '').trim(), per = String(p.period || '').trim();
    if (!date || !lab || !per) throw new Error('要移到的日期、教室、節次不完整。');
    if (cfg.rooms.indexOf(lab) < 0) throw new Error('「' + lab + '」不在設定的上課教室裡。');
    const pis = parsePeriods_(per, cfg.periods);
    if (!pis.length) throw new Error('節次不正確：' + per);
    const od = dateKey_(r[c['日期']]), olab = String(r[c['教室']]).trim(), oper = periodVal_(r[c['節次']]);
    if (od === date && olab === lab && oper === per) return { msg: '沒有移動。' };
    const conflicts = bookings_(date, date, cfg).filter(function (b) {
      return b.row !== row && b.lab === lab && b.pis.some(function (x) { return pis.indexOf(x) >= 0; });
    }).map(function (b) { return periodLabel_(b.periodText) + '：' + [b.content || b.type, b.cls, b.teacher].filter(String).join(' '); });
    if (conflicts.length && !p.force) return { conflicts: conflicts };
    fixPeriodCol_(t.sheet, c['節次'] + 1);
    t.sheet.getRange(row, c['日期'] + 1).setValue(dateValue_(date));
    t.sheet.getRange(row, c['節次'] + 1).setValue(per);
    t.sheet.getRange(row, c['教室'] + 1).setValue(lab);
    // 準備事項跟著移
    let moved = 0;
    const tsh = SpreadsheetApp.getActive().getSheetByName(typeof TODO_SHEET === 'string' ? TODO_SHEET : '準備事項');
    if (tsh && tsh.getLastRow() > 1) {
      const tt = getTable_(tsh.getName()), tc = tt.col;
      const cls = classLabel_(r[c['班級']], cfg.classNames);
      tt.rows.forEach(function (x, i) {
        const st = String(x[tc['狀態']]).trim() || '待準備';
        if ((st !== '待準備' && st !== '已準備') || dateKey_(x[tc['需要日期']]) !== od || String(x[tc['實驗室']]).trim() !== olab) return;
        if (periodVal_(x[tc['節次']]) !== oper) return;
        const xc = String(x[tc['班級']]).trim();
        if (xc && xc !== cls && xc !== String(r[c['班級']]).trim()) return;
        tsh.getRange(i + 2, tc['需要日期'] + 1).setValue(dateValue_(date));
        tsh.getRange(i + 2, tc['節次'] + 1).setNumberFormat('@').setValue(per);
        tsh.getRange(i + 2, tc['實驗室'] + 1).setValue(lab);
        moved++;
      });
    }
    return { msg: '已移到 ' + rocText_(date) + '（' + weekdayOf_(date) + '）' + periodLabel_(per) + ' ' + lab + (moved ? '，準備事項 ' + moved + ' 項也跟著移了' : '') + '。',
      undo: { date: od, period: oper, lab: olab }, key: [date, lab, per, cls_(r, c, cfg)].join('|'), row: row };
  });
}
function cls_(r, c, cfg) { return classLabel_(r[c['班級']], cfg.classNames); }
