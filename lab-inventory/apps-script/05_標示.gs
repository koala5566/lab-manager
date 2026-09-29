/**
 * 實驗室藥品器材耗材管理系統 — 藥品櫃標示（第 6 步）
 *
 * 選單「列印藥品櫃標示」→ 選櫃內版或櫃外版、選櫃子 → 在新分頁開啟 → 「🖨 列印」。
 * 依「品項」中藥品清單的櫃別與排序位置（例：A 櫃 3-2 ＝ 第 3 排第 2 個）自動排版，每格印品名與化學式。
 *   ・櫃內版：A4 直式，所有櫃子接著印，字較小
 *   ・櫃外版：A4 橫式，一櫃一頁，字較大，右上角印教室與更新日期
 * 排序位置寫「3-2旁」會排在 3-2 後面；「3-2/大冰箱」會在格子裡加註「大冰箱」。
 * 教室不同於該櫃的品項（例如放在生物準備室冰箱的）不會印在櫃子標示上。
 * 需要「01_基礎」「02_盤點」「03_手機盤點」「04_列印」。
 */

const LABEL_PER_LINE = 10;   // 一排超過 10 個時換行接著印

// ---------------------------------------------------------------- 選單

function printCabinetLabels() {
  const url = webAppUrl_();
  if (!url) return;
  const cabs = Object.keys(labelCabinets_()).sort(cabinetCompare_);
  if (!cabs.length) {
    SpreadsheetApp.getUi().alert('藥品清單中沒有「排序位置」是「排-位」格式（例：3-2）的品項，無法產生標示。');
    return;
  }
  const html = DIALOG_STYLE + `
    <label>版本</label>
    <label class="inline"><input type="radio" name="v" value="out" checked> 櫃外標示（A4 橫式，一櫃一頁，大字）</label>
    <label class="inline"><input type="radio" name="v" value="in"> 櫃內標示（A4 直式，全部接著印，小字）</label>
    <label>櫃子</label>
    <label class="inline"><input type="checkbox" id="all" checked onchange="toggle(this.checked)"> <b>全選</b></label>
    <div id="list"></div>
    <div class="btns"><button onclick="google.script.host.close()">取消</button>
      <button class="primary" onclick="go()">開啟列印頁</button></div>
    <script>
      const D = __DATA__;
      D.cabs.forEach(function (c) {
        const l = document.createElement('label'); l.className = 'inline'; l.style.display = 'inline-block'; l.style.width = '30%';
        const x = document.createElement('input'); x.type = 'checkbox'; x.checked = true; x.value = c;
        l.appendChild(x); l.appendChild(document.createTextNode(' ' + c + ' 櫃'));
        document.getElementById('list').appendChild(l);
      });
      function toggle(on) { document.querySelectorAll('#list input').forEach(function (c) { c.checked = on; }); }
      function go() {
        const c = Array.prototype.filter.call(document.querySelectorAll('#list input'), function (x) { return x.checked; })
          .map(function (x) { return x.value; });
        if (!c.length) { alert('請至少勾選一個櫃子。'); return; }
        const v = document.querySelector('input[name=v]:checked').value;
        window.open(D.url + '?page=labels&v=' + v + '&c=' + encodeURIComponent(c.join('|')), '_blank');
        google.script.host.close();
      }
    </script>`;
  showDialog_(html, { url: url, cabs: cabs }, '列印藥品櫃標示', 300 + Math.ceil(cabs.length / 3) * 28);
}

// ---------------------------------------------------------------- 列印頁（網址 ?page=labels）

function page_labels(p) {
  const outside = p.v !== 'in';
  const want = String(p.c || '').split('|').filter(String);
  const cabs = labelCabinets_();
  const names = Object.keys(cabs).sort(cabinetCompare_)
    .filter(function (c) { return !want.length || want.indexOf(c) >= 0; });
  const body = names.map(function (c) { return cabinetHtml_(c, cabs[c], outside); }).join('') ||
    '<p style="padding:20px">沒有可列印的櫃子。</p>';
  const title = outside ? '藥品櫃外標示' : '藥品櫃內標示';
  return HtmlService.createHtmlOutput(labelShell_(title, body, outside)).setTitle(title + '（列印）');
}

/**
 * 整理各櫃內容：{櫃別: {room, updated, rows: {排: [{c, order, name, spec, note}]}}}
 * 只取藥品清單、使用中、排序位置為「排-位」的品項。
 */
function labelCabinets_() {
  const items = getTable_('品項');
  need_(items, ['編號', '清單分區', '品名', '化學式或規格', '教室', '櫃別', '排序位置', '狀態', '最新盤點日期']);
  const ic = items.col;
  const list = [];
  items.rows.forEach(function (r) {
    if (String(r[ic['清單分區']]).trim() !== '藥品清單' || String(r[ic['狀態']]).trim() === '已淘汰') return;
    const cab = String(r[ic['櫃別']]).trim();
    const m = String(r[ic['排序位置']]).trim().match(/^(\d+)-(\d+)(.*)$/);
    if (!cab || !m) return;
    const rest = m[3].trim();
    list.push({
      cab: cab, row: Number(m[1]), c: Number(m[2]),
      order: rest ? 1 : 0,                               // 「3-2旁」排在「3-2」後面
      note: rest.replace(/^[\/／、]/, '').replace(/^旁$/, ''),
      name: String(r[ic['品名']]).trim(), spec: String(r[ic['化學式或規格']]).trim(),
      room: String(r[ic['教室']]).trim(), code: String(r[ic['編號']]).trim(),
      updated: String(r[ic['最新盤點日期']]).trim(),
    });
  });

  // 每櫃的教室＝該櫃品項最多的教室；教室不同的（例如在生物準備室冰箱）不印在櫃子標示上
  const byCab = {};
  list.forEach(function (x) { (byCab[x.cab] = byCab[x.cab] || []).push(x); });
  const out = {};
  Object.keys(byCab).forEach(function (cab) {
    const count = {};
    byCab[cab].forEach(function (x) { const k = x.room.split(/[\/／]/)[0]; count[k] = (count[k] || 0) + 1; });
    const room = Object.keys(count).sort(function (a, b) { return count[b] - count[a]; })[0];
    const rows = {};
    let updated = '';
    byCab[cab].forEach(function (x) {
      if (x.room.split(/[\/／]/)[0] !== room) return;
      (rows[x.row] = rows[x.row] || []).push(x);
      if (x.updated > updated) updated = x.updated;
    });
    Object.keys(rows).forEach(function (k) {
      rows[k].sort(function (a, b) { return a.c - b.c || a.order - b.order || naturalCompare_(a.code, b.code); });
    });
    out[cab] = { room: room, updated: updated, rows: rows };
  });
  return out;
}

/** 櫃別排序：A～Z 單一字母在前，其次 2-B、2-D… */
function cabinetCompare_(a, b) {
  const k = function (c) { return /^[A-Z]$/i.test(c) ? 0 : 1; };
  return k(a) - k(b) || naturalCompare_(a, b);
}

/** 把一排的品項依位置號碼排進格子；空號留空格，同號多個就並排。回傳格子陣列（每格 null 或品項）。 */
function layoutRow_(items) {
  const cells = [];
  const maxC = items.reduce(function (m, x) { return Math.max(m, x.c); }, 0);
  for (let c = 1; c <= maxC; c++) {
    const here = items.filter(function (x) { return x.c === c; });
    if (!here.length) cells.push(null);
    here.forEach(function (x) { cells.push(x); });
  }
  return cells;
}

/** 化學式的數字改成下標：NH4Cl → NH₄Cl（前面是字母或括號的數字才改；結晶水 ·4H2O 的 4 不改）。 */
function formulaHtml_(s) {
  const t = String(s || '');
  if (!/^[A-Za-z0-9()\[\]·∙.\s₀-₉]+$/.test(t) || !/[A-Za-z]/.test(t)) return esc_(t);
  return esc_(t).replace(/([A-Za-z)\]])(\d+)/g, '$1<sub>$2</sub>');
}

function cabinetHtml_(cab, info, outside) {
  const rowNos = Object.keys(info.rows).map(Number);
  const maxRow = Math.max(5, rowNos.length ? Math.max.apply(null, rowNos) : 0);
  const lines = [];   // 每一行：{label, cells}
  let width = 1;
  for (let r = 1; r <= maxRow; r++) {
    const cells = info.rows[r] ? layoutRow_(info.rows[r]) : [];
    if (!cells.length) { lines.push({ label: '第' + r + '排', cells: [] }); continue; }
    for (let i = 0; i < cells.length; i += LABEL_PER_LINE) {
      const part = cells.slice(i, i + LABEL_PER_LINE);
      lines.push({ label: '第' + r + '排' + (i ? '<br><small>（續）</small>' : ''), cells: part });
      width = Math.max(width, part.length);
    }
  }
  const head = '<div class="lh"><div class="cab">藥品 ' + esc_(cab) + ' 櫃</div><div class="rm">' + esc_(info.room) +
    (info.updated ? '<br><small>更新日期：' + esc_(info.updated) + '</small>' : '') + '</div></div>';
  const cols = '<col class="rowc">' + new Array(width + 1).join('<col>');
  const body = lines.map(function (ln) {
    let tds = '';
    for (let i = 0; i < width; i++) {
      const x = ln.cells[i];
      tds += x ? '<td><span class="pos">' + esc_(x.row + '-' + x.c) + '</span><div class="nm">' + esc_(x.name) +
        '</div><div class="fm">' + formulaHtml_(x.spec) + '</div>' +
        (x.note ? '<div class="nt">' + esc_(x.note) + '</div>' : '') + '</td>' : '<td class="empty">&nbsp;</td>';
    }
    return '<tr><th>' + ln.label + '</th>' + tds + '</tr>';
  }).join('');
  // 行數多的櫃子（例如 E 櫃換行後有 7 行）用較緊的版面，櫃外版才塞得進一頁
  const dense = lines.length > 5 ? ' dense' : '';
  return '<section class="cabinet ' + (outside ? 'out' : 'in') + dense + '">' + head +
    '<table class="lb"><colgroup>' + cols + '</colgroup><tbody>' + body + '</tbody></table></section>';
}

function labelShell_(title, body, outside) {
  return `<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><title>${esc_(title)}</title>
<style>
  @page { size: A4 ${outside ? 'landscape' : 'portrait'}; margin: 8mm; }
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; box-sizing: border-box; }
  body { margin: 0; color: #000; background: #e9ecef;
    font-family: "Noto Sans TC", "Microsoft JhengHei", "PMingLiU", "PingFang TC", sans-serif; }
  .bar { position: sticky; top: 0; z-index: 9; background: #fff8e1; border-bottom: 1px solid #e0c97a;
    padding: 10px 16px; font-size: 14px; display: flex; gap: 16px; align-items: center; flex-wrap: wrap; }
  .bar button { font-size: 16px; padding: 8px 20px; background: #1a73e8; color: #fff; border: none; border-radius: 6px; cursor: pointer; }
  .paper { width: ${outside ? '281mm' : '194mm'}; margin: 16px auto; background: #fff; padding: 6mm; box-shadow: 0 1px 4px rgba(0,0,0,.2); }
  section.cabinet { break-inside: avoid; page-break-inside: avoid; margin-bottom: 8mm; }
  section.cabinet.out + section.cabinet.out { break-before: page; page-break-before: always; }
  .lh { display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 3mm; }
  .cab { font-weight: bold; font-size: ${outside ? '26pt' : '15pt'}; }
  .rm { text-align: right; font-size: ${outside ? '16pt' : '11pt'}; font-weight: bold; }
  .rm small { font-size: ${outside ? '11pt' : '8pt'}; font-weight: normal; }
  table.lb { width: 100%; border-collapse: collapse; table-layout: fixed; }
  table.lb col.rowc { width: ${outside ? '20mm' : '14mm'}; }
  table.lb th, table.lb td { border: ${outside ? '1pt' : '0.6pt'} solid #000; text-align: center; vertical-align: middle;
    position: relative; padding: ${outside ? '5mm 1mm 3mm' : '3mm 0.5mm 1.5mm'}; overflow-wrap: anywhere; }
  table.lb th { background: #eee; font-size: ${outside ? '13pt' : '9pt'}; }
  .pos { position: absolute; top: 1px; left: 3px; font-size: ${outside ? '8pt' : '6pt'}; color: #666; }
  .nm { font-weight: bold; font-size: ${outside ? '15pt' : '9pt'}; line-height: 1.25; }
  .fm { font-size: ${outside ? '13pt' : '8pt'}; margin-top: 1mm; font-family: "Times New Roman", serif; }
  .nt { font-size: ${outside ? '9pt' : '6.5pt'}; color: #444; margin-top: 1mm; }
  td.empty { background: #fafafa; }
  section.out.dense table.lb th, section.out.dense table.lb td { padding: 4mm 1mm 1.5mm; }
  section.out.dense .nm { font-size: 12.5pt; }
  section.out.dense .fm { font-size: 11pt; }
  section.out.dense table.lb tr { height: 18mm !important; }
  section.out.dense .cab { font-size: 22pt; }
  table.lb tr { height: ${outside ? '22mm' : '14mm'}; }
  @media print {
    body { background: #fff; }
    .bar { display: none; }
    .paper { width: auto; margin: 0; padding: 0; box-shadow: none; }
  }
</style></head><body>
<div class="bar"><button onclick="window.print()">🖨 列印</button>
  <span>列印設定：紙張 A4、版面配置「<b>${outside ? '橫向' : '直向'}</b>」；「更多設定」取消勾選「頁首和頁尾」、勾選「背景圖形」。
  每格左上角的小字是排序位置（排-位）。</span></div>
<div class="paper">${body}</div>
</body></html>`;
}
