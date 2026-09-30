/**
 * 實驗室藥品器材耗材管理系統 — 網頁工作台（手機、平板、電腦共用）
 *
 * 開網頁（部署的網址）就是這個畫面；舊版手機網頁在網址後面加 ?page=old。
 *   ・今天：現在第幾節、哪幾間在上課、今天和下一個上課日的課、待準備、需補充（電腦、平板是儀表板）
 *   ・課表：各實驗室單日／整週／月曆／清單；點一堂課看詳細（器材夠不夠、準備事項、補實驗名稱、準備單）
 *   ・待辦：準備事項打勾、新增（可對到某堂課、品項搜尋）
 *   ・盤點：一次一項、大按鈕 −／＋、存好自動跳下一項
 *   ・更多：需補充、品項查詢、外觀（淺色／深色）
 * 畫面程式由 tools/webapp/ 的原始檔產生（build_webapp.py），請不要直接改下面的 APP_HTML。
 * 需要「01」～「15」。
 */

function page_app() {
  return HtmlService.createHtmlOutput(APP_HTML)
    .setTitle('實驗室管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

/** 舊版手機網頁（?page=old） */
function page_old(p) { return oldMobilePage_(p); }

/** 一開網頁要的資料，一次給齊（之後只在需要時補抓）。 */
function appInit() {
  const cfg = labConfig_();
  const s = getSettings_();
  const today = today_(), next = nextSchoolDay_(today);
  const books = bookings_(today, next, cfg);
  const byDay = function (k) {
    return books.filter(function (b) { return b.date === k; }).sort(function (a, b) {
      return (a.pis[0] === undefined ? 99 : a.pis[0]) - (b.pis[0] === undefined ? 99 : b.pis[0]) || naturalCompare_(a.lab, b.lab);
    });
  };
  const td = typeof todoData === 'function' ? todoData() : { list: [], items: [] };
  let restock = null;
  try { if (typeof restockItems_ === 'function') restock = restockItems_().length; } catch (e) { restock = null; }
  const count = appCountInfo_();
  return {
    school: String(s.params['學校名稱'] || ''), term: String(s.params['學年度學期'] || ''),
    rooms: cfg.rooms, periods: cfg.periods, colors: cfg.colors, types: cfg.types, classNames: cfg.classNames,
    today: { key: today, roc: rocText_(today), wd: weekdayOf_(today), list: byDay(today) },
    next: { key: next, roc: rocText_(next), wd: weekdayOf_(next), list: byDay(next) },
    todos: td.list, items: td.items,
    kpi: { restock: restock, loans: appSheetCount_(LOAN_SHEET, '狀態', ['借出中'], '預計歸還'),
      purchase: appSheetCount_(REQ_SHEET, '狀態', ['待處理', '已請購']), count: count },
    url: ScriptApp.getService().getUrl() || '',
  };
}

/** 某工作表某欄是某些值的列數；dueCol 有給時也算逾期幾筆。 */
function appSheetCount_(name, col, values, dueCol) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return { n: 0, late: 0 };
  const t = getTable_(name), c = t.col;
  if (!(col in c)) return { n: 0, late: 0 };
  const today = today_();
  let n = 0, late = 0;
  t.rows.forEach(function (r) {
    if (values.indexOf(String(r[c[col]]).trim()) < 0) return;
    n++;
    if (dueCol && dueCol in c) { const d = dateKey_(r[c[dueCol]]); if (d && d < today) late++; }
  });
  return { n: n, late: late };
}

/** 盤點表進度：{name, date, total, done} */
function appCountInfo_() {
  const info = JSON.parse(PropertiesService.getDocumentProperties().getProperty('count') || '{}');
  const sh = SpreadsheetApp.getActive().getSheetByName('盤點表');
  if (!sh || sh.getLastRow() < 2) return { name: info.name || '', total: 0, done: 0 };
  const t = getTable_('盤點表'), c = t.col;
  let total = 0, done = 0;
  t.rows.forEach(function (r) {
    if (!String(r[c['編號']]).trim()) return;
    total++;
    if (String(r[c['本次數量']]).trim() !== '' || String(r[c['本次說明']]).trim() !== '') done++;
  });
  return { name: info.name || '', date: info.date ? rocText_(info.date) : '', total: total, done: done };
}

/** 一堂課的器材檢查（實驗名稱要有對應的實驗套組）。 */
function appLessonKit(name, groups) {
  if (!name || typeof experimentCheck_ !== 'function') return null;
  const ck = experimentCheck_(name, Number(groups) || 0);
  if (!ck) return null;
  return { groups: ck.groups, rows: ck.rows.map(function (x) {
    return { name: x.name, per: x.per, need: x.need, have: x.haveText, unit: x.unit, loc: x.loc, short: x.short, note: x.note };
  }) };
}

/** 重新整理用：今天頁、待辦的資料 */
function appRefresh() { return appInit(); }

// ---------------------------------------------------------------- 畫面（自動產生，請改 tools/webapp/ 的原始檔）

const APP_HTML = `<!DOCTYPE html>
<html lang="zh-Hant"><head><meta charset="utf-8">
<meta name="theme-color" content="#1F5EDB">
<style>:root{
  --bg:#F4F6FA; --surface:#FFFFFF; --surface2:#F8FAFC; --line:#E3E8EF; --text:#1B2433; --sub:#5B6576; --mute:#8A94A6;
  --pri:#1F5EDB; --pri2:#3B82F6; --pri-weak:#E8EFFD; --ok:#1E8E4E; --ok-weak:#E6F4EC; --warn:#C96A04; --warn-weak:#FEF3E2;
  --bad:#D6363B; --bad-weak:#FDECEC; --r:14px; --shadow:0 1px 2px rgba(16,24,40,.06),0 2px 8px rgba(16,24,40,.06);
  --t-mix:0%;
}
html.dark{ --bg:#0F141C; --surface:#18202B; --surface2:#1D2633; --line:#263142; --text:#E6EAF1; --sub:#A3AEC2; --mute:#7C879B;
  --pri:#6D9BFF; --pri2:#8AB0FF; --pri-weak:#1D2B47; --ok:#4CC38A; --ok-weak:#16301F; --warn:#F0A94B; --warn-weak:#3A2A12; --bad:#FF7A7F; --bad-weak:#3A1A1C;
  --shadow:0 1px 2px rgba(0,0,0,.4); --t-mix:55%; color-scheme:dark; }
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
html,body{margin:0;height:100%}
body{font-family:-apple-system,"PingFang TC","Noto Sans TC","Microsoft JhengHei",sans-serif;background:var(--bg);color:var(--text);font-size:15px;-webkit-font-smoothing:antialiased;overscroll-behavior-y:contain}
button,input,select,textarea{font:inherit;color:inherit}
button{cursor:pointer;border:none;background:none;padding:0}
svg.i{width:20px;height:20px;stroke:currentColor;fill:none;stroke-width:1.9;stroke-linecap:round;stroke-linejoin:round;flex:none}
svg.i.s{width:16px;height:16px}
.art{display:block;flex:none}.bob{animation:bob 1.4s ease-in-out infinite}@keyframes bob{50%{transform:translateY(-6px)}}
.sp{flex:1}.hide{display:none!important}
/* 事件顏色：設定的顏色，深色模式時和底色混合 */
.ev-c{background:var(--c);background:color-mix(in srgb,var(--c) calc(100% - var(--t-mix)),var(--surface))}

/* ---------- 版面 ---------- */
.app{min-height:100vh}
.side{display:none}
.main{padding:0 0 96px;max-width:1500px;min-width:0}
.nav{position:fixed;left:0;right:0;bottom:0;background:var(--surface);border-top:1px solid var(--line);display:grid;grid-template-columns:repeat(5,1fr);padding:6px 6px calc(8px + env(safe-area-inset-bottom));z-index:30}
.nav button{display:flex;flex-direction:column;align-items:center;gap:3px;font-size:11.5px;color:var(--mute);padding:5px 0;border-radius:12px;position:relative}
.nav button.on{color:var(--pri);font-weight:700}.nav button.on svg{stroke-width:2.3}
.dotn{position:absolute;top:0;left:55%;background:var(--bad);color:#fff;font-size:10px;border-radius:9px;padding:0 5px;font-weight:700;line-height:16px;min-width:16px;text-align:center}
@media (min-width:900px){
  .app{display:grid;grid-template-columns:236px minmax(0,1fr)}
  .side{display:flex;flex-direction:column;gap:3px;position:sticky;top:0;height:100vh;background:var(--surface);border-right:1px solid var(--line);padding:16px 12px;overflow:auto}
  .nav{display:none}
  .main{padding:22px 28px 40px}
}
.brand{display:flex;gap:10px;align-items:center;padding:2px 8px 14px}.brand b{display:block;font-size:15px}.brand small{color:var(--sub);font-size:12px}
.brand .lg{width:42px;height:42px;border-radius:13px;background:linear-gradient(135deg,#DCE7FD,#E7F6F1);display:grid;place-items:center;overflow:hidden}
html.dark .brand .lg{background:linear-gradient(135deg,#23355A,#1C3A33)}
.side button.nv{display:flex;align-items:center;gap:12px;padding:10px 12px;border-radius:12px;color:var(--sub);font-weight:600;font-size:14.5px;text-align:left;width:100%}
.side button.nv:hover{background:var(--surface2)}.side button.nv.on{background:var(--pri-weak);color:var(--pri)}
.side .n{margin-left:auto;background:var(--bad);color:#fff;border-radius:10px;font-size:11px;padding:1px 7px}.side .n.w{background:var(--warn)}
.side .grp{font-size:11.5px;color:var(--mute);font-weight:700;margin:14px 12px 4px;letter-spacing:1px}
.side .foot{margin-top:auto;padding:10px 8px 0;font-size:12px;color:var(--mute)}

/* ---------- 共用元件 ---------- */
.top{position:sticky;top:0;z-index:20;background:color-mix(in srgb,var(--bg) 88%,transparent);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px);padding:14px 18px 10px;display:flex;align-items:center;gap:10px}
.top h1{font-size:26px;margin:0;font-weight:800;letter-spacing:.5px}.top .sub{color:var(--sub);font-size:14px;margin-top:2px}
@media (min-width:900px){.top{position:static;background:none;backdrop-filter:none;padding:0 0 16px}.top h1{font-size:25px}}
.icbtn{width:40px;height:40px;border-radius:12px;background:var(--surface);display:grid;place-items:center;box-shadow:var(--shadow);color:var(--sub)}
.icbtn:active{transform:scale(.96)}
.pad{padding:0 18px}@media (min-width:900px){.pad{padding:0}}
.h2{font-size:15px;color:var(--sub);font-weight:700;margin:18px 2px 8px;display:flex;align-items:center;gap:8px}
.h2 .more{margin-left:auto;font-size:13px;color:var(--pri);font-weight:600}
.card{background:var(--surface);border-radius:16px;box-shadow:var(--shadow);border:1px solid var(--line)}
@media (max-width:899px){.card{border:none}}
.chips{display:flex;gap:8px;padding:2px 18px 6px;overflow-x:auto;scrollbar-width:none}.chips::-webkit-scrollbar{display:none}
.chip{flex:none;display:flex;align-items:center;gap:6px;background:var(--surface);border-radius:999px;padding:7px 12px;font-size:13.5px;box-shadow:var(--shadow);color:var(--sub)}
.chip b{color:var(--text);font-size:15px}.chip.warn b{color:var(--warn)}.chip.bad b{color:var(--bad)}
.tag{display:inline-block;font-size:11.5px;padding:1px 7px;border-radius:6px;margin-right:6px;color:var(--text);font-weight:600;white-space:nowrap}
.badge{margin-left:auto;font-size:12px;font-weight:700;color:var(--warn);background:var(--warn-weak);padding:3px 8px;border-radius:8px;white-space:nowrap}
.btn{height:48px;border-radius:14px;display:flex;align-items:center;justify-content:center;gap:6px;font-weight:700;font-size:15.5px;padding:0 16px}
.btn.p{background:var(--pri);color:#fff;box-shadow:0 4px 12px rgba(31,94,219,.25)}.btn.g{background:var(--surface2);color:var(--sub);border:1px solid var(--line)}
.btn.s{height:40px;font-size:14px;border-radius:11px}.btn:disabled{opacity:.55}.btn:active{transform:scale(.98)}
.seg{display:flex;background:var(--surface);border-radius:12px;padding:3px;box-shadow:var(--shadow);border:1px solid var(--line)}
.seg button{flex:1;text-align:center;padding:7px 12px;border-radius:9px;font-size:13.5px;color:var(--sub);white-space:nowrap}.seg button.on{background:var(--pri-weak);color:var(--pri);font-weight:700}
.pill{background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:7px 12px;font-size:13.5px;color:var(--sub);display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.pill.on{background:var(--pri);color:#fff;border-color:var(--pri)}
.inp,select.inp,textarea.inp{width:100%;border:1.5px solid var(--line);border-radius:12px;padding:11px 12px;font-size:15px;background:var(--surface);outline:none}
.inp:focus{border-color:var(--pri);box-shadow:0 0 0 3px var(--pri-weak)}
.lbl{font-size:12.5px;color:var(--sub);font-weight:600;display:block;margin:12px 0 5px}
.ck{width:26px;height:26px;border-radius:8px;border:2px solid #C3CBD8;flex:none;display:grid;place-items:center;color:#fff;background:var(--surface)}
.ck.on{background:var(--ok);border-color:var(--ok)}
.empty{text-align:center;padding:26px 16px;color:var(--sub)}.empty .art{margin:0 auto 8px}.empty b{display:block;color:var(--text);font-size:16px;margin-bottom:4px}
.late{color:var(--bad);font-weight:700}.soon{color:var(--warn);font-weight:700}
.muted{color:var(--mute)}
/* 骨架 */
.sk{background:linear-gradient(90deg,var(--line) 25%,var(--surface2) 50%,var(--line) 75%);background-size:200% 100%;animation:sk 1.2s infinite;border-radius:12px}
@keyframes sk{to{background-position:-200% 0}}
/* 提示 */
.toast{position:fixed;left:50%;transform:translate(-50%,20px);bottom:calc(92px + env(safe-area-inset-bottom));background:#1B2433;color:#fff;border-radius:12px;padding:11px 16px;font-size:14px;display:flex;gap:8px;align-items:center;box-shadow:0 8px 24px rgba(0,0,0,.25);z-index:90;opacity:0;pointer-events:none;transition:.25s;max-width:92vw}
.toast.on{opacity:1;transform:translate(-50%,0)}.toast.err{background:#B42318}
@media (min-width:900px){.toast{bottom:28px}}
html.dark .toast{background:#E6EAF1;color:#111}
.fab{position:fixed;right:18px;bottom:calc(88px + env(safe-area-inset-bottom));width:58px;height:58px;border-radius:18px;background:var(--pri);color:#fff;display:grid;place-items:center;box-shadow:0 8px 20px rgba(31,94,219,.35);z-index:25}
.fab svg{width:26px;height:26px}@media (min-width:900px){.fab{bottom:28px;right:28px}}
/* 抽屜（手機從下面、平板電腦在中間） */
.scrim{position:fixed;inset:0;background:rgba(15,20,28,.45);z-index:60;opacity:0;transition:.2s}.scrim.on{opacity:1}
.sheet{position:fixed;left:0;right:0;bottom:0;max-height:92vh;overflow:auto;background:var(--surface);border-radius:22px 22px 0 0;padding:10px 18px calc(22px + env(safe-area-inset-bottom));z-index:61;transform:translateY(100%);transition:transform .25s ease}
.sheet.on{transform:none}.sheet .grab{width:40px;height:5px;border-radius:3px;background:var(--line);margin:0 auto 12px}
.sheet h3{margin:0 0 6px;font-size:19px;display:flex;align-items:center;gap:8px}.sheet .x{margin-left:auto;color:var(--mute)}
@media (min-width:900px){.sheet{left:50%;right:auto;bottom:auto;top:50%;width:560px;border-radius:20px;transform:translate(-50%,-46%);opacity:0;transition:.2s}.sheet.on{transform:translate(-50%,-50%);opacity:1}.sheet .grab{display:none}}

/* ---------- 今天（手機） ---------- */
.now{background:linear-gradient(135deg,#1F5EDB,#3B82F6);color:#fff;border-radius:20px;padding:14px 16px;box-shadow:0 6px 18px rgba(31,94,219,.28);display:flex;gap:10px}
html.dark .now{background:linear-gradient(135deg,#23408C,#2F5FC4)}
.now .h{display:flex;align-items:center;gap:8px;font-size:13px;opacity:.92}.now .dot{width:8px;height:8px;border-radius:50%;background:#7CFFB2;box-shadow:0 0 0 4px rgba(124,255,178,.25)}
.now .big{font-size:20px;font-weight:800;margin:6px 0 10px}.now .labs{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.now .lab{background:rgba(255,255,255,.15);border-radius:12px;padding:9px 10px;text-align:left;color:#fff}.now .lab small{opacity:.85;font-size:12px;display:block}.now .lab b{font-size:15px}
.now.off{background:linear-gradient(135deg,#7C8CA3,#A3B1C4)}
.tl{background:var(--surface);border-radius:var(--r);box-shadow:var(--shadow);overflow:hidden}
.tl .it{display:flex;gap:12px;padding:11px 14px;border-bottom:1px solid var(--line);align-items:center;width:100%;text-align:left}.tl .it:last-child{border:none}
.tl .p{width:36px;height:36px;border-radius:10px;background:var(--pri-weak);color:var(--pri);display:grid;place-items:center;font-weight:800;font-size:15px;flex:none}
.tl .it.past{opacity:.5}.tl .it.cur{box-shadow:inset 0 0 0 2px var(--pri);border-radius:14px}.tl .it.cur .p{background:var(--pri);color:#fff}
.tl .t{font-weight:700}.tl .s{color:var(--sub);font-size:13px;margin-top:2px}
.tl .bar{width:4px;align-self:stretch;border-radius:2px;flex:none}
.todo{background:var(--surface);border-radius:var(--r);box-shadow:var(--shadow);overflow:hidden}
.todo .it{display:flex;gap:12px;padding:11px 14px;align-items:flex-start;border-bottom:1px solid var(--line)}.todo .it:last-child{border:none}
.todo .t{font-weight:700}.todo .s{font-size:13px;color:var(--sub);margin-top:2px;line-height:1.45}.todo .it.done .t{text-decoration:line-through;color:var(--mute)}
.todo .x{margin-left:auto;font-size:12.5px;color:var(--mute);padding:4px 6px;border-radius:8px;white-space:nowrap}.todo .x.cf{color:#fff;background:var(--bad)}
.day{display:flex;align-items:center;gap:8px;font-weight:800;font-size:14px;margin:16px 4px 8px;color:var(--pri)}.day.late{color:var(--bad)}.day.soon{color:var(--warn)}

/* ---------- 今天（電腦、平板：儀表板） ---------- */
.hello{display:flex;align-items:center;gap:14px;margin-bottom:18px}.hello h1{margin:0;font-size:26px}.hello .sub{color:var(--sub);margin-top:3px}
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:16px}
.kpi{padding:14px 16px;text-align:left}.kpi small{color:var(--sub);font-size:13px;display:flex;align-items:center;gap:6px}
.kpi b{display:block;font-size:30px;margin:4px 0 0}.kpi span{font-size:12.5px;color:var(--mute)}.kpi:hover{border-color:var(--pri)}
.dash{display:grid;grid-template-columns:1.4fr 1fr;gap:16px}
@media (max-width:1180px){.kpis{grid-template-columns:repeat(3,1fr)}.dash{grid-template-columns:1fr}}
.ch{display:flex;align-items:center;gap:8px;padding:14px 16px 6px}.ch h3{margin:0;font-size:16px}.ch .more{margin-left:auto;color:var(--pri);font-size:13px;font-weight:600}
.labs6{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:10px;padding:8px 16px 14px}
.lt{border-radius:12px;padding:11px 12px;border:1px solid var(--line);text-align:left;min-height:84px}
.lt .n{font-size:12.5px;color:var(--sub);display:flex;justify-content:space-between;gap:6px}.lt b{display:block;font-size:15px;margin-top:5px}.lt span{font-size:12.5px;color:var(--sub)}
.lt.busy{border-color:transparent}.lt.free b{color:var(--mute);font-weight:600}
.st{font-size:11px;border-radius:6px;padding:0 6px;font-weight:700;white-space:nowrap;line-height:18px}.st.b{background:var(--pri);color:#fff}.st.f{background:var(--ok-weak);color:var(--ok)}
.heat{padding:4px 16px 16px;overflow-x:auto}.heat table{width:100%;border-collapse:separate;border-spacing:3px;min-width:520px;table-layout:fixed}
.heat td{height:28px;border-radius:6px;font-size:11.5px;text-align:center;color:var(--sub);background:var(--surface2);white-space:nowrap;overflow:hidden;max-width:80px;text-overflow:ellipsis}
.heat td.h{background:transparent;text-align:left;font-weight:700;width:48px}.heat tr.hd td{background:transparent;height:18px}
.heat td.x{color:var(--text);font-weight:600;cursor:pointer}.heat td.nw{outline:2px solid var(--pri)}
.lst .it{display:flex;gap:10px;align-items:center;padding:10px 16px;border-top:1px solid var(--line);font-size:14px;width:100%;text-align:left}.lst .it .r{margin-left:auto;color:var(--sub);font-size:13px;text-align:right}
.dt{width:10px;height:10px;border-radius:50%;flex:none}

/* ---------- 課表 ---------- */
.tbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:0 0 12px}
.labsel{display:flex;gap:6px;overflow-x:auto;scrollbar-width:none;padding-bottom:2px}.labsel::-webkit-scrollbar{display:none}
.split{display:grid;grid-template-columns:1fr;gap:16px;align-items:start}
@media (min-width:1100px){.split.has{grid-template-columns:1fr 340px}}
.wkwrap{overflow:auto}
table.wk{width:100%;border-collapse:separate;border-spacing:0;table-layout:fixed;min-width:540px}
table.wk th{font-size:12.5px;color:var(--sub);font-weight:700;padding:9px 4px;border-bottom:1px solid var(--line);background:var(--surface);position:sticky;top:0;z-index:2}
table.wk th.today{color:var(--pri)}table.wk th.today div{background:var(--pri);color:#fff;border-radius:8px;padding:1px 0;margin:2px 6px 0}
table.wk td{border-bottom:1px solid var(--line);border-right:1px solid var(--line);height:60px;padding:3px;vertical-align:top}table.wk td:last-child{border-right:none}
table.wk td.pr{font-size:11.5px;color:var(--sub);text-align:center;vertical-align:middle;background:var(--surface);position:sticky;left:0;z-index:1}table.wk td.pr b{display:block;color:var(--text);font-size:14px}
table.wk tr.nowr td.pr{background:var(--pri-weak)}table.wk tr.nowr td.pr b{color:var(--pri)}
.ev{display:block;width:100%;border-radius:9px;padding:5px 7px;font-size:12px;line-height:1.3;text-align:left;position:relative;margin-bottom:3px;color:var(--text)}
.ev b{display:block;font-size:12.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.ev.sel{outline:2.5px solid var(--pri);outline-offset:1px}
.ev .td{position:absolute;right:4px;top:3px;font-size:10.5px;background:var(--surface);border-radius:6px;padding:0 4px;color:var(--warn);font-weight:700}
.mon{width:100%;border-collapse:separate;border-spacing:4px;table-layout:fixed;min-width:560px}.mon th{font-size:12.5px;color:var(--sub);padding:4px}
.mon td{background:var(--surface);border-radius:10px;vertical-align:top;height:92px;padding:4px 5px;border:1px solid var(--line)}.mon td.out{opacity:.35}.mon td.today{border:2px solid var(--pri)}
.mon .dn{font-weight:800;font-size:12.5px;color:var(--sub)}.mon .e{display:block;width:100%;text-align:left;border-radius:5px;padding:1px 4px;margin-top:2px;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--text)}
.lrow{display:flex;gap:10px;align-items:center;padding:10px 14px;border-bottom:1px solid var(--line);width:100%;text-align:left}.lrow:last-child{border:none}
.lrow .d{width:74px;flex:none;font-size:13px;color:var(--sub)}.lrow .t{font-weight:700}.lrow .s{font-size:12.5px;color:var(--sub)}
.det{padding:16px}.det h3{margin:6px 0 2px;font-size:21px}.det .m{color:var(--sub);font-size:13.5px;margin-top:5px;display:flex;gap:6px;align-items:flex-start}
.det .blk{border-top:1px solid var(--line);margin-top:14px;padding-top:12px}.det .blk>b{font-size:13px;color:var(--sub);display:block;margin-bottom:8px}
.kit{display:flex;justify-content:space-between;gap:10px;font-size:13.5px;padding:6px 0;border-bottom:1px dashed var(--line)}.kit:last-child{border:none}.kit .ok{color:var(--ok);font-weight:700;white-space:nowrap}.kit .ng{color:var(--bad);font-weight:700;white-space:nowrap}
.mini{display:flex;gap:10px;align-items:center;padding:6px 0;font-size:14px}.mini.done span{text-decoration:line-through;color:var(--mute)}
.btns2{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:14px}
.nmrow{display:grid;grid-template-columns:1fr 76px;gap:8px}

/* ---------- 盤點 ---------- */
.prog{padding:12px 14px}.prog .l{display:flex;justify-content:space-between;font-size:13.5px;color:var(--sub);gap:8px}
.bar{height:8px;background:var(--line);border-radius:6px;margin-top:8px;overflow:hidden}.bar i{display:block;height:100%;background:var(--ok);border-radius:6px;transition:width .3s}
.cur{background:var(--surface);border-radius:20px;box-shadow:0 4px 20px rgba(16,24,40,.10);padding:16px;border:2px solid var(--pri)}
.cur .loc{font-size:12.5px;color:var(--sub);display:flex;gap:6px;align-items:center}.cur .nm{font-size:23px;font-weight:800;margin:6px 0 2px}.cur .fx{color:var(--sub);font-size:14px}
.lastq{display:flex;gap:8px;margin:12px 0;font-size:13px}.lastq div{flex:1;background:var(--bg);border-radius:10px;padding:8px 10px;color:var(--sub)}.lastq b{display:block;color:var(--text);font-size:16px;margin-top:2px}
.stepper{display:grid;grid-template-columns:64px 1fr 64px;gap:10px;align-items:center}
.stepper .b{height:62px;border-radius:16px;background:var(--pri-weak);color:var(--pri);display:grid;place-items:center;font-size:30px;font-weight:700}
.stepper .v{height:62px;border-radius:16px;border:2px solid var(--line);display:flex;align-items:center;justify-content:center;gap:6px;background:var(--surface)}
.stepper .v input{width:100%;border:none;background:none;text-align:center;font-size:30px;font-weight:800;outline:none;padding:0}.stepper .v small{font-size:15px;color:var(--sub);font-weight:600;padding-right:10px;white-space:nowrap}
.acts{display:grid;grid-template-columns:1fr 2fr;gap:10px;margin-top:12px}.acts .btn{white-space:nowrap;padding:0 10px}
.cwrap{max-width:620px}
.nx .it{display:flex;align-items:center;gap:10px;padding:10px 14px;border-bottom:1px solid var(--line);font-size:14px;width:100%;text-align:left}.nx .it:last-child{border:none}.nx .q{margin-left:auto;color:var(--sub);font-size:13px;white-space:nowrap}
@media (min-width:1100px){.cgrid{display:grid;grid-template-columns:minmax(0,620px) 1fr;gap:20px;align-items:start}}

/* ---------- 新增準備事項 ---------- */
.lesson{display:flex;gap:8px;overflow-x:auto;padding-bottom:2px}.lesson button{flex:none;border:1.5px solid var(--line);border-radius:10px;padding:7px 10px;font-size:13px;background:var(--surface)}
.lesson button.on{border-color:var(--pri);background:var(--pri-weak);color:var(--pri);font-weight:700}
.row3{display:grid;grid-template-columns:1fr 72px 62px 30px;gap:8px;align-items:center;margin-top:8px;position:relative}
.sug{position:absolute;left:0;right:0;top:48px;z-index:5;background:var(--surface);border:1px solid var(--line);border-radius:12px;overflow:hidden;box-shadow:0 8px 20px rgba(0,0,0,.12)}
.sug button{display:block;width:100%;text-align:left;padding:9px 12px;border-bottom:1px solid var(--line);font-size:14px}.sug button:last-child{border:none}.sug small{color:var(--sub)}
.info{grid-column:1/5;font-size:12px;color:var(--sub);margin-top:-2px}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:0 10px}

/* ---------- 更多 ---------- */
.menu{background:var(--surface);border-radius:var(--r);box-shadow:var(--shadow);overflow:hidden}
.menu button{display:flex;align-items:center;gap:12px;padding:14px 16px;border-bottom:1px solid var(--line);width:100%;text-align:left;font-size:15px}.menu button:last-child{border:none}
.menu .ic{width:36px;height:36px;border-radius:10px;display:grid;place-items:center;background:var(--pri-weak);color:var(--pri)}.menu .r{margin-left:auto;color:var(--mute);font-size:13px}
.res{background:var(--surface);border-radius:var(--r);box-shadow:var(--shadow);padding:12px 14px;margin-bottom:8px}.res .t{font-weight:700}.res .s{font-size:13px;color:var(--sub);margin-top:3px;line-height:1.5}
.rgrid{display:grid;grid-template-columns:1fr;gap:0 12px}@media (min-width:900px){.rgrid{grid-template-columns:1fr 1fr}}

@media (max-width:899px){
  table.wk{min-width:0}table.wk td{padding:2px;height:54px}.ev{padding:4px 5px;font-size:11px;border-radius:7px}.ev b{font-size:11.5px}
  table.wk th{font-size:11.5px;padding:7px 2px}table.wk td.pr{font-size:10.5px}
  .det h3{font-size:20px}
}

.buyb{display:inline-block;font-size:12px;font-weight:700;color:var(--warn);background:var(--warn-weak);border-radius:6px;padding:1px 7px;margin-left:6px}.buyb.ok{color:var(--ok);background:var(--ok-weak)}
.buyck{display:inline-flex;gap:5px;align-items:center;color:var(--warn);font-weight:600;font-size:12.5px;padding:4px 0}.buyck input{width:16px;height:16px}
</style></head><body>
<svg width="0" height="0" style="position:absolute" aria-hidden="true">
  <symbol id="home" viewBox="0 0 24 24"><path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V20h14V9.5"/><path d="M10 20v-5h4v5"/></symbol>
  <symbol id="cal" viewBox="0 0 24 24"><rect x="3" y="4.5" width="18" height="16" rx="3"/><path d="M3 9.5h18M8 3v3M16 3v3"/></symbol>
  <symbol id="todo" viewBox="0 0 24 24"><rect x="4" y="3.5" width="16" height="17" rx="3"/><path d="m8 11 2.5 2.5L16 8M8 17h8"/></symbol>
  <symbol id="count" viewBox="0 0 24 24"><path d="M9 4h6a1 1 0 0 1 1 1v1H8V5a1 1 0 0 1 1-1Z"/><path d="M8 5H6a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7a2 2 0 0 0-2-2h-2"/><path d="M8 12h8M8 16h5"/></symbol>
  <symbol id="grid" viewBox="0 0 24 24"><rect x="4" y="4" width="6.5" height="6.5" rx="2"/><rect x="13.5" y="4" width="6.5" height="6.5" rx="2"/><rect x="4" y="13.5" width="6.5" height="6.5" rx="2"/><rect x="13.5" y="13.5" width="6.5" height="6.5" rx="2"/></symbol>
  <symbol id="search" viewBox="0 0 24 24"><circle cx="11" cy="11" r="6.5"/><path d="m20 20-4.2-4.2"/></symbol>
  <symbol id="plus" viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></symbol>
  <symbol id="x" viewBox="0 0 24 24"><path d="M6 6l12 12M18 6 6 18"/></symbol>
  <symbol id="flask" viewBox="0 0 24 24"><path d="M9 3h6M10 3v6L4.5 18.5A1.8 1.8 0 0 0 6 21h12a1.8 1.8 0 0 0 1.5-2.5L14 9V3"/><path d="M7.5 15h9"/></symbol>
  <symbol id="box" viewBox="0 0 24 24"><path d="m12 3 8 4.5v9L12 21l-8-4.5v-9Z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/></symbol>
  <symbol id="cart" viewBox="0 0 24 24"><path d="M3 4h2.5l2.2 10.5h10.6L20.5 7H7"/><circle cx="9.5" cy="19" r="1.5"/><circle cx="17" cy="19" r="1.5"/></symbol>
  <symbol id="swap" viewBox="0 0 24 24"><path d="M7 7h12l-3-3M17 17H5l3 3"/></symbol>
  <symbol id="alert" viewBox="0 0 24 24"><path d="M12 4 2.8 19.5h18.4Z"/><path d="M12 10v4M12 17h.01"/></symbol>
  <symbol id="pin" viewBox="0 0 24 24"><path d="M12 21s-6.5-6-6.5-11a6.5 6.5 0 0 1 13 0c0 5-6.5 11-6.5 11Z"/><circle cx="12" cy="10" r="2.3"/></symbol>
  <symbol id="check" viewBox="0 0 24 24"><path d="m5 12.5 4.5 4.5L19 7"/></symbol>
  <symbol id="print" viewBox="0 0 24 24"><path d="M7 9V3.5h10V9"/><rect x="3.5" y="9" width="17" height="8" rx="2"/><path d="M7 14h10v6.5H7z"/></symbol>
  <symbol id="left" viewBox="0 0 24 24"><path d="m15 5-7 7 7 7"/></symbol>
  <symbol id="right" viewBox="0 0 24 24"><path d="m9 5 7 7-7 7"/></symbol>
  <symbol id="moon" viewBox="0 0 24 24"><path d="M20 14.5A8 8 0 1 1 9.5 4a6.5 6.5 0 0 0 10.5 10.5Z"/></symbol>
  <symbol id="clock" viewBox="0 0 24 24"><circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/></symbol>
  <symbol id="refresh" viewBox="0 0 24 24"><path d="M20 11a8 8 0 0 0-14.5-4.5L4 8M4 4v4h4M4 13a8 8 0 0 0 14.5 4.5L20 16M20 20v-4h-4"/></symbol>
</svg>
<div class="app"><aside class="side" id="side"></aside><main class="main" id="main"></main></div>
<nav class="nav" id="nav"></nav>
<button class="fab hide" id="fab" aria-label="新增準備事項"><svg class="i"><use href="#plus"/></svg></button>
<div class="toast" id="toast"></div>
<script>// 吉祥物（自己畫的 SVG，沒有版權問題）：無尾熊「考拉老師」、戴護目鏡的「實驗小鴨」
var ART = {
  // 無尾熊頭像（標誌、問候）
  koala: function (size) {
    return '<svg class="art" width="' + size + '" height="' + size + '" viewBox="0 0 120 120" aria-hidden="true">' +
      '<circle cx="27" cy="42" r="23" fill="#8E9AA8"/><circle cx="93" cy="42" r="23" fill="#8E9AA8"/>' +
      '<circle cx="28" cy="44" r="13" fill="#F6C9D3"/><circle cx="92" cy="44" r="13" fill="#F6C9D3"/>' +
      '<ellipse cx="60" cy="66" rx="40" ry="36" fill="#AEB9C6"/>' +
      '<ellipse cx="60" cy="80" rx="24" ry="17" fill="#C9D2DC"/>' +
      '<circle cx="43" cy="61" r="5" fill="#2B303A"/><circle cx="77" cy="61" r="5" fill="#2B303A"/>' +
      '<circle cx="44.6" cy="59.4" r="1.7" fill="#fff"/><circle cx="78.6" cy="59.4" r="1.7" fill="#fff"/>' +
      '<ellipse cx="60" cy="73" rx="10" ry="12.5" fill="#39404B"/><ellipse cx="56.5" cy="67.5" rx="3" ry="2" fill="#5E6674"/>' +
      '<path d="M54 88q6 4 12 0" stroke="#39404B" stroke-width="2.4" fill="none" stroke-linecap="round"/>' +
      '<ellipse cx="33" cy="76" rx="6.5" ry="4" fill="#F59DB0" opacity=".55"/><ellipse cx="87" cy="76" rx="6.5" ry="4" fill="#F59DB0" opacity=".55"/>' +
      '</svg>';
  },
  // 睡覺的無尾熊（今天沒有課）
  koalaSleep: function (size) {
    return '<svg class="art" width="' + size + '" height="' + size + '" viewBox="0 0 120 120" aria-hidden="true">' +
      '<circle cx="27" cy="46" r="21" fill="#8E9AA8"/><circle cx="93" cy="46" r="21" fill="#8E9AA8"/>' +
      '<circle cx="28" cy="48" r="12" fill="#F6C9D3"/><circle cx="92" cy="48" r="12" fill="#F6C9D3"/>' +
      '<ellipse cx="60" cy="70" rx="40" ry="34" fill="#AEB9C6"/><ellipse cx="60" cy="84" rx="24" ry="15" fill="#C9D2DC"/>' +
      '<path d="M37 64q6 5 12 0M71 64q6 5 12 0" stroke="#2B303A" stroke-width="2.6" fill="none" stroke-linecap="round"/>' +
      '<ellipse cx="60" cy="76" rx="9.5" ry="11.5" fill="#39404B"/>' +
      '<ellipse cx="33" cy="79" rx="6.5" ry="4" fill="#F59DB0" opacity=".55"/><ellipse cx="87" cy="79" rx="6.5" ry="4" fill="#F59DB0" opacity=".55"/>' +
      '<text x="92" y="22" font-size="16" font-weight="700" fill="#7C8CA3" font-family="sans-serif">Z</text>' +
      '<text x="104" y="12" font-size="11" font-weight="700" fill="#A3B1C4" font-family="sans-serif">z</text>' +
      '</svg>';
  },
  // 戴護目鏡的小鴨（全部準備好了）
  duck: function (size) {
    return '<svg class="art" width="' + size + '" height="' + size + '" viewBox="0 0 120 120" aria-hidden="true">' +
      '<ellipse cx="60" cy="108" rx="34" ry="5" fill="#000" opacity=".08"/>' +
      '<ellipse cx="50" cy="104" rx="9" ry="4" fill="#FF9B3D"/><ellipse cx="72" cy="104" rx="9" ry="4" fill="#FF9B3D"/>' +
      '<ellipse cx="64" cy="80" rx="36" ry="26" fill="#FFD34E"/>' +
      '<path d="M96 70q12-8 10 8q-4 6-12 4z" fill="#FFD34E"/>' +
      '<ellipse cx="72" cy="80" rx="16" ry="10" fill="#F6BD2A" transform="rotate(-12 72 80)"/>' +
      '<circle cx="50" cy="45" r="24" fill="#FFD34E"/>' +
      '<path d="M50 21q2-9 9-8q-4 3-3 9z" fill="#F6BD2A"/>' +
      '<path d="M24 50q-12 1-13 6q7 5 17 2z" fill="#FF9B3D"/>' +
      '<path d="M36 36q16-6 34 2" stroke="#2F6BE0" stroke-width="4" fill="none" stroke-linecap="round"/>' +
      '<circle cx="47" cy="44" r="9.5" fill="#CFEAFF" stroke="#2F6BE0" stroke-width="3"/>' +
      '<circle cx="47" cy="45" r="3.6" fill="#2B303A"/><circle cx="48.4" cy="43.6" r="1.3" fill="#fff"/>' +
      '<path d="M42 39.5l4-2" stroke="#fff" stroke-width="2" stroke-linecap="round"/>' +
      '<ellipse cx="60" cy="56" rx="5" ry="3" fill="#FF8FA3" opacity=".55"/>' +
      '</svg>';
  },
  // 小鴨拿燒瓶（讀取中）
  duckFlask: function (size) {
    return '<svg class="art bob" width="' + size + '" height="' + size + '" viewBox="0 0 120 120" aria-hidden="true">' +
      '<ellipse cx="58" cy="82" rx="32" ry="24" fill="#FFD34E"/>' +
      '<circle cx="48" cy="48" r="22" fill="#FFD34E"/>' +
      '<path d="M24 52q-11 1-12 6q7 4 16 2z" fill="#FF9B3D"/>' +
      '<circle cx="46" cy="46" r="3.6" fill="#2B303A"/><circle cx="47.3" cy="44.7" r="1.3" fill="#fff"/>' +
      '<ellipse cx="58" cy="58" rx="4.5" ry="2.6" fill="#FF8FA3" opacity=".55"/>' +
      '<path d="M82 52h12M85 52v14l-10 20a5 5 0 0 0 4.5 7h17a5 5 0 0 0 4.5-7l-10-20V52" fill="#E8F6FF" stroke="#2F6BE0" stroke-width="3" stroke-linejoin="round"/>' +
      '<path d="M78.5 80h24.5l3.5 7a3 3 0 0 1-2.7 4.3H77.7A3 3 0 0 1 75 87z" fill="#7DD3A8"/>' +
      '<circle cx="90" cy="44" r="2.5" fill="#7DD3A8"/><circle cx="96" cy="37" r="1.8" fill="#7DD3A8"/>' +
      '<ellipse cx="74" cy="80" rx="12" ry="8" fill="#F6BD2A" transform="rotate(-30 74 80)"/>' +
      '</svg>';
  },
};

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
  String(str || '').split(/[,，、\\s]+/).forEach(function (tok) {
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
    '<div class="s">' + where + (t.place || t.have !== '' ? '<br>' + ic('pin', 's') + ' ' + esc(t.place || '—') + (t.have !== '' && t.have != null ? '　庫存 ' + esc(t.have) + ' ' + esc(t.unit) : '') : '') +
    (t.short ? '　<span class="late">⚠ 可能不夠</span>' : '') + (t.note ? '<br>📝 ' + esc(t.note) : '') +
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
      var known = r.have !== '' && !isNaN(r.have);
      var st = r.short ? '<span class="ng">有 ' + esc(r.have) + ' ' + esc(r.unit) + ' ✘</span>' : known ? '<span class="ok">有 ' + esc(r.have) + ' ' + esc(r.unit) + ' ✔</span>' :
        '<span class="muted" style="white-space:nowrap">' + esc(r.have || '—') + '（請自行確認）</span>';
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
        toast(msg.split('\\n')[0]); S.week.key = '';
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
    s.innerHTML = hit.map(function (x, i) { return '<button data-h="' + i + '">' + esc(x[1]) + ' <small>' + esc(x[2]) + '　' + esc(x[5]) + '・庫存 ' + esc(x[4] || '—') + ' ' + esc(x[3]) + '</small></button>'; }).join('');
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
  var pa = String(a).match(/\\d+|\\D+/g) || [], pb = String(b).match(/\\d+|\\D+/g) || [];
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
        closeSheet(); toast(String(msg).split('\\n')[0] || '完成'); S.count.data = null; refresh(false);
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
</script>
</body></html>
`;
