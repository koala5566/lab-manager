"""依盤點紀錄推算建議安全存量，產出「安全存量建議.xlsx」讓承辦人勾選後匯入。

用法：python3 suggest_safety_stock.py <新系統初始資料.xlsx> <輸出.xlsx>

規則（只看藥品、耗材；器材不設）：
1. 有消耗：安全存量 ≈ 一學期的用量（歷次盤點「減少」的總和 ÷ 經過幾學期），
   不到 2 的取到 0.5，其餘無條件進位，最少 1。→ 剩下不夠用一學期時提醒。
2. 沒有消耗但只剩 1～2（瓶／包…）：設 1（剩不到 1 時提醒）；本來就不到 1 的設 0.5。
3. 備註或自己筆記寫「需購買／需補充／待買」：設成「目前整數＋1」，匯入後馬上列入需補充。
低可信度（預設不採用）：歷史有明顯打錯的數字、數量曾用「6包+50個」這類混合寫法。
"""
import collections
import math
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

src, dst = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
it = wb['品項']
H = [c.value for c in it[1]]
items = [dict(zip(H, [c.value for c in r])) for r in it.iter_rows(min_row=2)]
rec = wb['盤點紀錄']
RH = [c.value for c in rec[1]]
hist = collections.defaultdict(lambda: collections.defaultdict(float))
mixed = set()
for r in rec.iter_rows(min_row=2, values_only=True):
    d = dict(zip(RH, r))
    if isinstance(d['數量'], (int, float)):
        hist[d['編號']][d['盤點日期']] += d['數量']
    note = str(d['數量說明'] or '')
    if '+' in note and any(u in note for u in ('個', '包', '箱', '瓶', '支')):
        mixed.add(d['編號'])


def clean_series(vals):
    """去掉明顯打錯的數字：比其他值大 10 倍以上，或前後相同只有中間一次突然不同。"""
    out, dropped = list(vals), []
    others = sorted(vals)
    med = others[len(others) // 2] if others else 0
    for i, v in enumerate(vals):
        if med > 0 and v > med * 10:
            dropped.append(v)
            out[i] = None
    for i in range(1, len(vals) - 1):
        a, b, c = vals[i - 1], vals[i], vals[i + 1]
        if out[i] is not None and a == c and b != a and (b == 0 or a == 0):
            dropped.append(b)
            out[i] = None
    return [v for v in out if v is not None], dropped


rows = []
for x in items:
    if x['狀態'] == '已淘汰' or x['類別'] == '器材':
        continue
    h = sorted(hist[x['編號']].items())
    if not h:
        continue
    raw = [v for _, v in h]
    vals, dropped = clean_series(raw)
    dec = sum(max(0, a - b) for a, b in zip(vals, vals[1:]))
    inc = sum(max(0, b - a) for a, b in zip(vals, vals[1:]))
    span = (h[-1][0] - h[0][0]).days / 182.5 if len(h) > 1 else 0
    per = dec / span if span >= 0.5 else 0
    latest = raw[-1]
    note = str(x['備註'] or '') + str(x['自己筆記'] or '')
    wants = any(k in note for k in ('需購買', '需補充', '待買', '需買'))
    unit = x['單位'] or ('瓶' if x['類別'] == '藥品' else '')
    sug, why = None, ''
    if dec > 0 and per > 0:
        sug = max(1, math.ceil(per * 2) / 2 if per < 2 else math.ceil(per))
        why = f'約 {span:.1f} 學期共用掉 {dec:g}{unit}，每學期約 {per:.1f}'
        if inc > 0:
            why += f'（期間補進 {inc:g}）'
    elif 0 < latest <= 2:
        sug = 1 if latest >= 1 else 0.5
        why = f'沒有消耗紀錄，但只剩 {latest:g}{unit}'
    if wants and (sug is None or sug <= latest):
        sug = math.floor(latest) + 1
        why = (why + '；' if why else '') + '備註寫「需購買／需補充」'
    if sug is None:
        continue
    sug = int(sug) if float(sug).is_integer() else sug
    low = []
    if dropped:
        low.append('歷史有疑似打錯的數字（' + '、'.join(f'{v:g}' for v in dropped) + '），已排除')
    if x['編號'] in mixed:
        low.append('數量曾用混合寫法（例：6包+50個），數字可能不準')
    rows.append({
        '採用': not (x['編號'] in mixed), '編號': x['編號'], '類別': x['類別'], '清單分區': x['清單分區'],
        '品名': x['品名'], '單位': unit, '歷次數量（舊→新）': ' → '.join(f'{v:g}' for v in raw),
        '最新數量': latest, '每學期約用掉': round(per, 1) if per else '', '建議安全存量': sug,
        '理由': why, '請確認': '；'.join(low), '匯入後馬上列入需補充': '是' if latest < sug else '',
    })

rows.sort(key=lambda r: (r['類別'] != '藥品', r['匯入後馬上列入需補充'] != '是', r['理由'].startswith('沒有'), r['編號']))
cols = list(rows[0].keys())
out = openpyxl.Workbook()
ws = out.active
ws.title = '安全存量建議'
ws.append(cols)
for r in rows:
    ws.append([r[c] for c in cols])
for c in ws[1]:
    c.font = Font(bold=True)
    c.fill = PatternFill('solid', fgColor='DDEBF7')
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
yellow = PatternFill('solid', fgColor='FFF2CC')
for r in range(2, ws.max_row + 1):
    ws.cell(r, 1).fill = yellow
    ws.cell(r, cols.index('建議安全存量') + 1).fill = yellow
    if ws.cell(r, cols.index('請確認') + 1).value:
        ws.cell(r, cols.index('請確認') + 1).font = Font(color='C00000')
dv = DataValidation(type='list', formula1='"TRUE,FALSE"', allow_blank=False)
ws.add_data_validation(dv)
dv.add(f'A2:A{ws.max_row}')
for i, w in enumerate([6, 9, 6, 11, 24, 5, 34, 8, 8, 10, 36, 34, 10], 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
ws.freeze_panes = 'C2'
ws.auto_filter.ref = ws.dimensions
out.save(dst)

c = collections.Counter((r['類別'], '消耗' if r['理由'].startswith('約') else ('少量' if r['理由'].startswith('沒有') else '備註')) for r in rows)
print('建議', len(rows), dict(c))
print('匯入後馬上列入需補充', sum(r['匯入後馬上列入需補充'] == '是' for r in rows))
print('預設不採用', [r['品名'] for r in rows if not r['採用']])
print('請確認', [(r['品名'], r['請確認']) for r in rows if r['請確認']])
print('馬上列入：', [(r['品名'], r['最新數量'], r['建議安全存量']) for r in rows if r['匯入後馬上列入需補充'] == '是' and r['採用']])
