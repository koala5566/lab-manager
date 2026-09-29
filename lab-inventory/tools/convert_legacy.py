"""把三個舊 Excel 清單轉成新系統的「新系統初始資料.xlsx」。

用法：python3 convert_legacy.py <舊檔所在資料夾> <輸出檔路徑>

原則：品名、化學式、數字一律照原樣；疑似錯誤只列在「轉入檢查」，不自動更正。
"""
import datetime as dt
import glob
import os
import re
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

IMPORT_STAMP = '舊資料轉入'

# ---------------------------------------------------------------- 工具函式


def clean(v):
    """儲存格文字：換行改空白、去頭尾空白；數字 29.0 → 29。"""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).replace('\n', ' ').replace('\r', ' ')
    s = re.sub(r' {2,}', ' ', s)
    return s.strip()


def raw(v):
    """保留原字串（含前後空白），只把換行改空白；用於品名，讓多餘空白能被列入檢查。"""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r'\s*\n\s*', ' ', str(v))


def roc(s):
    """'114.2.8' / '115.08.26' → date。"""
    y, m, d = [int(x) for x in s.split('.')]
    return dt.date(y + 1911, m, d)


def roc_text(d):
    return f'{d.year - 1911}.{d.month:02d}.{d.day:02d}'


def term_label(d):
    """依日期給一個好認的盤點名稱。"""
    y = d.year - 1911
    if d.month >= 8:
        sy, sem = y, 1
    elif d.month == 1:
        sy, sem = y - 1, 1
    else:
        sy, sem = y - 1, 2
    if sem == 1:
        stage = {8: '期初', 9: '期初', 10: '期中', 11: '期中', 12: '期中', 1: '期末'}[d.month]
    else:
        stage = {2: '期初', 3: '期初', 4: '期中', 5: '期中', 6: '期末', 7: '期末'}[d.month]
    return f'{sy}-{sem}{stage}'


NUM = r'\d+(?:\.\d+)?'
FORM_WORDS = ('粉狀', '固體顆粒', '固', '液')


def parse_qty(v):
    """回傳 (數量數字或 None, 數量說明)。純數字時說明為空。"""
    if v is None or v == '':
        return None, ''
    if isinstance(v, (int, float)):
        n = float(v)
        return (int(n) if n.is_integer() else n), ''
    text = clean(v)
    if re.fullmatch(NUM, text):
        n = float(text)
        return (int(n) if n.is_integer() else n), ''
    # a*b 先相乘：100g*2+50g
    parts = re.findall(r'(' + NUM + r')\s*([^\d\s+*()（）]*)\s*(?:\*\s*(' + NUM + r'))?', text)
    if not parts:
        return None, text
    units = [u for _, u, _ in parts]
    vals = [float(a) * (float(b) if b else 1) for a, _, b in parts]
    if len(set(units)) == 1 or all(u in FORM_WORDS for u in units):
        n = sum(vals)
    else:
        n = vals[0]
    n = round(n, 3)
    return (int(n) if float(n).is_integer() else n), text


def pos_key(s):
    """排序位置自然排序：'3-10' 排在 '3-9' 後面。"""
    return [int(t) if t.isdigit() else t for t in re.split(r'(\d+)', str(s))]


# ---------------------------------------------------------------- 資料容器

items = []        # dict，欄位見 ITEM_COLS
records = []      # (date, 編號, 品名, 存放處, 數量, 說明)
moves = []        # 異動紀錄
checks = []       # 轉入檢查
kits = []         # 實驗套組
slides_need = []  # 玻片需求

ITEM_COLS = ['編號', '類別', '科別', '清單分區', '品名', '化學式或規格', '單位', '教室', '櫃別',
             '排序位置', '分處存放', '安全存量', 'SDS', '危險物品', '危險分類', '狀態', '備註',
             '自己筆記', '原清單序號', '最新盤點日期', '最新數量', '最新數量說明']

counters = {'藥品': 0, '器材': 0, '耗材': 0}
PREFIX = {'藥品': '藥', '器材': '器', '耗材': '耗'}


def new_item(**kw):
    it = {c: '' for c in ITEM_COLS}
    it.update(kw)
    it['狀態'] = it['狀態'] or '使用中'
    counters[it['類別']] += 1
    it['編號'] = f"{PREFIX[it['類別']]}{counters[it['類別']]:04d}"
    items.append(it)
    return it


def add_rec(d, it, qty_cell, place=''):
    n, note = parse_qty(qty_cell)
    if n is None and not note:
        return
    records.append([d, it['編號'], it['品名'].strip(), place, n, note])


def add_check(kind, it, field, current, suggest, note, apply=False):
    checks.append({
        '類型': kind, '編號': it['編號'] if it else '', '品名': it['品名'] if it else '',
        '欄位': field, '目前內容': current, '建議內容': suggest, '說明': note,
        '可套用': apply,
    })


def join_notes(*xs):
    return '；'.join(clean(x) for x in xs if clean(x))


# ---------------------------------------------------------------- 找檔案

src = sys.argv[1]
out_path = sys.argv[2]
books = {}
for f in glob.glob(os.path.join(src, '*.xlsx')):
    wb = openpyxl.load_workbook(f, data_only=True)
    names = wb.sheetnames
    if '115-1藥品清單最新' in names:
        books['drug'] = wb
    elif '化準114器材區' in names:
        books['chem'] = wb
    elif '耗材' in names and '顯微鏡觀察' in names:
        books['bio'] = wb
assert set(books) == {'drug', 'chem', 'bio'}, books.keys()

# ================================================================ 藥品清單

drug = books['drug']
ws = drug['115-1藥品清單最新']
old114_1 = drug['114-1藥品清單最新']   # 114.8.21 在 J 欄
old114_2 = drug['114-2藥品清單最新']   # 115.06.25 在 J 欄
old113_1 = drug['113-1藥品清單+SDS']   # 113.12.5 在 G 欄，SDS 在 H 欄
old113_2 = drug['113-2藥品清單(減少)']


def by_name(sheet, first_row, cols):
    out = {}
    for r in range(first_row, sheet.max_row + 1):
        n = raw(sheet.cell(r, 2).value).strip()
        if n and sheet.cell(r, 1).value is not None:
            out[n] = [sheet.cell(r, c).value for c in cols]
    return out


h_1128 = by_name(old113_1, 3, [7, 8])       # 113.12.5 數量, SDS
h_0821 = by_name(old114_1, 4, [10])         # 114.8.21
h_0625 = by_name(old114_2, 4, [10])         # 115.06.25

# 生物耗材中與藥品重複、合併為一筆的品項（承辦人決定：合併，化準與生準冰箱分處存放）
MERGE_BIO = {'抗A血清', '抗B血清', '腎上腺素', '油性椒紅色素', '雙縮脲試劑A', '雙縮脲試劑B'}
merged = {}   # 品名 → item
# 一列兩種藥品，承辦人決定拆成兩筆：原品名 → [(新品名, 化學式或規格)]
SPLIT = {
    '順/反 丁烯二酸': [('順丁烯二酸', 'C4H4O4'), ('反丁烯二酸', 'C4H4O4')],
    '雙縮脲試劑': [('雙縮脲試劑A', 'A劑'), ('雙縮脲試劑B', 'B劑')],
}


def per_part(v):
    """拆開後每一筆的數量：「各1」→ 1、「各10小瓶」→ 10小瓶。"""
    if isinstance(v, str) and clean(v).startswith('各'):
        return clean(v)[1:]
    return v

D_1208, D_0208, D_0513, D_0710 = roc('113.12.5'), roc('114.2.8'), roc('114.5.13'), roc('114.7.10')
D_0821, D_0625, D_0826 = roc('114.8.21'), roc('115.06.25'), roc('115.08.26')


def drug_row(it, name, orig, split, g, h, i, j, k):
    if name in MERGE_BIO:
        it['科別'] = '生物'
        it['分處存放'] = '化準、生準冰箱'
        merged[name] = it
    if name == '硫酸鉻鉀':
        it['SDS'] = '無'
        add_check('轉入說明', it, 'SDS', '', '無', '115-1 新增，舊檔沒有 SDS 紀錄；承辦人不確定，先記為「無」')

    # 115.08.26：K 欄有數字時，K 才是 115.08.26 的數量（J 為上次的數字）
    latest = j
    if k is not None and (isinstance(k, (int, float)) or re.fullmatch(NUM, str(k).strip())):
        latest = k
        add_check('轉入說明', it, '盤點數量', f'J欄 {clean(j)}／K欄 {clean(k)}', clean(k),
                  '依承辦人說明：K 欄是 115.08.26 最新數量，J 欄為上次數字')
    hist = [(D_1208, h_1128.get(orig, [None])[0]), (D_0208, g), (D_0513, h), (D_0710, i),
            (D_0821, h_0821.get(orig, [None])[0]), (D_0625, h_0625.get(orig, [None])[0]),
            (D_0826, latest)]
    for d, v in hist:
        if split:
            v = per_part(v)
        if name == '油性椒紅色素' and isinstance(v, str) and '化準' in v:
            m = re.fullmatch(r'化準(\S+?)生準(\S+)', clean(v))
            add_rec(d, it, m.group(1), '化準')
            add_rec(d, it, m.group(2), '生準冰箱')
        elif name in MERGE_BIO:
            add_rec(d, it, v, '生準冰箱')
        else:
            add_rec(d, it, v)
    if isinstance(k, str) and '新購' in k:
        m = re.match(r'(\d+\.\d+\.\d+)新購(\d+)', k.strip())
        if m:
            moves.append([roc(m.group(1)), it['編號'], name, '新購', int(m.group(2)), clean(k), ''])


for r in range(4, ws.max_row + 1):
    seq = ws.cell(r, 1).value
    name_raw = raw(ws.cell(r, 2).value)
    if seq is None or not name_raw.strip():
        continue
    name = name_raw.strip()
    formula = raw(ws.cell(r, 3).value)
    room, cab, pos = clean(ws.cell(r, 4).value), clean(ws.cell(r, 5).value), clean(ws.cell(r, 6).value)
    g, h, i, j, k = [ws.cell(r, c).value for c in (7, 8, 9, 10, 11)]
    sds = ''
    if name in h_1128:
        sds = '有' if clean(h_1128[name][1]) else '無'
    note = ''
    if isinstance(k, str) and not re.fullmatch(NUM, k.strip()):
        note = clean(k)
    parts = SPLIT.get(name, [(name_raw, formula)])
    for part_name, part_formula in parts:
        it = new_item(類別='藥品', 科別='化學', 清單分區='藥品清單', 品名=part_name, 化學式或規格=part_formula,
                      教室=room, 櫃別=cab, 排序位置=pos, SDS=sds, 備註=note, 原清單序號=seq)
        if len(parts) > 1:
            add_check('轉入說明', it, '品名', name, part_name, f'依承辦人決定，原「{name}」拆成兩筆')
        drug_row(it, part_name.strip(), name, len(parts) > 1, g, h, i, j, k)

# 甲醛：113-2 以後已移除，保留為「已淘汰」與歷史數字
for r in range(3, old113_1.max_row + 1):
    if clean(old113_1.cell(r, 2).value) == '甲醛':
        it = new_item(類別='藥品', 科別='化學', 清單分區='藥品清單', 品名='甲醛',
                      化學式或規格=clean(old113_1.cell(r, 3).value), 教室=clean(old113_1.cell(r, 4).value),
                      櫃別=clean(old113_1.cell(r, 5).value), 排序位置=clean(old113_1.cell(r, 6).value),
                      SDS='有' if clean(old113_1.cell(r, 8).value) else '無', 狀態='已淘汰',
                      自己筆記='113-2 起不在藥品清單，轉入時保留歷史紀錄', 原清單序號=old113_1.cell(r, 1).value)
        add_rec(D_1208, it, old113_1.cell(r, 7).value)
        add_check('轉入說明', it, '狀態', '', '已淘汰', '113-2 起清單已無甲醛，設為已淘汰（不會出現在盤點表與列印），歷史數字保留')

# ================================================================ 化學器材與耗材

chem = books['chem']
CONSUMABLE_WORDS = ('手套', '口罩', '菜瓜布', '抹布', '濾網', '洗衣粉', '小蘇打', '清潔', '補充', '塑膠滴管',
                    '夾鏈袋', '秤量紙', '毛細管', 'TLC', '濾紙', '棉線', '試紙', '砂紙', '筆芯', '電池',
                    '黑紙', '衛生紙', '擦手紙', '洗碗精', '洗手乳')


def chem_cat(name, zone):
    if zone in ('耗材區', '文具區', '多元課程', '食品區'):
        return '耗材'
    if any(w in name for w in CONSUMABLE_WORDS):
        return '耗材'
    return '器材'


def date_from_header(v):
    m = re.search(r'(\d{3})\.(\d{1,2})\.(\d{1,2})', clean(v))
    if not m:
        return None
    # 承辦人確認：化準器材區、化準綜合區、科學館共用 標題寫 115.06.17 的欄位，實際是 115.08.26 盤點
    return D_0826 if m.group(0) == '115.06.17' else roc(m.group(0))


def chem_sheet(sheet_name, zone_name, room_default, qty_cols, note_col, memo_cols, zone_is_cab=True):
    sh = chem[sheet_name]
    header_row = next(r for r in range(1, 8) if clean(sh.cell(r, 1).value) == '序')
    dates = [date_from_header(sh.cell(header_row, c).value) for c in qty_cols]
    room = room_default
    for r in range(header_row + 1, sh.max_row + 1):
        seq = sh.cell(r, 1).value
        name_raw = raw(sh.cell(r, 3).value)
        if not isinstance(seq, (int, float)) or not name_raw.strip():
            continue
        loc = clean(sh.cell(r, 2).value)
        if zone_name == '化1化2':
            if loc.startswith(('化1', '化一')):
                room = '化一'
            elif loc.startswith(('化2', '化二')):
                room = '化二'
        memo = join_notes(*[sh.cell(r, c).value for c in memo_cols])
        if zone_name == '化1化2':
            loc = re.sub(r'^化[1一2二]', '', loc)
        it = new_item(類別=chem_cat(name_raw, loc), 科別='化學', 清單分區=zone_name, 品名=name_raw,
                      單位=clean(sh.cell(r, 4).value), 教室=room, 櫃別=loc, 排序位置=int(seq),
                      備註=clean(sh.cell(r, note_col).value), 自己筆記=memo, 原清單序號=int(seq))
        # 承辦人決定：兩筆「10mL分度吸量管」合併為一筆，各次數量相加
        prev = items[-2] if len(items) > 1 else None
        if prev and prev['清單分區'] == zone_name and clean(prev['品名']) == clean(name_raw):
            items.pop()
            counters[it['類別']] -= 1
            prev['自己筆記'] = join_notes(prev['自己筆記'], f'原序{int(seq)}併入', it['備註'], it['自己筆記'])
            add_check('轉入說明', prev, '品名', f'序{prev["原清單序號"]}、序{int(seq)} 同名', clean(name_raw),
                      '依承辦人決定合併為一筆，各次盤點數量相加；原序號與備註移到自己筆記')
            it = prev
        for d, c in zip(dates, qty_cols):
            add_rec(d, it, sh.cell(r, c).value)
    return dates


chem_sheet('化準114器材區', '化準器材區', '化準', [5, 6, 7, 8, 9], 10, [11, 12])
chem_sheet('化1化2', '化1化2', '化一', [5, 6], 7, [])
chem_sheet('化準114綜合區', '化準綜合區', '化準', [5, 6, 7, 8, 9], 10, [11, 12, 13])
chem_sheet('114化3後櫃', '化3後櫃', '化三', [5, 6, 7, 8, 9], 10, [11, 12, 13])

# ================================================================ 生物

bio = books['bio']

# 1 科學館共用
sh = bio['1科學館共用']
dates = [date_from_header(sh.cell(3, c).value) for c in range(4, 9)]
for r in range(4, sh.max_row + 1):
    if not isinstance(sh.cell(r, 1).value, (int, float)):
        continue
    it = new_item(類別='耗材', 科別='科學館共用', 清單分區='科學館共用', 品名=raw(sh.cell(r, 2).value),
                  化學式或規格=clean(sh.cell(r, 9).value), 單位=clean(sh.cell(r, 3).value),
                  教室='生3', 櫃別='後儲物空間', 排序位置=int(sh.cell(r, 1).value),
                  原清單序號=int(sh.cell(r, 1).value))
    for d, c in zip(dates, range(4, 9)):
        add_rec(d, it, sh.cell(r, c).value)

# 2 生物器材（生1、生2 分處）
sh = bio['2.114年生物器材物品(A格式)']
for r in range(5, sh.max_row + 1):
    if not isinstance(sh.cell(r, 1).value, (int, float)) or not clean(sh.cell(r, 2).value):
        continue
    name = raw(sh.cell(r, 2).value)
    it = new_item(類別='器材', 科別='生物', 清單分區='生物器材', 品名=name, 單位=clean(sh.cell(r, 3).value),
                  教室='生1、生2', 分處存放='生1、生2', 排序位置=int(sh.cell(r, 1).value),
                  原清單序號=int(sh.cell(r, 1).value))
    add_rec(D_0826, it, sh.cell(r, 4).value, '生1')
    add_rec(D_0826, it, sh.cell(r, 5).value, '生2')

# 生物耗材
sh = bio['耗材']
D_0506 = roc('114.5.6')
LOCS = ['木櫃', '防潮櫃', '冰箱']
for r in range(5, sh.max_row + 1):
    if not isinstance(sh.cell(r, 1).value, (int, float)):
        continue
    name = clean(sh.cell(r, 2).value)
    unit = clean(sh.cell(r, 3).value)
    locq = [sh.cell(r, c).value or 0 for c in (4, 5, 6)]
    h_0506, h_0821, j = sh.cell(r, 8).value, sh.cell(r, 9).value, sh.cell(r, 10).value
    if isinstance(h_0506, str) and h_0506.startswith('='):
        h_0506 = None
    note, buy_when, buy_qty = sh.cell(r, 11).value, sh.cell(r, 12).value, sh.cell(r, 13).value
    is_buy = buy_when is not None and re.fullmatch(r'\d{3}\.\d{1,2}', clean(buy_when)) and isinstance(buy_qty, (int, float))
    memo = '' if is_buy else join_notes(buy_when, buy_qty)
    seq = int(sh.cell(r, 1).value)

    if name in merged:
        it = merged[name]
        drug_latest = [rec for rec in records if rec[1] == it['編號'] and rec[0] == D_0826]
        if clean(j) and all(clean(rec[4]) != clean(j) for rec in drug_latest if rec[3] == '生準冰箱'):
            dl = '、'.join(f'{rec[3]} {rec[5] or rec[4]}' for rec in drug_latest)
            # 承辦人決定：暫用藥品清單的數字，下次盤點確認
            memo = join_notes(memo, f'生物耗材清單 115.08.26 記為 {clean(j)}')
            add_check('需確認', it, '盤點數量', f'藥品清單 115.08.26：{dl}', f'生物耗材 115.08.26：{clean(j)}',
                      '兩份清單數字不同，依承辦人決定暫用藥品清單的數字，下次盤點確認')
        it['自己筆記'] = join_notes(it['自己筆記'], f'原生物耗材清單序{seq}', note, memo)
        continue

    nonzero = [LOCS[x] for x in range(3) if locq[x]]
    total = sum(locq)
    it = new_item(類別='耗材', 科別='生物', 清單分區='生物耗材', 品名=raw(sh.cell(r, 2).value), 單位=unit,
                  教室='生準', 排序位置=seq, 原清單序號=seq)
    if name == '單面刀片':
        it['化學式或規格'] = clean(note)
        it['自己筆記'] = memo
    else:
        it['備註'] = clean(note)
        it['自己筆記'] = memo
    if is_buy:
        m = re.fullmatch(r'(\d{3})\.(\d{1,2})', clean(buy_when))
        moves.append([dt.date(int(m.group(1)) + 1911, int(m.group(2)), 1), it['編號'], name, '新購',
                      buy_qty, f'原資料購買時間 {clean(buy_when)}（無日，暫記為1日）', ''])

    add_rec(D_0506, it, h_0506)
    add_rec(roc('114.8.21'), it, h_0821)
    jn, _ = parse_qty(j)
    if len(nonzero) > 1 and total == jn:
        it['分處存放'] = '、'.join(nonzero)
        it['櫃別'] = '、'.join(nonzero)
        for x in range(3):
            if locq[x]:
                add_rec(D_0826, it, locq[x], LOCS[x])
    else:
        it['櫃別'] = '、'.join(nonzero) if nonzero else clean(note)
        add_rec(D_0826, it, j)
        if total != jn:
            where = '、'.join(f'{LOCS[x]}{clean(locq[x])}' for x in range(3) if locq[x]) or '三處皆 0'
            add_check('需確認', it, '盤點數量', f'存放處合計 {clean(total)}（{where}）', f'115.08.26 總數 {clean(j)}',
                      '承辦人指示先以最新總數為準；請下次盤點時確認' +
                      ('；櫃別暫填備註中的位置' if not nonzero else ''))

# 顯微鏡觀察
sh = bio['顯微鏡觀察']
MICRO_CONSUMABLE = ('蓋玻片', '載玻片', '拭鏡紙')
for r in range(4, sh.max_row + 1):
    if not isinstance(sh.cell(r, 1).value, (int, float)):
        continue
    name = raw(sh.cell(r, 2).value)
    if clean(name) == '永久玻片':
        continue   # 改由「永久玻片」清單分區逐種登記
    it = new_item(類別='耗材' if clean(name) in MICRO_CONSUMABLE else '器材', 科別='生物', 清單分區='顯微鏡觀察',
                  品名=name, 單位=clean(sh.cell(r, 3).value), 教室=clean(sh.cell(r, 5).value),
                  排序位置=int(sh.cell(r, 1).value), 備註=clean(sh.cell(r, 6).value),
                  原清單序號=int(sh.cell(r, 1).value))
    places = [x for x in re.split(r'[、,，/]', it['教室']) if x]
    if len(places) > 1:
        # 承辦人決定：分處盤點；舊檔只有總數，這次先記總數
        it['分處存放'] = '、'.join(places)
        add_check('需確認', it, '盤點數量', f'115.08.26 總數 {clean(sh.cell(r, 4).value)}', '',
                  '改為分處盤點，舊檔只有總數，先記總數；下次盤點時各處分別填寫')
    add_rec(D_0826, it, sh.cell(r, 4).value)

# 解剖用具（115.08.26）＋ 114.2.18 舊版數字
old = bio['解剖用具']
old_q = {clean(old.cell(r, 2).value): old.cell(r, 4).value for r in range(4, 13)}
sh = bio['解剖+實驗套組']
for r in range(4, 14):
    if not isinstance(sh.cell(r, 1).value, (int, float)):
        continue
    name = raw(sh.cell(r, 2).value)
    it = new_item(類別='耗材' if '大頭針' in name else '器材', 科別='生物', 清單分區='解剖用具', 品名=name,
                  單位=clean(sh.cell(r, 3).value), 教室=clean(sh.cell(r, 5).value),
                  排序位置=int(sh.cell(r, 1).value), 備註=clean(sh.cell(r, 6).value),
                  原清單序號=int(sh.cell(r, 1).value))
    if clean(name) in old_q:
        add_rec(roc('114.2.18'), it, old_q[clean(name)])
    add_rec(D_0826, it, sh.cell(r, 4).value)

# 實驗套組（第二階段表，先轉入）
bio_items = {clean(it['品名']): it for it in items if it['科別'] == '生物'}
kit_name = None
for r in range(17, 33):
    a = clean(sh.cell(r, 1).value)
    m = re.match(r'高1生物：(.+?)。每組器材清單\(共(\d+)組\)', a)
    if m:
        kit_name, kit_n, kit_place = m.group(1), int(m.group(2)), ''
        continue
    if isinstance(sh.cell(r, 1).value, (int, float)) and kit_name:
        name = clean(sh.cell(r, 2).value)
        q, qn = parse_qty(sh.cell(r, 3).value)
        kit_place = clean(sh.cell(r, 5).value) or kit_place
        ref = bio_items.get(name)
        kits.append([kit_name, '高1生物', kit_n, ref['編號'] if ref else '', name, q, qn,
                     clean(sh.cell(r, 4).value), kit_place])

# 永久玻片
sh = bio['永久玻片清單(期中)']
D_0507 = roc('114.5.7')
grade = course = ''
for r in range(4, 30):
    if clean(sh.cell(r, 1).value):
        grade = clean(sh.cell(r, 1).value)
    if clean(sh.cell(r, 2).value):
        course = clean(sh.cell(r, 2).value)
    name = raw(sh.cell(r, 4).value)
    if not clean(name) or not isinstance(sh.cell(r, 3).value, (int, float)):
        continue
    need, b1, p1, b2, p2, plan, have, n1, n2 = [sh.cell(r, c).value for c in range(5, 14)]
    have_n, have_note = parse_qty(have)
    ref = ''
    if have is not None or clean(course) == '選手培訓用':
        it = new_item(類別='器材', 科別='生物', 清單分區='永久玻片', 品名=name, 單位='片' if course != '選手培訓用' else '組',
                      教室='生準', 櫃別='木櫃及防潮箱', 排序位置=len([x for x in items if x['清單分區'] == '永久玻片']) + 1,
                      備註=clean(n1),
                      原清單序號=int(sh.cell(r, 3).value))
        ref = it['編號']
        records.append([D_0507, it['編號'], clean(name), '', have_n, have_note])
    if course != '選手培訓用':
        slides_need.append([ref, clean(name), grade, course, need, b1, p1, b2, p2, plan, join_notes(n1, n2)])

# 同一次盤點、同一品項、同一存放處有兩筆（合併品項造成）→ 數量相加
combined = {}
for rec in records:
    key = (rec[0], rec[1], rec[3])
    if key in combined:
        prev = combined[key]
        prev[4] = round((prev[4] or 0) + (rec[4] or 0), 3)
        prev[5] = join_notes(prev[5], rec[5])
    else:
        combined[key] = rec
records[:] = list(combined.values())

# ================================================================ 疑似錯誤（不自動更正）

by_name_zone = {}
for it in items:
    by_name_zone.setdefault(clean(it['品名']), []).append(it)


def find(name, zone=None):
    for it in by_name_zone.get(name, []):
        if zone is None or it['清單分區'] == zone:
            return it
    return None


def sug(kind, name, field, suggest, note, zone=None, apply=False, current=None):
    it = find(name, zone)
    if not it:
        raise KeyError(name)
    cur = current if current is not None else it[field]
    add_check(kind, it, field, cur, suggest, note, apply)


def fix(name, field, new, note, zone=None):
    """承辦人已同意的更正：直接改，並在轉入檢查留紀錄。"""
    it = find(name, zone)
    if not it:
        raise KeyError(name)
    add_check('已更正', it, field, it[field], new, note)
    it[field] = new


for it in items:
    if it['品名'] != it['品名'].strip():
        add_check('已更正', it, '品名', repr(it['品名']), it['品名'].strip(), '品名前後有多餘空白')
        it['品名'] = it['品名'].strip()

fix('氟化納', '品名', '氟化鈉', '「納」應為「鈉」')
fix('Dcpip', '品名', 'DCPIP', '英文縮寫大寫')
fix('溴瑞香草藍', '品名', '溴瑞香草酚藍', 'Bromothymol blue（BTB）')
fix('本式液', '品名', '本氏液', 'Benedict\'s 試液')
fix('斐林式液A', '品名', '斐林氏液A', 'Fehling\'s 試液（或稱斐林試液）')
fix('斐林式液B', '品名', '斐林氏液B', 'Fehling\'s 試液（或稱斐林試液）')
for p in ('3', '4', '5', '7', '9', '10'):
    fix(f'酸鹼緩衝液(PH={p})', '品名', f'酸鹼緩衝液(pH={p})', 'pH 的 p 小寫')
fix('12烷基硫酸鈉', '品名', '十二烷基硫酸鈉', '中文數字')
sug('可不改', '磷酸氫鈉', '品名', '磷酸氫二鈉', 'Na2HPO4 正式名稱（可不改）')
sug('可不改', '磷酸氫鉀', '品名', '磷酸氫二鉀', 'K2HPO4 正式名稱（可不改）')
fix('硝酸鋁', '化學式或規格', 'Al(NO3)3', '大寫 I 應為小寫 l（鋁 Al）')
fix('硫酸鋁鉀 (鉀鋁礬)(明礬)', '化學式或規格', 'KAl(SO4)2', '大寫 I 應為小寫 l（鋁 Al）')
sug('待查瓶身', '氧化鐵(III)', '化學式或規格', 'Fe2O3', 'Fe3O4 是四氧化三鐵；承辦人決定先照原檔，之後看瓶身')
fix('鉬酸銨', '化學式或規格', '(NH4)6Mo7O24', '鉬是 Mo（o 小寫）')
fix('麥芽糖', '化學式或規格', 'C12H22O11', '雙醣 C12H22O11')
sug('待查瓶身', '甲基纖維', '化學式或規格', '', '此化學式是羧甲基纖維素鈉（CMC），與品名「甲基纖維素」不符；先照原檔')
sug('待查瓶身', '碳酸氫鈣', '化學式或規格', '', '碳酸氫鈣只存在於水溶液，沒有固體試劑；先照原檔')
sug('可不改', '聚乙烯醇', '化學式或規格', '(C2H4O)n', '其他聚合物都寫 n（可不改）')
sug('可不改', '氨水', '化學式或規格', 'NH3(aq)', '較新寫法（可不改）')
sug('可不改', 'Bacto Agar (BD214010)', '化學式或規格', '', '化學式欄填「寒天」，非化學式；欄位名稱是「化學式或規格」，照原樣也可以')
sug('可不改', 'LB Broth (sigma L3522)', '化學式或規格', '', '化學式欄填「蛋白腖」，非化學式；照原樣也可以')
sug('可不改', '牛血清白蛋白', '化學式或規格', '', '化學式欄填英文名；照原樣也可以')
sug('可不改', '固綠', '化學式或規格', '', '化學式欄填英文名；照原樣也可以')
fix('250mL 圓底燒杯', '品名', '250mL 圓底燒瓶', '圓底應為燒瓶')
fix('500mL 圓底燒杯', '品名', '500mL 圓底燒瓶', '圓底應為燒瓶')
for c in ('綠色盒子', '橘色盒子', '藍色盒子', '藍色塑膠盒PH-5011'):
    fix(f'ph pen ({c})', '品名', f'pH pen ({c})', 'pH 寫法')
fix('手趴雞用手套', '品名', '手扒雞用手套', '「趴」應為「扒」')
fix('砂紙粗)', '品名', '砂紙(粗)', '少左括號')
fix('布式漏斗', '品名', '布氏漏斗', 'Büchner 漏斗')
fix('Parafilm封口臘膜', '品名', 'Parafilm封口蠟膜', '「臘」應為「蠟」')
fix('美工刀', '備註', '含庫存區新的：20', '多一個「含」', zone='化準綜合區')
fix('毛莨根', '品名', '毛茛根', '「莨」應為「茛」')
add_check('轉入說明', None, '', '教室寫法', '化學準備室、化學實驗室一…', '依承辦人決定，教室一律寫全名（化準→化學準備室、生1→生物實驗室一…）；化1化2 的櫃別去掉教室字首（化1A櫃→A櫃）')
add_check('轉入說明', None, '', '化準器材區、化準綜合區、科學館共用 最後一欄', '115.08.26',
          '標題寫 115.06.17，依承辦人確認記為 115.08.26')
# 承辦人確認：半乳糖 113.12.5 舊檔寫 205，應為 2.5
galactose = find('半乳糖')
for rec in records:
    if rec[1] == galactose['編號'] and rec[0] == D_1208 and rec[4] == 205:
        rec[4] = 2.5
        add_check('已更正', galactose, '盤點數量', '113.12.05：205', '2.5', '承辦人確認舊檔少打小數點')
for k in kits:
    if k[4] == '6mm漏斗':
        k[4] = '6cm漏斗'
        add_check('已更正', None, '', '實驗套組 DNA粗萃取「6mm漏斗」', '6cm漏斗', '承辦人確認是 6cm')

# 數量文字（照原文列印；這裡只列出數字是怎麼取的）
last_date = {}
for rec in records:
    last_date[rec[1]] = max(last_date.get(rec[1], rec[0]), rec[0])
for rec in records:
    d, code, name, place, n, note = rec
    if note and d == last_date[code]:
        if len(re.findall(NUM, note)) > 1 or n is None:
            add_check('數量文字', {'編號': code, '品名': name}, '數量', note, clean(n),
                      '數量說明照原文列印；右邊數字是系統取來比對安全存量用的，不對請改盤點紀錄的「數量」')

# ================================================================ 教室寫全名（承辦人決定）

ROOM_FULL = {'準': '準備室', '一': '實驗室一', '二': '實驗室二', '三': '實驗室三',
             '1': '實驗室一', '2': '實驗室二', '3': '實驗室三'}


def room_full(text):
    """化準→化學準備室、生1→生物實驗室一；其他文字不動。"""
    return re.sub(r'([化生])([準一二三123])',
                  lambda m: ('化學' if m.group(1) == '化' else '生物') + ROOM_FULL[m.group(2)], str(text))


for it in items:
    for f in ('教室', '分處存放', '櫃別'):
        if it[f]:
            it[f] = room_full(it[f])
for rec in records:
    rec[3] = room_full(rec[3])
for k in kits:
    k[8] = room_full(k[8])

# ================================================================ 計算最新數量

item_by_code = {it['編號']: it for it in items}
latest = {}
for d, code, name, place, n, note in records:
    latest.setdefault(code, {}).setdefault(d, []).append((place, n, note))
for code, by_date in latest.items():
    d = max(by_date)
    rows = by_date[d]
    it = item_by_code[code]
    it['最新盤點日期'] = roc_text(d)
    nums = [n for _, n, _ in rows if n is not None]
    it['最新數量'] = round(sum(nums), 3) if nums else ''
    if isinstance(it['最新數量'], float) and it['最新數量'].is_integer():
        it['最新數量'] = int(it['最新數量'])
    if len(rows) == 1:
        it['最新數量說明'] = rows[0][2]
    elif any(note for _, _, note in rows):
        it['最新數量說明'] = '、'.join(f'{p} {note or clean(n)}' for p, n, note in rows)

# ================================================================ 輸出

wb = openpyxl.Workbook()
HEAD_FILL = PatternFill('solid', fgColor='DDEBF7')
AUTO_FILL = PatternFill('solid', fgColor='F2F2F2')
BOLD = Font(bold=True)
thin = Side(style='thin', color='BFBFBF')


def write_sheet(ws, header, rows, widths=None, auto_cols=()):
    ws.append(header)
    for c in range(1, len(header) + 1):
        cell = ws.cell(1, c)
        cell.font = BOLD
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for row in rows:
        ws.append(row)
    ws.freeze_panes = 'A2'
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    for c in auto_cols:
        for r in range(2, ws.max_row + 1):
            ws.cell(r, c).fill = AUTO_FILL


# 設定
ws = wb.active
ws.title = '設定'
params = [
    ('學校名稱', '臺北市立中山女子高級中學'),
    ('學年度學期', '115-1'),
    ('核章欄', '實驗室管理員,設備組長,主任,校長'),
    ('藥品清單列印盤點次數', 4),
    ('器材耗材清單列印盤點次數', 1),
    ('手機網頁網址', '（第4步部署後填入）'),
]
lists = {
    '類別': ['藥品', '器材', '耗材'],
    '科別': ['化學', '生物', '科學館共用'],
    '狀態': ['使用中', '已淘汰'],
    '有無': ['有', '無'],
    '是否': ['是', '否'],
    '異動類型': ['新購', '報廢', '領用', '移位'],
    '單位': sorted({it['單位'] for it in items if it['單位']} | {'瓶', '罐', 'g', 'mL', '公升'}),
}
zones = [
    # 清單分區, 列印標題, 列印副標題, 存放地點, 列印順序, 列印格式, 分處欄位
    ('藥品清單', '化學與生物實驗室藥品清單', '', '', 1, '藥品', ''),
    ('化準器材區', '3樓化學實驗室清單', '器材區物品與耗材', '化學準備室', 2, '一般', ''),
    ('化1化2', '3樓化學實驗室清單', '物品與耗材', '化學實驗室一、化學實驗室二', 3, '一般', ''),
    ('化準綜合區', '3樓化學實驗室清單', '綜合區(耗材、工具、文具)+配藥區', '化學準備室', 4, '一般', ''),
    ('化3後櫃', '3樓化學實驗室清單', '多元課程、食品區、各類食具', '化學實驗室三', 5, '一般', ''),
    ('科學館共用', '科學館共用(含探究與實作)', '', '生物實驗室三後儲物空間', 6, '一般', ''),
    ('生物器材', '生物(玻璃)器材與物品清單', '', '', 7, '分處', '生1、生2'),
    ('生物耗材', '生物實驗室耗材清單', '', '生物準備室', 8, '分處', '木櫃、防潮櫃、冰箱'),
    ('顯微鏡觀察', '顯微鏡觀察用物品清單', '', '', 9, '一般', ''),
    ('解剖用具', '生物解剖相關用具', '', '', 10, '一般', ''),
    ('永久玻片', '生物永久玻片清單', '', '生物準備室', 11, '一般', ''),
]
rooms = [('化學準備室', '化準'), ('化學實驗室一', '化一'), ('化學實驗室二', '化二'),
         ('化學實驗室三', '化三'), ('生物準備室', '生準'), ('生物實驗室一', '生1'),
         ('生物實驗室二', '生2'), ('生物實驗室三', '生3')]

ws['A1'], ws['B1'] = '項目', '值'
for i, (k, v) in enumerate(params, 2):
    ws.cell(i, 1, k)
    ws.cell(i, 2, v)
col = 4
list_cols = {}
for key, vals in lists.items():
    ws.cell(1, col, key)
    for i, v in enumerate(vals, 2):
        ws.cell(i, col, v)
    list_cols[key] = get_column_letter(col)
    col += 1
col += 1
ws.cell(1, col, '教室')
ws.cell(1, col + 1, '舊簡稱')
for i, (a, b) in enumerate(rooms, 2):
    ws.cell(i, col, a)
    ws.cell(i, col + 1, b)
list_cols['教室'] = get_column_letter(col)
col += 3
zone_head = ['清單分區', '列印標題', '列印副標題', '存放地點', '列印順序', '列印格式', '分處欄位']
for j, h in enumerate(zone_head):
    ws.cell(1, col + j, h)
for i, z in enumerate(zones, 2):
    for j, v in enumerate(z):
        ws.cell(i, col + j, v)
list_cols['清單分區'] = get_column_letter(col)
for c in range(1, col + len(zone_head)):
    ws.cell(1, c).font = BOLD
    ws.cell(1, c).fill = HEAD_FILL
    ws.column_dimensions[get_column_letter(c)].width = 12
ws.column_dimensions['A'].width = 24
ws.column_dimensions['B'].width = 34
ws.freeze_panes = 'A2'

# 品項：依清單分區列印順序、原順序排列
zone_order = {z[0]: z[4] for z in zones}
items.sort(key=lambda it: (zone_order[it['清單分區']], it['狀態'] == '已淘汰', pos_key(it['原清單序號'])))
ws = wb.create_sheet('品項')
write_sheet(ws, ITEM_COLS, [[it[c] for c in ITEM_COLS] for it in items],
            [9, 6, 10, 11, 26, 22, 6, 10, 13, 8, 14, 8, 6, 8, 10, 7, 26, 26, 8, 11, 9, 20],
            auto_cols=(1, 20, 21, 22))
for r in range(2, ws.max_row + 1):
    ws.cell(r, 6).number_format = '@'
    ws.cell(r, 10).number_format = '@'
n_items = ws.max_row


def dv_list(ws, col_letter, rng, key, allow_other=False):
    dv = DataValidation(type='list', formula1=f"'設定'!${list_cols[key]}$2:${list_cols[key]}$40",
                        allow_blank=True, showErrorMessage=not allow_other)
    ws.add_data_validation(dv)
    dv.add(f'{col_letter}2:{col_letter}{rng}')


dv_list(ws, 'B', 2000, '類別')
dv_list(ws, 'C', 2000, '科別')
dv_list(ws, 'D', 2000, '清單分區')
dv_list(ws, 'G', 2000, '單位', allow_other=True)
dv_list(ws, 'H', 2000, '教室', allow_other=True)
dv_list(ws, 'M', 2000, '有無')
dv_list(ws, 'N', 2000, '是否')
dv_list(ws, 'P', 2000, '狀態')
ws.auto_filter.ref = f'A1:{get_column_letter(len(ITEM_COLS))}{n_items}'

# 盤點紀錄
records.sort(key=lambda x: (x[0], x[1]))
ws = wb.create_sheet('盤點紀錄')
write_sheet(ws, ['盤點日期', '盤點名稱', '編號', '品名', '存放處', '數量', '數量說明', '登錄時間'],
            [[d, term_label(d), c, item_by_code[c]['品名'], p, q, t, IMPORT_STAMP] for d, c, n, p, q, t in records],
            [12, 11, 9, 26, 10, 8, 24, 16])
for r in range(2, ws.max_row + 1):
    ws.cell(r, 1).number_format = 'yyyy/mm/dd'

# 盤點表（程式產生，先放表頭）
ws = wb.create_sheet('盤點表')
write_sheet(ws, ['編號', '清單分區', '教室', '櫃別', '排序位置', '品名', '化學式或規格', '單位', '存放處',
                 '上次日期', '上次數量', '上次說明', '本次數量', '本次說明', '已盤'],
            [], [9, 11, 8, 12, 8, 24, 18, 6, 9, 11, 9, 16, 10, 16, 6])

ws = wb.create_sheet('需補充清單')
ws['A1'] = '（第 7 步安裝程式後自動產生）'
ws = wb.create_sheet('查詢')
ws['A1'] = '（第 7 步安裝程式後自動產生）'

# 異動紀錄
ws = wb.create_sheet('異動紀錄')
write_sheet(ws, ['日期', '編號', '品名', '類型', '數量', '說明', '經手人'], sorted([m[:2] + [item_by_code[m[1]]['品名']] + m[3:] for m in moves]),
            [12, 9, 22, 8, 8, 40, 10])
for r in range(2, ws.max_row + 1):
    ws.cell(r, 1).number_format = 'yyyy/mm/dd'
dv_list(ws, 'D', 2000, '異動類型')

# 實驗套組
ws = wb.create_sheet('實驗套組')
write_sheet(ws, ['套組名稱', '課程年級', '組數', '編號', '品名', '每組數量', '每組數量說明', '備註', '存放位置'],
            kits, [14, 10, 6, 9, 16, 8, 12, 16, 10])

# 玻片需求
ws = wb.create_sheet('玻片需求')
write_sheet(ws, ['編號', '品名', '年級', '課本／課程', '需求數量(112年)', '購買數量(第1批)', '單價(第1批)',
                 '購買數量(第2批)', '單價(第2批)', '預計購買數量', '備註'],
            slides_need, [9, 18, 6, 18, 10, 10, 9, 10, 9, 10, 24])

# 轉入檢查
ws = wb.create_sheet('轉入檢查')
order = {'需確認': 1, '待查瓶身': 2, '可不改': 3, '已更正': 4, '轉入說明': 5, '數量文字': 6}
checks.sort(key=lambda c: order[c['類型']])
rows = []
for i, c in enumerate(checks, 1):
    if c['類型'] in ('可不改', '待查瓶身'):
        c['說明'] += '（承辦人：先照原檔，之後有空再確認）'
    rows.append([i, c['類型'], c['編號'], clean(c['品名']), c['欄位'], c['目前內容'], c['建議內容'], c['說明'],
                 False if c['可套用'] else '', ''])
write_sheet(ws, ['項次', '類型', '編號', '品名', '欄位', '目前內容', '建議內容', '說明', '採用建議', '您的回覆'],
            rows, [5, 8, 9, 22, 12, 26, 22, 50, 8, 20])
for r in range(2, ws.max_row + 1):
    for c in (6, 7, 8):
        ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical='top')

wb.save(out_path)

# 摘要
from collections import Counter
print('品項', len(items), Counter(it['類別'] for it in items))
print('分區', Counter(it['清單分區'] for it in items))
print('盤點紀錄', len(records), '日期', sorted({roc_text(r[0]) for r in records}))
print('異動', len(moves), '套組', len(kits), '玻片需求', len(slides_need))
print('轉入檢查', len(checks), Counter(c['類型'] for c in checks))
