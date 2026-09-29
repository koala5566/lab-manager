"""依品名、化學式建議危險分類，產出「危險分類建議.xlsx」讓承辦人勾選後匯入。

用法：python3 suggest_hazard.py <新系統初始資料.xlsx> <輸出.xlsx>

只是依常見性質的初步建議（參考一般 SDS 的 GHS 分類），實際請以各藥品 SDS 為準。
分類：易燃液體、易燃固體、禁水性、氧化性、腐蝕性、毒性、刺激／有害
"""
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

import re

# 規則用「品名主體」（去掉括號、空白後的名稱）比對，避免「硫酸鈉」被當成「硫酸」。
EXACT = [
    ('易燃液體', ['乙醇', '燃燒酒精', '甲醇', '正丙醇', '異丙醇', '正丁醇', '異丁醇', '第三丁醇', '正戊醇', '丙酮', '丁酮', '乙醚',
              '石油醚', '正己烷', '環己烷', '環己烯', '甲苯', '二甲苯', '乙酸乙酯', '乙酸異戊酯', '丙醛', '己二醯氯正己烷溶液'],
     '揮發性有機溶劑，閃火點低'),
    ('易燃固體', ['鎂帶', '鋁粒', '鋅粉', '硫粉', '萘', '天然樟腦粉', '鐵粉', '銻粉'], '可燃性粉末或固體'),
    ('禁水性', ['鈉', '碳化鈣'], '遇水產生可燃氣體'),
    ('氧化性', ['過錳酸鉀', '氯酸鉀', '碘酸鉀', '過硫酸鉀', '35%雙氧水', '亞硝酸鈉', '硝酸'], '強氧化劑，勿與可燃物、還原劑同放'),
    ('腐蝕性', ['鹽酸', '硫酸', '硝酸', '醋酸', '甲酸', '丙酸', '丁酸', '氨水', '氫氧化鈉', '氫氧化鉀', '乙酐', '乙二胺', '溴水', '氯水',
             '己二胺氫氧化鈉溶液', '己二醯氯正己烷溶液', '硫酸氫鈉', '硫酸氫鉀', '氧化鈣', '硝酸銀', '35%雙氧水'],
     '強酸、強鹼或具腐蝕性'),
    ('毒性', ['氯化鋇', '硝酸鋇', '氫氧化鋇', '氟化鈉', '亞硝酸鈉', '甲醇', '銻粉', '硫氰化鉀'], '急毒性物質'),
    ('刺激／有害', ['碘', '碘液', '硫酸銅', '氯化銅', '硝酸銅', '氯化鐵', '硼酸', '硼酸鈉', '水楊酸', '結晶紫', '剛果紅', '亞甲藍',
                 '番紅', '甲基橙', '酚酞', '硫化銨', '亞硫酸鈉', '亞硫酸氫鈉', '焦亞硫酸鈉', '低亞硫酸鈉', '赤血鹽', '黃血鹽',
                 '甲苯', '二甲苯', '萘', '乙醚'], '對皮膚、眼睛有刺激性或吞食有害'),
]
# 用「開頭」比對的（一整類）
PREFIX = [
    ('氧化性', '硝酸', '硝酸鹽為氧化劑'),
    ('刺激／有害', '草酸', '草酸及草酸鹽吞食有害'),
]
# 用「包含」比對的重金屬
HEAVY = (['汞', '鉛', '鎳', '鈷', '鉻'], '毒性', '重金屬化合物，請確認是否屬列管毒性化學物質')


def base_name(name):
    return re.sub(r'[\s(（].*$', '', name).strip()


def classify(name):
    b = base_name(name)
    out = {}
    for cls, names, why in EXACT:
        if b in names:
            out.setdefault(cls, why)
    for cls, pre, why in PREFIX:
        if b.startswith(pre):
            out.setdefault(cls, why)
    if any(k in b for k in HEAVY[0]) and b not in ('鉻黑T',):   # 鉻黑T 是指示劑染料，不是鉻化合物
        out.setdefault(HEAVY[1], HEAVY[2])
    return out


src, dst = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
ws = wb['品項']
H = [c.value for c in ws[1]]
rows = []
for r in ws.iter_rows(min_row=2, values_only=True):
    x = dict(zip(H, r))
    if x['類別'] != '藥品' or x['狀態'] == '已淘汰':
        continue
    name, formula = str(x['品名'] or ''), str(x['化學式或規格'] or '')
    found = classify(name)
    if not found:
        continue
    order = ['易燃液體', '易燃固體', '禁水性', '氧化性', '腐蝕性', '毒性', '刺激／有害']
    cls = [c for c in order if c in found]
    why = [c + '：' + found[c] for c in cls]
    rows.append([True, x['編號'], name, formula, x['教室'], x['櫃別'], '、'.join(cls), '；'.join(why)])

out = openpyxl.Workbook()
o = out.active
o.title = '危險分類建議'
cols = ['採用', '編號', '品名', '化學式', '教室', '櫃別', '建議危險分類', '依據（僅供參考，請以 SDS 為準）']
o.append(cols)
for r in rows:
    o.append(r)
for c in o[1]:
    c.font = Font(bold=True)
    c.fill = PatternFill('solid', fgColor='DDEBF7')
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
y = PatternFill('solid', fgColor='FFF2CC')
for i in range(2, o.max_row + 1):
    o.cell(i, 1).fill = y
    o.cell(i, 7).fill = y
dv = DataValidation(type='list', formula1='"TRUE,FALSE"')
o.add_data_validation(dv)
dv.add(f'A2:A{o.max_row}')
for i, w in enumerate([6, 9, 24, 20, 12, 7, 26, 60], 1):
    o.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
o.freeze_panes = 'C2'
o.auto_filter.ref = o.dimensions
out.save(dst)

from collections import Counter
cnt = Counter(c for r in rows for c in r[6].split('、'))
print('建議', len(rows), '項', dict(cnt))
