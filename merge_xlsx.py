# -*- coding: utf-8 -*-
"""把 xlsx 数据合并进 index.html 的 INITIAL_DATA，跳过重复记录"""
import openpyxl
import re

HTML = 'index.html'
XLSX = '2023年江西鄱阳湖水体微塑料数据.xlsx'

# ---------- 1. 解析 xlsx ----------
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))

def to_num(v):
    if v in (None, ''):
        return 0
    f = float(v)
    return int(f) if f.is_integer() else f

def fmt_coord(v):
    return ('%.6f' % round(float(v), 6)).rstrip('0').rstrip('.')

def fmt_abd(v):
    return '%.2f' % round(float(v), 2)

def fmt_vol(v):
    f = float(v)
    return str(int(f)) if f.is_integer() else ('%g' % f)

pat = re.compile(r'^(.+?)\uff08([^\uff0c,]+)[\uff0c,](.+?)\uff09\s*$')
# 同时兼容全角和半角括号/逗号
pat2 = re.compile(r'^(.+?)[(\uff08]([^\uff0c,]+)[\uff0c,](.+?)[)\uff09]\s*$')
records = []
period_raw = None
for i in range(3, len(rows)):
    r = rows[i]
    if r[0] not in (None, ''):
        period_raw = str(r[0]).strip()
    if r[1] in (None, ''):
        continue
    m = pat2.match(str(r[1]).strip())
    if not m:
        continue
    if all(v in (None, '') for v in r[4:26]):
        continue
    rive_name, abbr, num = m.group(1), m.group(2), m.group(3)
    pm = re.match(r'^(\u4e30\u6c34\u671f|\u67af\u6c34\u671f)(.*)$', period_raw or '')
    period = pm.group(1) if pm else ''
    stime = pm.group(2) if pm else (period_raw or '')
    records.append({
        'period': period, 'sampling_time': stime,
        'rive_name': rive_name, 'rive_abbreviation': abbr, 'river_number': num,
        'latitude': round(float(r[2]), 6), 'longitude': round(float(r[3]), 6),
        'plastics_splinter_type': to_num(r[4]), 'plastics_fibre_type': to_num(r[5]),
        'plastics_grain_type': to_num(r[6]), 'plastics_membrane_type': to_num(r[7]),
        'plastics_foam_type': to_num(r[8]),
        'plastics_red_color': to_num(r[9]), 'plastics_black_color': to_num(r[10]),
        'plastics_white_color': to_num(r[11]), 'plastics_transparent_color': to_num(r[12]),
        'plastics_blue_color': to_num(r[13]), 'plastics_other_color': to_num(r[14]),
        'fifty_radius': to_num(r[15]), 'one_hundred_radius': to_num(r[16]),
        'three_hundred_radius': to_num(r[17]), 'five_hundred_radius': to_num(r[18]),
        'one_thousand_radius': to_num(r[19]), 'two_thousand_radius': to_num(r[20]),
        'three_thousand_radius': to_num(r[21]), 'five_thousand_radius': to_num(r[22]),
        'plastics_total': to_num(r[23]), 'water_volume': to_num(r[24]),
        'plastics_abundance': round(float(r[25]), 2),
    })

print('xlsx parsed rows:', len(records))

# ---------- 2. 解析 index.html 中现有 INITIAL_DATA ----------
with open(HTML, encoding='utf-8') as f:
    html = f.read()
start = html.index('const INITIAL_DATA = [')
end = html.index('];', start)
block = html[start:end]

line_re = re.compile(r"\{id:\s*(\d+),\s*period:\s*'([^']*)',\s*sampling_time:\s*'([^']*)'.*?\}")

def parse_line(line):
    d = {}
    for k, v in re.findall(r"(\w+):\s*('(?:[^']*)'|[-\d.]+)", line):
        d[k] = v.strip("'") if v.startswith("'") else float(v)
    return d

existing = [parse_line(m.group(0)) for m in line_re.finditer(block)]
print('existing records:', len(existing), 'max id:', max(int(e['id']) for e in existing))

existing_keys = {(e['river_number'], e['period'], e['sampling_time']) for e in existing}

# ---------- 3. 去重 + 一致性校验 ----------
new_records, skipped, mismatches, seen = [], [], [], set()
for rec in records:
    key = (rec['river_number'], rec['period'], rec['sampling_time'])
    if key in seen:
        skipped.append((key, 'dup-in-xlsx'))
        continue
    seen.add(key)
    if key in existing_keys:
        ex = next(e for e in existing if (e['river_number'], e['period'], e['sampling_time']) == key)
        diff = []
        for k, v in rec.items():
            ev = ex.get(k)
            if isinstance(v, float) and isinstance(ev, float):
                if abs(v - ev) > 0.005:
                    diff.append((k, v, ev))
            elif v != ev:
                diff.append((k, v, ev))
        if diff:
            mismatches.append((key, diff))
        skipped.append((key, 'exists'))
        continue
    new_records.append(rec)

print('skipped duplicates:', len(skipped), ' value mismatches among skipped:', len(mismatches))
for k, d in mismatches:
    print('  MISMATCH:', k, d)

# ---------- 4. 生成新记录行并插入 ----------
ORDER = ['id', 'period', 'sampling_time', 'rive_name', 'rive_abbreviation', 'river_number', 'latitude', 'longitude',
         'plastics_splinter_type', 'plastics_fibre_type', 'plastics_grain_type', 'plastics_membrane_type', 'plastics_foam_type',
         'plastics_red_color', 'plastics_black_color', 'plastics_white_color', 'plastics_transparent_color', 'plastics_blue_color', 'plastics_other_color',
         'fifty_radius', 'one_hundred_radius', 'three_hundred_radius', 'five_hundred_radius', 'one_thousand_radius', 'two_thousand_radius', 'three_thousand_radius', 'five_thousand_radius',
         'plastics_total', 'water_volume', 'plastics_abundance']

def js_val(k, v):
    if k in ('period', 'sampling_time', 'rive_name', 'rive_abbreviation', 'river_number'):
        return "'" + str(v) + "'"
    if k in ('latitude', 'longitude'):
        return fmt_coord(v)
    if k == 'plastics_abundance':
        return fmt_abd(v)
    if k == 'water_volume':
        return fmt_vol(v)
    return str(v)

next_id = max(int(e['id']) for e in existing) + 1
lines = []
for rec in new_records:
    rec2 = dict(rec)
    rec2['id'] = next_id
    next_id += 1
    body = ', '.join('%s: %s' % (k, js_val(k, rec2[k])) for k in ORDER)
    lines.append('            {' + body + '},')

new_html = html[:end]
if not new_html.endswith('\n'):
    new_html += '\n'
new_html += '\n'.join(lines) + '\n        ' + html[end:]

# localStorage 键升级，避免浏览器旧缓存遮住新种子数据
new_html = new_html.replace("localStorage.getItem('water_data')", "localStorage.getItem('water_data_v2')")
new_html = new_html.replace("localStorage.setItem('water_data',", "localStorage.setItem('water_data_v2',")

with open(HTML, 'w', encoding='utf-8') as f:
    f.write(new_html)

print('added records:', len(new_records), 'id', max(int(e['id']) for e in existing) + 1, '~', next_id - 1)
by_period = {}
for r in new_records:
    by_period[r['period']] = by_period.get(r['period'], 0) + 1
print('added by period:', by_period)
print('localStorage key -> water_data_v2')
