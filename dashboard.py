# -*- coding: utf-8 -*-
"""区域末销 / 存销比看板 (P0)

输入: 平台库存末销数据盘点 + carflow + 经销商主数据 + config/targets.json
输出: 单文件 HTML (内嵌数据, 自绘图表, 离线可开)

口径:
  月末销 = 平台「末销时间」口径
  在库   = 经销商已入库 + 进口商已入库, 且未末销
  在途   = 在途=是 且未末销
  月均 A = 截至上月近 6 个完整月（最近半年）的末销均值; 无销量时回退 carflow 当月目标
  存销比 = 库存 / A, 在库/在途/合计各算
"""
import datetime as dt
import json
import os
import sys
from collections import Counter, defaultdict

import openpyxl

from vehicle_flow import (BASE, INBOX, OUT_DIR, CONFIG_DIR, find_snapshot,
                          load_snapshot, state_of, parse_date, load_config,
                          active_dealers)

DASH_DIR = os.path.join(OUT_DIR, 'dashboard')
TARGETS_PATH = os.path.join(BASE, 'config', 'targets.json')

# 月均 A 的计算窗口: 最近半年 (6 个完整月)。代码常量, 不读 targets.json,
# 保证已安装用户的本地旧 config 不会覆盖此口径。
WINDOW_MONTHS = 6

ALIAS = {}  # 国家别名映射 (如快照里同一国家有多种写法时填这里)


# ---------------------------------------------------------------- 工具

def add_months(ym, n):
    y, m = int(ym[:4]), int(ym[5:7])
    m += n
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return '%04d-%02d' % (y, m)


def month_key(value):
    if not value:
        return ''
    if isinstance(value, (dt.datetime, dt.date)):
        return '%04d-%02d' % (value.year, value.month)
    s = str(value).strip()
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S'):
        try:
            d = dt.datetime.strptime(s, fmt)
            return '%04d-%02d' % (d.year, d.month)
        except ValueError:
            pass
    return s[:7] if len(s) >= 7 else ''


def norm_country(name):
    name = (name or '').strip()
    return ALIAS.get(name, name)


def load_targets():
    cfg = {'window_months': WINDOW_MONTHS,
           'default': {'target_ratio': 2.0, 'band': 0.3, 'min_keep_units': 1},
           'by_country': {}}
    if os.path.exists(TARGETS_PATH):
        with open(TARGETS_PATH, encoding='utf-8') as f:
            cfg.update(json.load(f))
    return cfg


def save_targets(cfg):
    os.makedirs(os.path.dirname(TARGETS_PATH), exist_ok=True)
    with open(TARGETS_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def target_for(cfg, country):
    c = cfg.get('by_country', {}).get(country, {})
    base = dict(cfg.get('default', {}))
    base.update(c)
    return base


def load_carflow(path):
    """返回 {国家: {'YYYY-MM': 台数}}, 只取「末销」项目。"""
    out = defaultdict(Counter)
    if not path or not os.path.exists(path):
        return out
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.worksheets[0]
    header = list(next(ws.iter_rows(max_row=1, values_only=True)))
    idx = {h: i for i, h in enumerate(header) if h}
    if '国家' not in idx or '项目' not in idx or '1月' not in idx:
        wb.close()
        return out
    year = dt.date.today().year
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[idx['国家']] or str(r[idx['项目']]).strip() != '末销':
            continue
        country = norm_country(str(r[idx['国家']]))
        for m in range(1, 13):
            v = r[idx['%d月' % m]] if '%d月' % m in idx else 0
            if v:
                out[country]['%04d-%02d' % (year, m)] += int(v)
    wb.close()
    return out


# ---------------------------------------------------------------- 数据

def build_data(snapshot, carflow_path, month, targets):
    rows = load_snapshot(snapshot)
    # 快照截止日: 取末销/入库/创建时间里最新的日期, 上限为今天
    dates = []
    for r in rows:
        for v in (r['sold_at'], r['dealer_in'], r['created']):
            if v:
                try:
                    dates.append(parse_date(v))
                except Exception:
                    pass
    cutoff = max(dates) if dates else dt.date.today()
    cutoff = min(cutoff, dt.date.today())

    month = month or '%04d-%02d' % (cutoff.year, cutoff.month)
    n = WINDOW_MONTHS
    windows = [add_months(month, -i) for i in range(n, 0, -1)]
    trend_months = [add_months(month, -i) for i in range(11, -1, -1)]
    flow = load_carflow(carflow_path)

    # 经销商主数据
    status = {}
    cfgs = {}
    for imp in sorted({r['imp'] for r in rows if r['imp']}):
        try:
            cfgs[imp] = load_config(imp, rows)
        except Exception as e:
            print('⚠ 进口商 %s 主数据加载失败，已跳过（不影响其余进口商）：%s' % (imp, e))
            continue
        for d in cfgs[imp]['dealers']:
            status[(imp, d['code'])] = d
    # 补齐未在主数据里的经销商
    for r in rows:
        key = (r['imp'], r['dealer'])
        if r['dealer'] and key not in status:
            status[key] = {'code': r['dealer'], 'name': r['dealer_name'],
                           'status': '未知', 'outlets': 1}

    countries = {}
    dealers = {}
    models = defaultdict(lambda: {'sold_ytd': 0, 'in_stock': 0, 'in_transit': 0})

    for r in rows:
        imp = r['imp']
        if not imp:
            continue
        country = norm_country(r['country']) or imp
        c = countries.setdefault(country, {
            'code': imp, 'name': country,
            'in_stock': 0, 'in_transit': 0, 'importer_stock': 0,
            'sold_ytd': 0, 'sold_mtd': 0, 'sold_total': 0,
            'months': Counter(), 'models': Counter(), 'last_sold': None,
        })
        st = state_of(r)
        if st == 'sold':
            c['sold_total'] += 1
            if r['sold_at']:
                sd = parse_date(r['sold_at'])
                if c['last_sold'] is None or sd > c['last_sold']:
                    c['last_sold'] = sd
            mk = month_key(r['sold_at'])
            if mk:
                c['months'][mk] += 1
                if mk[:4] == month[:4] and mk <= month:
                    c['sold_ytd'] += 1
                if mk == month:
                    c['sold_mtd'] += 1
            models[r['model']]['sold_ytd'] += 1 if (mk[:4] == month[:4] and mk <= month) else 0
        elif st == 'transit':
            c['in_transit'] += 1
            models[r['model']]['in_transit'] += 1
        elif st == 'importer':
            c['in_stock'] += 1
            c['importer_stock'] += 1
            models[r['model']]['in_stock'] += 1
        else:
            c['in_stock'] += 1
            models[r['model']]['in_stock'] += 1

        if r['dealer']:
            d = dealers.setdefault((imp, r['dealer']), {
                'dealer': r['dealer'], 'name': r['dealer_name'], 'country': country,
                'in_stock': 0, 'in_transit': 0, 'sold_mtd': 0, 'sold_ytd': 0,
                'months': Counter(),
            })
            if st == 'dealer' or st == 'importer':
                d['in_stock'] += 1
            elif st == 'transit':
                d['in_transit'] += 1
            elif st == 'sold':
                mk = month_key(r['sold_at'])
                if mk:
                    d['months'][mk] += 1
                    if mk[:4] == month[:4] and mk <= month:
                        d['sold_ytd'] += 1
                    if mk == month:
                        d['sold_mtd'] += 1

    def avg_of(counter):
        vals = [counter.get(m, 0) for m in windows]
        if sum(vals) > 0:
            # 首笔末销之前的月份不计入分母, 避免新市场被零销月拉低
            first = min(m for m in counter if counter[m] > 0 and m <= windows[-1])
            used = [m for m in windows if m >= first]
            return sum(counter[m] for m in used) / float(len(used)), '近%d月' % len(used)
        hist_start = add_months(month, -12)
        hist = [counter[m] for m in sorted(counter) if hist_start <= m < month]
        if hist and sum(hist) > 0:
            return sum(hist) / float(len(hist)), '历史均值'
        return 0.0, '无'

    out_countries = []
    for name, c in countries.items():
        a, a_src = avg_of(c['months'])
        target_mtd = flow.get(name, {}).get(month, 0) or 0
        window_detail = [{'m': m, 'v': c['months'].get(m, 0)} for m in windows]
        if a == 0 and target_mtd:
            a, a_src = float(target_mtd), 'carflow目标'
        t = target_for(targets, name)
        band = float(t.get('band', 0.3))
        r_star = float(t.get('target_ratio', 2.0))
        has_stock = (c['in_stock'] + c['in_transit']) > 0
        has_recent = sum(d['v'] for d in window_detail) > 0
        r_stock = r_transit = r_total = None
        if not has_stock and not has_recent:
            # 快照里既没库存也没近期销量: 不参与体检, 避免用陈旧销量或 carflow 目标硬算
            state, advice, gap = 'na', '快照无该市场库存与近期销量，暂不体检', 0
        else:
            r_stock = (c['in_stock'] / a) if a else None
            r_transit = (c['in_transit'] / a) if a else None
            r_total = ((c['in_stock'] + c['in_transit']) / a) if a else None
            if r_total is None:
                state, advice, gap = 'na', '无销量基准，先补历史或 carflow 目标', 0
            elif r_total > r_star + band:
                state = 'high'
                gap = (r_total - r_star) * a
                advice = '偏高，本月建议多出 %d 台或暂缓入库' % round(gap)
            elif r_total < r_star - band:
                state = 'low'
                gap = (r_star - r_total) * a
                advice = '偏低，建议补 %d 台或在途尽快分店' % round(gap)
            else:
                state, advice, gap = 'ok', '健康，按计划进货与末销', 0

        active = [d for d in cfgs.get(c['code'], {}).get('dealers', [])
                  if d['status'] == '正常运营']
        if not active:
            active = [d for d in cfgs.get(c['code'], {}).get('dealers', [])
                      if d['code'] == c['code']]
        with_stock = [d for d in active
                      if dealers.get((c['code'], d['code']), {}).get('in_stock', 0) > 0]
        coverage = '%d/%d' % (len(with_stock), len(active)) if active else '—'

        # 全期月均: 自首笔末销月至上月, 用于判断近 N 月是否被数据滞后压低
        past = sorted(m for m in c['months'] if m < month)
        long_sold = sum(c['months'][m] for m in past)
        avg_long = 0.0
        if past and state != 'na':
            f, l = past[0], add_months(month, -1)
            span = (int(l[:4]) - int(f[:4])) * 12 + int(l[5:7]) - int(f[5:7]) + 1
            if span > 0:
                avg_long = long_sold / float(span)
        ratio_long = ((c['in_stock'] + c['in_transit']) / avg_long) if avg_long else None

        out_countries.append({
            'name': name, 'code': c['code'],
            'in_stock': c['in_stock'], 'in_transit': c['in_transit'],
            'importer_stock': c['importer_stock'],
            'sold_ytd': c['sold_ytd'], 'sold_mtd': c['sold_mtd'],
            'avg': round(a, 1), 'avg_src': a_src,
            'avg_long': round(avg_long, 1),
            'ratio_long': round(ratio_long, 2) if ratio_long is not None else None,
            'ratio_stock': round(r_stock, 2) if r_stock is not None else None,
            'ratio_transit': round(r_transit, 2) if r_transit is not None else None,
            'ratio_total': round(r_total, 2) if r_total is not None else None,
            'target': r_star, 'band': band, 'state': state, 'advice': advice,
            'gap_units': round(gap), 'target_mtd': target_mtd,
            'window_detail': window_detail,
            'months_map': dict(c['months']),
            'last_sold': c['last_sold'].isoformat() if c['last_sold'] else None,
            'achieve': round(c['sold_mtd'] / target_mtd * 100, 1) if target_mtd else None,
            'coverage': coverage,
        })

    data_month = max((c['last_sold'] for c in out_countries if c['last_sold']),
                     default=None)

    warnings = []
    for c in out_countries:
        if c['state'] == 'na':
            if c['target_mtd']:
                warnings.append('%s 快照无库存、近 %d 月无末销，但 carflow %s 有末销目标 %d 台，'
                                '请确认是否漏导该市场数据' %
                                (c['name'], n, month, c['target_mtd']))
            continue
        if c['last_sold']:
            ls = parse_date(c['last_sold'])
            ls_month = '%04d-%02d' % (ls.year, ls.month)
            if ls_month < data_month[:7]:
                warnings.append('%s 末销数据滞后：最新一笔 %s，%s 之后无末销记录；'
                                '近 %d 月分母被低估、存销比偏高' %
                                (c['name'], c['last_sold'], ls_month, n))
            elif ls.day < 20 and ls_month >= windows[0]:
                warnings.append('%s 末月（%s）末销仅到 %s，当月未完整；'
                                '该月若在均值窗口内会拉低分母、抬高存销比' %
                                (c['name'], ls_month, c['last_sold']))
        sold_months = [d['m'] for d in c['window_detail'] if d['v'] > 0]
        zeros = [d['m'] for d in c['window_detail']
                 if d['v'] == 0 and sold_months and d['m'] > sold_months[0]]
        if not zeros:
            continue
        if c['avg_src'].startswith('近'):
            warnings.append('%s 均值窗口内有 %d 个零销月（%s），月均被拉低、存销比偏高；'
                            '若为平台末销数据导出延迟，待数据补全后重算' %
                            (c['name'], len(zeros), '、'.join(zeros)))
        else:
            warnings.append('%s 近 %d 月无末销记录，月均 A 暂用「%s」= %.1f 台，'
                            '存销比为估算值' % (c['name'], n, c['avg_src'], c['avg']))

    out_dealers = []
    for (imp, code), d in dealers.items():
        if d['in_stock'] == 0 and d['in_transit'] == 0 and d['sold_ytd'] == 0:
            continue
        a, a_src = avg_of(d['months'])
        country = d['country']
        cinfo = next((x for x in out_countries if x['name'] == country), None)
        if a == 0 and cinfo:
            a, a_src = float(cinfo['avg']), '国家均值'
        ratio = (d['in_stock'] / a) if a else None
        t = target_for(targets, country)
        r_star, band = float(t['target_ratio']), float(t['band'])
        if ratio is None:
            state, advice = 'na', '无销量基准'
        elif ratio > r_star + band:
            state, advice = 'high', '多出 %d 台' % round((ratio - r_star) * a)
        elif ratio < r_star - band:
            state, advice = 'low', '补 %d 台' % round((r_star - ratio) * a)
        else:
            state, advice = 'ok', '正常'
        meta = status.get((imp, code), {})
        st = meta.get('status', '未知')
        if code == imp:
            st = '进口商自营'
        out_dealers.append({
            'country': country, 'code': code, 'name': d['name'] or meta.get('name', ''),
            'status': st, 'outlets': meta.get('outlets', 1),
            'in_stock': d['in_stock'], 'in_transit': d['in_transit'],
            'mtd': d['sold_mtd'], 'ytd': d['sold_ytd'],
            'avg': round(a, 1), 'avg_src': a_src,
            'ratio': round(ratio, 2) if ratio is not None else None,
            'state': state, 'advice': advice,
        })

    out_models = []
    for m, v in models.items():
        if not m:
            continue
        total = v['sold_ytd'] + v['in_stock'] + v['in_transit']
        if total == 0:
            continue
        out_models.append({'model': m, 'sold_ytd': v['sold_ytd'],
                           'in_stock': v['in_stock'], 'in_transit': v['in_transit'],
                           'total': total})
    out_models.sort(key=lambda x: -x['total'])

    def region_avg(cinfo):
        return sum(x['avg'] for x in out_countries if x['avg']) or 0

    tot = {
        'sold_ytd': sum(x['sold_ytd'] for x in out_countries),
        'sold_mtd': sum(x['sold_mtd'] for x in out_countries),
        'in_stock': sum(x['in_stock'] for x in out_countries),
        'in_transit': sum(x['in_transit'] for x in out_countries),
        'avg': round(sum(x['avg'] for x in out_countries), 1),
        'ratio': round(sum(x['in_stock'] + x['in_transit'] for x in out_countries) /
                       sum(x['avg'] for x in out_countries), 2) if sum(x['avg'] for x in out_countries) else None,
        'models': len(out_models),
    }

    trend = []
    region_months = Counter()
    for c in countries.values():
        region_months.update(c['months'])
    for m in trend_months:
        trend.append({'m': m, 'v': region_months.get(m, 0)})

    missing = [k for k in flow if k not in countries and
               sum(flow[k].values()) > 0]
    if missing:
        warnings.append('carflow 有末销目标、但快照无库存数据的国家共 %d 个（%s 等），详见页脚' %
                        (len(missing), '、'.join(sorted(missing)[:3])))

    all_months = sorted({m for c in countries.values() for m in c['months']} | {month},
                        reverse=True)

    return {
        'meta': {
            'generated': dt.datetime.now().strftime('%Y-%m-%d %H:%M'),
            'cutoff': cutoff.isoformat(), 'month': month,
            'windows': windows, 'snapshot': os.path.basename(snapshot),
            'window_months': n,
        },
        'totals': tot, 'countries': out_countries,
        'dealers': out_dealers, 'models': out_models[:12],
        'trend': trend, 'missing': sorted(missing), 'warnings': warnings,
        'all_months': all_months, 'carflow_flow': flow,
    }


# ---------------------------------------------------------------- HTML

TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>区域末销 / 存销比看板</title>
<style>
:root{--bg:#f5f6f8;--card:#fff;--ink:#1b2430;--sub:#66707d;--line:#e6e9ee;
--blue:#1f6feb;--green:#16a34a;--amber:#d97706;--red:#dc2626;--violet:#7c3aed;--teal:#0d9488}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 -apple-system,"PingFang SC","Microsoft YaHei",sans-serif}
.wrap{max-width:1280px;margin:0 auto;padding:24px 20px 48px}
h1{font-size:22px;margin:0 0 4px}
.meta{color:var(--sub);font-size:12px}
.note{color:var(--sub);font-size:12px;margin-top:6px}
.grid{display:grid;gap:14px}
.kpis{grid-template-columns:repeat(6,1fr);margin:18px 0}
@media(max-width:980px){.kpis{grid-template-columns:repeat(3,1fr)}}
@media(max-width:640px){.kpis{grid-template-columns:repeat(2,1fr)}}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px 18px}
.kpi .k{color:var(--sub);font-size:12px}
.kpi .v{font-size:26px;font-weight:600;letter-spacing:-.5px}
.kpi .s{color:var(--sub);font-size:11px;margin-top:2px}
.sec{margin-top:22px}
.sec h2{font-size:15px;margin:0 0 10px;display:flex;align-items:center;gap:8px}
.sec h2 span{font-weight:400;color:var(--sub);font-size:12px}
.pill{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;white-space:nowrap}
.pill.ok{background:#e8f7ee;color:#147a3d}.pill.high{background:#fdeaea;color:#b91c1c}
.pill.low{background:#fff4e5;color:#b45309}.pill.na{background:#eef0f3;color:#5b6470}
.row{display:grid;grid-template-columns:120px 1fr 150px;gap:12px;align-items:center;padding:9px 0;border-top:1px dashed var(--line)}
.row:first-of-type{border-top:0}
.row .nm{font-weight:600}
.row .sub{color:var(--sub);font-size:11px}
.bars{display:flex;flex-direction:column;gap:5px}
.bar{height:14px;border-radius:7px;background:#eef1f5;position:relative;overflow:hidden}
.bar i{position:absolute;left:0;top:0;bottom:0;border-radius:7px}
.bar .lbl{position:absolute;right:6px;top:-1px;font-size:10px;color:#39424e}
.legend{display:flex;gap:14px;color:var(--sub);font-size:12px;margin:4px 0 10px;flex-wrap:wrap}
.legend b{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:5px}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{padding:8px 10px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap}
th{color:var(--sub);font-weight:500;cursor:pointer;user-select:none;font-size:12px}
th.l,td.l{text-align:left}
tr:hover td{background:#fafbfc}
.tblwrap{overflow:auto}
select{border:1px solid var(--line);background:#fff;border-radius:8px;padding:5px 8px;font-size:13px}
.flex{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap}
.stack{display:flex;height:16px;border-radius:8px;overflow:hidden;background:#eef1f5}
.stack div{height:16px}
.foot{margin-top:26px;color:var(--sub);font-size:12px;line-height:1.7}
.warn{margin:14px 0 0;padding:10px 14px;border:1px solid #f3d2a2;background:#fff8ee;
  border-radius:10px;color:#8a4b08;font-size:12.5px;line-height:1.8}
.warn b{color:#b45309}
svg{width:100%;height:auto}
.small{font-size:11px;color:var(--sub)}
</style>
</head>
<body>
<div class="wrap">
  <h1>区域末销 / 存销比看板</h1>
  <div class="meta" id="meta"></div>
  <div style="margin:8px 0 12px;display:flex;align-items:center;gap:8px;flex-wrap:wrap">
    <label style="font-size:13px;color:var(--sub)">统计月</label>
    <select id="monthPick" style="padding:5px 8px;border:1px solid #d6dde6;border-radius:6px;font-size:13px;background:#fff"></select>
    <span class="small">切换后按所选月份重算近 N 月均值与存销比；在库 / 在途为快照当前值</span>
  </div>
  <div id="warn"></div>

  <div class="grid kpis" id="kpis"></div>

  <div class="sec card">
    <h2>分国家 <span>末销累计 vs 当前库存 · 存销比与目标</span></h2>
    <div class="legend"><span><b style="background:var(--blue)"></b>累计末销</span>
      <span><b style="background:var(--teal)"></b>在库</span>
      <span><b style="background:var(--amber)"></b>在途</span></div>
    <div id="countries"></div>
  </div>

  <div class="sec card">
    <h2>分车型 <span>按累计末销 + 当前库存排序（前 12 款）</span></h2>
    <div class="legend"><span><b style="background:var(--blue)"></b>累计末销</span>
      <span><b style="background:var(--teal)"></b>在库</span>
      <span><b style="background:var(--amber)"></b>在途</span></div>
    <div id="models"></div>
  </div>

  <div class="sec card">
    <h2>月度末销趋势 <span>近 12 个月（区域合计）</span></h2>
    <div id="trend"></div>
  </div>

  <div class="sec card">
    <div class="flex"><h2 style="margin:0">单店明细 <span>末销口径 = 平台末销时间</span></h2>
      <div><select id="cfilter"></select></div></div>
    <div class="tblwrap" style="margin-top:10px"><table id="dealerTbl"></table></div>
  </div>

  <div class="sec card">
    <h2>存销比体检 <span>在库 / 在途 / 合计，对比国家目标带</span></h2>
    <div class="tblwrap"><table id="ratioTbl"></table></div>
  </div>

  <div class="foot" id="foot"></div>
</div>

<script>
const DATA = __DATA__;
const COL = {blue:'#1f6feb',teal:'#0d9488',amber:'#d97706',green:'#16a34a',red:'#dc2626',violet:'#7c3aed'};
const $ = s => document.querySelector(s);
const nf = v => v==null ? '—' : v.toLocaleString('zh-CN');
const rf = v => v==null ? '—' : v.toFixed(2);
const pill = (st,txt) => `<span class="pill ${st}">${txt}</span>`;
let V = {month: DATA.meta.month, windows: DATA.meta.windows.slice(), countries: DATA.countries};
const N = DATA.meta.window_months;
function addMonthsYM(ym,n){ const [y,m]=ym.split('-').map(Number); const d=new Date(y,m-1+n,1);
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'); }
// 镜像服务端 avg_of + 国家存销比口径, 保证默认月与生成时一致, 切月时前端重算
function calcCountry(c, month){
  const mm = c.months_map || {};
  const windows = [];
  for(let i=N;i>0;i--) windows.push(addMonthsYM(month,-i));
  const vals = windows.map(m=>mm[m]||0);
  let a=0, a_src='无';
  if(vals.reduce((s,v)=>s+v,0) > 0){
    const cands = Object.keys(mm).filter(m=>mm[m]>0 && m <= windows[windows.length-1]);
    const first = cands.length ? Math.min(...cands) : null;
    const used = first ? windows.filter(m=>m>=first) : windows.slice();
    a = used.reduce((s,m)=>s+(mm[m]||0),0)/used.length; a_src='近'+used.length+'月';
  } else {
    const hs = addMonthsYM(month,-12);
    const hist = Object.keys(mm).filter(m=>m>=hs && m<month).map(m=>mm[m]||0);
    if(hist.length && hist.reduce((s,v)=>s+v,0)>0){ a=hist.reduce((s,v)=>s+v,0)/hist.length; a_src='历史均值'; }
  }
  const flow = DATA.carflow_flow || {};
  const target_mtd = (flow[c.name] && flow[c.name][month]) || 0;
  if(a===0 && target_mtd){ a=target_mtd; a_src='carflow目标'; }
  const r_star=c.target, band=c.band, in_stock=c.in_stock, in_transit=c.in_transit;
  const r_stock = a ? in_stock/a : null;
  const r_transit = a ? in_transit/a : null;
  const r_total = a ? (in_stock+in_transit)/a : null;
  let state, advice, gap=0;
  if(r_total==null){ state='na'; advice='无销量基准，先补历史'; }
  else if(r_total > r_star+band){ state='high'; gap=(r_total-r_star)*a; advice='偏高，本月建议多出 '+Math.round(gap)+' 台或暂缓入库'; }
  else if(r_total < r_star-band){ state='low'; gap=(r_star-r_total)*a; advice='偏低，建议补 '+Math.round(gap)+' 台或在途尽快分店'; }
  else { state='ok'; advice='健康，按计划进货与末销'; }
  const year=month.slice(0,4);
  const sold_ytd = Object.keys(mm).filter(m=>m.slice(0,4)===year && m<=month).reduce((s,m)=>s+mm[m],0);
  const sold_mtd = mm[month]||0;
  return {avg:+a.toFixed(1), avg_src: a_src,
    ratio_stock:r_stock==null?null:+r_stock.toFixed(2),
    ratio_transit:r_transit==null?null:+r_transit.toFixed(2),
    ratio_total:r_total==null?null:+r_total.toFixed(2),
    state, advice, gap_units:Math.round(gap),
    sold_ytd, sold_mtd, window_detail:windows.map(m=>({m,v:mm[m]||0})), target_mtd};
}
function renderAll(selMonth, recompute){
  V.month = selMonth;
  V.windows = []; for(let i=N;i>0;i--) V.windows.push(addMonthsYM(selMonth,-i));
  V.countries = recompute ? DATA.countries.map(c=>Object.assign({}, c, calcCountry(c, selMonth))) : DATA.countries;
  meta(); kpis(); countries(); ratioTable(); foot();
}

function meta(){
  const m = DATA.meta;
  $('#meta').innerHTML = `数据截止 ${m.cutoff} · 统计月 ${V.month} · 月均窗口 ${V.windows.join(' / ')} · 来源 ${m.snapshot} · 生成 ${m.generated}`;
}
function warn(){
  const list = DATA.warnings || [];
  if(!list.length) return;
  $('#warn').innerHTML = `<div class="warn"><b>数据质量提示</b>` +
    list.map(t=>`<br>· ${t}`).join('') + `</div>`;
}
function kpis(){
  const cs = V.countries, t = DATA.totals;
  const ytd = cs.reduce((s,c)=>s+(c.sold_ytd||0),0);
  const mtd = cs.reduce((s,c)=>s+(c.sold_mtd||0),0);
  const avgs = cs.reduce((s,c)=>s+(c.avg||0),0);
  const ratio = avgs ? +((t.in_stock+t.in_transit)/avgs).toFixed(2) : null;
  const items = [
    ['累计末销', nf(ytd), '台 · '+V.month.slice(0,4)+' 年至今'],
    ['本月末销', nf(mtd), '台 · '+V.month],
    ['在库', nf(t.in_stock), '台（含进口商库存）'],
    ['在途', nf(t.in_transit), '台'],
    ['合计存销比', rf(ratio), '总库存 / 月均 '+nf(avgs)+' 台'],
    ['在营国家 / 车型', cs.length+' / '+t.models, '个 / 款'],
  ];
  $('#kpis').innerHTML = items.map(([k,v,s])=>`<div class="card kpi"><div class="k">${k}</div>
    <div class="v">${v}</div><div class="s">${s}</div></div>`).join('');
}
function bar(v,max,color){
  const w = max ? Math.max(2, v/max*100) : 0;
  return `<div class="bar"><i style="width:${w}%;background:${color}"></i>
    <span class="lbl">${nf(v)}</span></div>`;
}
function countries(){
  const max = Math.max(...V.countries.map(c=>Math.max(c.sold_ytd,c.in_stock,c.in_transit)),1);
  $('#countries').innerHTML = V.countries.map(c=>`
    <div class="row">
      <div><div class="nm">${c.name}</div><div class="sub">${c.code} · 覆盖 ${c.coverage}</div></div>
      <div class="bars">
        ${bar(c.sold_ytd,max,COL.blue)}${bar(c.in_stock,max,COL.teal)}${bar(c.in_transit,max,COL.amber)}
      </div>
      <div>
        ${pill(c.state, c.ratio_total==null?'无基准':'存销比 '+rf(c.ratio_total)+' / 目标 '+c.target)}
        <div class="sub" style="margin-top:4px">在库 ${rf(c.ratio_stock)} · 在途 ${rf(c.ratio_transit)}</div>
        <div class="sub">${c.advice}</div>
        <div class="sub">近 ${N} 月：${(c.window_detail||[]).map(d=>d.m.slice(5)+'月 '+d.v).join(' / ')}</div>
      </div>
    </div>`).join('');
}
function models(){
  const max = Math.max(...DATA.models.map(m=>m.total),1);
  $('#models').innerHTML = DATA.models.map(m=>`
    <div class="row" style="grid-template-columns:80px 1fr 220px">
      <div><div class="nm">${m.model}</div><div class="sub">合计 ${nf(m.total)}</div></div>
      <div class="stack">
        <div style="width:${m.sold_ytd/max*100}%;background:var(--blue)"></div>
        <div style="width:${m.in_stock/max*100}%;background:var(--teal)"></div>
        <div style="width:${m.in_transit/max*100}%;background:var(--amber)"></div>
      </div>
      <div class="sub">末销 ${nf(m.sold_ytd)} · 在库 ${nf(m.in_stock)} · 在途 ${nf(m.in_transit)}</div>
    </div>`).join('');
}
function trend(){
  const d = DATA.trend, W=1100, H=220, P={l:46,r:16,t:14,b:28};
  const max = Math.max(...d.map(x=>x.v),1);
  const xs = i => P.l + i*(W-P.l-P.r)/Math.max(d.length-1,1);
  const ys = v => H-P.b - (v/max)*(H-P.t-P.b);
  const path = d.map((x,i)=>`${i?'L':'M'}${xs(i)},${ys(x.v)}`).join(' ');
  const grid = [0,.25,.5,.75,1].map(r=>{
    const y=H-P.b-r*(H-P.t-P.b);
    return `<line x1="${P.l}" y1="${y}" x2="${W-P.r}" y2="${y}" stroke="#eef1f5"/>
      <text x="${P.l-8}" y="${y+4}" text-anchor="end" font-size="10" fill="#8a93a0">${Math.round(max*r)}</text>`;
  }).join('');
  const dots = d.map((x,i)=>`<circle cx="${xs(i)}" cy="${ys(x.v)}" r="3" fill="#1f6feb"><title>${x.m}: ${x.v}</title></circle>`).join('');
  const labels = d.map((x,i)=>`<text x="${xs(i)}" y="${H-8}" text-anchor="middle" font-size="10" fill="#8a93a0">${x.m.slice(2)}</text>`).join('');
  $('#trend').innerHTML = `<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet">
    ${grid}<path d="${path}" fill="none" stroke="#1f6feb" stroke-width="2.5"/>${dots}${labels}</svg>`;
}
function dealerRows(){
  const c = $('#cfilter').value;
  let rows = DATA.dealers.filter(d=>!c || d.country===c);
  rows.sort((a,b)=>b.ytd-a.ytd);
  const html = `<tr><th class="l">经销商</th><th class="l">国家</th><th class="l">状态</th>
    <th>网点</th><th>本月末销</th><th>累计末销</th><th>在库</th><th>在途</th><th>月均</th><th>存销比</th><th class="l">建议</th></tr>` +
    rows.map(d=>`<tr><td class="l">${d.code} ${d.name?'· '+d.name:''}</td><td class="l">${d.country}</td>
      <td class="l">${d.status}</td><td>${d.outlets}</td><td>${nf(d.mtd)}</td><td>${nf(d.ytd)}</td>
      <td>${nf(d.in_stock)}</td><td>${nf(d.in_transit)}</td><td>${d.avg}</td><td>${rf(d.ratio)}</td>
      <td class="l">${pill(d.state,d.advice)}</td></tr>`).join('');
  $('#dealerTbl').innerHTML = html;
}
function ratioTable(){
  const head = `<tr><th class="l">国家</th><th class="l">近 ${N} 月末销（月/台）</th>
    <th>在库</th><th>在库比</th><th>在途</th><th>在途比</th>
    <th>合计</th><th>合计比</th><th>目标</th><th>Δ</th><th class="l">建议</th></tr>`;
  const rows = V.countries.map(c=>{
    const delta = c.ratio_total==null ? null : (c.ratio_total - c.target);
    const wd = (c.window_detail||[]).map(d=>`${d.m.slice(5)}月 <b>${d.v}</b>`).join(' · ');
    return `<tr><td class="l"><b>${c.name}</b><div class="small">均值 ${c.avg} 台/月（${c.avg_src}）
      · 全期 ${c.avg_long} 台/月 → 比 ${rf(c.ratio_long)}</div></td>
      <td class="l small">${wd}</td><td>${nf(c.in_stock)}</td><td>${rf(c.ratio_stock)}</td>
      <td>${nf(c.in_transit)}</td><td>${rf(c.ratio_transit)}</td>
      <td>${nf(c.in_stock+c.in_transit)}</td><td>${rf(c.ratio_total)}</td>
      <td>${c.target} ± ${c.band}</td><td>${delta==null?'—':(delta>0?'+':'')+delta.toFixed(2)}</td>
      <td class="l">${pill(c.state,c.advice)}</td></tr>`;
  }).join('');
  $('#ratioTbl').innerHTML = head + rows;
}
function foot(){
  const m = DATA.meta;
  $('#foot').innerHTML = `口径：累计末销按平台末销时间统计；在库含经销商已入库与进口商已入库；在途为在途=是。<br>
    存销比分母为近 ${m.window_months} 个完整月（${V.windows.join('、')}）的月均末销；无销量时回退 carflow 当月末销目标。<br>
    carflow 中有目标但快照暂无库存数据的国家：${DATA.missing.length?DATA.missing.join('、'):'无'}。`;
}
(function init(){
  const mp = $('#monthPick');
  mp.innerHTML = (DATA.all_months||[]).map(m=>`<option value="${m}" ${m===DATA.meta.month?'selected':''}>${m}</option>`).join('');
  mp.onchange = ()=>renderAll(mp.value, true);
  renderAll(DATA.meta.month, false);
  warn();models();trend();
  const sel = $('#cfilter');
  sel.innerHTML = `<option value="">全部国家</option>` + DATA.countries.map(c=>`<option>${c.name}</option>`).join('');
  sel.onchange = dealerRows;
  dealerRows();
})();
</script>
</body>
</html>
"""


def render(data):
    payload = json.dumps(data, ensure_ascii=False).replace('</', '<\\/')
    return TEMPLATE.replace('__DATA__', payload)


def cmd_dashboard(args):
    snapshot = find_snapshot(args.snapshot)
    targets = load_targets()
    carflow = args.carflow or os.path.join(INBOX, '月度末销目标.xlsx')
    data = build_data(snapshot, carflow, args.month, targets)

    if not os.path.exists(TARGETS_PATH):
        save_targets(targets)
        print('已生成 %s（分国家目标可在此调整）' % TARGETS_PATH)

    month = data['meta']['month']
    out_path = args.out or os.path.join(DASH_DIR, '看板_%s.html' % month)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(render(data))
    print('看板已生成: %s' % out_path)
    print('  累计末销 %d 台 · 在库 %d · 在途 %d · 合计存销比 %s' % (
        data['totals']['sold_ytd'], data['totals']['in_stock'],
        data['totals']['in_transit'], data['totals']['ratio']))
    for c in sorted(data['countries'], key=lambda x: -(x['ratio_total'] or 0)):
        print('  %-12s 末销 %5d | 在库 %4d | 在途 %4d | 存销比 %s | %s' % (
            c['name'], c['sold_ytd'], c['in_stock'], c['in_transit'],
            c['ratio_total'] if c['ratio_total'] is not None else '—', c['advice']))
    return 0


def main():
    import argparse
    p = argparse.ArgumentParser(description='区域末销/存销比看板')
    p.add_argument('--snapshot')
    p.add_argument('--carflow')
    p.add_argument('--month', help='统计月, 如 2026-11, 默认按快照截止日')
    p.add_argument('--out')
    args = p.parse_args()
    sys.exit(cmd_dashboard(args) or 0)


if __name__ == '__main__':
    main()
