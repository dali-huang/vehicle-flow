# -*- coding: utf-8 -*-
"""车辆入库 / 末销录入工具 v0.2

数据基准:
  平台库存快照 (data/inbox/车辆库存_*.xlsx, 31 列)
  经销商主数据 (config/countries/<进口商>.json, 首次运行自动从快照生成)
  本地台账     (data/ledger/<进口商>_ledger.json, 记录已提交过的 VIN)

用法:
  python3 vehicle_flow.py init   --importer IMP001
  python3 vehicle_flow.py intake --vin-file data/inbox/vin.txt --date 2026-11-02
  python3 vehicle_flow.py intake --vin "VIN1,VIN2" --date 2026-11-02 --preview
  python3 vehicle_flow.py check  out/2026-11-02/入库_20261102.xlsx
"""
import argparse
import datetime as dt
import glob
import json
import os
import re
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
INBOX = os.path.join(BASE, 'data', 'inbox')
LEDGER_DIR = os.path.join(BASE, 'data', 'ledger')
CONFIG_DIR = os.path.join(BASE, 'config', 'countries')
OUT_DIR = os.path.join(BASE, 'out')

HEADER = ['VIN', '车型', '进口商编码', '进口商入库时间',
          '经销商编码', '经销商入库时间', '末销时间', '备注']
VIN_RE = re.compile(r'^[A-Z0-9]{17}$')
DATE_FMT = 'yyyy-mm-dd'

THIN = Side(style='thin', color='FF000000')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal='center', vertical='center')
WIDTHS = [22, 10, 14, 18, 14, 18, 14, 28]

STATE_LABEL = {'transit': '在途', 'importer': '进口商已入库',
               'dealer': '经销商已入库', 'sold': '已末销'}


# ---------------------------------------------------------------- 基础

def current_snapshot_file():
    return os.path.join(BASE, 'data', '.current_snapshot')


def find_snapshot(path=None):
    if path:
        if os.path.exists(path):
            return path
        # 传入路径不存在时仍交给后续, 让 load_snapshot 报明确错误
        return path
    # 1) 界面记录的「当前快照」
    cp = current_snapshot_file()
    if os.path.exists(cp):
        try:
            with open(cp, encoding='utf-8') as f:
                p = f.read().strip()
            if p and os.path.exists(p):
                return p
        except Exception:
            pass
    # 2) inbox 里最新的 xlsx / xls (不限文件名前缀)
    files = sorted(glob.glob(os.path.join(INBOX, '*.xlsx')) +
                   glob.glob(os.path.join(INBOX, '*.xls')), key=os.path.getmtime)
    if files:
        return files[-1]
    sys.exit('未找到库存快照：请先在网页「库存快照」处上传平台导出的车辆台账（.xlsx / .xls）')


def load_snapshot(path, data_only=True):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=data_only)
    ws = wb.active
    header = list(next(ws.iter_rows(max_row=1, values_only=True)))
    idx = {h: i for i, h in enumerate(header) if h}
    missing = [h for h in ('VIN', '车型', '进口商编码', '经销商编码',
                           '运营状态', '经销商入库时间', '末销时间', '在途') if h not in idx]
    if missing:
        sys.exit('快照缺列: %s' % missing)

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[idx['VIN']]:
            continue
        vin = str(r[idx['VIN']]).strip()
        rows.append({
            'vin': vin,
            'model': str(r[idx['车型']] or '').strip(),
            'imp': str(r[idx['进口商编码']] or '').strip(),
            'imp_name': str(r[idx['进口商名称']] or '').strip(),
            'dealer': str(r[idx['经销商编码']] or '').strip(),
            'dealer_name': str(r[idx['经销商名称']] or '').strip(),
            'status': str(r[idx['运营状态']] or '').strip(),
            'imp_in': r[idx['进口商入库时间']],
            'dealer_in': r[idx['经销商入库时间']],
            'sold_at': r[idx['末销时间']],
            'transit': str(r[idx['在途']] or '').strip(),
            'imp_done': str(r[idx['进口商已入库']] or '').strip(),
            'dealer_done': str(r[idx['经销商已入库']] or '').strip(),
            'sold_done': str(r[idx['已末销出库']] or '').strip(),
            'country': str(r[idx['国家/地区名称']] or '').strip(),
            'created': r[idx['创建时间']],
            'prod_date': r[idx['生产日期']] if '生产日期' in idx else None,
            'age': r[idx['车龄（月）']] if '车龄（月）' in idx else None,
        })
    wb.close()
    return rows


def state_of(r):
    if r['sold_done'] == '是' or r['sold_at']:
        return 'sold'
    if r['dealer_done'] == '是':
        return 'dealer'
    if r['imp_done'] == '是':
        return 'importer'
    if r['transit'] == '否':
        return 'dealer'
    return 'transit'


def parse_date(s):
    if isinstance(s, dt.datetime):
        return s.date()
    if isinstance(s, dt.date):
        return s
    s = str(s).strip()
    for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S'):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError('无法解析日期: %s' % s)


# ---------------------------------------------------------------- 主数据

def config_path(importer):
    return os.path.join(CONFIG_DIR, '%s.json' % importer)


def build_config(rows, importer):
    """从快照生成经销商主数据; 之后由人工维护(失效/新增/网点数)。"""
    sub = [r for r in rows if r['imp'] == importer]
    if not sub:
        sys.exit('快照里没有进口商 %s' % importer)

    country = Counter(r['country'] for r in sub if r['country']).most_common(1)[0][0]
    dealers = {}
    for r in sub:
        code = r['dealer']
        if not code:
            continue
        name = r['dealer_name']
        st = r['status']
        if code == importer:
            status = '失效'                      # 进口商自身不是经销网点
            note = '编码同进口商, 视为进口商自有库存'
        elif st == '正常运营':
            status = '正常运营'
            note = ''
        elif st == '终止运营':
            status = '失效'
            note = '快照标记终止运营'
        else:
            status = '未知'
            note = '运营状态为空, 待确认'
        old = dealers.get(code)
        if old and old['status'] == '正常运营':
            continue
        dealers[code] = {'code': code, 'name': name, 'status': status,
                         'outlets': 1, 'carflow_code': None,
                         'note': note, 'in_stock': 0}

    for r in sub:
        if r['dealer'] in dealers and state_of(r) == 'dealer':
            dealers[r['dealer']]['in_stock'] += 1

    cfg = {
        'importer': importer,
        'importer_name': Counter(r['imp_name'] for r in sub if r['imp_name']).most_common(1)[0][0],
        'country': country,
        'dealers': sorted(dealers.values(), key=lambda d: d['code']),
        'rule': {'weight': 'outlets', 'keep_months': 2.5, 'min_keep_units': 1},
    }
    return cfg


def load_config(importer, rows):
    p = config_path(importer)
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            cfg = json.load(f)
    else:
        cfg = build_config(rows, importer)
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(p, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        print('已生成主数据: %s（可手动补 outlets / carflow_code / 失效状态）' % p)
    return cfg


def load_models():
    """读 config/models.json, 返回 {VIN前缀: DMS车型码}。
    不存在则返回空 dict, 让程序只靠快照自动推断。
    """
    p = os.path.join(BASE, 'config', 'models.json')
    m = {}
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            data = json.load(f)
        for row in data.get('models', []):
            code = row.get('code')
            for pf in row.get('prefixes', []):
                m[pf.upper()] = code
    return m


def extract_project_name(path_or_str):
    """从 PL 文件名/项目号字符串提取项目标识。
    例: 'PL Attachment PR26AB-101XY002 100UNITS.xls' -> 'PR26AB-101XY002'
        'PR26AB-101XY002'                           -> 'PR26AB-101XY002'
    """
    s = os.path.basename(str(path_or_str)) if os.path.sep in str(path_or_str) else str(path_or_str)
    s = re.sub(r'\.[A-Za-z0-9]+$', '', s)          # 去掉扩展名
    m = re.search(r'[A-Z]{2}\d{2}[A-Z]{2}[-\dA-Z]{4,}', s, re.I)   # 项目号: 2字母+2数字+2字母+连字符段
    if m:
        return m.group(0).upper()
    m = re.search(r'[A-Z]{2}[A-Z]{0,2}[-\dA-Z]{6,}', s, re.I)
    if m:
        return m.group(0).upper()
    # 最后兜底: 整串去掉 PL Attachment / 数量 等词
    s = re.sub(r'PL\s*Attachment\s*', '', s, flags=re.I)
    s = re.sub(r'[_\s]+[0-9]+\s*(UNIT|UNITS|台)?$', '', s, flags=re.I).strip()
    return s[:30] if s else ''


def active_dealers(cfg):
    return [d for d in cfg['dealers'] if d['status'] == '正常运营']


# ---------------------------------------------------------------- 台账

def ledger_path(importer):
    return os.path.join(LEDGER_DIR, '%s_ledger.json' % importer)


def load_ledger(importer):
    p = ledger_path(importer)
    if not os.path.exists(p):
        return {'importer': importer, 'records': []}
    with open(p, encoding='utf-8') as f:
        return json.load(f)


def save_ledger(importer, ledger):
    os.makedirs(LEDGER_DIR, exist_ok=True)
    with open(ledger_path(importer), 'w', encoding='utf-8') as f:
        json.dump(ledger, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------- 输入解析

def parse_vin_input(text):
    """支持 VIN / VIN,车型 / 一行多个 VIN; 逗号、空格、换行均可分隔。

    规则: 命中 17 位 VIN 视为新记录, 紧随其后且不是 VIN 的词作为车型。
    """
    items = []
    for line in text.replace('\r', '\n').split('\n'):
        line = line.strip().replace(',', ' ').replace('，', ' ').replace('\t', ' ')
        for tok in [t for t in line.split(' ') if t]:
            tok = tok.upper()
            if VIN_RE.match(tok) or not items or items[-1][1]:
                items.append([tok, ''])
            elif len(tok) <= 10:
                items[-1][1] = tok
            else:
                items.append([tok, ''])
    return [(v, m) for v, m in items]


def _sheet_rows(path):
    """统一读出表格式文件的所有单元格 (支持 .xls / .xlsx)。"""
    if path.lower().endswith('.xls'):
        import xlrd
        wb = xlrd.open_workbook(path)
        sh = wb.sheet_by_index(0)
        return [[sh.cell_value(i, j) for j in range(sh.ncols)] for i in range(sh.nrows)]
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def read_vin_file(path, known_models=(), prefix_map=None):
    """读取 VIN 清单。支持 txt/csv/xlsx/xls。

    表格式文件自动找含 VIN 的表头行; 车型按 车型列 -> 商品名称 -> VIN 前缀 依次推断。
    """
    prefix_map = prefix_map or {}
    if path.lower().endswith(('.xlsx', '.xlsm', '.xls')):
        rows = _sheet_rows(path)

        hi, vcol, mcol = None, 0, None
        for i, row in enumerate(rows[:10]):
            cells = [str(c).strip().upper() if c is not None else '' for c in row]
            for key in ('VIN', 'VIN#', 'VIN.#', 'VIN编码'):
                if key in cells:
                    hi, vcol = i, cells.index(key)
                    break
            if hi is not None:
                raw_cells = [str(c).strip() if c is not None else '' for c in row]
                for key in ('车型', '商品名称', 'Description'):
                    if key in raw_cells:
                        mcol = raw_cells.index(key)
                        break
                break

        if hi is None:                      # 无表头: 按第一列 VIN、第二列车型
            out = []
            for row in rows:
                if not row or not row[0]:
                    continue
                v = str(row[0]).strip().upper()
                if not VIN_RE.match(v):
                    continue
                m = str(row[1]).strip().upper() if len(row) > 1 and row[1] else ''
                out.append((v, m))
            return out

        out = []
        for row in rows[hi + 1:]:
            if not row or vcol >= len(row) or not row[vcol]:
                continue
            v = str(row[vcol]).strip().upper()
            if not VIN_RE.match(v):
                continue
            m = ''
            if mcol is not None and mcol < len(row) and row[mcol]:
                raw = str(row[mcol]).strip().upper()
                if raw in known_models:
                    m = raw
                else:
                    for km in sorted(known_models, key=len, reverse=True):
                        if raw.startswith(km.upper()):
                            m = km
                            break
            if not m:                       # 商品名称不带车型时, 按 VIN 前缀匹配
                m = prefix_map.get(v[:6]) or prefix_map.get(v[:5]) or ''
            out.append((v, m))
        return out
    with open(path, encoding='utf-8') as f:
        return parse_vin_input(f.read())


# ---------------------------------------------------------------- 分配

def allocate_intake(batch, cfg, stock):
    """按 网点权重 + 车型结构 逐台落位。

    batch: [(vin, model)] ; stock: {dealer: Counter(model)}
    规则: ① 本批按网点权重均分, 不因历史库存多少而集中; ② 各店车型结构贴近全局;
          ③ 已压该车型的店少分。
    """
    dealers = active_dealers(cfg)
    if not dealers:
        sys.exit('没有状态为「正常运营」的经销商')

    weight = {d['code']: max(1, int(d.get('outlets') or 1)) for d in dealers}
    if cfg['rule'].get('weight') == 'equal':
        weight = {k: 1 for k in weight}
    wsum = sum(weight.values())

    cur = {d: Counter(stock.get(d, Counter())) for d in weight}
    cur_total = {d: sum(cur[d].values()) for d in weight}

    global_share = Counter()
    for d in weight:
        global_share.update(cur[d])
    n_all = sum(global_share.values()) or 1
    share = {m: global_share[m] / n_all for m in global_share}

    n_batch = len(batch)
    target_batch = {d: n_batch * weight[d] / wsum for d in weight}

    assigned = {d: Counter() for d in weight}
    result = []
    for vin, model in sorted(batch, key=lambda x: (x[1], x[0])):
        best, best_score = None, None
        gm = share.get(model, 0)
        for d in sorted(weight):
            load = (sum(assigned[d].values()) + 1) / (target_batch[d] + 1)
            base = cur_total[d] + sum(assigned[d].values())
            after = (cur[d][model] + assigned[d][model] + 1) / (base + 1)
            model_dev = max(0.0, after - gm) / (gm + 0.4)
            score = load + 0.5 * model_dev
            if best_score is None or score < best_score:
                best, best_score = d, score
        assigned[best][model] += 1
        result.append((vin, model, best))
    return result


def stock_by_dealer(rows, cfg):
    codes = {d['code'] for d in cfg['dealers']}
    stock = defaultdict(Counter)
    for r in rows:
        if r['imp'] != cfg['importer']:
            continue
        if r['dealer'] in codes and state_of(r) == 'dealer':
            stock[r['dealer']][r['model']] += 1
    return stock


# ---------------------------------------------------------------- 校验

def validate(records, cfg, rows, ledger):
    errors, warns = [], []
    by_vin = {r['vin']: r for r in rows}
    active = {d['code'] for d in active_dealers(cfg)}
    known_models = {r['model'] for r in rows if r['model']}
    ledger_vins = {x['vin'] for x in ledger['records']}

    seen = set()
    for rec in records:
        vin = rec['vin']
        if not VIN_RE.match(vin):
            errors.append('%s: VIN 不是 17 位字母数字' % vin)
        if vin in seen:
            errors.append('%s: 批次内重复' % vin)
        seen.add(vin)
        if vin in ledger_vins:
            errors.append('%s: 台账已提交过（%s）' % (vin, next(x['dealer'] for x in ledger['records'] if x['vin'] == vin)))
        if rec['model'] not in known_models:
            errors.append('%s: 车型 %s 不在品牌车型表内' % (vin, rec['model']))
        if rec['dealer'] not in active:
            errors.append('%s: 经销商 %s 非正常运营' % (vin, rec['dealer']))
        if rec['dealer_in'] < rec['imp_in']:
            errors.append('%s: 经销商入库时间早于进口商入库时间' % vin)

        src = by_vin.get(vin)
        if src:
            st = state_of(src)
            if st != 'transit':
                errors.append('%s: 平台状态已是「%s」，不能重新入库' % (vin, STATE_LABEL[st]))
            if src['model'] and src['model'] != rec['model']:
                errors.append('%s: 车型与快照不一致（%s vs %s）' % (vin, rec['model'], src['model']))
        else:
            warns.append('%s: 快照中未找到，按提供车型 %s 处理' % (vin, rec['model']))
    return errors, warns


# ---------------------------------------------------------------- 输出

def write_template(path, records):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '模板'
    for j, h in enumerate(HEADER, start=1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = Font(name='Calibri', size=12)
        c.alignment = CENTER
        c.border = BORDER
        ws.column_dimensions[get_column_letter(j)].width = WIDTHS[j - 1]

    for i, rec in enumerate(records, start=2):
        sold = rec.get('sold_at')
        sold_s = sold.isoformat() if hasattr(sold, 'isoformat') else (sold or None)
        remark = rec.get('remark') or ''
        vals = [rec['vin'], rec['model'], rec['imp'], rec['imp_in'].isoformat(),
                rec['dealer'], rec['dealer_in'].isoformat(), sold_s, remark]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.border, c.alignment = BORDER, CENTER
            c.font = Font(name='Calibri', size=10)
            if j == 1:
                c.number_format = '@'
            if j in (4, 6, 7, 8):
                c.number_format = '@'
    ws.freeze_panes = 'A2'
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def check_file(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    header = [c.value for c in ws[1][:len(HEADER)]]
    errors = []
    if header != HEADER:
        errors.append('表头不符: %s' % header)
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]
    seen = set()
    for i, r in enumerate(rows, start=2):
        vin = str(r[0])
        if not VIN_RE.match(vin):
            errors.append('第%d行 VIN 非法: %s' % (i, vin))
        if vin in seen:
            errors.append('第%d行 VIN 重复: %s' % (i, vin))
        seen.add(vin)
        if not r[1] or not r[2] or not r[3] or not r[4] or not r[5]:
            errors.append('第%d行必填字段缺失' % i)
        try:
            di, dd = parse_date(r[3]), parse_date(r[5])
            if di > dd:
                errors.append('第%d行 进口商入库时间晚于经销商入库时间' % i)
            if len(r) > 6 and r[6]:
                if parse_date(r[6]) < dd:
                    errors.append('第%d行 末销时间早于经销商入库时间' % i)
        except Exception as e:
            errors.append('第%d行日期解析失败: %s' % (i, e))
    wb.close()
    print('文件: %s' % path)
    print('数据行: %d, 唯一 VIN: %d' % (len(rows), len(seen)))
    if errors:
        print('校验不通过 %d 项:' % len(errors))
        for e in errors[:20]:
            print('  -', e)
        return 1
    print('校验通过：表头/必填/日期顺序/VIN 唯一 均正常')
    return 0


# ---------------------------------------------------------------- 主流程

def cmd_init(args):
    rows = load_snapshot(find_snapshot(args.snapshot))
    imports = Counter(r['imp'] for r in rows if r['imp'])
    print('快照 %d 行, 进口商: %s' % (len(rows), dict(imports)))
    for imp in ([args.importer] if args.importer else imports):
        cfg = build_config(rows, imp)
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(config_path(imp), 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        print('\n[%s] %s | %s' % (imp, cfg['country'], cfg['importer_name']))
        for d in cfg['dealers']:
            print('  %-10s %-4s 网点%-2s 在库%-3s %s' % (
                d['code'], d['status'], d['outlets'], d['in_stock'], d['name']))
        print('  主数据 -> %s' % config_path(imp))


def cmd_intake(args):
    snapshot = find_snapshot(args.snapshot)
    rows = load_snapshot(snapshot)

    known_models = {r['model'] for r in rows if r['model']}
    # 1) 从快照统计 VIN 前缀 -> 车型 (最贴近真实数据)
    pm = defaultdict(Counter)
    for r in rows:
        if r['model']:
            pm[r['vin'][:6]][r['model']] += 1
    prefix_map = {p: c.most_common(1)[0][0] for p, c in pm.items()}
    # 2) 合并 config/models.json 里的官方车型表 (有则优先, 无则快照推断)
    for p, code in load_models().items():
        prefix_map.setdefault(p, code)

    # 从文件名提取项目号 (PR26AB-101XY002 这种), 统一写到备注列
    project = ''
    if args.vin_file:
        vfiles = args.vin_file if isinstance(args.vin_file, list) else [args.vin_file]
        batch, parts = [], []
        for f in vfiles:
            if not project:
                project = extract_project_name(f)
            part = read_vin_file(f, known_models, prefix_map)
            print('  读取 %s: %d 台' % (os.path.basename(f), len(part)))
            batch.extend(part)
            parts.append((f, part))
    elif args.vin:
        batch = parse_vin_input(args.vin)
    else:
        sys.exit('需要 --vin 或 --vin-file')

    batch = [(v, m) for v, m in batch if v]
    if not batch:
        sys.exit('没有解析到 VIN')

    by_vin = {r['vin']: r for r in rows}
    importer = args.importer
    if not importer:
        importer = Counter(by_vin[v]['imp'] for v, _ in batch if v in by_vin).most_common(1)
        importer = importer[0][0] if importer else ''
    if not importer:
        # 入库的都是新 VIN, 快照里本来就不会有; 只有一个进口商时直接用它
        imps = sorted({r['imp'] for r in rows if r['imp']})
        importer = imps[0] if len(imps) == 1 else ''
    if not importer:
        sys.exit('这批是新车(VIN 不在快照里属正常)，请指定进口商: '
                 '网页里选「进口商」, 命令行用 --importer')

    cfg = load_config(importer, rows)
    dealers = active_dealers(cfg)
    print('进口商 %s | %s | 在营经销商 %d 家: %s' % (
        importer, cfg['country'], len(dealers),
        ', '.join(d['code'] for d in dealers)))

    # 补齐车型: 快照有则用快照, 否则用输入里带的
    fixed = []
    unknown = []
    for vin, model in batch:
        src = by_vin.get(vin)
        if src and src['model']:
            fixed.append((vin, src['model']))
        elif model:
            fixed.append((vin, model))
        else:
            unknown.append(vin)
    if unknown:
        print('以下 VIN 在快照中找不到且未提供车型，已跳过: %s' % ', '.join(unknown))

    # 去重
    seen, dedup = set(), []
    for v, m in fixed:
        if v in seen:
            print('批次内重复已忽略: %s' % v)
            continue
        seen.add(v)
        dedup.append((v, m))

    stock = stock_by_dealer(rows, cfg)
    result = allocate_intake(dedup, cfg, stock)

    dealer_in = parse_date(args.date) if args.date else dt.date.today()
    imp_in = parse_date(args.importer_date) if args.importer_date else dealer_in
    if imp_in > dealer_in:
        sys.exit('进口商入库时间不能晚于经销商入库时间')

    records = [{'vin': v, 'model': m, 'imp': importer, 'imp_in': imp_in,
                'dealer': d, 'dealer_in': dealer_in, 'sold_at': None,
                'remark': project}
               for v, m, d in result]

    ledger = load_ledger(importer)
    errors, warns = validate(records, cfg, rows, ledger)
    if warns:
        print('\n提示:')
        for w in warns[:5]:
            print('  -', w)
        if len(warns) > 5:
            print('  ... 同类提示共 %d 条' % len(warns))
    if errors:
        print('\n校验不通过，未生成文件:')
        for e in errors:
            print('  -', e)
        return 1

    print('\n分配结果 (入库日期 %s):' % dealer_in.isoformat())
    per_dealer = defaultdict(Counter)
    for rec in records:
        per_dealer[rec['dealer']][rec['model']] += 1
    for d in sorted(per_dealer):
        c = per_dealer[d]
        print('  %-10s %3d 台  %s' % (d, sum(c.values()),
                                   ' '.join('%s×%d' % (m, c[m]) for m in sorted(c))))
    model_total = Counter(rec['model'] for rec in records)
    print('  合计 %d 台, 车型: %s' % (len(records),
                                 ', '.join('%s×%d' % (m, model_total[m]) for m in sorted(model_total))))

    if args.preview:
        print('\n(预览模式，未写文件和台账)')
        return 0

    tag = args.tag or project or dealer_in.strftime('%Y%m%d')
    out_path = args.out or os.path.join(
        OUT_DIR, dealer_in.isoformat(),
        '%s_入库_%s.xlsx' % (cfg['country'], tag))
    write_template(out_path, records)

    ledger['records'].extend(
        {'vin': r['vin'], 'model': r['model'], 'dealer': r['dealer'],
         'importer_in': r['imp_in'].isoformat(), 'dealer_in': r['dealer_in'].isoformat(),
         'batch': tag, 'file': os.path.relpath(out_path, BASE)} for r in records)
    save_ledger(importer, ledger)

    print('\n已生成: %s' % out_path)
    print('台账已追加 %d 条 -> %s' % (len(records), ledger_path(importer)))
    return check_file(out_path)


def cmd_sale(args):
    """月度末销: 结合 carflow 目标与存销比给建议量, 再分店 FIFO 出库。

    三数对比: carflow 目标 T / 系统建议 S* = Q0 + I - r* × A / 你报的量 U
    """
    import dashboard as D                      # 局部导入避免循环
    snapshot = find_snapshot(args.snapshot)
    rows = load_snapshot(snapshot)
    if not args.importer:
        sys.exit('需要 --importer')
    cfg = load_config(args.importer, rows)
    active = {d['code'] for d in active_dealers(cfg)}
    floor = int(cfg['rule'].get('min_keep_units', 1))
    targets = D.load_targets()
    t = D.target_for(targets, cfg['country'])
    r_star = float(t.get('target_ratio', 2.0))
    band = float(t.get('band', 0.3))
    n = int(targets.get('window_months', 3))

    # 统计月: 默认取该进口商末销数据的最新月
    mc = Counter()
    for r in rows:
        if r['imp'] == args.importer and state_of(r) == 'sold' and r['sold_at']:
            mc[D.month_key(r['sold_at'])] += 1
    month = args.month or (max(mc) if mc else dt.date.today().strftime('%Y-%m'))
    windows = [D.add_months(month, -i) for i in range(n, 0, -1)]

    # 月均 A: 近 n 月, 首笔末销之前的零销月不计入
    first = min((m for m in mc if mc[m] > 0 and m <= windows[-1]), default=None)
    used = [m for m in windows if first and m >= first]
    avg = (sum(mc[m] for m in used) / len(used)) if used and sum(mc[m] for m in used) else 0.0

    # 库存
    in_transit = sum(1 for r in rows
                     if r['imp'] == args.importer and state_of(r) == 'transit')
    in_stock = sum(1 for r in rows
                   if r['imp'] == args.importer and state_of(r) in ('dealer', 'importer'))
    stock = defaultdict(list)
    for r in rows:
        if r['imp'] == args.importer and r['dealer'] in active and state_of(r) == 'dealer':
            stock[r['dealer']].append(r)
    feasible = sum(len(v) for v in stock.values()) - floor * len(stock)

    # carflow 当月目标
    flow = D.load_carflow(args.carflow or os.path.join(INBOX, '月度末销目标.xlsx'))
    target_t = flow.get(D.norm_country(cfg['country']), {}).get(month, 0)

    # 建议量 S* = Q0 + I - r* × A
    s_star = max(0, round(in_stock + in_transit - r_star * avg)) if avg else None

    print('进口商 %s | %s | 统计月 %s' % (args.importer, cfg['country'], month))
    print('  库存: 在库 %d + 在途 %d | 月均 A=%.1f（窗口 %s）| 目标存销比 %.1f±%.1f'
          % (in_stock, in_transit, avg, '/'.join(windows), r_star, band))
    zero_m = [m for m in used if mc[m] == 0]
    if zero_m and avg:
        print('  注意: 窗口内 %s 末销为 0（疑导出延迟），A 被低估、建议量与推演仅供参考'
              % '、'.join(zero_m))
    print('  三数: carflow 目标 %s | 系统建议 %s | 可执行上限 %d'
          % (target_t or '无', s_star if s_star is not None else '—', max(0, feasible)))
    if not feasible or feasible <= 0:
        sys.exit('没有可末销的在库车辆（每店需保留 %d 台）' % floor)

    want = None
    if args.count and str(args.count).strip():
        want = int(args.count)
    elif args.source == 'carflow' and target_t:
        want = int(target_t)
    elif s_star is not None:
        want = s_star
    if not want or want <= 0:
        sys.exit('请用 --count 报量, 或 --source carflow 按目标, 否则按系统建议量(当前为 0)')
    if want > feasible:
        print('  报量 %d 超出可执行上限, 已封顶为 %d 台' % (want, feasible))
        want = feasible

    # 分店: 按网点权重水位法, 每店留 floor 台
    weight = {d: max(1, int(next((x.get('outlets') for x in cfg['dealers']
                                  if x['code'] == d), 1) or 1)) for d in stock}
    avail = {d: max(0, len(stock[d]) - floor) for d in stock}
    quota = {d: 0 for d in stock}
    left = want
    while left > 0:
        live = [d for d in stock if quota[d] < avail[d]]
        if not live:
            break
        ws = sum(weight[d] for d in live)
        for d in live:
            give = min(avail[d] - quota[d], max(1, round(left * weight[d] / ws)))
            if d == live[-1]:
                give = min(avail[d] - quota[d], left)
            quota[d] += give
            left -= give

    # 末销日期均匀铺到当月工作日 (周末除外)
    import calendar
    sale_month_date = parse_date(month + '-01')
    last_day = calendar.monthrange(sale_month_date.year, sale_month_date.month)[1]
    workdays = [dt.date(sale_month_date.year, sale_month_date.month, d)
                for d in range(1, last_day + 1)
                if dt.date(sale_month_date.year, sale_month_date.month, d).weekday() < 5]
    # 如果统计月是已结束月份, 只用 ≤ 今天的日子
    if sale_month_date.month != dt.date.today().month or sale_month_date.year != dt.date.today().year:
        workdays = [d for d in workdays if d <= dt.date.today()] or workdays
    if args.date:
        workdays = [parse_date(args.date)]
    if not workdays:
        workdays = [sale_month_date]

    # 均匀分配: want 台平均铺到 workdays, 每一天分到的台数差 <= 1, 同店不同日
    # 先按经销商排序后轮转, 再把每台车分配到日槽
    day_count = [0] * len(workdays)
    picked_sorted = []
    for d in sorted(quota):
        cand = sorted(stock[d], key=lambda r: (str(r['prod_date'] or r['created'] or ''), r['vin']))
        cutoff = args.date and parse_date(args.date) or None
        if cutoff:
            cand = [r for r in cand if parse_date(r['dealer_in']) <= cutoff]
        picked_sorted.extend(cand[:quota[d]])
    # 均匀分日: 逐台选当日台数最少的那个工作日
    per_day = {}
    for i, r in enumerate(picked_sorted):
        # 同店尽量分散在不同天: day_count 按 (天数, 同店当日已分配数) 综合选
        dealer_today = [sum(1 for x in per_day.values() if x[0] == d_idx and x[1] == r['dealer'])
                        for d_idx in range(len(workdays))]
        best_idx = min(range(len(workdays)),
                       key=lambda j: (day_count[j], dealer_today[j]))
        per_day[id(r)] = (best_idx, r['dealer'])
        day_count[best_idx] += 1

    records = []
    for r in picked_sorted:
        d_idx = per_day[id(r)][0]
        sd = workdays[d_idx]
        records.append({
            'vin': r['vin'], 'model': r['model'], 'imp': args.importer,
            'imp_in': parse_date(r['imp_in']) if r['imp_in'] else parse_date(r['dealer_in']),
            'dealer': r['dealer'], 'dealer_in': parse_date(r['dealer_in']),
            'sold_at': sd,
        })

    # 输出
    used_days = sorted({workdays[x[0]] for x in per_day.values()})
    print('末销 %s 共 %d 台, 铺到 %d 个工作日 (日均 %.1f):' % (
        month, len(records), len(used_days), len(records) / max(1, len(used_days))))
    for day in used_days:
        cnt = sum(1 for x in per_day.values() if workdays[x[0]] == day)
        print('  %s  %d 台' % (day.isoformat(), cnt))
    per = defaultdict(Counter)
    for rec in records:
        per[rec['dealer']][rec['model']] += 1
    for d in sorted(per):
        c = per[d]
        print('  %-10s %3d 台  %s' % (d, sum(c.values()),
                                   ' '.join('%s×%d' % (m, c[m]) for m in sorted(c))))
    if avg:
        after = (in_stock + in_transit - len(records)) / avg
        print('  推演: 卖完后存销比 %.2f（目标 %.1f±%.1f）' % (after, r_star, band))

    if args.preview:
        print('\n(预览模式，未写文件)')
        return 0

    tag = args.tag or month.replace('-', '')
    out_path = args.out or os.path.join(
        OUT_DIR, month, '%s_末销_%s.xlsx' % (cfg['country'], tag))
    write_template(out_path, records)

    p = os.path.join(LEDGER_DIR, '%s_sales.json' % args.importer)
    data = {'importer': args.importer, 'records': []}
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            data = json.load(f)
    data['records'].extend({'vin': r['vin'], 'dealer': r['dealer'],
                            'sold_at': r['sold_at'].isoformat(), 'batch': tag} for r in records)
    os.makedirs(LEDGER_DIR, exist_ok=True)
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print('\n已生成: %s' % out_path)
    return check_file(out_path)


# ---------------------------------------------------------------- 真实末销表识别

SALE_COLS = {
    'vin': ('vin', '车架号', 'vin码', 'vin编码'),
    'model': ('车型', '商品名称', '车型代号', '车型编码'),
    'dealer': ('经销商编码', '经销商代码', '经销商', 'dms编码', '门店'),
    'sold_at': ('末销时间', '末销日期', '出库日期', '销售日期'),
    'qty': ('数量', '台数', 'qty'),
}


def _pick_col(header, keys):
    cells = [str(c or '').strip().lower().replace(' ', '') for c in header]
    for i, s in enumerate(cells):
        for k in keys:
            if k in s:
                return i
    return None


def read_sale_file(path):
    """读出真实末销表, 返回 (items, notes)。

    items: [{'vin','model','dealer','sold_at','qty','row'}]
    表头自适应: VIN/车架号, 车型/商品名称, 经销商/门店, 末销时间/出库日期, 数量
    """
    if path.lower().endswith('.csv'):
        import csv
        with open(path, encoding='utf-8-sig', newline='') as f:
            rows = [list(r) for r in csv.reader(f)]
    else:
        rows = _sheet_rows(path)
    hi, cols = None, {}
    for i, row in enumerate(rows[:15]):
        c = {k: _pick_col(row, v) for k, v in SALE_COLS.items()}
        if c['vin'] is not None or (c['dealer'] is not None and c['qty'] is not None):
            hi, cols = i, c
            break
    if hi is None:
        return [], ['未识别到表头（需要含 VIN 列，或 经销商 + 数量 列）']

    items, notes = [], []
    for n, row in enumerate(rows[hi + 1:], start=hi + 2):
        if not row or all(not x for x in row):
            continue

        def cell(key):
            j = cols.get(key)
            return row[j] if j is not None and j < len(row) else None

        vin = str(cell('vin') or '').strip().upper()
        model = str(cell('model') or '').strip().upper()
        dealer = str(cell('dealer') or '').strip()
        sold_at = cell('sold_at')
        qty = 1
        try:
            if cell('qty'):
                qty = max(1, int(float(cell('qty'))))
        except (TypeError, ValueError):
            qty = 1

        if vin and not VIN_RE.match(vin):
            notes.append('第%d行 VIN 格式不合法，已跳过: %s' % (n, vin))
            continue
        if not vin and not (dealer and qty):
            continue
        items.append({'vin': vin, 'model': model, 'dealer': dealer,
                      'sold_at': sold_at, 'qty': qty, 'row': n})
    return items, notes


def cmd_sale_file(args):
    """按真实末销表识别出库: 逐条核对库存状态后输出标准导入文件。"""
    snapshot = find_snapshot(args.snapshot)
    rows = load_snapshot(snapshot)
    items, notes = read_sale_file(args.file)
    if not items:
        sys.exit('末销表里没有可识别记录%s'
                 % ('（%s）' % '；'.join(notes) if notes else ''))

    by_vin = {r['vin']: r for r in rows}
    stock = defaultdict(list)
    for r in rows:
        if state_of(r) == 'dealer':
            stock[r['dealer']].append(r)

    used = set()
    recs = []
    errs = []
    default_day = parse_date(args.date) if args.date else dt.date.today()

    # 第一轮: 带 VIN 的直接核对快照
    for it in items:
        vin = it['vin']
        if not vin:
            continue
        r = by_vin.get(vin)
        if not r:
            errs.append('第%d行 VIN 不在库存快照: %s' % (it['row'], vin))
            continue
        st = state_of(r)
        if st not in ('dealer', 'importer'):
            errs.append('第%d行 VIN 当前状态「%s」，不能末销: %s'
                        % (it['row'], STATE_LABEL[st], vin))
            continue
        if vin in used:
            errs.append('第%d行 VIN 重复: %s' % (it['row'], vin))
            continue
        used.add(vin)
        sold = parse_date(it['sold_at']) if it['sold_at'] else default_day
        recs.append({
            'vin': vin, 'model': r['model'], 'imp': r['imp'],
            'imp_in': parse_date(r['imp_in']) if r['imp_in'] else parse_date(r['dealer_in']),
            'dealer': r['dealer'], 'dealer_in': parse_date(r['dealer_in']),
            'sold_at': sold, 'remark': it.get('remark') or '',
        })

    # 第二轮: 只给经销商+数量的, 从未占用在库车里 FIFO 取
    for it in items:
        if it['vin']:
            continue
        dealer, qty = it['dealer'], it['qty']
        cand = [x for x in stock.get(dealer, []) if x['vin'] not in used]
        if it['model']:
            hit = [x for x in cand if x['model'] == it['model']]
            cand = hit or cand
        cand.sort(key=lambda r: (str(r['prod_date'] or r['created'] or ''), r['vin']))
        if len(cand) < qty:
            errs.append('第%d行 %s 库存不足: 需 %d 台，可用 %d 台'
                        % (it['row'], dealer, qty, len(cand)))
        take = cand[:qty]
        sold = parse_date(it['sold_at']) if it['sold_at'] else default_day
        for r in take:
            used.add(r['vin'])
            recs.append({
                'vin': r['vin'], 'model': r['model'], 'imp': r['imp'],
                'imp_in': parse_date(r['imp_in']) if r['imp_in'] else parse_date(r['dealer_in']),
                'dealer': r['dealer'], 'dealer_in': parse_date(r['dealer_in']),
                'sold_at': sold, 'remark': it.get('remark') or '',
            })

    for n in notes:
        print('  提示:', n)
    print('末销表 %s: 识别 %d 条, 成单 %d 台' % (os.path.basename(args.file),
                                            len(items), len(recs)))
    per = defaultdict(Counter)
    for rec in recs:
        per[rec['dealer']][rec['model']] += 1
    for d in sorted(per):
        print('  %-10s %3d 台  %s' % (d, sum(per[d].values()),
                                  ' '.join('%s×%d' % (m, per[d][m]) for m in sorted(per[d]))))
    if not recs:
        sys.exit('没有可输出的记录')
    if errs:
        print('  有 %d 条未能成单:' % len(errs))
        for e in errs[:20]:
            print('   -', e)

    if args.preview:
        print('\n(预览模式，未写文件)')
        return 0

    month = (parse_date(min(r['sold_at'] for r in recs)).strftime('%Y-%m'))
    tag = args.tag or month.replace('-', '')
    imp = Counter(r['imp'] for r in recs).most_common(1)[0][0]
    cfg = load_config(imp, rows)
    out_path = args.out or os.path.join(
        OUT_DIR, month, '%s_末销识别_%s.xlsx' % (cfg['country'], tag))
    write_template(out_path, recs)

    p = os.path.join(LEDGER_DIR, '%s_sales.json' % imp)
    data = {'importer': imp, 'records': []}
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            data = json.load(f)
    data['records'].extend({'vin': r['vin'], 'dealer': r['dealer'],
                            'sold_at': r['sold_at'].isoformat(), 'batch': tag}
                           for r in recs)
    os.makedirs(LEDGER_DIR, exist_ok=True)
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print('\n已生成: %s' % out_path)
    return check_file(out_path)



def main():
    p = argparse.ArgumentParser(description='车辆入库/末销录入工具')
    p.add_argument('--snapshot', help='库存快照路径 (默认取 data/inbox 最新)')
    sub = p.add_subparsers(dest='cmd', required=True)

    q = sub.add_parser('init', help='生成/查看经销商主数据')
    q.add_argument('--importer', help='进口商编码, 缺省为全部')
    q.set_defaults(func=cmd_init)

    q = sub.add_parser('intake', help='VIN 入库分配并输出录入文件')
    q.add_argument('--vin', help='VIN 文本, 逗号/空格/换行分隔')
    q.add_argument('--vin-file', help='VIN 文件 (txt/csv/xlsx)')
    q.add_argument('--importer', help='进口商编码, 默认按 VIN 自动识别')
    q.add_argument('--date', help='经销商入库日期, 默认今天')
    q.add_argument('--importer-date', help='进口商入库日期, 默认同经销商入库日期')
    q.add_argument('--tag', help='批次标识, 默认日期')
    q.add_argument('--out', help='输出路径')
    q.add_argument('--preview', action='store_true', help='只预览不写文件')
    q.set_defaults(func=cmd_intake)

    q = sub.add_parser('sale', help='月度末销: 结合 carflow 目标给建议并出库')
    q.add_argument('--importer', required=True)
    q.add_argument('--count', help='你报的台数(优先); 不报则按 --source 或系统建议')
    q.add_argument('--month', help='统计月, 如 2026-11, 默认取末销数据最新月')
    q.add_argument('--source', choices=['carflow'], help='按 carflow 当月末销目标')
    q.add_argument('--carflow', help='carflow 文件路径, 默认 data/inbox 下自动找')
    q.add_argument('--date', help='末销日期, 默认统计月月末')
    q.add_argument('--tag', help='批次标识, 默认日期')
    q.add_argument('--out', help='输出路径')
    q.add_argument('--preview', action='store_true')
    q.set_defaults(func=cmd_sale)

    q = sub.add_parser('sale-file', help='按真实末销表识别出库, 输出标准导入文件')
    q.add_argument('--file', required=True, help='真实末销表 (xlsx/xls/csv)')
    q.add_argument('--snapshot')
    q.add_argument('--date', help='表里没有末销日期时用的默认日期, 默认今天')
    q.add_argument('--tag', help='批次标识, 默认月份')
    q.add_argument('--out', help='输出路径')
    q.add_argument('--preview', action='store_true')
    q.set_defaults(func=cmd_sale_file)

    q = sub.add_parser('dashboard', help='区域末销/存销比看板 (HTML)')
    q.add_argument('--snapshot')
    q.add_argument('--carflow')
    q.add_argument('--month', help='统计月, 如 2026-11')
    q.add_argument('--out')
    q.set_defaults(func=lambda a: __import__('dashboard').cmd_dashboard(a))

    q = sub.add_parser('check', help='校验生成的录入文件')
    q.add_argument('file')
    q.set_defaults(func=lambda a: check_file(a.file))

    args = p.parse_args()
    sys.exit(args.func(args) or 0)


if __name__ == '__main__':
    main()
