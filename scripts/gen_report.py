# -*- coding: utf-8 -*-
"""按 douyin-product-report skill 口径生成双期报告 Excel。"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsxio import read_sheets, num
import xlsxwrite as X
from xlsxwrite import BANNER, HEADER, LABEL, NUM0, NUM2, PCT, CELL, NOTE

_HERE = os.path.dirname(os.path.abspath(__file__))
# 默认取 skill 自带的 assets；可用 SKILL_ASSETS 覆盖
SKILL = os.environ.get('SKILL_ASSETS') or os.path.normpath(os.path.join(_HERE, '..', 'assets'))
# 源文件由 run_report.py 注入，或使用 SRC_PREV / SRC_CURR 环境变量
SRC = {'prev': os.environ.get('SRC_PREV', ''), 'curr': os.environ.get('SRC_CURR', '')}
CATS = ['主板', '显卡', '显示器', '键盘', '鼠标', '耳机', '特价', '机箱', '电源', '散热', '周边']
MONITOR_SUB = {'XG显示器', 'VG显示器', 'PG显示器', 'PA显示器'}
OUT = os.environ.get('OUT_XLSX') or os.path.join(os.getcwd(), '抖音双期报告.xlsx')


def load_mapping():
    path = os.environ.get('MAP_PATH') or os.path.join(SKILL, 'category-mapping.xlsx')
    rows = read_sheets(path)['Sheet1']
    m = {}
    for r in rows[1:]:
        if len(r) < 3 or not r[0]:
            continue
        code = str(r[0]).strip()
        if code.endswith('.0'):
            code = code[:-2]
        cat = str(r[1]).strip()
        if cat in MONITOR_SUB:
            cat = '显示器'
        m[code] = (cat, str(r[2]).strip())
    return m


def load_source(path):
    rows = read_sheets(path)['全部']
    col = {str(h).strip(): i for i, h in enumerate(rows[0]) if h}
    out = {}
    for r in rows[1:]:
        if not r or len(r) <= col['商品编码'] or not r[col['商品编码']]:
            continue
        get = lambda k: r[col[k]] if col.get(k) is not None and col[k] < len(r) else None
        code = str(get('商品编码')).strip()
        if code.endswith('.0'):
            code = code[:-2]
        out[code] = {
            'name': get('商品名称'),
            'period': str(get('统计周期')).strip(),
            'gmv': num(get('成交金额')) or 0.0,
            'orders': num(get('成交订单数')) or 0.0,
            'refund_amt': num(get('退款金额（支付时间）')) or 0.0,
            'refund_orders': num(get('退款订单数（支付时间）')) or 0.0,
            'imp': num(get('商品曝光人数')) or 0.0,
            'clk': num(get('商品点击人数')) or 0.0,
            'good': num(get('好评数')) or 0.0,
        }
    return out


def agg_period(ds, mapping):
    agg = {c: dict(imp=0., clk=0., gmv=0., orders=0., good=0., refund_amt=0., refund_orders=0.) for c in CATS}
    st = dict(gmv=0., orders=0., refund_amt=0., refund_orders=0., imp=0., clk=0.)
    excl = dict(gmv=0., n=0)
    for code, it in ds.items():
        for k in ('gmv', 'orders', 'refund_amt', 'refund_orders', 'imp', 'clk'):
            st[k] += it[k]
        cat = mapping.get(code, (None, None))[0]
        if cat in agg:
            for k in agg[cat]:
                agg[cat][k] += it[k]
        else:
            excl['gmv'] += it['gmv']
            excl['n'] += 1
    return agg, st, excl


def div(a, b):
    return a / b if b else None


def ratio(cur, prev):
    return (cur / prev - 1) if prev else None


def short_label(ds):
    """从源数据的统计周期生成标签：2026/11/02-2026/11/08 -> 11/02—11/08。"""
    try:
        s = str(next(iter(ds.values()))['period']).strip()
        a, b = s.replace('—', '-').split('-')[:2]
        ay, am, ad = a.strip().split('/')
        by, bm, bd = b.strip().split('/')
        return '%s/%s—%s/%s' % (am.zfill(2), ad.zfill(2), bm.zfill(2), bd.zfill(2))
    except Exception:
        return str(next(iter(ds.values()))['period']).strip()


def main():
    mapping = load_mapping()
    p = load_source(SRC['prev'])
    c = load_source(SRC['curr'])
    ap, sp, ep = agg_period(p, mapping)
    ac, sc, ec = agg_period(c, mapping)

    bk = X.Book()
    ws = bk.add('品类汇总', widths={1: 20, **{i: 15 for i in range(2, 14)}}, heights={1: 28})

    period_label = {'prev': short_label(p), 'curr': short_label(c)}

    def add_block(agg, st, start, label, with_mom=False):
        rows = ws['rows']
        while len(rows) < start - 1:
            rows.append([])
        rows.append([(BANNER, label)] + [(0, None)] * 12)
        rows.append([(HEADER, None)] + [(HEADER, x) for x in CATS])
        def line(name, fn, style):
            rows.append([(LABEL, name)] + [(style, v) for v in [fn(agg[k]) for k in CATS]])
        line('曝光', lambda a: a['imp'], NUM0)
        line('点击人数', lambda a: a['clk'], NUM0)
        line('点击率', lambda a: div(a['clk'], a['imp']), PCT)
        line('成交金额', lambda a: a['gmv'], NUM2)
        line('成交订单数量', lambda a: a['orders'], NUM0)
        line('成交订单占比', lambda a: div(a['orders'], st['orders']), PCT)
        line('转化率', lambda a: div(a['orders'], a['clk']), PCT)
        line('品类流量占比', lambda a: div(a['clk'], st['clk']), PCT)
        line('好评数', lambda a: a['good'], NUM0)
        line('好评率', lambda a: div(a['good'], a['orders'] - a['refund_orders']), PCT)
        line('退款金额', lambda a: a['refund_amt'], NUM2)
        line('退款订单数', lambda a: a['refund_orders'], NUM0)
        line('退款率', lambda a: div(a['refund_orders'], a['orders']), PCT)
        rows.append([(LABEL, '实际成交单数')] + [(NUM0, a['orders'] - a['refund_orders']) for a in [agg[k] for k in CATS]])
        if with_mom:
            rows.append([(LABEL, '实际成交单数环比')] + [
                (PCT, ratio(agg[k]['orders'] - agg[k]['refund_orders'],
                            ap[k]['orders'] - ap[k]['refund_orders'])) for k in CATS])

    add_block(ap, sp, 1, period_label['prev'])
    add_block(ac, sc, 18, period_label['curr'], with_mom=True)

    # 店铺大盘
    rows = ws['rows']
    rows.append([])
    rows.append([(BANNER, '整体店铺大盘数据')] + [(0, None)] * 12)
    heads = ['成交金额', '退款金额', '实际金额', '成交订单数', '成交人数', '退款订单数', '退款率',
             '商品曝光人数', '商品点击人数', '商品点击率', '商品点击-支付转化率(次数)', '千次曝光成交金额']
    rows.append([(HEADER, None)] + [(HEADER, h) for h in heads])

    def store_row(st, label):
        vals = [st['gmv'], st['refund_amt'], st['gmv'] - st['refund_amt'], st['orders'], None,
                st['refund_orders'], div(st['refund_orders'], st['orders']), None, None, None, None, None]
        rows.append([(LABEL, label)] + [(NUM2 if i in (0, 1, 2) else NUM0 if i in (3, 4, 5) else PCT, v)
                                        for i, v in enumerate(vals)])

    store_row(sp, period_label['prev'])
    store_row(sc, period_label['curr'])
    mom = [ratio(sc['gmv'], sp['gmv']), ratio(sc['refund_amt'], sp['refund_amt']),
           ratio(sc['gmv'] - sc['refund_amt'], sp['gmv'] - sp['refund_amt']),
           ratio(sc['orders'], sp['orders']), None, ratio(sc['refund_orders'], sp['refund_orders']),
           ratio(div(sc['refund_orders'], sc['orders']), div(sp['refund_orders'], sp['orders'])),
           None, None, None, None, None]
    rows.append([(LABEL, '环比')] + [(PCT, v) for v in mom])

    # 口径说明
    rows.append([])
    note_lines = [
        '品类按商品编码参考表归类，XG/VG/PG/PA 显示器合并为“显示器”；商品曝光、点击为商品行求和，不代表店铺去重人数。',
        '好评率=好评数÷实际成交单数；退款率=退款订单数÷成交订单数；实际成交单数=成交订单数−退款订单数；实际金额=成交金额−退款金额。退款相关字段均为支付时间口径。',
        '店铺大盘包含全部商品；品类页不含路由器及未匹配商品。未展示成交金额：前周 %s 元（%d 款），后周 %s 元（%d 款）。'
        % ('{:,.2f}'.format(ep['gmv']), ep['n'], '{:,.2f}'.format(ec['gmv']), ec['n']),
    ]
    start = len(rows) + 1
    for i, t in enumerate(note_lines):
        rows.append([(LABEL if i == 0 else 0, '口径说明' if i == 0 else None), (NOTE, t)])
    ws['merges'] = ['A1:M1', 'A18:M18', 'A36:M36'] + \
                   ['A%d:A%d' % (start, start + 2)] + ['B%d:M%d' % (start + i, start + i) for i in range(3)]
    ws['heights'][start] = 30
    for i in range(1, 3):
        ws['heights'][start + i] = 30

    # 商品对比
    ws2 = bk.add('商品对比', widths={1: 10, 2: 20, 3: 56, **{i: 14 for i in range(4, 16)}})
    ws2['rows'].append([(HEADER, h) for h in ['品类', '型号', '商品名称', '前周成交金额', '后周成交金额', '金额环比',
                                              '点击人数环比', '前周成交订单', '后周成交订单', '订单环比',
                                              '前周退款订单', '后周退款订单', '前周实际订单', '后周实际订单', '实际订单环比']])
    items = []
    for code in sorted(set(p) | set(c)):
        cat, model = mapping.get(code, (None, None))
        if cat not in CATS:
            continue
        a = p.get(code) or {'gmv': 0., 'clk': 0., 'orders': 0., 'refund_orders': 0., 'name': None}
        b = c.get(code) or {'gmv': 0., 'clk': 0., 'orders': 0., 'refund_orders': 0., 'name': None}
        items.append({'cat': cat, 'model': model, 'code': code,
                      'name': (b.get('name') or a.get('name') or '').strip(),
                      'pa': a, 'cb': b})
    items.sort(key=lambda x: (CATS.index(x['cat']), -x['cb']['gmv'], -x['pa']['gmv'], x['name'], x['code']))
    for it in items:
        a, b = it['pa'], it['cb']
        ws2['rows'].append([
            (CELL, it['cat']), (CELL, it['model']), (CELL, it['name']),
            (NUM2, a['gmv']), (NUM2, b['gmv']), (PCT, ratio(b['gmv'], a['gmv'])),
            (PCT, ratio(b['clk'], a['clk'])),
            (NUM0, a['orders']), (NUM0, b['orders']), (PCT, ratio(b['orders'], a['orders'])),
            (NUM0, a['refund_orders']), (NUM0, b['refund_orders']),
            (NUM0, a['orders'] - a['refund_orders']), (NUM0, b['orders'] - b['refund_orders']),
            (PCT, ratio(b['orders'] - b['refund_orders'], a['orders'] - a['refund_orders'])),
        ])
    ws2['rows'].append([])
    idle = sum(1 for it in items if it['pa']['gmv'] + it['cb']['gmv'] <= 0.004)
    src = '两周商品明细；按后周成交金额排序。实际订单=成交订单数−退款订单数（支付时间）；空白环比表示前周无可比值。'
    if idle:
        src += '两期均无成交的商品 %d 款仅在明细表保留，不进入 PPT 商品页。' % idle
    ws2['rows'].append([(LABEL, '来源')] + [(NOTE, src)])

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    bk.save(OUT)
    print('写出:', OUT)
    print('前期排除 %d 款 %s' % (ep['n'], ep['gmv']))
    print('商品对比行数:', len(items))
    print('商品数按品类:', {k: sum(1 for i in items if i['cat'] == k) for k in CATS})


if __name__ == '__main__':
    main()
