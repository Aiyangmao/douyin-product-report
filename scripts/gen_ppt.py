# -*- coding: utf-8 -*-
"""按 douyin-product-report skill 的参考模板，用两期源数据正式出稿 PPT。"""
import copy
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Emu

import gen_report as G

TPL = os.environ.get('SKILL_TPL') or os.path.join(G.SKILL, 'slides-reference.pptx')
OUT = os.environ.get('OUT_PPTX') or os.path.join(os.getcwd(), '抖音店铺运营周报.pptx')
CAP = 12  # 商品页每页最多行数
ONLY_ACTIVE = True  # 商品页只展示「两期成交金额合计 > 0」的商品，纯零成交款不占页
CATS = G.CATS


# ---------- 基础工具 ----------
def set_text_frame(tf, text):
    """只替换第一个 run 的文字，保留原字体/字号/颜色；多余的 run 和段落删掉。"""
    p = tf.paragraphs[0]
    runs = p.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p.text = text
    for extra in tf.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)


def set_cell(cell, text):
    set_text_frame(cell.text_frame, text)


def set_shape_text(shape, text):
    set_text_frame(shape.text_frame, text)


def set_rows(tbl, n):
    """把表格调整到 n 行（数据行，不含表头由调用方决定）。"""
    trs = tbl._tbl.tr_lst
    while len(trs) > n:
        tbl._tbl.remove(trs[-1])
        trs = tbl._tbl.tr_lst
    while len(trs) < n:
        new = copy.deepcopy(trs[-1])
        for tc in new.tc_lst:
            for p in tc.txBody.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
                for t in p.findall('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
                    t.text = ''
        tbl._tbl.append(new)
        trs = tbl._tbl.tr_lst


def move_slide(prs, old_idx, new_idx):
    lst = prs.slides._sldIdLst
    items = list(lst)
    el = items[old_idx]
    lst.remove(el)
    lst.insert(new_idx, el)


def clone_slide(prs, src_idx):
    """复制一页（含所有形状），返回新页索引。"""
    src = prs.slides[src_idx]
    new = prs.slides.add_slide(src.slide_layout)
    for shp in list(new.shapes):
        shp._element.getparent().remove(shp._element)
    for shp in src.shapes:
        new.shapes._spTree.append(copy.deepcopy(shp._element))
    return len(prs.slides) - 1


def delete_slide(prs, idx):
    lst = prs.slides._sldIdLst
    el = list(lst)[idx]
    rId = el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    prs.part.drop_rel(rId)
    lst.remove(el)


# ---------- 格式 ----------
def money(v):
    return '' if v is None else '{:,.2f}'.format(v)


def integer(v):
    return '' if v is None else '{:,.0f}'.format(v)


def pct(v):
    return '' if v is None else '{:.2f}%'.format(v * 100)


def period_label(s):
    """2026/10/26-2026/11/01 -> 2026.10.26—11.01"""
    s = str(s).strip()
    try:
        a, b = s.split('-') if s.count('-') == 1 else (s[:10], s[11:])
        y1, m1, d1 = [int(x) for x in a.replace('/', '-').split('-')]
        y2, m2, d2 = [int(x) for x in b.replace('/', '-').split('-')]
        return '%d.%02d.%02d—%02d.%02d' % (y1, m1, d1, m2, d2) if y1 == y2 else \
               '%d.%02d.%02d—%d.%02d.%02d' % (y1, m1, d1, y2, m2, d2)
    except Exception:
        return s


def short_label(s):
    """2026/10/26-2026/11/01 -> 10/26—11/01"""
    s = str(s).strip()
    try:
        a, b = s.split('-') if s.count('-') == 1 else (s[:10], s[11:])
        m1, d1 = [int(x) for x in a.replace('/', '-').split('-')[1:]]
        m2, d2 = [int(x) for x in b.replace('/', '-').split('-')[1:]]
        return '%02d/%02d—%02d/%02d' % (m1, d1, m2, d2)
    except Exception:
        return s


def div(a, b):
    return a / b if b else None


def ratio(x, y):
    return (x / y - 1) if y else None


# ---------- 主流程 ----------
def main():
    mapping = G.load_mapping()
    p = G.load_source(G.SRC['prev'])
    c = G.load_source(G.SRC['curr'])
    ap, sp, ep = G.agg_period(p, mapping)
    ac, sc, ec = G.agg_period(c, mapping)
    lab_p = short_label(next(iter(p.values()))['period'])
    lab_c = short_label(next(iter(c.values()))['period'])

    prs = Presentation(TPL)
    slides = list(prs.slides)

    # 1) 封面副标题
    cover = slides[0]
    sub = [sh for sh in cover.shapes if sh.has_text_frame and '—' in sh.text_frame.text]
    if sub:
        set_shape_text(sub[0], '%s  /  %s' % (period_label(next(iter(p.values()))['period']),
                                              period_label(next(iter(c.values()))['period'])))

    # 2) 品类对比页（slide3，33x12）
    t3 = [sh for sh in slides[2].shapes if sh.has_table][0].table
    blocks = [(lab_p, ap, sp), (lab_c, ac, sc)]
    metric_rows = ['曝光', '点击人数', '点击率', '成交金额', '成交订单数量', '成交订单占比',
                   '转化率', '品类流量占比', '好评数', '好评率', '退款金额', '退款订单数', '退款率']
    for bi, (label, agg, st) in enumerate(blocks):
        base = bi * 16
        set_cell(t3.cell(base, 0), label)
        for ci, cat in enumerate(CATS):
            set_cell(t3.cell(base + 1, ci + 1), cat)
        for mi, name in enumerate(metric_rows):
            a = None
            for ci, cat in enumerate(CATS):
                a = agg[cat]
                if name == '曝光':
                    v = integer(a['imp'])
                elif name == '点击人数':
                    v = integer(a['clk'])
                elif name == '点击率':
                    v = pct(div(a['clk'], a['imp']))
                elif name == '成交金额':
                    v = money(a['gmv'])
                elif name == '成交订单数量':
                    v = integer(a['orders'])
                elif name == '成交订单占比':
                    v = pct(div(a['orders'], st['orders']))
                elif name == '转化率':
                    v = pct(div(a['orders'], a['clk']))
                elif name == '品类流量占比':
                    v = pct(div(a['clk'], st['clk']))
                elif name == '好评数':
                    v = integer(a['good'])
                elif name == '好评率':
                    v = pct(div(a['good'], a['orders'] - a['refund_orders']))
                elif name == '退款金额':
                    v = money(a['refund_amt'])
                elif name == '退款订单数':
                    v = integer(a['refund_orders'])
                else:
                    v = pct(div(a['refund_orders'], a['orders']))
                set_cell(t3.cell(base + 2 + mi, ci + 1), v)
        for ci, cat in enumerate(CATS):
            a = agg[cat]
            set_cell(t3.cell(base + 15, ci + 1), integer(a['orders'] - a['refund_orders']))
    for ci, cat in enumerate(CATS):
        set_cell(t3.cell(32, ci + 1),
                 pct(ratio(ac[cat]['orders'] - ac[cat]['refund_orders'],
                           ap[cat]['orders'] - ap[cat]['refund_orders'])))

    # 3) 大盘页（slide4，5x13）
    t4 = [sh for sh in slides[3].shapes if sh.has_table][0].table
    store = [(lab_p, sp), (lab_c, sc)]
    for ri, (label, st) in enumerate(store):
        row = [label, money(st['gmv']), money(st['refund_amt']), money(st['gmv'] - st['refund_amt']),
               integer(st['orders']), '', integer(st['refund_orders']),
               pct(div(st['refund_orders'], st['orders'])), '', '', '', '', '']
        for i, v in enumerate(row):
            set_cell(t4.cell(2 + ri, i), v)
    mom = ['环比', pct(ratio(sc['gmv'], sp['gmv'])), pct(ratio(sc['refund_amt'], sp['refund_amt'])),
           pct(ratio(sc['gmv'] - sc['refund_amt'], sp['gmv'] - sp['refund_amt'])),
           pct(ratio(sc['orders'], sp['orders'])), '', pct(ratio(sc['refund_orders'], sp['refund_orders'])),
           pct(ratio(div(sc['refund_orders'], sc['orders']), div(sp['refund_orders'], sp['orders']))),
           '', '', '', '', '']
    for i, v in enumerate(mom):
        set_cell(t4.cell(4, i), v)

    # 4) 商品清单与筛选（口径说明要用到零成交款数，故先算）
    items = []
    for code in sorted(set(p) | set(c)):
        cat, model = mapping.get(code, (None, None))
        if cat not in CATS:
            continue
        a = p.get(code) or {'gmv': 0., 'clk': 0., 'orders': 0., 'refund_orders': 0., 'name': None}
        b = c.get(code) or {'gmv': 0., 'clk': 0., 'orders': 0., 'refund_orders': 0., 'name': None}
        items.append({'cat': cat, 'model': model, 'code': code,
                      'name': (b.get('name') or a.get('name') or '').strip(), 'a': a, 'b': b})
    all_items = list(items)
    dropped = []
    if ONLY_ACTIVE:
        # 两期成交金额合计为 0 的款（键帽、防滑贴等纯零成交周边）不占商品页
        kept, dropped = [], []
        for it in items:
            (kept if (it['a']['gmv'] + it['b']['gmv']) > 0.004 else dropped).append(it)
        items = kept
    items.sort(key=lambda x: (CATS.index(x['cat']), -x['b']['gmv'], -x['a']['gmv'], x['name'], x['code']))

    pages = {}
    empty_cats = []
    for cat in CATS:
        sub_ = [x for x in items if x['cat'] == cat]
        pg = [sub_[k:k + CAP] for k in range(0, len(sub_), CAP)]
        if not pg:
            # 该品类两期均无成交：仍保留一页表头，避免整类从目录里消失
            pg = [[]]
            empty_cats.append(cat)
        pages[cat] = pg

    # 5) 口径说明（含动态数字）
    for sh in slides[3].shapes:
        if sh.has_text_frame and '排除路由器' in sh.text_frame.text:
            note = '品类页排除路由器及未匹配商品：前周成交金额 %s 元，后周 %s 元' % (money(ep['gmv']), money(ec['gmv']))
            if dropped:
                note += '；商品页略去两期均无成交的商品 %d 款' % len(dropped)
            set_shape_text(sh, note + '。')

    # 模板中商品页的页码索引（1-based 幻灯片序号）
    slide_cat = {5: '主板', 6: '主板', 7: '显卡', 8: '显示器', 9: '显示器', 10: '显示器',
                 11: '键盘', 12: '键盘', 13: '键盘', 14: '鼠标', 15: '鼠标', 16: '耳机', 17: '耳机',
                 18: '特价', 19: '机箱', 20: '机箱', 21: '电源', 22: '散热',
                 23: '周边', 24: '周边', 25: '周边', 26: '周边', 27: '周边', 28: '周边'}
    tpl_pages = {}
    for sn, cat in slide_cat.items():
        tpl_pages.setdefault(cat, []).append(sn)

    # 页数不一致时就地增删：按品类逆序处理，先算好各品类在文档中的区块位置
    blocks = []
    cursor = 4
    for cat in CATS:
        have = len(tpl_pages[cat])
        blocks.append((cat, cursor, cursor + have))
        cursor += have
    for cat, start, end in reversed(blocks):
        need, have = len(pages[cat]), end - start
        if need < have:
            for k in range(have - need):
                delete_slide(prs, end - 1 - k)
        elif need > have:
            for k in range(need - have):
                new_idx = clone_slide(prs, end - 1)
                move_slide(prs, new_idx, end + k)

    # 重新建立 幻灯片序号 -> 品类 的映射
    slides = list(prs.slides)
    header = ['型号 / 商品', '前周成交金额', '后周成交金额', '金额环比', '点击环比',
              '前周实际订单', '后周实际订单', '订单环比']
    cursor = 4  # 前 4 页为封面/分节/品类/大盘；其后按品类顺序排列
    for cat in CATS:
        total = len(pages[cat])
        for pi, page in enumerate(pages[cat]):
            s = slides[cursor]
            cursor += 1
            tbl = [sh for sh in s.shapes if sh.has_table][0].table
            title = '%s周环比' % cat + ('' if total == 1 else '（%d/%d）' % (pi + 1, total))
            if not page:
                title = '%s周环比（两期均无成交）' % cat
            for sh in s.shapes:
                if sh.has_text_frame and '周环比' in sh.text_frame.text:
                    set_shape_text(sh, title)
                    break
            set_rows(tbl, len(page) + 1)
            for ci in range(8):
                set_cell(tbl.cell(0, ci), header[ci])
            for ri, it in enumerate(page):
                a, b = it['a'], it['b']
                vals = [it['model'] or it['name'][:24],
                        money(a['gmv']), money(b['gmv']), pct(ratio(b['gmv'], a['gmv'])),
                        pct(ratio(b['clk'], a['clk'])),
                        integer(a['orders'] - a['refund_orders']),
                        integer(b['orders'] - b['refund_orders']),
                        pct(ratio(b['orders'] - b['refund_orders'], a['orders'] - a['refund_orders']))]
                for ci, v in enumerate(vals):
                    set_cell(tbl.cell(ri + 1, ci), v)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    prs.save(OUT)
    print('已生成:', OUT)
    print('页数:', len(prs.slides))
    print('各品类页数:', {c: len(pages[c]) for c in CATS})
    print('商品总数 %d，上版 %d 款' % (len(items), len(items) + len(dropped)))
    if dropped:
        print('已过滤零成交 %d 款：%s' % (len(dropped), '、'.join(x['model'] or x['name'][:10] for x in dropped[:6]) + ' ...'))
    if empty_cats:
        print('两期无成交品类（保留空表头页）:', empty_cats)
    print('未展示:', ep['n'], money(ep['gmv']), '/', ec['n'], money(ec['gmv']))


if __name__ == '__main__':
    main()
