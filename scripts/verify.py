# -*- coding: utf-8 -*-
"""出稿后校验：源数据重算 vs PPT 表格，加版式与商品页检查。

用法：
  python verify.py --prev 前周.xlsx --curr 后周.xlsx --pptx 周报.pptx [--map 映射.xlsx]

检查项：
  1. 分品类成交金额 + 排除项 = 店铺大盘（勾稽为 0）
  2. PPT 品类对比页的两期成交金额行、大盘页两期金额（按文字定位，不依赖固定行号）
  3. 版面：表格是否超出页面、标题/脚注是否与表格重叠、是否有空表页
  4. 商品页：是否残留两期零成交行、每页行数是否超过模板容量

退出码 0 表示全部通过，1 表示发现差异。
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--prev', required=True)
    ap.add_argument('--curr', required=True)
    ap.add_argument('--pptx', required=True)
    ap.add_argument('--map', default='')
    ap.add_argument('--cap', type=int, default=12)
    a = ap.parse_args()

    if a.map:
        os.environ['MAP_PATH'] = a.map

    import gen_report as G
    import gen_ppt as P
    from pptx import Presentation

    G.SRC['prev'] = a.prev
    G.SRC['curr'] = a.curr
    mapping = G.load_mapping()
    p = G.load_source(a.prev)
    c = G.load_source(a.curr)
    ap_, sp, ep = G.agg_period(p, mapping)
    ac, sc, ec = G.agg_period(c, mapping)

    prs = Presentation(a.pptx)
    H = prs.slide_height
    fails = []

    # ---- 1. 勾稽 ----
    print('== 勾稽 ==')
    print('分品类前周 %.2f + 排除 %.2f = %.2f，大盘 %.2f' % (
        sum(ap_[k]['gmv'] for k in G.CATS), ep['gmv'],
        sum(ap_[k]['gmv'] for k in G.CATS) + ep['gmv'], sp['gmv']))
    print('分品类后周 %.2f + 排除 %.2f = %.2f，大盘 %.2f' % (
        sum(ac[k]['gmv'] for k in G.CATS), ec['gmv'],
        sum(ac[k]['gmv'] for k in G.CATS) + ec['gmv'], sc['gmv']))
    for tag, tot, agg, ex in (('前周', sp['gmv'], ap_, ep), ('后周', sc['gmv'], ac, ec)):
        d = abs(sum(agg[k]['gmv'] for k in G.CATS) + ex['gmv'] - tot)
        if d > 0.01:
            fails.append('%s勾稽差 %.2f' % (tag, d))
    print('排除项：前周 %d 款 / %.2f，后周 %d 款 / %.2f' % (ep['n'], ep['gmv'], ec['n'], ec['gmv']))

    # ---- 2. PPT 数值 ----
    def tables(slide):
        return [sh for sh in slide.shapes if sh.has_table]

    def find_col_header(tbl):
        for ri, row in enumerate(tbl.rows):
            texts = [x.text.strip() for x in row.cells]
            hits = {t: i for i, t in enumerate(texts) if t in G.CATS}
            if len(hits) >= 8:
                return ri, hits
        return None, None

    def label_rows(tbl, label):
        return [ri for ri, row in enumerate(tbl.rows) if row.cells[0].text.strip() == label]

    print('\n== PPT 品类对比页 ==')
    cat_slide = None
    for s in prs.slides:
        ts = tables(s)
        if ts and find_col_header(ts[0].table)[0] is not None:
            cat_slide = s
            break
    if cat_slide is None:
        fails.append('未找到品类对比页')
    else:
        tbl = tables(cat_slide)[0].table
        _, cols = find_col_header(tbl)
        rows = label_rows(tbl, '成交金额')
        if len(rows) < 2:
            fails.append('品类页未找到两行「成交金额」')
        else:
            for idx, (ri, agg) in enumerate(zip(rows[:2], (ap_, ac))):
                bad = 0
                for cat in G.CATS:
                    ci = cols.get(cat)
                    if ci is None:
                        fails.append('品类页缺列 %s' % cat)
                        continue
                    got = tbl.cell(ri, ci).text.strip()
                    exp = P.money(agg[cat]['gmv'])
                    if got != exp:
                        bad += 1
                        if bad <= 3:
                            print('  差异 %s %s：PPT=%s 应为=%s' % (cat, '前周' if idx == 0 else '后周', got, exp))
                print('%s成交金额行 11 格，差异 %d' % ('前周' if idx == 0 else '后周', bad))
                if bad:
                    fails.append('%s品类成交金额 %d 格不一致' % ('前周' if idx == 0 else '后周', bad))

    print('\n== PPT 大盘页 ==')
    dash = None
    for s in prs.slides:
        for sh in tables(s):
            tb = sh.table
            for ri, row in enumerate(tb.rows):
                texts = [x.text.strip() for x in row.cells]
                if '成交金额' in texts and '退款金额' in texts:
                    dash = (tb, ri, texts.index('成交金额'))
                    break
            if dash:
                break
        if dash:
            break
    if dash is None:
        fails.append('未找到大盘页')
    else:
        tb, hri, ci_gmv = dash
        for off, (tag, st) in enumerate((('前周', sp), ('后周', sc)), start=1):
            got = tb.cell(hri + off, ci_gmv).text.strip() if hri + off < len(tb.rows) else ''
            exp = P.money(st['gmv'])
            ok = got == exp
            print('%s成交金额 PPT=%s 应为=%s %s' % (tag, got, exp, 'OK' if ok else '不一致'))
            if not ok:
                fails.append('%s大盘成交金额不一致' % tag)

    # ---- 3. 版式 ----
    print('\n== 版式 ==')
    for i, s in enumerate(prs.slides, 1):
        for sh in tables(s):
            tb = sh.table
            bottom = sh.top + sum(r.height for r in tb.rows)
            if bottom > H:
                fails.append('第 %d 页表格超出页面 %.2f in' % (i, (bottom - H) / 914400))
            if i > 4 and len(tb.rows) < 2:
                fails.append('第 %d 页为空表页' % i)
            for o in s.shapes:
                if o is sh or not o.has_text_frame or not o.text_frame.text.strip():
                    continue
                if o.top + o.height > sh.top and o.top < bottom:
                    fails.append('第 %d 页「%s」与表格重叠' % (i, o.text_frame.text[:10]))
    print('检查 %d 页，问题 %d 项' % (len(prs.slides), len([f for f in fails if '页' in f])))

    # ---- 4. 商品页 ----
    print('\n== 商品页 ==')
    def num(s):
        s = s.replace(',', '').replace('元', '').strip()
        try:
            return float(s) if s not in ('', '-', '—') else 0.0
        except ValueError:
            return 0.0

    total = 0
    over = []
    zero = []
    for i in range(5, len(prs.slides) + 1):
        ts = tables(prs.slides[i - 1])
        if not ts:
            continue
        tb = ts[0].table
        n = 0
        for row in list(tb.rows)[1:]:
            cs = [x.text for x in row.cells]
            if not cs[0].strip():
                continue
            n += 1
            total += 1
            if num(cs[1]) + num(cs[2]) <= 0:
                zero.append((i, cs[0]))
        if n > a.cap:
            over.append((i, n))
    print('商品行合计 %d，单页超容量 %s，零成交行 %d 行' % (total, over or '无', len(zero)))
    if over:
        fails.append('部分商品页超过 %d 行：%s' % (a.cap, over))
    if zero:
        fails.append('商品页仍有 %d 行两期零成交（前 3：%s）' % (len(zero), zero[:3]))

    print('\n== 结论 ==')
    if fails:
        for f in fails:
            print('  失败：' + f)
        sys.exit(1)
    print('  全部通过')


if __name__ == '__main__':
    main()
