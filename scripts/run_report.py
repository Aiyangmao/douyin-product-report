# -*- coding: utf-8 -*-
"""双期报告一键出稿：核定 Excel + 模板 PPT。

用法：
  python run_report.py --prev 前周.xlsx --curr 后周.xlsx
  python run_report.py --prev A.xlsx --curr B.xlsx --map 映射.xlsx --out-dir ./out --suffix _v2 --all-products

参数：
  --prev/--curr   两期经营版商品明细（含「商品编码」「统计周期」「退款订单数（支付时间）」）
  --map           商品编码→品类/型号映射表，默认用 skill 自带 assets/category-mapping.xlsx
  --out-dir       输出目录，默认当前目录
  --name          文件名主干，默认按两期周期自动生成，如 抖音双期报告_20261102-1115
  --suffix        文件名后缀，如 _v2（用于文件被占用时另存）
  --all-products  商品页保留两期零成交商品（默认只留有成交的）
  --cap           商品页每页最多数据行，默认 12
  --excel-only / --ppt-only  只出其中一种
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--prev', required=True)
    ap.add_argument('--curr', required=True)
    ap.add_argument('--map', default='')
    ap.add_argument('--out-dir', default='.')
    ap.add_argument('--name', default='')
    ap.add_argument('--suffix', default='')
    ap.add_argument('--all-products', action='store_true')
    ap.add_argument('--cap', type=int, default=12)
    ap.add_argument('--excel-only', action='store_true')
    ap.add_argument('--ppt-only', action='store_true')
    a = ap.parse_args()

    if a.map:
        os.environ['MAP_PATH'] = a.map

    import gen_report as G
    import gen_ppt as P

    G.SRC['prev'] = a.prev
    G.SRC['curr'] = a.curr

    # 先读一遍源数据，用于推导文件名和校验字段
    prev = G.load_source(a.prev)
    curr = G.load_source(a.curr)
    for tag, ds in (('前周', prev), ('后周', curr)):
        if not ds:
            sys.exit('%s 未读到商品行，检查工作表名是否为「全部」、表头是否含「商品编码」' % tag)
        miss = [k for k in ('gmv', 'orders') if k not in next(iter(ds.values()))]
        if miss:
            sys.exit('%s 缺少字段 %s' % (tag, miss))
        if '退款订单数' not in str(next(iter(ds.values())).get('refund_orders')) and \
                not any(v.get('refund_orders') for v in ds.values()):
            print('警告：%s 未读到退款订单数，实际成交单数与退款率将为空' % tag)

    def period_tag(ds):
        s = str(next(iter(ds.values())).get('period') or '')
        n = re.findall(r'\d+', s)
        return n[:3], n[3:6]

    if a.name:
        name = a.name
    else:
        p1, _ = period_tag(prev)
        _, p2 = period_tag(curr)
        p2 = p2 or p1
        if len(p1) >= 3 and len(p2) >= 3:
            name = '抖音双期报告_%s%s%s-%s%s' % (p1[0], p1[1], p1[2], p2[1], p2[2])
        else:
            name = '抖音双期报告'

    out_dir = os.path.abspath(a.out_dir)
    os.makedirs(out_dir, exist_ok=True)

    if not a.ppt_only:
        G.OUT = os.path.join(out_dir, '%s%s.xlsx' % (name, a.suffix))
    if not a.excel_only:
        P.OUT = os.path.join(out_dir, '%s%s.pptx' % (name.replace('双期报告', '店铺运营周报'), a.suffix))
        P.ONLY_ACTIVE = not a.all_products
        P.CAP = a.cap
        P.main()

    if not a.ppt_only:
        G.main()

    print('\n完成。源文件：')
    print('  前周', os.path.basename(a.prev), '共 %d 款' % len(prev))
    print('  后周', os.path.basename(a.curr), '共 %d 款' % len(curr))
    if a.excel_only:
        print('  Excel：', G.OUT)
    elif a.ppt_only:
        print('  PPT  ：', P.OUT)
    else:
        print('  Excel：', G.OUT)
        print('  PPT  ：', P.OUT)


if __name__ == '__main__':
    main()
