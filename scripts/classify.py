# -*- coding: utf-8 -*-
"""给未匹配商品补品类/型号，输出更新版映射表 + 待确认清单。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

import gen_report as G

SKILL_MAP = os.path.join(G.SKILL, 'category-mapping.xlsx')
OUT = os.environ.get('OUT_MAP') or os.path.join(os.getcwd(), 'category-mapping-补充版.xlsx')

# (名称关键词, 品类, 型号, 归类依据)  按关键词长度从长到短匹配，先命中先得
ASSIGN = [
    ('特惠DIY组合套包', '特价', '特价DIY', '套包链接，沿用映射表「特价DIY」先例'),
    ('外设键盘特惠链接', '特价', '特价外设', '外设特惠链接，沿用「特价外设」先例'),
    ('特价二手数码配件', '特价', '特价外设', '特价清仓链接，沿用「特价外设」先例'),
    ('瑕疵显示器特价专区', '特价', '特价显示器', '特价专区链接，归特价'),
    ('瑕疵特价清库', '特价', '特价外设', '特价清库链接，归特价'),
    ('秒杀价专用链接', '特价', '特价外设', '秒杀/补差链接，归特价'),
    ('9820周年键帽', '周边', '键帽', '键帽，沿用「键帽→周边」先例'),
    ('夜魔Ex SE风格复刻键帽', '周边', '键帽', '键帽，沿用「键帽→周边」先例'),
    ('透明键帽', '周边', '键帽', '键帽，沿用「键帽→周边」先例'),
    ('ZOMO原创猫爪键帽', '周边', '键帽', '单颗键帽，沿用「键帽→周边」先例'),
    ('定制ROG金属键帽', '周边', '键帽', '键帽，沿用「键帽→周边」先例'),
    ('EVA 2号机RX键帽', '周边', '键帽', '联名替换键帽，沿用「键帽→周边」先例'),
    ('克罗心键帽', '周边', '键帽', '键帽，名称可辨'),
    ('仿碳键帽', '周边', '键帽', '键帽，名称可辨'),
    ('零度指令', '周边', '键帽', 'CNC 金属键帽，名称可辨'),
    ('孔雀翎', '周边', '键帽', '客制化键帽，名称可辨'),
    ('雾透白', '周边', '键帽', '键帽配色款（适配87配列），名称可辨'),
    ('碳纤维黑适配87', '周边', '键帽', '键帽配色款（适配87配列），名称可辨'),
    ('防尘盖', '周边', '键盘配件', '键盘防尘盖，归周边配件'),
    ('CNC铝合金键盘外壳', '周边', '键盘配件', '键盘外壳改装件，归周边配件'),
    ('改装铝合金金属外壳上盖底壳', '周边', '键盘配件', '键盘外壳改装件，归周边配件'),
    ('屏幕保护膜', '周边', '屏幕保护膜', '键屏保护膜，归周边配件'),
    ('防滑贴', '周边', '防滑贴', '沿用「防滑贴→周边」先例'),
    ('NX轴体30颗', '周边', '轴体', '轴体配件，键帽同类归周边'),
    ('10颗轴体', '周边', '轴体', '轴体配件，键帽同类归周边'),
    ('拆机风扇', '周边', '风扇', '沿用「风扇→周边」先例（如 TUF TF120/F120）'),
    ('GR120', '周边', '风扇', 'ARGB 风扇，沿用「风扇→周边」先例'),
    ('Modgo鼠标垫', '周边', '鼠标垫', '鼠标垫，沿用先例'),
    ('一个垫子', '周边', '鼠标垫', '垫类（推断为鼠标垫），已标注推断'),
    ('氮化镓快速充电器', '周边', '充电器', '充电器，沿用「a豆充电器→周边」先例'),
    ('手提包', '周边', '包袋', '包袋，沿用先例'),
    ('零钱包', '周边', '包袋', '包袋，沿用先例'),
    ('一个华硕包', '周边', '包袋', '包袋，名称可辨'),
    ('一个杯子', '周边', '杯具', '杯具，沿用「马克杯/不锈钢杯→周边」先例'),
    ('杯碟套装', '周边', '杯具', '杯具，沿用先例'),
    ('典藏足球', '周边', '周边杂项', 'ROG 周边纪念品'),
    ('冰箱贴', '周边', '周边杂项', '周边摆件'),
    ('初音未来 挂件', '周边', '周边杂项', '周边挂件'),
    ('毛绒挂件', '周边', '周边杂项', '周边挂件'),
    ('玩家桶', '周边', '周边杂项', '盲盒周边'),
    ('快乐纸仓', '周边', '周边杂项', '亚克力周边'),
    ('积木熊', '周边', '周边杂项', '联名积木熊周边'),
    ('多巴胺撞色', '周边', '周边杂项', 'ROG 周边'),
    ('蓝光', '周边', '周边杂项', '外置光驱，归周边'),
    ('PCE-AX58BT', '周边', '周边杂项', '无线网卡，归周边'),
    ('PCIE延长线', '周边', '周边杂项', '显卡竖装延长线，归周边'),
    ('MESH侧板', '周边', '周边杂项', '机箱替换侧板，归周边'),
    ('装机手套', '周边', '周边杂项', '周边用品'),
    ('大力神支架', '周边', '大力神支架', '沿用「大力神支架→周边」先例'),
    ('耳机支架', '周边', '周边杂项', '支架类，沿用「支架→周边」先例'),
    ('XIN120', '周边', '机箱灯板', '机箱灯板，周边'),
    ('雷切2 Pro', '周边', '雷切2', '沿用映射表「雷切2→周边」先例'),
    ('飞龙4代', '散热', '飞龙4', '一体式水冷散热器，沿用「飞龙→散热」先例'),
    ('龙王3代', '散热', '龙王3', '一体式水冷散热器，沿用「龙王3→散热」先例'),
    ('龙神2 360一体式', '散热', '龙神2', '一体式水冷散热器，沿用「龙神→散热」先例'),
    ('GR801幻世神机箱', '机箱', 'GR801', '机箱，名称可辨'),
    ('PRIME大师750W', '电源', '大师电源', '电源，沿用「大师电源」先例'),
    ('巨齿鲨750W', '电源', '巨齿鲨电源', '电源，名称可辨'),
    ('雷神3代', '电源', '雷神电源', '电源，沿用「雷神电源」先例'),
    ('X870E-A', '主板', 'X870系列', 'X870E 主板，归 X870 系列'),
    ('B760-I', '主板', '760系列', 'B760 主板，归 760 系列'),
    ('术士98', '键盘', '术士98', '机械键盘，名称可辨'),
    ('夜魔Extreme无线三模机械键盘', '键盘', '夜魔Extreme', '机械键盘，夜魔系列'),
    ('夜魔客制化铝坨坨', '键盘', '夜魔', '改装键盘，归夜魔'),
    ('TX75 磁轴键盘', '键盘', 'TX75', '磁轴键盘，名称可辨'),
    ('游侠RXTKL', '键盘', '游侠RX', '光轴键盘，归游侠RX'),
    ('一个魔导士75HE', '键盘', '魔导士75 HE', '键盘，名称可辨'),
    ('降临OWS', '耳机', '降临OWS', '耳机，沿用「降临」系列先例'),
    ('影刃2', '鼠标', '影刃2', '鼠标，沿用「影刃→鼠标」先例'),
    ('一个鼠标', '鼠标', '其他鼠标', '名称即鼠标，0 成交'),
    ('PG32UCWM', '显示器', 'PG32UCWM', '显示器，名称可辨'),
    ('PG27UCWM', '显示器', 'PG27UCWM', '显示器，名称可辨'),
    ('XG32UQDMS', '显示器', 'XG32UQDMS', '显示器，名称可辨'),
    ('XG27UQDMS', '显示器', 'XG27UQDMS', '显示器，名称可辨'),
    ('XG32UCWG', '显示器', 'XG32UCWG', '显示器，名称可辨'),
    ('XG256Q', '显示器', 'XG256Q', '显示器，名称可辨'),
    ('VG27AQL3A', '显示器', 'VG27AQL3A', '显示器（瑕疵特价），归显示器'),
    ('VG279QL1A', '显示器', 'VG279QL1A', '显示器，名称可辨'),
    ('VG258QM', '显示器', 'VG258QM', '显示器，名称可辨'),
    ('VG34WQML5A', '显示器', 'VG34WQML5A', '显示器，名称可辨'),
    ('VP32UQ', '显示器', 'VP32UQ', '显示器，名称可辨'),
    ('MB169CK', '显示器', 'MB169CK', '便携显示器，名称可辨'),
    ('路由器', '路由器', '路由器', '路由器，按口径不进品类页'),
    ('分布式子母路由', '路由器', '路由器', '路由器，按口径不进品类页'),
    # ↓ 2026-11 由用户确认的两条
    ('Womier量子蜂巢', '周边', '键帽', '键帽套装，用户确认为键帽'),
    ('魔导士Ace HFX', '键盘', '魔导士Ace HFX', '磁轴键盘（镂空外壳为改装件），用户确认为键盘'),
]

# 名称无法判断时进这里，等用户确认；确认后请把结论移入 ASSIGN，避免每次重复问
PENDING = {
}


def main():
    mapping = G.load_mapping()
    p = G.load_source(G.SRC['prev'])
    c = G.load_source(G.SRC['curr'])
    rows = []
    for code in set(p) | set(c):
        if mapping.get(code, (None, None))[0] in G.CATS:
            continue
        a = p.get(code) or {'gmv': 0.}
        b = c.get(code) or {'gmv': 0.}
        rows.append({'code': code, 'name': (b.get('name') or a.get('name') or '').strip(),
                     'pa': a['gmv'], 'cb': b['gmv']})
    rows.sort(key=lambda r: (-(r['pa'] + r['cb']), r['code']))

    frags = sorted(ASSIGN, key=lambda x: -len(x[0]))
    pend = sorted(PENDING, key=lambda x: -len(x))

    resolved, pending = [], []
    for r in rows:
        hit = None
        for frag, cat, model, why in frags:
            if frag in r['name']:
                hit = (cat, model, why)
                break
        if hit is None:
            for frag in pend:
                if frag in r['name']:
                    hit = ('待确认', '', PENDING[frag])
                    break
        if hit is None:
            hit = ('待确认', '', '名称无品类信息')
        r.update(cat=hit[0], model=hit[1], why=hit[2])
        (pending if hit[0] == '待确认' else resolved).append(r)

    print('未匹配 %d 条：可补 %d，待确认 %d' % (len(rows), len(resolved), len(pending)))
    res_prev = sum(r['pa'] for r in resolved)
    res_curr = sum(r['cb'] for r in resolved)
    print('可补金额 前周 %.2f 后周 %.2f' % (res_prev, res_curr))
    print('待确认金额 前周 %.2f 后周 %.2f' % (sum(r['pa'] for r in pending), sum(r['cb'] for r in pending)))
    router = [r for r in resolved if r['cat'] == '路由器']
    print('其中路由器 %d 款（按口径不进品类页）' % len(router))

    # --- 写更新版映射表 ---
    wb = load_workbook(SKILL_MAP)
    ws = wb['Sheet1']
    start = ws.max_row + 1
    for i, r in enumerate(resolved + pending):
        cell = ws.cell(row=start + i, column=1, value=r['code'])
        cell.number_format = '@'
        ws.cell(row=start + i, column=2, value=r['cat'])
        ws.cell(row=start + i, column=3, value=r['model'])

    # 明细页
    if '补充明细' in wb.sheetnames:
        del wb['补充明细']
    ws2 = wb.create_sheet('补充明细')
    head = ['商品编码', '品类', '型号', '合计成交金额', '前周成交金额', '后周成交金额', '商品名称', '归类依据', '状态']
    ws2.append(head)
    for r in resolved + pending:
        ws2.append([r['code'], r['cat'], r['model'], round(r['pa'] + r['cb'], 2), round(r['pa'], 2),
                    round(r['cb'], 2), r['name'], r['why'],
                    '待确认' if r['cat'] == '待确认' else '已补'])
    for row in ws2.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.font = Font(bold=True, color='FFFFFFFF')
            cell.fill = PatternFill('solid', fgColor='FF1F4E3D')
            cell.alignment = Alignment(horizontal='center', vertical='center')
    widths = [22, 10, 16, 14, 14, 14, 62, 40, 10]
    for i, w in enumerate(widths, 1):
        ws2.column_dimensions[ws2.cell(row=1, column=i).column_letter].width = w
    for row in ws2.iter_rows(min_row=2):
        for cell in row[:1]:
            cell.number_format = '@'
        row[3].number_format = '#,##0.00'
        row[4].number_format = '#,##0.00'
        row[5].number_format = '#,##0.00'
        row[6].alignment = Alignment(wrap_text=True, vertical='center')
        row[7].alignment = Alignment(wrap_text=True, vertical='center')
    ws2.freeze_panes = 'A2'

    # 待确认页
    if '待确认' in wb.sheetnames:
        del wb['待确认']
    ws3 = wb.create_sheet('待确认')
    ws3.append(['商品编码', '商品名称', '前周成交金额', '后周成交金额', '待确认原因'])
    for r in sorted(pending, key=lambda x: -(x['pa'] + x['cb'])):
        ws3.append([r['code'], r['name'], round(r['pa'], 2), round(r['cb'], 2), r['why']])
    for cell in ws3[1]:
        cell.font = Font(bold=True)
    for i, w in enumerate([22, 60, 14, 14, 34], 1):
        ws3.column_dimensions[ws3.cell(row=1, column=i).column_letter].width = w
    for row in ws3.iter_rows(min_row=2):
        row[0].number_format = '@'
        row[2].number_format = '#,##0.00'
        row[3].number_format = '#,##0.00'

    wb.save(OUT)
    print('写出:', OUT)
    print('映射表原有 %d 行，新增 %d 行' % (start - 2, len(resolved) + len(pending)))


if __name__ == '__main__':
    main()
