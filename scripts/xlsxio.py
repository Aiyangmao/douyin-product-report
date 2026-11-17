"""最小 xlsx 读写工具（仅标准库），用于校验 skill 口径。"""
import zipfile, re
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def col_to_idx(ref):
    m = re.match(r'([A-Z]+)', ref)
    s = m.group(1)
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def read_sheets(path):
    """返回 {sheet_name: [ [cellvalue,...], ... ]}"""
    out = {}
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall(NS + 'si'):
                shared.append(''.join(t.text or '' for t in si.iter(NS + 't')))
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        relmap = {r.get('Id'): r.get('Target') for r in rels}
        for sh in wb.find(NS + 'sheets'):
            name = sh.get('name')
            rid = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            target = relmap[rid].lstrip('/')
            if not target.startswith('xl/'):
                target = 'xl/' + target
            rows = []
            sroot = ET.fromstring(z.read(target))
            for r in sroot.iter(NS + 'row'):
                cells = {}
                for c in r.findall(NS + 'c'):
                    ref = c.get('r')
                    idx = col_to_idx(ref)
                    t = c.get('t')
                    v = c.find(NS + 'v')
                    if t == 's':
                        val = shared[int(v.text)] if v is not None else None
                    elif t == 'inlineStr':
                        is_ = c.find(NS + 'is')
                        val = ''.join(x.text or '' for x in is_.iter(NS + 't')) if is_ is not None else None
                    elif t == 'str':
                        val = v.text if v is not None else None
                    else:
                        if v is None:
                            val = None
                        else:
                            txt = v.text
                            try:
                                val = float(txt)
                            except (TypeError, ValueError):
                                val = txt
                    cells[idx] = val
                if cells:
                    width = max(cells) + 1
                    rows.append([cells.get(i) for i in range(width)])
                else:
                    rows.append([])
            out[name] = rows
    return out


def num(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(',', ''))
    except ValueError:
        return None
