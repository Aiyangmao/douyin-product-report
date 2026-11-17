# -*- coding: utf-8 -*-
"""极简 xlsx 写出器（标准库），支持少量样式。"""
import zipfile
from xml.sax.saxutils import escape

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
{sheets}
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="3">
<numFmt numFmtId="200" formatCode="#,##0.00"/>
<numFmt numFmtId="201" formatCode="0.00%"/>
<numFmt numFmtId="202" formatCode="#,##0"/>
</numFmts>
<fonts count="5">
<font><sz val="11"/><color rgb="FF252931"/><name val="宋体"/></font>
<font><b/><sz val="14"/><color rgb="FFFF0000"/><name val="微软雅黑"/></font>
<font><b/><sz val="11"/><color rgb="FF1F4E3D"/><name val="微软雅黑"/></font>
<font><b/><sz val="11"/><color rgb="FF252931"/><name val="宋体"/></font>
<font><i/><sz val="10"/><color rgb="FF7F7F7F"/><name val="宋体"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDCF2E6"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF4F4F4"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="9">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="200" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="202" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="201" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

TEXT, BANNER, HEADER, LABEL, NUM2, NUM0, PCT, CELL, NOTE = range(9)


def colname(i):
    s = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


class Book:
    def __init__(self):
        self.sheets = []  # (name, rows, widths, merges, heights)

    def add(self, name, widths=None, merges=None, heights=None):
        self.sheets.append({'name': name, 'rows': [], 'widths': widths or {},
                            'merges': merges or [], 'heights': heights or {}})
        return self.sheets[-1]

    @staticmethod
    def cell(ref, style, value):
        if value is None or value == '':
            return '<c r="%s" s="%d"/>' % (ref, style) if style else ''
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, style, repr(float(value)))
        return '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
            ref, style, escape(str(value)))

    def save(self, path):
        n = len(self.sheets)
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            over = '\n'.join(
                '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
                for i in range(n))
            z.writestr('[Content_Types].xml', CT.format(sheets=over))
            z.writestr('_rels/.rels', RELS)
            sheets_xml = ''.join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(s['name']), i + 1, i + 1)
                                 for i, s in enumerate(self.sheets))
            z.writestr('xl/workbook.xml',
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                       '<sheets>%s</sheets></workbook>' % sheets_xml)
            rels = ''.join(
                '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1)
                for i in range(n))
            rels += ('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (n + 1))
            z.writestr('xl/_rels/workbook.xml.rels',
                       '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>' % rels)
            z.writestr('xl/styles.xml', STYLES)
            for i, s in enumerate(self.sheets):
                cols = ''
                if s['widths']:
                    cols = '<cols>' + ''.join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (k, k, w)
                                              for k, w in s['widths'].items()) + '</cols>'
                body = []
                for r_idx, row in enumerate(s['rows'], start=1):
                    ht = s['heights'].get(r_idx)
                    ht_attr = ' ht="%s" customHeight="1"' % ht if ht else ''
                    cells = ''.join(self.cell('%s%d' % (colname(c), r_idx), st, v)
                                    for c, (st, v) in enumerate(row))
                    body.append('<row r="%d"%s>%s</row>' % (r_idx, ht_attr, cells))
                merges = ''
                if s['merges']:
                    merges = '<mergeCells count="%d">%s</mergeCells>' % (
                        len(s['merges']), ''.join('<mergeCell ref="%s"/>' % m for m in s['merges']))
                z.writestr('xl/worksheets/sheet%d.xml' % (i + 1),
                           '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                           '<sheetFormatPr defaultRowHeight="15"/>%s<sheetData>%s</sheetData>%s'
                           '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
                           '</worksheet>' % (cols, ''.join(body), merges))
