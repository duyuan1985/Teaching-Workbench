from docx import Document
from docx.oxml.ns import qn
from lxml import etree

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)
t4 = doc.tables[4]
tbl = t4._tbl

# 打印R1C0的完整XML（有文本的单元格）
rows = tbl.findall(qn('w:tr'))
tc = rows[1].findall(qn('w:tc'))[0]
xml_str = etree.tostring(tc, pretty_print=True).decode()
print('=== R1C0 (商务数据分析) XML ===')
print(xml_str[:1000])

# 打印R1C2的完整XML
tc2 = rows[1].findall(qn('w:tc'))[2]
xml_str2 = etree.tostring(tc2, pretty_print=True).decode()
print('\n=== R1C2 (签到) XML ===')
print(xml_str2[:1000])

# 检查表格grid定义
tblPr = tbl.find(qn('w:tblPr'))
tblGrid = tbl.find(qn('w:tblGrid'))
if tblGrid is not None:
    gridCols = tblGrid.findall(qn('w:gridCol'))
    print(f'\n=== tblGrid: {len(gridCols)} cols ===')
    for gi, gc in enumerate(gridCols):
        w = gc.get(qn('w:w'))
        print(f'  col {gi}: w={w}')

# 检查字体颜色
print('\n=== 字体颜色检查 ===')
for ri, tr in enumerate(rows):
    tcs = tr.findall(qn('w:tc'))
    for ci, tc in enumerate(tcs):
        for r in tc.findall('.//' + qn('w:r')):
            rPr = r.find(qn('w:rPr'))
            color_info = ''
            if rPr is not None:
                color = rPr.find(qn('w:color'))
                if color is not None:
                    color_info = f' color={color.get(qn("w:val"))}'
                else:
                    color_info = ' no color tag'
            t_elem = r.find(qn('w:t'))
            text = t_elem.text if t_elem is not None else ''
            if text:
                print(f'  R{ri}C{ci}: "{text[:20]}"{color_info}')
