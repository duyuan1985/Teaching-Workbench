from docx import Document
from docx.oxml.ns import qn

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

# 遍历body元素，查看段落和表格的顺序
body = doc.element.body
table_idx = 0
para_idx = 0
for elem in body:
    if elem.tag == qn('w:p'):
        txt = ''.join([t.text or '' for t in elem.findall('.//' + qn('w:t'))]).strip()
        if 35 <= para_idx <= 90:
            if txt:
                print(f'P{para_idx}: {txt[:100]}')
            else:
                print(f'P{para_idx}: (empty)')
        para_idx += 1
    elif elem.tag == qn('w:tbl'):
        # Get first cell text
        first_cell = ''
        tcs = elem.findall('.//' + qn('w:tc'))
        if tcs:
            first_cell = ''.join([t.text or '' for t in tcs[0].findall('.//' + qn('w:t'))]).strip()
        rows = elem.findall(qn('w:tr'))
        if 35 <= para_idx <= 90:
            print(f'  [Table {table_idx}: {len(rows)} rows, first_cell="{first_cell[:30]}"]')
        table_idx += 1
