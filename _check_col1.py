from docx import Document
from docx.oxml.ns import qn
from lxml import etree

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)
t4 = doc.tables[4]
tbl = t4._tbl
rows = tbl.findall(qn('w:tr'))

print(f'表格行数: {len(rows)}')
print()
for ri, tr in enumerate(rows):
    tcs = tr.findall(qn('w:tc'))
    for ci, tc in enumerate(tcs):
        texts = []
        for t in tc.findall('.//' + qn('w:t')):
            texts.append(t.text or '')
        full_text = ''.join(texts)

        tcPr = tc.find(qn('w:tcPr'))
        vmerge_info = ''
        if tcPr is not None:
            vm = tcPr.find(qn('w:vMerge'))
            if vm is not None:
                val = vm.get(qn('w:val'))
                vmerge_info = f' [vMerge={"restart" if val=="restart" else "continue"}]'

        print(f'  R{ri}C{ci}: "{full_text[:50]}"{vmerge_info}')
    print()
