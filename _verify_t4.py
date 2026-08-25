from docx import Document
from docx.oxml.ns import qn

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)
t4 = doc.tables[4]
tbl = t4._tbl
rows = tbl.findall(qn('w:tr'))

print('=== 合并结构验证 ===')
for ri, tr in enumerate(rows):
    tcs = tr.findall(qn('w:tc'))
    for ci, tc in enumerate(tcs):
        tcPr = tc.find(qn('w:tcPr'))
        vmerge = ''
        if tcPr is not None:
            vm = tcPr.find(qn('w:vMerge'))
            if vm is not None:
                val = vm.get(qn('w:val'))
                vmerge = ' [vMerge=restart]' if val == 'restart' else ' [vMerge=continue]'
        text = ''.join([t.text or '' for t in tc.findall('.//' + qn('w:t'))])[:30]
        print(f'  R{ri}C{ci}: "{text}"{vmerge}')
    print()
