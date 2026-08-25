from docx import Document
from docx.oxml.ns import qn

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

# 检查Table 5的R8结构
t = doc.tables[5]
r8 = t.rows[8]
tcs = r8._tr.findall(qn('w:tc'))
print(f'R8: {len(tcs)} tc elements')
for ci, tc in enumerate(tcs):
    txt = ''
    for p in tc.findall(qn('w:p')):
        for r in p.findall(qn('w:r')):
            for elem in r:
                if elem.tag == qn('w:t'):
                    txt += elem.text or ''
    tcPr = tc.find(qn('w:tcPr'))
    gs = ''
    vm = ''
    if tcPr is not None:
        gs_elem = tcPr.find(qn('w:gridSpan'))
        if gs_elem is not None:
            gs = f' gridSpan={gs_elem.get(qn("w:val"))}'
        vm_elem = tcPr.find(qn('w:vMerge'))
        if vm_elem is not None:
            vm = f' vMerge={vm_elem.get(qn("w:val")) or "continue"}'
    print(f'  TC{ci}: [{txt[:80]}]{gs}{vm}')
