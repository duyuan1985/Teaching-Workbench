from docx import Document
from docx.oxml.ns import qn

fp = r'原始资料\模板\模板5：教学设计\模板5：教学设计 模板（2023-2024）.docx'
doc = Document(fp)

t = doc.tables[5]
print(f'=== 模板 Table 5: {len(t.rows)} rows ===')

# R9-R15的tc级别结构
for ri in range(9, 16):
    row = t.rows[ri]
    tcs = row._tr.findall(qn('w:tc'))
    print(f'\nR{ri}: {len(tcs)} tc elements')
    for ci, tc in enumerate(tcs):
        txt = ''
        for p in tc.findall(qn('w:p')):
            for r in p.findall(qn('w:r')):
                for elem in r:
                    if elem.tag == qn('w:t'):
                        txt += elem.text or ''
        # 检查gridSpan和vMerge
        tcPr = tc.find(qn('w:tcPr'))
        gs = ''
        vm = ''
        if tcPr is not None:
            gs_elem = tcPr.find(qn('w:gridSpan'))
            if gs_elem is not None:
                gs = f' gridSpan={gs_elem.get(qn("w:val"))}'
            vm_elem = tcPr.find(qn('w:vMerge'))
            if vm_elem is not None:
                vm_val = vm_elem.get(qn('w:val'))
                vm = f' vMerge={vm_val or "continue"}'
        print(f'  TC{ci}: [{txt[:30]}]{gs}{vm}')
