from docx import Document
from docx.oxml.ns import qn

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

t = doc.tables[5]
print(f'=== 当前 Table 5: {len(t.rows)} rows ===')

for ri in range(9, 15):
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
        print(f'  TC{ci}: [{txt[:40]}]{gs}{vm}')
