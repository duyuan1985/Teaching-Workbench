"""
检查表0的合并单元格情况
"""
from docx import Document

fp = r"生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx"
doc = Document(fp)
t0 = doc.tables[0]

print(f"表0: {len(t0.rows)}行 x {len(t0.rows[0].cells)}列")
for ri in range(len(t0.rows)):
    cells = t0.rows[ri].cells
    print(f"\nR{ri} ({len(cells)}个cells):")
    for ci in range(len(cells)):
        t = cells[ci].text.strip()
        # 检查是否是合并单元格
        tc = cells[ci]._tc
        gridspan = tc.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gridSpan')
        vmerge = tc.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}vMerge')
        gs_val = gridspan.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if gridspan is not None else None
        vm_val = vmerge.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if vmerge is not None else ('continue' if vmerge is not None else None)
        print(f"  col{ci}: '{t[:30]}' gridSpan={gs_val} vMerge={vm_val}")
