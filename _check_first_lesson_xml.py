from docx import Document
from docx.oxml.ns import qn
from lxml import etree

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

# 检查P109-P119的XML
for pi in range(109, 120):
    p = doc.paragraphs[pi]
    txt = p.text.strip()
    # 检查是否有run
    runs = p.runs
    # 检查字体颜色
    color_info = ''
    for r in runs:
        rPr = r._element.find(qn('w:rPr'))
        if rPr is not None:
            color = rPr.find(qn('w:color'))
            if color is not None:
                color_info = f' color={color.get(qn("w:val"))}'
            vanish = rPr.find(qn('w:vanish'))
            if vanish is not None:
                color_info += ' VANISH!'
    print(f'P{pi} [{len(txt)}] runs={len(runs)}{color_info}: {txt[:100]}')
