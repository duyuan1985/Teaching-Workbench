from docx import Document

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

# 找到第一节课设计梗概段落
in_section = False
for pi, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if '九、第一节课设计梗概' in txt:
        in_section = True
    if in_section:
        print(f'P{pi} [{len(txt)}]: {txt[:300]}')
    if in_section and '十、' in txt:
        break
