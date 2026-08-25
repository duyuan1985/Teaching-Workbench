from docx import Document

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

in_section = False
for pi, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if '（五）课程教学进度表设计' in txt:
        in_section = True
    if in_section:
        print(f'P{pi}: [{len(txt):3d}] {txt[:200]}')
    if in_section and '五、考核方案' in txt:
        break
