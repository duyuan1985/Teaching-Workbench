from docx import Document

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

print('=== 教学设计 P12-P36 ===')
for pi, p in enumerate(doc.paragraphs):
    if 12 <= pi <= 36:
        txt = p.text.strip()
        print(f'P{pi}: {txt[:150]}')
