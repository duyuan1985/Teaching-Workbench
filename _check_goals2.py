from docx import Document

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

# 找到课程目标设计开始的段落
start = False
count = 0
for pi, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if '课程目标设计' in txt:
        start = True
    if start:
        print(f'P{pi}: {txt[:200]}')
        count += 1
        if count > 25:
            break
