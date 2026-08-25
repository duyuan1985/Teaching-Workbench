from docx import Document
import re

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

in_section = False
total = 0
for pi, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if '九、第一节课' in txt:
        in_section = True
    if in_section and txt:
        print(f'P{pi}: {txt[:120]}')
        m = re.search(r'(\d+)分钟', txt)
        if m:
            total += int(m.group(1))
    if in_section and '十、' in txt:
        break

print(f'\n总时长: {total}分钟')
