from docx import Document

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)
t3 = doc.tables[3]
print(f'Table 3: {len(t3.rows)} rows, {len(t3.columns)} cols')
for ri in range(len(t3.rows)):
    texts = [c.text.strip()[:30] for c in t3.rows[ri].cells]
    print(f'  R{ri}: {texts}')
