from docx import Document

# 查看授课计划表0中的章节信息
fp = r'生成结果\精修版\2023-2024-2《商务数据分析》授课计划 杜媛.docx'
doc = Document(fp)

# 找到表0（日程表）
t = doc.tables[0]
print(f'=== 授课计划 Table 0: {len(t.rows)} rows ===')
for ri in range(len(t.rows)):
    cells = t.rows[ri].cells
    texts = [c.text.strip()[:30] for c in cells]
    print(f'R{ri}: {texts[:5]}')
