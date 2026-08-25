from docx import Document
from docx.oxml.ns import qn

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

# Table 5(seq=1), 7(seq=2), 9(seq=3), 11(seq=4) 都是第一章
table_indices = [5, 7, 9, 11]
labels = ['seq=1', 'seq=2', 'seq=3', 'seq=4']

for idx, ti in enumerate(table_indices):
    t = doc.tables[ti]
    
    # R3 教学任务
    r3_txt = ''
    for tc in t.rows[3]._tr.findall(qn('w:tc')):
        txt = ''
        for p in tc.findall(qn('w:p')):
            for r in p.findall(qn('w:r')):
                for elem in r:
                    if elem.tag == qn('w:t'):
                        txt += elem.text or ''
                    elif elem.tag == qn('w:br'):
                        txt += ' | '
        if txt and '第' in txt:
            r3_txt = txt
    
    # R4-R7: 知识目标、能力目标、思政目标、素质目标
    goals = {}
    for ri in range(4, 8):
        label = t.rows[ri].cells[1].text.strip()
        content = t.rows[ri].cells[2].text.strip()
        goals[label] = content
    
    print(f'\n=== Table {ti} ({labels[idx]}) ===')
    print(f'教学任务: {r3_txt}')
    for k, v in goals.items():
        print(f'{k}: {v}')
