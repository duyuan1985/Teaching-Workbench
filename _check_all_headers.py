from docx import Document
from docx.oxml.ns import qn

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

table_indices = list(range(5, 63, 2))
print(f'检查 {len(table_indices)} 个表格\n')

for ti in table_indices:
    t = doc.tables[ti]
    r9 = t.rows[9]
    r9_tcs = r9._tr.findall(qn('w:tc'))
    
    # 获取R3教学任务文本
    r3_txt = ''
    for tc in t.rows[3]._tr.findall(qn('w:tc')):
        for p in tc.findall(qn('w:p')):
            for r in p.findall(qn('w:r')):
                for elem in r:
                    if elem.tag == qn('w:t'):
                        r3_txt += elem.text or ''
                    elif elem.tag == qn('w:br'):
                        r3_txt += '|'
    
    # 检查R9各TC的文本
    r9_texts = []
    for tc in r9_tcs:
        txt = ''
        for p in tc.findall(qn('w:p')):
            for r in p.findall(qn('w:r')):
                for t_elem in r.findall(qn('w:t')):
                    txt += t_elem.text or ''
        r9_texts.append(txt[:20])
    
    # 检查是否有"知识＆技能"
    has_knowledge = any('知识' in t and '技能' in t for t in r9_texts)
    has_key = any('重点' in t for t in r9_texts)
    has_diff = any('难点' in t for t in r9_texts)
    
    status = 'OK' if (has_knowledge and has_key and has_diff) else 'FAIL'
    
    print(f'Table {ti}: [{status}] R3={r3_txt[:40]}')
    if status == 'FAIL':
        print(f'  R9: {r9_texts}')
