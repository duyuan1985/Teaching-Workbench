from docx import Document
from docx.oxml.ns import qn

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

# 验证3个任务的R5/R6/R7内容长度
for idx, ti in enumerate([6, 8, 20]):
    seq = idx + 1 if ti == 6 else (idx + 2 if ti == 8 else 10)
    t = doc.tables[ti]
    print(f'\n=== Table {ti} (seq~{seq}) ===')
    for ri in [5, 6, 7]:
        row = t.rows[ri]
        tcs = row._tr.findall(qn('w:tc'))
        for tc in tcs:
            txt = ''
            for p in tc.findall(qn('w:p')):
                for r in p.findall(qn('w:r')):
                    for elem in r:
                        if elem.tag == qn('w:t'):
                            txt += elem.text or ''
            if txt and len(txt) > 50:
                print(f'  R{ri}: {len(txt)}字符 | {txt[:120]}...')
                break
