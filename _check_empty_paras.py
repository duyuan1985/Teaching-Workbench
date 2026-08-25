from docx import Document
from docx.oxml.ns import qn

fp = r'生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx'
doc = Document(fp)

# 检查Table 6的R0-R10各单元格的段落结构
t = doc.tables[6]
for ri in range(11):
    row = t.rows[ri]
    tcs = row._tr.findall(qn('w:tc'))
    for ci, tc in enumerate(tcs):
        paras = tc.findall(qn('w:p'))
        if len(paras) > 1:
            texts = []
            for p in paras:
                txt = ''
                for r in p.findall(qn('w:r')):
                    for elem in r:
                        if elem.tag == qn('w:t'):
                            txt += elem.text or ''
                texts.append(txt[:30] if txt else '(空段落)')
            print(f'R{ri} TC{ci}: {len(paras)}个段落')
            for pi, t_text in enumerate(texts):
                print(f'  P{pi}: {t_text}')
