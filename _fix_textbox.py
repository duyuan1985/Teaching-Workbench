"""
填充教学设计封面文本框内容
模板文本框包含：
  学年  第    学期 → 2023——2024学年  第二学期
  课程名称： → 商务数据分析
  班    级： → 2022电商教学班
  教    材： → 商务数据分析与应用（人民邮电出版社）
  授课教师： → 杜媛
"""
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

fp = r"生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx"
doc = Document(fp)

body = doc.element.body

# 找所有txbxContent
txbx_list = body.findall('.//' + qn('w:txbxContent'))
print(f"找到{len(txbx_list)}个文本框")

for ti, txbx in enumerate(txbx_list):
    texts = []
    for p in txbx.findall(qn('w:p')):
        para_text = ''
        for r in p.findall(qn('w:r')):
            for t in r.findall(qn('w:t')):
                para_text += t.text or ''
        if para_text.strip():
            texts.append(para_text.strip())
    
    if not texts:
        continue
    
    print(f"\nTextBox{ti}: {texts}")
    
    # 填充内容
    for p in txbx.findall(qn('w:p')):
        # 获取段落所有run的文本
        full_text = ''
        for r in p.findall(qn('w:r')):
            for t in r.findall(qn('w:t')):
                full_text += t.text or ''
        full_text = full_text.strip()
        
        # 替换内容
        new_text = None
        if '学年' in full_text and '学期' in full_text and '课程名称' not in full_text:
            new_text = '2023——2024学年  第二学期'
        elif '课程名称' in full_text:
            new_text = '课程名称：商务数据分析'
        elif '班    级' in full_text or '班 级' in full_text:
            new_text = '班    级：2022电商教学班'
        elif '教    材' in full_text or '教 材' in full_text:
            new_text = '教    材：商务数据分析与应用（人民邮电出版社）'
        elif '授课教师' in full_text:
            new_text = '授课教师：杜媛'
        
        if new_text:
            # 清除旧run文字，只保留第一个run并替换
            runs = p.findall(qn('w:r'))
            if runs:
                # 保留第一个run，清空其文字
                for t in runs[0].findall(qn('w:t')):
                    runs[0].remove(t)
                # 删除后续run
                for r in runs[1:]:
                    p.remove(r)
                # 在第一个run中写入新文字
                t = OxmlElement('w:t')
                t.text = new_text
                t.set(qn('xml:space'), 'preserve')
                runs[0].append(t)
                print(f"  → 填充: {new_text}")

doc.save(fp)
print(f"\n保存完成: {fp}")
