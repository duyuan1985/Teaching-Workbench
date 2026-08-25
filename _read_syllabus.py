"""
读取教学大纲.doc文件，查看教材信息
"""
import os

doc_path = r'原始资料\教材\商务数据分析\大数据分析方法项目实战\01 教学大纲\《大数据分析方法项目实战》教学大纲.doc'

# 尝试用python-docx读取（可能不支持.doc）
try:
    from docx import Document
    doc = Document(doc_path)
    for p in doc.paragraphs[:30]:
        txt = p.text.strip()
        if txt:
            print(f'P: {txt[:200]}')
except Exception as e:
    print(f'python-docx无法读取: {e}')

# 尝试用二进制读取文本内容
try:
    with open(doc_path, 'rb') as f:
        content = f.read()
    # 尝试提取可读文本
    text = content.decode('utf-8', errors='ignore')
    # 找到中文文本
    import re
    chinese_texts = re.findall(r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]+', text)
    for t in chinese_texts[:50]:
        if len(t) > 5:
            print(f'提取: {t[:200]}')
except Exception as e:
    print(f'二进制读取失败: {e}')

# 尝试用olefile
try:
    import olefile
    ole = olefile.OleFileIO(doc_path)
    if ole.exists('WordDocument'):
        stream = ole.openstream('WordDocument')
        data = stream.read()
        # 尝试GBK解码
        text = data.decode('gbk', errors='ignore')
        import re
        chinese_texts = re.findall(r'[\u4e00-\u9fff]+[^\x00-\x7f\u4e00-\u9fff]*[\u4e00-\u9fff]+', text)
        for t in chinese_texts[:30]:
            if len(t) > 10:
                print(f'OLE: {t[:200]}')
    ole.close()
except Exception as e:
    print(f'olefile读取失败: {e}')
