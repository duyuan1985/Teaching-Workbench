"""
读取.doc文件（OLE格式 D0 CF 11 E0）
"""
import olefile
import re

doc_path = r'原始资料\教材\商务数据分析\大数据分析方法项目实战\01 教学大纲\《大数据分析方法项目实战》教学大纲.doc'

ole = olefile.OleFileIO(doc_path)

# 读取WordDocument流
if ole.exists('WordDocument'):
    stream = ole.openstream('WordDocument')
    data = stream.read()
    print(f'WordDocument size: {len(data)} bytes')
    
    # 尝试GBK解码
    text = data.decode('gbk', errors='ignore')
    
    # 提取中文文本片段
    chinese_texts = re.findall(r'[\u4e00-\u9fff\uff00-\uffef\u3000-\u303f][\u4e00-\u9fff\uff00-\uffef\u3000-\u303f\w\s.,;:()（）、。，；：""''！？\-—…·]+', text)
    
    for t in chinese_texts:
        if len(t) > 10:
            print(t[:300])

# 也尝试1Table或0Table
for table_name in ['1Table', '0Table']:
    if ole.exists(table_name):
        stream = ole.openstream(table_name)
        data = stream.read()
        text = data.decode('gbk', errors='ignore')
        chinese_texts = re.findall(r'[\u4e00-\u9fff][\u4e00-\u9fff\w\s.,;:()（）、。，；：""''！？\-—…·]+', text)
        for t in chinese_texts:
            if '十三五' in t or '规划' in t or '教材' in t or '出版' in t or '大学' in t:
                print(f'[{table_name}] {t[:300]}')

ole.close()
