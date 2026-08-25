"""
分析模板中单元教学设计的文字描述部分结构
模板中P116-P177是第一个任务的文字描述模板
"""
from docx import Document
from docx.oxml.ns import qn

template_fp = r"原始资料\模板\模板5：教学设计\模板5：教学设计 模板（2023-2024）.docx"
doc = Document(template_fp)

print("=== 模板段落结构（P116开始）===")
for i, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if i >= 114 and i <= 210:
        # 检查格式
        runs_info = ""
        if p.runs:
            r = p.runs[0]
            fn = r.font.name or ""
            sz = r.font.size.pt if r.font.size else ""
            bold = r.font.bold
            runs_info = f"[{fn},{sz},bold={bold}]"
        print(f"  P{i}: {runs_info} {text[:70]}")
