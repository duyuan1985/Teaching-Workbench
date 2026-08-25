"""
检查表2结构和修复职业能力训练表
"""
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from copy import deepcopy
import store

fp = r"生成结果\精修版\2023-2024-2《商务数据分析》教学设计 杜媛.docx"
doc = Document(fp)

# 检查表2结构
t2 = doc.tables[2]
print(f"表2: {len(t2.rows)}行x{len(t2.rows[0].cells)}列")
for ri in range(min(5, len(t2.rows))):
    cells = t2.rows[ri].cells
    texts = [c.text.strip()[:25] for c in cells]
    print(f"  R{ri}: {texts}")

# 检查任务和单元数据
tasks = store.rows("SELECT * FROM tasks WHERE offering_id=20 ORDER BY seq")
units = store.rows("SELECT * FROM curriculum_units WHERE offering_id=20 ORDER BY seq")

print(f"\n单元数: {len(units)}")
print(f"任务数: {len(tasks)}")

for u in units[:3]:
    print(f"  unit: id={u.get('id')}, title={u.get('title')}")
for t in tasks[:3]:
    print(f"  task: id={t.get('id')}, title={t.get('title')}, unit_id={t.get('unit_id')}, seq={t.get('seq')}")
