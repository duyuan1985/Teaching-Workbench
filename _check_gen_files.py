import generate, inspect

# 检查output_dir默认值
src = inspect.getsource(generate.generate_plan)
print("=== generate_plan 前几行 ===")
lines = src.split('\n')
for i, line in enumerate(lines[:15]):
    print(f"  {line}")

src2 = inspect.getsource(generate.generate_standard)
print("\n=== generate_standard 前几行 ===")
lines2 = src2.split('\n')
for i, line in enumerate(lines2[:15]):
    print(f"  {line}")

src3 = inspect.getsource(generate.generate_design)
print("\n=== generate_design 前几行 ===")
lines3 = src3.split('\n')
for i, line in enumerate(lines3[:15]):
    print(f"  {line}")

# 检查模板路径
import store
import os

# 查看模板文件
templates = store.rows("SELECT * FROM template_files WHERE offering_id=20")
print("\n=== ID=20 模板文件 ===")
for t in templates:
    exists = os.path.exists(t['template_path']) if t.get('template_path') else False
    print(f"  {t['document_type']}: {t['template_path']} | 存在={exists}")

# 查看其他课程的模板
templates18 = store.rows("SELECT * FROM template_files WHERE offering_id=18")
print("\n=== ID=18 模板文件 ===")
for t in templates18:
    exists = os.path.exists(t['template_path']) if t.get('template_path') else False
    print(f"  {t['document_type']}: {t['template_path']} | 存在={exists}")
