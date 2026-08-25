"""
重新生成新模板课程的课程标准文档
用于修复 2024-2025-2、2025-2026 模板课程标准的生成错误
"""
import store
import generate
import importlib
import sys
import time

# 强制重新加载已修改的 gen_standard 模块
if 'gen_standard' in sys.modules:
    importlib.reload(sys.modules['gen_standard'])
if 'generate' in sys.modules:
    importlib.reload(sys.modules['generate'])

# 需要重新生成课程标准的课程（新模板，6个表格）
# 2024-2025-2: ID=24, 25, 26
# 2025-2026-1: ID=27, 28
# 2025-2026-2: ID=29, 30, 31, 32
# 2026-2027-1: ID=33, 34
new_template_ids = [24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34]

print("=== 重新生成新模板课程的课程标准 ===")
print(f"共 {len(new_template_ids)} 门课程需要处理")
print("=" * 60)
sys.stdout.flush()

for oid in new_template_ids:
    o = store.rows('SELECT * FROM offerings WHERE id=?', [oid])
    if not o:
        print(f"  ID={oid}: 未找到，跳过")
        continue
    o = o[0]
    print(f"\nID={oid} {o['course_name']} ({o['term']})")
    sys.stdout.flush()

    try:
        results = generate.generate_all(oid)
        for dt, r in results.items():
            if 'path' in r:
                print(f"  {dt}: OK - {r['path']}")
            elif 'error' in r:
                print(f"  {dt}: ERROR - {r['error']}")
    except Exception as e:
        print(f"  失败: {e}")

    sys.stdout.flush()

print("\n" + "=" * 60)
print("课程标准重新生成完成!")
