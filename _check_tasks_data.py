"""
从数据库获取任务数据，检查完整内容
"""
import store

# 获取任务数据
tasks = store.rows("SELECT * FROM tasks WHERE offering_id=20 ORDER BY seq")
print(f"任务数: {len(tasks)}")
if tasks:
    print(f"字段: {list(tasks[0].keys())}")
    for t in tasks[:3]:
        print(f"\n--- 任务 {t.get('seq')} ---")
        for k, v in t.items():
            if v and isinstance(v, str) and len(v) > 0:
                print(f"  {k}: {str(v)[:100]}")
