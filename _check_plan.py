"""
1. 查看授课计划中的章节信息
2. 修复所有表格R9/R10标题行
3. 修复R3章节名称
"""
import store

# 查看数据库中的任务信息
tasks = store.rows("SELECT seq, title FROM tasks WHERE offering_id=20 ORDER BY seq")
print("=== 数据库中的任务 ===")
for t in tasks:
    print(f"  seq={t['seq']}: {t['title']}")
