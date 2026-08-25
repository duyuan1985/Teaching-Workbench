"""
1. 确认所有curriculum_units状态为"已确认"
2. 调用build_tasks()创建教学任务
"""
import store
import task_builder

# 获取所有需要处理的课程（排除实训课程和已完成的id=20）
offerings = store.rows("SELECT * FROM offerings WHERE offering_kind != '实训课程' AND id != 20 ORDER BY id")
print(f"共 {len(offerings)} 个课程需要处理\n")

# 步骤1: 确认所有curriculum_units
print("=== 步骤1: 确认curriculum_units ===")
for o in offerings:
    oid = o['id']
    units = store.rows("SELECT * FROM curriculum_units WHERE offering_id=? AND approval_status='待确认'", [oid])
    if units:
        with store.connect() as db:
            db.execute("UPDATE curriculum_units SET approval_status='已确认' WHERE offering_id=? AND approval_status='待确认'", [oid])
            db.commit()
        print(f"  ID={oid} {o['course_name']} ({o['term']}): 确认了{len(units)}个units")
    else:
        print(f"  ID={oid} {o['course_name']} ({o['term']}): 无待确认units")

# 步骤2: 调用build_tasks()创建教学任务
print("\n=== 步骤2: 创建教学任务 ===")
for o in offerings:
    oid = o['id']
    try:
        existing = store.rows("SELECT COUNT(*) as cnt FROM tasks WHERE offering_id=?", [oid])
        if existing[0]['cnt'] > 0:
            print(f"  ID={oid} {o['course_name']} ({o['term']}): 已有{existing[0]['cnt']}个任务，跳过")
            continue
        
        task_count = task_builder.build_tasks(o)
        print(f"  ID={oid} {o['course_name']} ({o['term']}): 创建了{task_count}个任务 ✅")
    except Exception as e:
        print(f"  ID={oid} {o['course_name']} ({o['term']}): 失败 - {e}")

# 最终统计
print("\n=== 最终统计 ===")
all_tasks = store.rows("SELECT offering_id, COUNT(*) as cnt FROM tasks GROUP BY offering_id ORDER BY offering_id")
for t in all_tasks:
    offering = store.rows("SELECT course_name, term FROM offerings WHERE id=?", [t['offering_id']])[0]
    print(f"  ID={t['offering_id']} {offering['course_name']} ({offering['term']}): {t['cnt']}个任务")
