import store

# 查找所有多班课程
offerings = store.rows("SELECT * FROM offerings WHERE offering_kind != '实训课程' ORDER BY id")
print("=== 多班课程 ===")
for o in offerings:
    tc = o.get('teaching_class', '') or ''
    classes = [c.strip() for c in tc.replace('；', ';').split(';') if c.strip()]
    if len(classes) > 1:
        print(f"\nID={o['id']} {o['course_name']} ({o['term']})")
        print(f"  班级: {classes}")
        
        # 检查各班级的sessions
        for cn in classes:
            sessions = store.rows("SELECT COUNT(*) as cnt, SUM(hours) as h, status FROM sessions WHERE offering_id=? AND class_name=? GROUP BY status", [o['id'], cn])
            for s in sessions:
                print(f"    {cn}: {s['cnt']}条 {s['h']}学时 状态={s['status']}")
