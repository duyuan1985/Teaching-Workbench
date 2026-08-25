"""
1. 确认所有待确认sessions
2. 为学时不足的班级添加18周补课
3. 处理多班课程的学时问题
"""
import store

def get_confirmed_hours_by_class(oid, class_name):
    sessions = store.rows("SELECT * FROM sessions WHERE offering_id=? AND class_name=? AND status='已确认' AND session_type<>'停课'", [oid, class_name])
    return sum(s['hours'] for s in sessions)

# 所有多班课程
multi_class_oids = [21, 23, 24, 26, 28, 29, 30, 31, 33, 34]

for oid in multi_class_oids:
    offering = store.rows("SELECT * FROM offerings WHERE id=?", [oid])[0]
    tc = offering.get('teaching_class', '') or ''
    classes = [c.strip() for c in tc.replace('；', ';').split(';') if c.strip()]
    total_hours = offering['total_hours']
    
    print(f"\n=== ID={oid} {offering['course_name']} ({offering['term']}) 总学时={total_hours} ===")
    
    # 步骤1: 确认所有待确认sessions
    pending = store.rows("SELECT * FROM sessions WHERE offering_id=? AND status='待确认'", [oid])
    if pending:
        with store.connect() as db:
            db.execute("UPDATE sessions SET status='已确认' WHERE offering_id=? AND status='待确认'", [oid])
            db.commit()
        print(f"  确认了{len(pending)}条待确认sessions")
    
    # 步骤2: 检查每个班级的学时
    for cn in classes:
        confirmed = get_confirmed_hours_by_class(oid, cn)
        deficit = total_hours - confirmed
        if deficit > 0:
            # 添加18周补课
            with store.connect() as db:
                remaining = deficit
                while remaining > 0:
                    h = min(4, remaining)
                    if remaining - h < 2 and h > 2:
                        h = remaining
                    db.execute("""
                        INSERT INTO sessions (offering_id, week_no, lesson_date, hours, status, session_type, class_name)
                        VALUES (?, 18, '', ?, '已确认', '补课', ?)
                    """, [oid, h, cn])
                    remaining -= h
                db.commit()
            print(f"  {cn}: {confirmed}→{confirmed+deficit} (补{deficit}学时)")
        elif deficit < 0:
            print(f"  {cn}: {confirmed}学时 (超出{-deficit}学时，需检查)")
        else:
            print(f"  {cn}: {confirmed}学时 ✅")

print("\n完成!")
