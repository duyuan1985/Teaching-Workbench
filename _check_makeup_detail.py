import store

# 8个需要补课的课程
need_makeup = [18, 19, 22, 25, 27, 32, 33, 34]

for oid in need_makeup:
    offering = store.rows("SELECT * FROM offerings WHERE id=?", [oid])[0]
    sessions = store.rows("SELECT * FROM sessions WHERE offering_id=? ORDER BY week_no, lesson_date, id", [oid])
    
    # 统计已确认的非停课学时
    confirmed_hours = sum(s['hours'] for s in sessions if s['status'] == '已确认' and s.get('session_type') != '停课')
    deficit = offering['total_hours'] - confirmed_hours
    
    # 找最大周次
    max_week = max((s.get('week_no') or 0) for s in sessions) if sessions else 0
    
    # 检查已有补课
    makeup = [s for s in sessions if s.get('session_type') == '补课']
    
    print(f"\n=== ID={oid} {offering['course_name']} ({offering['term']}) ===")
    print(f"  总学时={offering['total_hours']} 已确认={confirmed_hours} 差={deficit} 最大周={max_week}")
    print(f"  已有补课: {len(makeup)}条")
    print(f"  sessions总数: {len(sessions)}")
    
    # 显示最后几条sessions
    if sessions:
        print(f"  最后3条:")
        for s in sessions[-3:]:
            print(f"    week={s.get('week_no','')} date={s.get('lesson_date','')} hours={s.get('hours','')} status={s.get('status','')} type={s.get('session_type','')}")
    
    # 显示sessions状态分布
    statuses = {}
    for s in sessions:
        key = f"{s['status']}/{s.get('session_type','')}"
        statuses[key] = statuses.get(key, 0) + 1
    print(f"  状态分布: {statuses}")
