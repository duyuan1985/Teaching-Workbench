import store

# 检查ID=33和34的sessions是否有class_name
for oid in [33, 34]:
    offering = store.rows("SELECT * FROM offerings WHERE id=?", [oid])[0]
    sessions = store.rows("SELECT * FROM sessions WHERE offering_id=? ORDER BY week_no, id", [oid])
    
    print(f"\n=== ID={oid} {offering['course_name']} ({offering['term']}) ===")
    print(f"  teaching_class: {offering.get('teaching_class','')}")
    print(f"  total_hours: {offering['total_hours']}")
    print(f"  sessions: {len(sessions)}")
    
    # 按class_name统计
    by_class = {}
    for s in sessions:
        cn = s.get('class_name', '') or '(空)'
        if cn not in by_class:
            by_class[cn] = {'count': 0, 'hours': 0}
        by_class[cn]['count'] += 1
        by_class[cn]['hours'] += s.get('hours', 0)
    
    print(f"  按班级统计:")
    for cn, stats in by_class.items():
        print(f"    {cn}: {stats['count']}条, {stats['hours']}学时")
    
    # 显示前10条
    print(f"  前10条:")
    for s in sessions[:10]:
        print(f"    week={s.get('week_no','')} date={s.get('lesson_date','')} hours={s.get('hours','')} status={s.get('status','')} type={s.get('session_type','')} class={s.get('class_name','')}")
