import store

# 检查id=20的sessions完整情况
sessions = store.rows("SELECT * FROM sessions WHERE offering_id=20 ORDER BY week_no, lesson_date")
print(f"=== ID=20 商务数据分析 sessions: {len(sessions)} ===")
for s in sessions:
    print(f"  week={s.get('week_no','')} date={s.get('lesson_date','')} hours={s.get('hours','')} status={s.get('status','')} type={s.get('session_type','')} class={s.get('class_name','')}")

# 检查是否有18周补课
print(f"\n=== 检查18周 ===")
for s in sessions:
    if s.get('week_no') == 18:
        print(f"  week=18 date={s.get('lesson_date','')} hours={s.get('hours','')} status={s.get('status','')} type={s.get('session_type','')}")

# 检查session_type字段的所有值
types = store.rows("SELECT session_type, COUNT(*) as cnt FROM sessions WHERE offering_id=20 GROUP BY session_type")
print(f"\nsession_type值: {[(t['session_type'], t['cnt']) for t in types]}")

# 检查其他课程的sessions是否需要补课
print("\n=== 各课程sessions统计 ===")
for oid in [18, 19, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34]:
    offering = store.rows("SELECT course_name, term, total_hours FROM offerings WHERE id=?", [oid])
    if not offering:
        continue
    o = offering[0]
    sessions = store.rows("SELECT status, session_type, COUNT(*) as cnt, SUM(hours) as total_hours FROM sessions WHERE offering_id=? GROUP BY status, session_type", [oid])
    confirmed_hours = 0
    for s in sessions:
        if s['status'] == '已确认' and s['session_type'] != '停课':
            confirmed_hours += s['total_hours'] or 0
    diff = o['total_hours'] - confirmed_hours
    status = '✅' if diff == 0 else f'❌ 差{diff}学时'
    print(f"  ID={oid} {o['course_name']} ({o['term']}): 总学时={o['total_hours']} 已确认={confirmed_hours} {status}")
