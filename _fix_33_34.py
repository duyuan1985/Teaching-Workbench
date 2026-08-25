import store

# ID=33: 全媒体258601有60学时，总学时64，差4学时
oid = 33
offering = store.rows("SELECT * FROM offerings WHERE id=?", [oid])[0]
class_name = '全媒体258601'

with store.connect() as db:
    db.execute("""
        INSERT INTO sessions (offering_id, week_no, lesson_date, hours, status, session_type, class_name)
        VALUES (?, 18, '', 4, '已确认', '补课', ?)
    """, [oid, class_name])
    db.commit()

# 验证
sessions = store.rows("SELECT * FROM sessions WHERE offering_id=? AND class_name=? AND status='已确认' AND session_type<>'停课'", [oid, class_name])
confirmed = sum(s['hours'] for s in sessions)
print(f"ID={oid} {offering['course_name']}: 班级={class_name} 已确认={confirmed} 总学时={offering['total_hours']} 差={offering['total_hours']-confirmed}")

# ID=34验证
oid = 34
offering = store.rows("SELECT * FROM offerings WHERE id=?", [oid])[0]
class_name = '农商255701'
sessions = store.rows("SELECT * FROM sessions WHERE offering_id=? AND class_name=? AND status='已确认' AND session_type<>'停课'", [oid, class_name])
confirmed = sum(s['hours'] for s in sessions)
print(f"ID={oid} {offering['course_name']}: 班级={class_name} 已确认={confirmed} 总学时={offering['total_hours']} 差={offering['total_hours']-confirmed}")
