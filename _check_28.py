import store

sessions = store.rows("SELECT * FROM sessions WHERE offering_id=28 AND class_name=? AND status='已确认' AND session_type<>'停课' ORDER BY week_no", ['农商255701'])
print(f'农商255701: {len(sessions)}条')
for s in sessions:
    print(f'  week={s["week_no"]} date={s.get("lesson_date","")} hours={s["hours"]} type={s.get("session_type","")}')
total = sum(s['hours'] for s in sessions)
print(f'总学时: {total} (课程总学时48, 差={48-total})')
