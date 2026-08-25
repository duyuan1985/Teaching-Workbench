import store

# 取消农商255701的第17周记录（多出的4学时）
with store.connect() as db:
    db.execute("UPDATE sessions SET status='已取消' WHERE offering_id=28 AND class_name='农商255701' AND week_no=17 AND status='已确认' AND session_type='正常排课'")
    db.commit()

# 验证
sessions = store.rows("SELECT * FROM sessions WHERE offering_id=28 AND class_name=? AND status='已确认' AND session_type<>'停课'", ['农商255701'])
total = sum(s['hours'] for s in sessions)
print(f'农商255701: {len(sessions)}条, {total}学时 (总学时48, 差={48-total})')
