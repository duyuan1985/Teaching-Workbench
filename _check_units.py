import store

# 检查curriculum_units是否有章节信息
for oid in [18, 19, 20, 21, 22]:
    units = store.rows("SELECT * FROM curriculum_units WHERE offering_id=? ORDER BY id", [oid])
    offering = store.rows("SELECT course_name FROM offerings WHERE id=?", [oid])[0]
    print(f"\n=== ID={oid} {offering['course_name']}: {len(units)} units ===")
    for u in units[:10]:
        print(f"  unit_id={u['id']} name={str(u.get('name',''))[:40]} approval={u.get('approval_status','')} action={u.get('review_action','')}")
