import store

# 查看各课程的curriculum_units详情
for oid in [18, 19, 21, 22]:
    offering = store.rows("SELECT course_name, term, total_hours FROM offerings WHERE id=?", [oid])[0]
    units = store.rows("SELECT * FROM curriculum_units WHERE offering_id=? ORDER BY seq, id", [oid])
    print(f"\n=== ID={oid} {offering['course_name']} ({offering['term']}) 总学时={offering['total_hours']} ===")
    print(f"units: {len(units)}个")
    for u in units:
        print(f"  seq={u.get('seq','?')} id={u['id']} | project={str(u.get('project_title',''))[:30]} | hours={u.get('suggested_hours','?')} | approval={u.get('approval_status','')} | skills={str(u.get('source_skills',''))[:50]}")
    
    # 检查sessions状态
    sessions = store.rows("SELECT status, COUNT(*) as cnt FROM sessions WHERE offering_id=? GROUP BY status", [oid])
    print(f"  sessions状态: {[(s['status'], s['cnt']) for s in sessions]}")

# 对比id=20
print(f"\n=== ID=20 商务数据分析 (参考) ===")
units20 = store.rows("SELECT * FROM curriculum_units WHERE offering_id=20 ORDER BY seq", [20])
for u in units20:
    print(f"  seq={u.get('seq','?')} | project={u.get('project_title','')} | hours={u.get('suggested_hours','?')} | approval={u.get('approval_status','')}")
sessions20 = store.rows("SELECT status, COUNT(*) as cnt FROM sessions WHERE offering_id=20 GROUP BY status", [20])
print(f"  sessions状态: {[(s['status'], s['cnt']) for s in sessions20]}")
