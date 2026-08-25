import store
units = store.rows("SELECT * FROM curriculum_units WHERE offering_id=20 ORDER BY seq")
for u in units:
    skills = u.get("source_skills", "")
    sks = [s.strip() for s in skills.split("；") if s.strip()] if skills else []
    pt = u["project_title"]
    print(f"{pt}: {len(sks)}个子任务")
    for s in sks:
        print(f"  - {s}")
