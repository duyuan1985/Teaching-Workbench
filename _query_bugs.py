import sqlite3, json

conn = sqlite3.connect(r'e:\开发\AIGC\教学档案工作台\data\workbench.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

# 1. Get template_file_id for offering 20
rows = c.execute(
    "SELECT id, offering_id, document_type FROM template_files WHERE offering_id=20 AND document_type='课程标准'"
).fetchall()
print("=== Template files for offering 20 ===")
for r in rows:
    print(f"  template_file id={r['id']}, offering={r['offering_id']}, type={r['document_type']}")
    tfid = r['id']

    # 2. Table 0 slots (goals table)
    slots = c.execute(
        "SELECT slot_key, field_name, locator, content_kind FROM template_slots WHERE template_file_id=? AND slot_key LIKE 'table:0%' ORDER BY slot_key",
        (tfid,)
    ).fetchall()
    print(f"\n--- Table 0 slots ({len(slots)}) ---")
    for s in slots:
        print(f"  slot_key={s['slot_key']}, field={s['field_name']}, locator={s['locator']}, kind={s['content_kind']}")

    # 3. Table 1 slots (content/hours table)
    slots1 = c.execute(
        "SELECT slot_key, field_name, locator, content_kind FROM template_slots WHERE template_file_id=? AND slot_key LIKE 'table:1%' ORDER BY slot_key",
        (tfid,)
    ).fetchall()
    print(f"\n--- Table 1 slots ({len(slots1)}) ---")
    for s in slots1:
        print(f"  slot_key={s['slot_key']}, field={s['field_name']}, locator={s['locator']}, kind={s['content_kind']}")

    # 4. Table 2 slots (assessment table)
    slots2 = c.execute(
        "SELECT slot_key, field_name, locator, content_kind FROM template_slots WHERE template_file_id=? AND slot_key LIKE 'table:2%' ORDER BY slot_key",
        (tfid,)
    ).fetchall()
    print(f"\n--- Table 2 slots ({len(slots2)}) ---")
    for s in slots2:
        print(f"  slot_key={s['slot_key']}, field={s['field_name']}, locator={s['locator']}, kind={s['content_kind']}")

    # 5. Paragraph slots (especially around paragraph 39 / assessment)
    pslots = c.execute(
        "SELECT slot_key, field_name, locator, content_kind FROM template_slots WHERE template_file_id=? AND slot_key LIKE 'paragraph:%' ORDER BY CAST(slot_key.split(':')[1] AS INTEGER)",
        (tfid,)
    ).fetchall()
    # fallback if above fails
    if not pslots:
        pslots = c.execute(
            "SELECT slot_key, field_name, locator, content_kind FROM template_slots WHERE template_file_id=? AND slot_key LIKE 'paragraph:%' ORDER BY slot_key",
            (tfid,)
        ).fetchall()
    print(f"\n--- Paragraph slots ({len(pslots)}) ---")
    for s in pslots:
        print(f"  slot_key={s['slot_key']}, field={s['field_name']}, locator={s['locator']}, kind={s['content_kind']}")

# 6. Authored sections for offering 20, assessment
print("\n=== Authored sections for offering 20 (assessment) ===")
asections = c.execute(
    "SELECT section_key, repeat_key, content_json FROM authored_sections WHERE offering_id=20 AND document_type='课程标准' AND (section_key LIKE '%assessment%' OR section_key LIKE '%考核%' OR section_key LIKE '%评价%')"
).fetchall()
for a in asections:
    print(f"  section_key={a['section_key']}, repeat_key={a['repeat_key']}")
    content = json.loads(a['content_json'])
    print(f"  content_json={json.dumps(content, ensure_ascii=False, indent=2)[:2000]}")

# 7. All authored sections keys
print("\n=== All authored sections for offering 20 ===")
allsec = c.execute(
    "SELECT section_key, repeat_key FROM authored_sections WHERE offering_id=20 AND document_type='课程标准' ORDER BY section_key",
).fetchall()
for a in allsec:
    print(f"  section_key={a['section_key']}, repeat_key={a['repeat_key']}")

# 8. Curriculum units for offering 20
print("\n=== Curriculum units for offering 20 ===")
units = c.execute(
    "SELECT id, seq, project_title, suggested_hours, source_skills FROM curriculum_units WHERE offering_id=20 AND approval_status='已确认' AND review_action<>'删除' ORDER BY seq"
).fetchall()
for u in units:
    print(f"  seq={u['seq']}, title={u['project_title']}, hours={u['suggested_hours']}, skills={str(u['source_skills'])[:100]}")

conn.close()
