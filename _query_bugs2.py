import sqlite3, json

conn = sqlite3.connect(r'e:\开发\AIGC\教学档案工作台\data\workbench.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

tfid = 9  # template_file_id for offering 20

# 1. Paragraph slots (especially around paragraph 39 / assessment)
pslots = c.execute(
    "SELECT slot_key, field_name, locator, content_kind FROM template_slots WHERE template_file_id=? AND slot_key LIKE 'paragraph:%' ORDER BY slot_key",
    (tfid,)
).fetchall()
print(f"--- Paragraph slots ({len(pslots)}) ---")
for s in pslots:
    print(f"  slot_key={s['slot_key']}, field={s['field_name']}, locator={s['locator']}, kind={s['content_kind']}")

# 2. Section slots
sslots = c.execute(
    "SELECT slot_key, field_name, locator, content_kind, structure_json FROM template_slots WHERE template_file_id=? AND slot_key LIKE 'section:%' ORDER BY slot_key",
    (tfid,)
).fetchall()
print(f"\n--- Section slots ({len(sslots)}) ---")
for s in sslots:
    print(f"  slot_key={s['slot_key']}, field={s['field_name']}, locator={s['locator']}, kind={s['content_kind']}")
    if s['structure_json']:
        print(f"    structure={s['structure_json'][:200]}")

# 3. Authored sections for offering 20
print("\n=== All authored sections for offering 20 ===")
allsec = c.execute(
    "SELECT section_key, repeat_key, content_json FROM authored_sections WHERE offering_id=20 AND document_type='课程标准' ORDER BY section_key, repeat_key",
).fetchall()
for a in allsec:
    print(f"  section_key={a['section_key']}, repeat_key={a['repeat_key']}")
    content = json.loads(a['content_json'])
    content_str = json.dumps(content, ensure_ascii=False)
    if len(content_str) > 500:
        print(f"    content (truncated): {content_str[:500]}...")
    else:
        print(f"    content: {content_str}")

# 4. Specifically the assessment section
print("\n=== Assessment authored section ===")
asections = c.execute(
    "SELECT section_key, repeat_key, content_json FROM authored_sections WHERE offering_id=20 AND document_type='课程标准' AND (section_key LIKE '%assessment%' OR section_key LIKE '%考核%' OR section_key LIKE '%评价%')",
).fetchall()
for a in asections:
    print(f"  section_key={a['section_key']}, repeat_key={a['repeat_key']}")
    content = json.loads(a['content_json'])
    print(f"  content_json={json.dumps(content, ensure_ascii=False, indent=2)}")

# 5. Curriculum units for offering 20
print("\n=== Curriculum units for offering 20 ===")
units = c.execute(
    "SELECT id, seq, project_title, suggested_hours, source_skills FROM curriculum_units WHERE offering_id=20 AND approval_status='已确认' AND review_action<>'删除' ORDER BY seq"
).fetchall()
for u in units:
    print(f"  seq={u['seq']}, title={u['project_title']}, hours={u['suggested_hours']}")
    print(f"    skills={str(u['source_skills'])[:120]}")

# 6. Contract JSON for template analysis (table structure)
print("\n=== Template analysis contract ===")
contracts = c.execute(
    "SELECT contract_json FROM template_analyses WHERE template_file_id=?",
    (tfid,)
).fetchall()
for ct in contracts:
    contract = json.loads(ct['contract_json'])
    tables = contract.get('tables', [])
    print(f"  Tables in contract: {len(tables)}")
    for t in tables[:6]:
        print(f"    index={t.get('index')}, role={t.get('role')}, rows={t.get('row_count')}, cols={t.get('col_count')}")
        # Print merge info if available
        merges = t.get('merges', t.get('merged_cells', []))
        if merges:
            print(f"    merges: {json.dumps(merges, ensure_ascii=False)[:300]}")
    # Print full contract for first few tables
    print(f"\n  Full contract tables (first 3):")
    for t in tables[:3]:
        print(f"    {json.dumps(t, ensure_ascii=False)[:500]}")

conn.close()
