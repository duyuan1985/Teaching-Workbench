import sqlite3, json

conn = sqlite3.connect(r'e:\开发\AIGC\教学档案工作台\data\workbench.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

tfid = 9  # template_file_id for offering 20

# Check template_slots columns
cols = c.execute("PRAGMA table_info(template_slots)").fetchall()
print("=== template_slots columns ===")
for col in cols:
    print(f"  {col['name']} ({col['type']})")

# All slots for this template
all_slots = c.execute(
    "SELECT * FROM template_slots WHERE template_file_id=? ORDER BY slot_key",
    (tfid,)
).fetchall()
print(f"\n=== All slots ({len(all_slots)}) ===")
for s in all_slots:
    print(f"  slot_key={s['slot_key']}")
    for k in s.keys():
        if k != 'slot_key' and s[k] is not None:
            val = str(s[k])
            if len(val) > 200:
                val = val[:200] + "..."
            print(f"    {k}={val}")

# Authored sections for offering 20
print("\n=== Authored sections for offering 20 ===")
allsec = c.execute(
    "SELECT section_key, repeat_key, content_json FROM authored_sections WHERE offering_id=20 AND document_type='课程标准' ORDER BY section_key, repeat_key",
).fetchall()
for a in allsec:
    print(f"\n  section_key={a['section_key']}, repeat_key={a['repeat_key']}")
    content = json.loads(a['content_json'])
    content_str = json.dumps(content, ensure_ascii=False, indent=2)
    if len(content_str) > 1500:
        print(f"  content (truncated):\n{content_str[:1500]}...")
    else:
        print(f"  content:\n{content_str}")

# Curriculum units for offering 20
print("\n=== Curriculum units for offering 20 ===")
units = c.execute(
    "SELECT id, seq, project_title, suggested_hours, source_skills FROM curriculum_units WHERE offering_id=20 AND approval_status='已确认' AND review_action<>'删除' ORDER BY seq"
).fetchall()
for u in units:
    print(f"  seq={u['seq']}, title={u['project_title']}, hours={u['suggested_hours']}")
    print(f"    skills={str(u['source_skills'])[:150]}")

# Contract JSON for table structure
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
        print(f"\n  Table index={t.get('index')}, role={t.get('role')}")
        print(f"    rows={t.get('row_count')}, cols={t.get('col_count')}")
        # Print merges if available
        for k in ('merges', 'merged_cells', 'cell_merges', 'merge_info'):
            if k in t:
                print(f"    {k}: {json.dumps(t[k], ensure_ascii=False)[:400]}")
        # Print all keys for this table
        print(f"    keys: {list(t.keys())}")

conn.close()
