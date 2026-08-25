import sqlite3, json

conn = sqlite3.connect(r'e:\开发\AIGC\教学档案工作台\data\workbench.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

tfid = 9

# Full contract JSON
contracts = c.execute(
    "SELECT contract_json FROM template_analyses WHERE template_file_id=?",
    (tfid,)
).fetchall()
for ct in contracts:
    contract = json.loads(ct['contract_json'])
    tables = contract.get('tables', [])
    for t in tables:
        print(f"\n=== Table index={t.get('index')}, role={t.get('role')} ===")
        print(f"  rows: {json.dumps(t.get('rows'), ensure_ascii=False)[:500]}")
        print(f"  columns: {json.dumps(t.get('columns'), ensure_ascii=False)[:500]}")
        print(f"  header_text: {json.dumps(t.get('header_text'), ensure_ascii=False)[:500]}")
        print(f"  merged_cells: {json.dumps(t.get('merged_cells'), ensure_ascii=False)}")
        print(f"  repeat_mode: {t.get('repeat_mode')}")

# Also get the offering info
print("\n=== Offering 20 info ===")
offering = c.execute("SELECT * FROM offerings WHERE id=20").fetchone()
if offering:
    for k in offering.keys():
        print(f"  {k}={offering[k]}")

# Get scheme for offering 20
print("\n=== Assessment scheme ===")
try:
    from assessment_scheme import get_scheme, component_text
    scheme = get_scheme(20)
    print(f"  scheme: {json.dumps(scheme, ensure_ascii=False, indent=2)}")
    comp_text = component_text(20)
    print(f"  component_text: {comp_text}")
except Exception as e:
    print(f"  Error: {e}")

conn.close()
