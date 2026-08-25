import store

# 检查一个课程（ID=18 H5设计与制作）的sessions和教材信息
oid = 18

sessions = store.rows("SELECT * FROM sessions WHERE offering_id=? ORDER BY week_no, lesson_date", [oid])
print(f'=== Offering {oid} sessions: {len(sessions)} ===')
for s in sessions[:5]:
    print(f"  week={s['week_no']} date={s.get('lesson_date','')} content={str(s.get('content',''))[:60]}")

# 检查教材内容
offering = store.rows("SELECT * FROM offerings WHERE id=?", [oid])[0]
print(f"\n教材: {offering.get('textbook_version','')}")
print(f"教材路径: {offering.get('textbook_path','')}")

# 检查教材目录
import os
tb_path = offering.get('textbook_path', '')
if tb_path and os.path.exists(tb_path):
    files = os.listdir(tb_path)
    print(f"\n教材目录({len(files)}个文件):")
    for f in sorted(files)[:15]:
        print(f"  {f}")

# 检查ID=20的tasks作为对比
tasks20 = store.rows("SELECT seq, title, hours FROM tasks WHERE offering_id=20 ORDER BY seq LIMIT 5")
print(f"\n=== Offering 20 tasks (参考) ===")
for t in tasks20:
    print(f"  seq={t['seq']}: {t['title']} ({t['hours']}h)")
