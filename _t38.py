import store

store.initialize()
with store.connect() as db:
    cur = db.execute("DELETE FROM offerings WHERE term LIKE '%-一' OR term LIKE '%-二'")
    db.commit()
    print(f"删除错误term记录 {cur.rowcount} 条")

print("\n=== 开课记录 ===")
for r in store.rows(
    "SELECT id,term,course_name,major,teaching_class,credits,total_hours,teacher_name,textbook_version,assessment_type,course_nature FROM offerings ORDER BY term,course_name"
):
    print(f"[{r['id']}] {r['term']} {r['course_name']} ({r['major']}) 班:{r['teaching_class'][:20]} "
          f"学分:{r['credits']} 学时:{r['total_hours']} 师:{r['teacher_name']} 考核:{r['assessment_type']} 类别:{r['course_nature']}")
    print(f"      教材: {r['textbook_version']}")

print("\n=== 排课数 ===")
for r in store.rows(
    "SELECT o.term,o.course_name,COUNT(*) n,SUM(CASE WHEN s.lesson_date<>'' THEN 1 ELSE 0 END) dated "
    "FROM offerings o JOIN sessions s ON s.offering_id=o.id GROUP BY o.id ORDER BY o.term"
):
    print(f"{r['term']} {r['course_name']}: {r['n']} 次课（{r['dated']} 次有日期）")
