import store
tasks = store.rows("SELECT seq,title,chapter FROM tasks WHERE offering_id=20 ORDER BY seq")
for t in tasks:
    print(f'{t["seq"]}. [{t["chapter"]}] {t["title"]}')
