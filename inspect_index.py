import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, r"E:\开发\AIGC\教学档案工作台")

import store

for item in store.rows(
    "SELECT title,file_path,content_excerpt FROM resource_items WHERE offering_id=1 AND resource_type=? ORDER BY file_path",
    ("PPT课件",),
):
    print("TITLE", item["title"])
    print("PATH", item["file_path"])
    print(item["content_excerpt"][:1200])
    print()
