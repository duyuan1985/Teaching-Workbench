from pathlib import Path

import store


store.initialize()
rows = store.rows(
    "SELECT * FROM offerings WHERE course_name=? AND term=?",
    ("H5设计与制作", "2023-2024-2"),
)
if not rows:
    raise SystemExit("未找到第一个课程实例。")
offering_id = rows[0]["id"]
root = Path(r'E:\工作\05-教学辅助\04-模板库\教学档案资料模板、标准')
with store.connect() as db:
    db.execute(
        "UPDATE offerings SET template_version=? WHERE id=?",
        ("课程标准2023-2024＋教学设计2023-2024＋通用授课计划", offering_id),
    )
    db.execute(
        "DELETE FROM template_files WHERE offering_id=? AND document_type IN ('课程标准','教学设计')",
        (offering_id,),
    )
    db.execute(
        """INSERT INTO template_files
        (offering_id,document_type,template_name,template_path,required,output_name_pattern,notes)
        VALUES (?,?,?,?,?,?,?)""",
        (
            offering_id, "课程标准", "模板3：课程标准 模板（2023-2024）",
            str(root / "模板3：课程标准、岗位实习标准" / "模板3：课程标准 模板（2023-2024）.docx"),
            1, "{学期}《{课程}》课程标准 {教师}.docx", "本学期对应官方模板",
        ),
    )
    db.execute(
        """INSERT INTO template_files
        (offering_id,document_type,template_name,template_path,required,output_name_pattern,notes)
        VALUES (?,?,?,?,?,?,?)""",
        (
            offering_id, "教学设计", "模板5：教学设计 模板（2023-2024）",
            str(root / "模板5：教学设计" / "模板5：教学设计 模板（2023-2024）.docx"),
            1, "{学期}《{课程}》教学设计 {教师}.docx", "本学期对应官方模板",
        ),
    )
    db.commit()
print("fixed", offering_id)
