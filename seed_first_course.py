from pathlib import Path

import importers
import store


ROOT = Path(r'E:\工作\05-教学辅助\04-模板库\教学档案资料模板、标准')
store.initialize()
existing = store.rows(
    "SELECT * FROM offerings WHERE course_name=? AND term=?",
    ("H5设计与制作", "2023-2024-2"),
)
if existing:
    offering_id = existing[0]["id"]
else:
    offering_id = store.create_offering({
        "course_name": "H5设计与制作",
        "term": "2023-2024-2",
        "major": "电子商务",
        "course_code": "A400118",
        "course_type": "专业核心课",
        "credits": 3,
        "total_hours": 60,
        "weekly_hours": 4,
        "template_version": "课程标准2023-2024＋教学设计2023-2024＋通用授课计划",
        "textbook_version": "《HTML5与CSS3项目实战》（南开大学出版社）",
        "textbook_path": r"E:\工作\02-课程资源\H5设计与制作\HTML5与CSS3项目实战",
        "template_path": str(ROOT),
        "schedule_path": r"E:\工作\90-历史学期归档\2023-2024-2\课表与进程表\2023-2024-2进程表（经贸系）.xlsx",
        "notes": "2022电商教学班；教学安排表名义排课56学时，规定60学时，需结合校历核实停课并补足。",
    })

templates = [
    ("课程标准", "模板3：课程标准 模板（2023-2024）", ROOT / "模板3：课程标准、岗位实习标准" / "模板3：课程标准 模板（2023-2024）.docx", 1, "{学期}《{课程}》课程标准 {教师}.docx"),
    ("授课计划", "模板4：授课计划", ROOT / "模板4：授课计划 模板.docx", 1, "{学期}《{课程}》授课计划 {教师}.docx"),
    ("教学设计", "模板5：教学设计 模板（2023-2024）", ROOT / "模板5：教学设计" / "模板5：教学设计 模板（2023-2024）.docx", 1, "{学期}《{课程}》教学设计 {教师}.docx"),
    ("实训资料", "模板8：《XXX》实训资料", ROOT / "模板8：《XXX》实训资料.docx", 0, "{学期}《{课程}》实训资料 {教师}.docx"),
    ("成绩分析", "模板6：成绩分析", ROOT / "模板6：成绩分析.doc", 0, "{学期}《{课程}》成绩分析 {教师}.doc"),
]
for document_type, name, path, required, pattern in templates:
    if not store.rows(
        "SELECT id FROM template_files WHERE offering_id=? AND document_type=? AND template_path=?",
        (offering_id, document_type, str(path)),
    ):
        store.create_template_file(offering_id, {
            "document_type": document_type,
            "template_name": name,
            "template_path": str(path),
            "required": str(required),
            "output_name_pattern": pattern,
            "notes": "本学期对应官方模板" if required else "按本学期归档要求选用",
        })

sources = [
    ("教学安排表", r"E:\开发\AIGC\教学安排表20260816115742.xlsx", 1, "教务系统导出的杜媛历学期教学安排"),
    ("学期进程表", r"E:\工作\90-历史学期归档\2023-2024-2\课表与进程表\2023-2024-2进程表（经贸系）.xlsx", 1, "2023-2024-2班级教学进程"),
    ("教材目录", r"E:\工作\02-课程资源\H5设计与制作\HTML5与CSS3项目实战", 1, "本学期教材、PPT和实训资源"),
]
for source_type, source_path, required, notes in sources:
    if not store.rows(
        "SELECT id FROM source_files WHERE offering_id=? AND source_type=? AND source_path=?",
        (offering_id, source_type, source_path),
    ):
        store.create_source_file(offering_id, {
            "source_type": source_type,
            "source_path": source_path,
            "required": str(required),
            "notes": notes,
        })

offering = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))[0]
arrangement = importers.import_teaching_arrangement(
    offering,
    r"E:\开发\AIGC\教学安排表20260816115742.xlsx",
)
progress = importers.import_progress_table(offering, offering["schedule_path"])
print("offering", offering_id, "arrangement", arrangement, "progress", progress)
