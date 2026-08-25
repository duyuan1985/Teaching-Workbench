import json
import re
import subprocess
import tempfile
import zipfile
from copy import deepcopy
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

import store


ROOT = Path(__file__).parent
TEMPLATE_DIR = ROOT / "原始资料" / "模板" / "模板：听课记录表"
DOC_SCRIPT = ROOT / "generate_listening_doc.ps1"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
ET.register_namespace("", NS)


def templates():
    if not TEMPLATE_DIR.exists():
        return []
    return sorted(path for path in TEMPLATE_DIR.iterdir() if path.is_file() and path.suffix.lower() in {".xlsx", ".doc"})


def _task_for_session(offering, session):
    tasks = store.rows(
        """SELECT * FROM tasks WHERE offering_id=? ORDER BY
        CASE WHEN lesson_date<>'' AND lesson_date=? THEN 0 WHEN week_no IS NOT NULL AND week_no=? THEN 1 ELSE 2 END,seq""",
        (offering["id"], session.get("lesson_date", ""), session.get("week_no")),
    )
    if tasks:
        return tasks[0]
    units = store.rows("SELECT * FROM curriculum_units WHERE offering_id=? AND review_action<>'删除' ORDER BY seq", (offering["id"],))
    if not units:
        return {"chapter": offering["course_name"], "title": "课程项目任务", "knowledge_goal": "", "ability_goal": "", "ideological_goal": ""}
    week = int(session.get("week_no") or 1)
    index = min(len(units) - 1, max(0, (week - 1) * len(units) // 18))
    unit = units[index]
    return {
        "chapter": unit["project_title"], "title": unit["source_skills"],
        "knowledge_goal": unit.get("source_objectives", ""), "ability_goal": unit.get("source_skills", ""),
        "ideological_goal": "规范操作、诚实记录、团队协作和安全责任",
    }


def _weekday_label(value):
    mapping = {"周一": "星期一", "周二": "星期二", "周三": "星期三", "周四": "星期四", "周五": "星期五", "周六": "星期六", "周日": "星期日"}
    return mapping.get(value, value or "")


def _data(offering, session):
    task = _task_for_session(offering, session)
    chapter = str(task.get("chapter") or "").strip()
    task_title = str(task.get("title") or "").strip()
    title = task_title if chapter and task_title.startswith(chapter) else "：".join(part for part in (chapter, task_title) if part)
    title = title or offering["course_name"]
    enrollment = store.rows("SELECT enrollment_count FROM offering_classes WHERE offering_id=? AND class_name=?", (offering["id"], session.get("class_name", "")))
    count = int(enrollment[0]["enrollment_count"]) if enrollment else 0
    date_value = session.get("lesson_date", "")
    try:
        parsed_date = date.fromisoformat(date_value)
        date_label = f"{parsed_date.year}年{parsed_date.month}月{parsed_date.day}日"
    except ValueError:
        date_label = date_value or "待定"
    weekday = _weekday_label(session.get("weekday", ""))
    periods = session.get("periods", "")
    compact_date = f"{parsed_date.month}.{parsed_date.day}" if date_value and 'parsed_date' in locals() else date_label
    time_label = f"{compact_date} {periods}".strip()
    word_time_label = f"{date_label}{weekday} 第{periods}".strip()
    title_parts = [part.strip() for part in re.split(r"[：:、，,；;]", title) if part.strip()]
    planned_short = "、".join(title_parts[:4])
    if len(planned_short) > 32:
        planned_short = planned_short[:31].rstrip("、，：:；;") + "…"
    knowledge = str(task.get("knowledge_goal") or "").strip() or f"理解并掌握“{title}”的核心知识、技术步骤和质量要求。"
    ability = str(task.get("ability_goal") or "").strip() or f"能够按照任务要求完成“{title}”的分析、操作、检查和成果表达。"
    ideology = str(task.get("ideological_goal") or "").strip() or "教学中融入规范操作、诚实守信、安全责任、团队协作和精益求精要求。"
    record = (
        f"1. 课程导入：结合实际案例提出问题，明确“{title}”的任务与成果要求。\n"
        f"2. 理论讲解：讲解本节核心概念、技术方法、操作步骤和质量标准。\n"
        f"3. 实操训练：教师示范关键步骤，学生完成分步练习、任务实施和结果检查。\n"
        f"4. 讨论分享：展示代表性成果，分析典型问题、易错点及改进方法。\n"
        f"5. 课堂总结：归纳知识与操作流程，强调规范操作、安全责任和诚实记录。"
    )
    attendance_rows = store.rows(
        "SELECT status,COUNT(*) AS count FROM attendance_records ar JOIN students st ON st.id=ar.student_id "
        "WHERE ar.offering_id=? AND st.class_name=? AND ar.lesson_date=? GROUP BY status",
        (offering["id"], session.get("class_name", ""), session.get("lesson_date", "")),
    )
    attendance_counts = {row["status"]: int(row["count"]) for row in attendance_rows}
    actual = sum(attendance_counts.values()) if attendance_counts else ""
    late = attendance_counts.get("迟到", "") if attendance_rows else ""
    actual_label = str(actual) if attendance_rows else "  "
    late_label = str(late) if attendance_rows else "  "
    return {
        "teacher": store.get_setting("teacher_name", "杜媛"), "department": store.get_setting("department", "经济贸易系"),
        "course_name": offering["course_name"], "class_name": session.get("class_name", ""),
        "classroom": session.get("classroom", ""), "date_label": date_label, "time_label": time_label, "word_time_label": word_time_label,
        "planned": planned_short or title, "record": record,
        "attendance": f"应到{count if count else '  '}人，实到{actual_label}人，迟到{late_label}人",
        "remarks": "本节教学内容与授课计划和课程教学任务相对应。",
        "overall": f"教学准备充分，目标明确，围绕“{title}”组织理论讲解与实践训练；任务设计符合学生基础，示范步骤清晰，课堂组织有序，能够将知识学习、技能训练和职业素养培养有机结合。",
        "suggestion": "建议继续加强分层指导和学生成果展示，适当增加对典型错误的对比分析与即时反馈，并通过过程记录检验学生独立完成任务和解决问题的能力。",
    }


def _xlsx_set_cell(root, reference, value):
    row_number = int(re.search(r"\d+", reference).group())
    sheet_data = root.find(f"{{{NS}}}sheetData")
    row = next((item for item in sheet_data.findall(f"{{{NS}}}row") if int(item.get("r")) == row_number), None)
    if row is None:
        row = ET.SubElement(sheet_data, f"{{{NS}}}row", {"r": str(row_number)})
    cell = next((item for item in row.findall(f"{{{NS}}}c") if item.get("r") == reference), None)
    if cell is None:
        cell = ET.SubElement(row, f"{{{NS}}}c", {"r": reference})
    for child in list(cell):
        cell.remove(child)
    cell.set("t", "inlineStr")
    inline = ET.SubElement(cell, f"{{{NS}}}is")
    text = ET.SubElement(inline, f"{{{NS}}}t")
    text.set(f"{{{XML_NS}}}space", "preserve")
    text.text = str(value)


def _generate_xlsx(template_path, output_path, data):
    replacements = {
        "B2": data["course_name"], "D2": data["class_name"], "F2": data["time_label"],
        "B3": data["planned"], "B4": data["record"], "B5": data["remarks"],
        "B6": data["overall"], "B7": data["suggestion"],
    }
    with zipfile.ZipFile(template_path, "r") as source, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            payload = source.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                root = ET.fromstring(payload)
                for reference, value in replacements.items():
                    _xlsx_set_cell(root, reference, value)
                payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            target.writestr(deepcopy(item), payload)


def generate_listening_record(session_id, template_path):
    sessions = store.rows("SELECT * FROM sessions WHERE id=?", (session_id,))
    if not sessions:
        raise ValueError("所选课次不存在。")
    session = sessions[0]
    offerings = store.rows("SELECT * FROM offerings WHERE id=?", (session["offering_id"],))
    if not offerings:
        raise ValueError("课程实例不存在。")
    offering = offerings[0]
    if offering.get("offering_kind") == "实训课程":
        raise ValueError("实训课程不生成听课记录。")
    template = Path(template_path).resolve()
    allowed = {path.resolve() for path in templates()}
    if template not in allowed:
        raise ValueError("请选择听课记录模板目录中的有效模板。")
    data = _data(offering, session)
    output_dir = Path(store.get_setting("output_root", ROOT / "生成结果")) / offering["term"] / offering["course_name"] / "听课记录"
    output_dir.mkdir(parents=True, exist_ok=True)
    date_part = session.get("lesson_date") or f"第{session.get('week_no') or ''}周"
    output = output_dir / f"{date_part}《{offering['course_name']}》{session.get('class_name','')}听课记录{template.suffix.lower()}"
    if template.suffix.lower() == ".xlsx":
        _generate_xlsx(template, output, data)
    else:
        with tempfile.TemporaryDirectory(prefix="listening-record-") as temp_dir:
            data_path = Path(temp_dir) / "data.json"
            data_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8-sig")
            result = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(DOC_SCRIPT),
                 "-TemplatePath", str(template), "-OutputPath", str(output), "-DataPath", str(data_path)],
                capture_output=True, text=True, encoding="utf-8",
            )
            if result.returncode != 0:
                raise RuntimeError((result.stderr or result.stdout or "Word听课记录套版失败").strip())
    store.execute(
        "INSERT INTO listening_records(offering_id,session_id,template_path,output_path) VALUES (?,?,?,?)",
        (offering["id"], session_id, str(template), str(output)),
    )
    return output
