import json
from collections import Counter

import store


GENERIC_PHRASES = ("相关岗位", "相关任务", "本领域主流技术", "现行职业规范", "等等", "……", "待填写", "的作用的作用", "能够完成能够", "企业级卓越人才培养（信息类专业集群）")

REQUIRED_TEMPLATE_FIELDS = {
    "课程标准": {"课程基本信息", "课程类型", "学时学分", "开设学期", "课程性质", "知识目标", "能力目标", "思政目标", "素质目标", "课程设计总体思路", "课程内容与学时", "考核评价", "教师知识能力要求", "教学资源"},
    "授课计划": {"授课班级", "授课日期", "教学环境", "授课教师", "教学任务"},
    "教学设计": {"授课班级", "授课日期", "教学环境", "授课教师", "教学任务", "知识目标", "能力目标", "素质目标", "教材学情分析及教育理念", "教学场景设计", "教学资源", "教学活动流程", "教法学法", "达成目标", "教学时间", "课堂小结", "课后作业", "教学反思"},
}

# 模板结构变更时，部分必需栏目可能被拆分或合并为其他栏目，这些等价栏目可替代必需栏目
FIELD_ALIASES = {
    "学时学分": {"总学时", "学分"},
    "思政目标": {"素质目标"},
}


def run_foundation_quality_check(offering_id):
    models = store.rows("SELECT model_json FROM course_content_models WHERE offering_id=?", (offering_id,))
    if not models:
        raise ValueError("尚未生成课程语义模型。")
    model = json.loads(models[0]["model_json"])
    if not isinstance(model, dict):
        raise ValueError(f"课程语义模型JSON不是对象: {type(model).__name__}")
    issues = []
    required = ("identity", "projects", "knowledge_system", "ability_outcomes", "tools_technology", "work_process")
    for field in required:
        if not model.get(field):
            issues.append(("语义模型", "错误", "MISSING_FIELD", field, f"课程语义模型缺少{field}。"))
    for project in model.get("projects", []):
        if not project.get("resources") and project.get("title") != "综合评价与课程总结":
            issues.append(("资源映射", "提醒", "PROJECT_WITHOUT_RESOURCE", project.get("title", ""), "项目尚未关联到PPT、源码或实训文档。"))
        if not project.get("knowledge_skills"):
            issues.append(("资源映射", "错误", "PROJECT_WITHOUT_SKILL", project.get("title", ""), "项目没有提取出知识技能。"))
    tasks = store.rows("SELECT * FROM tasks WHERE offering_id=? ORDER BY seq", (offering_id,))
    for field in ("knowledge_goal", "ability_goal", "ideological_goal", "quality_goal"):
        values = [str(task.get(field, "")).strip() for task in tasks if str(task.get(field, "")).strip()]
        for value, count in Counter(values).items():
            if count >= 3:
                issues.append(("课程任务", "错误", "REPEATED_GOAL", field, f"同一目标完整重复{count}次：{value[:80]}"))
    combined = json.dumps(model, ensure_ascii=False)
    for phrase in GENERIC_PHRASES:
        if phrase in combined:
            issues.append(("内容质量", "提醒", "GENERIC_PHRASE", phrase, f"模型包含空泛表达“{phrase}”，撰写正文前必须具体化。"))
    templates = store.rows("SELECT id,document_type FROM template_files WHERE offering_id=?", (offering_id,))
    for template in templates:
        slots = store.rows("SELECT id,field_name FROM template_slots WHERE template_file_id=?", (template["id"],))
        if not slots:
            issues.append(("模板合同", "错误", "NO_TEMPLATE_SLOTS", template["document_type"], "模板尚未建立填写槽位。"))
            continue
        fields = {slot["field_name"] for slot in slots}
        for field in sorted(REQUIRED_TEMPLATE_FIELDS.get(template["document_type"], set()) - fields):
            aliases = FIELD_ALIASES.get(field)
            if aliases and (aliases & fields):
                continue
            issues.append(("模板合同", "错误", "MISSING_TEMPLATE_FIELD", template["document_type"], f"模板填写合同尚未识别必需栏目：{field}。"))
    authored = store.rows("SELECT document_type,section_key,repeat_key,content_json,evidence_json FROM authored_sections WHERE offering_id=?", (offering_id,))
    required_drafts = {"课程标准", "授课计划", "教学设计"}
    present_drafts = {item["document_type"] for item in authored}
    for document_type in sorted(required_drafts - present_drafts):
        issues.append(("内容撰写", "错误", "MISSING_DRAFT", document_type, f"{document_type}尚未形成可审查内容草稿。"))
    required_sections = {
        "课程标准": {"course_nature", "course_goals", "course_design", "content_hours", "learning_scenarios", "assessment", "teacher_requirements", "course_resources"},
        "授课计划": {"schedule_rows"},
        "教学设计": {"course_goals", "implementation_conditions", "first_lesson_outline", "unit_design"},
    }
    section_pairs = {(item["document_type"], item["section_key"]) for item in authored}
    for document_type, keys in required_sections.items():
        for key in sorted(keys - {section for doc, section in section_pairs if doc == document_type}):
            issues.append(("内容撰写", "错误", "MISSING_DRAFT_SECTION", f"{document_type}:{key}", f"{document_type}缺少内容草稿栏目：{key}。"))
    for item in authored:
        if not json.loads(item["evidence_json"] or "[]"):
            issues.append(("内容撰写", "错误", "DRAFT_WITHOUT_EVIDENCE", f"{item['document_type']}:{item['section_key']}:{item['repeat_key']}", "内容草稿没有记录资料依据。"))
        content = item["content_json"]
        for phrase in GENERIC_PHRASES:
            if phrase in content:
                issues.append(("内容撰写", "提醒", "GENERIC_DRAFT", f"{item['document_type']}:{item['section_key']}:{item['repeat_key']}", f"内容草稿包含空泛表达“{phrase}”。"))
    with store.connect() as db:
        db.execute("DELETE FROM quality_issues WHERE offering_id=?", (offering_id,))
        db.executemany(
            "INSERT INTO quality_issues (offering_id,stage,severity,issue_code,location,message) VALUES (?,?,?,?,?,?)",
            [(offering_id, *issue) for issue in issues],
        )
        db.commit()
    return issues
