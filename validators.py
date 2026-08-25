import json
from datetime import date
from pathlib import Path

import store


PLACEHOLDERS = ("相关教研室", "相关行业企业", "等等", "……", "待填写", "XXX", "********")


def validate(offering, tasks, sessions=None, calendar_events=None, template_files=None, curriculum_units=None):
    sessions = sessions or []
    calendar_events = calendar_events or []
    template_files = template_files or []
    curriculum_units = curriculum_units or []
    issues = []
    total = sum(int(task["hours"]) for task in tasks)
    theory = sum(int(task["theory_hours"]) for task in tasks)
    practice = sum(int(task["practice_hours"]) for task in tasks)
    expected = int(offering["total_hours"])
    if tasks and total != expected:
        issues.append(("错误", f"任务学时合计为 {total}，与课程总学时 {expected} 不一致。"))
    if tasks and theory + practice != total:
        issues.append(("错误", f"理论 {theory} + 实践 {practice} = {theory + practice}，与任务学时 {total} 不一致。"))
    if sessions:
        confirmed = [s for s in sessions if s["status"] == "已确认" and s["session_type"] != "停课"]
        planned = [s for s in sessions if s["status"] != "已取消" and s["session_type"] != "停课"]
        cancelled = [s for s in sessions if s["status"] == "已取消" or s["session_type"] == "停课"]
        class_names = sorted({s.get("class_name", "").strip() for s in sessions if s.get("class_name", "").strip()})
        scope_names = class_names or [""]
        scheduled_by_class = {
            name: sum(int(s["hours"]) for s in confirmed if not name or s.get("class_name", "").strip() == name)
            for name in scope_names
        }
        planned_by_class = {
            name: sum(int(s["hours"]) for s in planned if not name or s.get("class_name", "").strip() == name)
            for name in scope_names
        }
        scheduled_hours = min(scheduled_by_class.values())
        cancelled_hours = sum(int(s["hours"]) for s in cancelled)
        mismatches = [(name, hours) for name, hours in planned_by_class.items() if hours != expected]
        if mismatches:
            for name, hours in mismatches:
                gap = expected - hours
                if gap > 0:
                    issues.append(("提醒", f"{name or '课程'}名义排课 {hours} 学时，距规定 {expected} 学时还差 {gap} 学时。"))
                else:
                    issues.append(("错误", f"{name or '课程'}名义排课 {hours} 学时，超过规定总学时 {-gap} 学时。"))
        else:
            detail = "；".join(f"{name or '课程'}{hours}学时" for name, hours in planned_by_class.items())
            issues.append(("通过", f"各班名义排课与规定总学时一致（{detail}）。"))
        pending_normal_hours = sum(
            int(s["hours"]) for s in sessions
            if s["status"] == "待确认" and s["session_type"] == "正常排课"
        )
        if pending_normal_hours:
            issues.append(("提醒", f"其中 {pending_normal_hours} 学时正常排课仍待进程表或校历确认，当前已确认 {scheduled_hours} 学时。"))
        if cancelled_hours:
            for name in scope_names:
                class_rows = [s for s in sessions if not name or s.get("class_name", "").strip() == name]
                class_cancelled = sum(
                    int(s["hours"]) for s in class_rows
                    if s["status"] == "已取消" or s["session_type"] == "停课"
                )
                if not class_cancelled:
                    continue
                confirmed_replacement = sum(
                    int(s["hours"]) for s in class_rows
                    if s["status"] == "已确认" and s["session_type"] in ("补课", "调课")
                )
                pending_replacement = sum(
                    int(s["hours"]) for s in class_rows
                    if s["status"] == "待确认" and s["session_type"] in ("补课", "调课")
                )
                label = name or "课程"
                if planned_by_class[name] == expected and not pending_replacement:
                    issues.append(("通过", f"{label}有 {class_cancelled} 学时停课记录，已通过 {confirmed_replacement} 学时调课或补课处理，有效学时已达到 {expected}。"))
                elif planned_by_class[name] == expected and pending_replacement:
                    issues.append(("提醒", f"{label}有 {class_cancelled} 学时停课记录，系统已生成 {pending_replacement} 学时待补安排；填写日期、节次和地点并确认后即可补足。"))
                else:
                    remaining = max(0, expected - planned_by_class[name])
                    issues.append(("提醒", f"{label}有 {class_cancelled} 学时停课记录，已确认调课或补课 {confirmed_replacement} 学时，仍缺 {remaining} 学时。"))
    else:
        issues.append(("提醒", "尚未导入或录入排课记录，无法计算正常排课与补课学时。"))
    for event in calendar_events:
        replacement = event["replacement_date"]
        target_week = event["target_week_no"]
        target_day = event["target_weekday"]
        if replacement and (not target_week or not target_day):
            issues.append(("错误", f"校历事件“{event['event_name']}”设置了补课日期，但未明确补第几周星期几的课程。"))
    document_types = {item["document_type"] for item in template_files}
    for required_type in ("课程标准", "授课计划", "教学设计"):
        if required_type not in document_types:
            issues.append(("提醒", f"模板套件中尚未登记“{required_type}”模板。"))
    if curriculum_units:
        curriculum_hours = sum(int(unit["suggested_hours"]) for unit in curriculum_units if unit["review_action"] != "删除")
        if curriculum_hours != expected:
            issues.append(("错误", f"课程蓝本学时合计为 {curriculum_hours}，与规定总学时 {expected} 不一致。"))
        pending = sum(1 for unit in curriculum_units if unit["approval_status"] != "已确认")
        if pending:
            issues.append(("提醒", f"课程蓝本还有 {pending} 个项目未确认，暂不能生成最终文档。"))
        elif not tasks:
            issues.append(("提醒", "课程蓝本已确认，下一步应由系统生成子任务，不需要手工录入。"))
    else:
        issues.append(("提醒", "尚未生成教材内容审查表和课程蓝本。"))
    for task in tasks:
        prefix = f"任务{task['seq']}“{task['title']}”"
        if int(task["hours"]) <= 0:
            issues.append(("错误", f"{prefix}学时必须大于0。"))
        if int(task["theory_hours"]) + int(task["practice_hours"]) != int(task["hours"]):
            issues.append(("错误", f"{prefix}的理论、实践学时之和不等于任务学时。"))
        if not task["lesson_date"]:
            issues.append(("提醒", f"{prefix}尚未安排授课日期。"))
        else:
            try:
                date.fromisoformat(task["lesson_date"])
            except ValueError:
                issues.append(("错误", f"{prefix}日期格式无效，应为 YYYY-MM-DD。"))
        try:
            refs = json.loads(task["resource_refs"])
        except json.JSONDecodeError:
            refs = []
            issues.append(("错误", f"{prefix}资源字段无法解析。"))
        if not refs:
            issues.append(("提醒", f"{prefix}尚未关联教材、PPT或实训资源。"))
        combined = " ".join(str(task.get(key, "")) for key in (
            "knowledge_goal", "ability_goal", "ideological_goal", "quality_goal"
        ))
        for placeholder in PLACEHOLDERS:
            if placeholder in combined:
                issues.append(("错误", f"{prefix}包含占位或空泛文本“{placeholder}”。"))
                break
    return issues or [("通过", "基础数据检查通过，可以进入模板套版阶段。")]


def generation_readiness(offering, tasks, sessions, template_files, curriculum_units, document_types=None):
    """检查生成前置条件。document_types 为 None 时检查全部三类文档；
    指定时只检查所选文档类型依赖的条件（排课检查限于授课计划/教学设计，
    内容模型审查限于课程标准/教学设计，模板规则按所选类型逐项检查）。"""
    if document_types:
        selected = [dt for dt in ("课程标准", "授课计划", "教学设计") if dt in document_types]
    else:
        selected = ["课程标准", "授课计划", "教学设计"]
    blockers = []
    expected = int(offering["total_hours"])
    if sum(int(item["hours"]) for item in tasks) != expected:
        blockers.append("课程任务学时尚未与总学时一致")
    if "授课计划" in selected or "教学设计" in selected:
        active_sessions = [item for item in sessions if item["status"] == "已确认" and item["session_type"] != "停课"]
        class_names = sorted({item.get("class_name", "").strip() for item in active_sessions if item.get("class_name", "").strip()})
        totals = {
            name: sum(int(item["hours"]) for item in active_sessions if item.get("class_name", "").strip() == name)
            for name in class_names
        } if class_names else {"课程": sum(int(item["hours"]) for item in active_sessions)}
        if any(hours != expected for hours in totals.values()):
            blockers.append("至少一个班级的已确认排课学时尚未与总学时一致")
    active_units = [item for item in curriculum_units if item["review_action"] != "删除"]
    if not active_units or any(item["approval_status"] != "已确认" for item in active_units):
        blockers.append("教材课程蓝本尚未全部确认")
    if "课程标准" in selected or "教学设计" in selected:
        content_models = store.rows(
            "SELECT generation_status,review_status FROM course_content_models WHERE offering_id=?",
            (offering["id"],),
        )
        if not content_models:
            blockers.append("尚未根据教材资源生成课程内容模型")
        elif content_models[0]["review_status"] != "已确认":
            blockers.append("课程定位、岗位方向、课程目标和教师要求尚未审查确认")
    templates_by_type = {item["document_type"]: item for item in template_files}
    for document_type in selected:
        template = templates_by_type.get(document_type)
        if not template:
            blockers.append(f"缺少{document_type}模板")
        elif template.get("template_path") and not Path(template["template_path"]).exists():
            blockers.append(f"{document_type}模板文件不存在：{template['template_path']}")
        else:
            analysis = store.rows(
                "SELECT analysis_status FROM template_analyses WHERE template_file_id=?",
                (template["id"],),
            )
            if not analysis or analysis[0].get("analysis_status") != "已确认":
                blockers.append(f"{document_type}模板规则尚未确认")
    return blockers
