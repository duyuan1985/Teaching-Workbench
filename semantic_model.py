import hashlib
import json
import re

import store
import talent_plan
from assessment_scheme import get_scheme


def _parts(text):
    result = []
    for value in re.split(r"[；;、\n]+", str(text or "")):
        value = value.strip()
        if value and value not in result:
            result.append(value)
    return result


def _evidence(db, offering_id, key, evidence_type, value, source_type, locator, confidence=1.0):
    db.execute(
        "INSERT INTO course_evidence (offering_id,evidence_key,evidence_type,value_json,source_type,source_locator,confidence) VALUES (?,?,?,?,?,?,?)",
        (offering_id, key, evidence_type, json.dumps(value, ensure_ascii=False), source_type, locator, confidence),
    )


def build_semantic_model(offering_id):
    offering = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))[0]
    units = store.rows(
        "SELECT * FROM curriculum_units WHERE offering_id=? AND approval_status='已确认' AND review_action<>'删除' ORDER BY seq",
        (offering_id,),
    )
    if not units:
        raise ValueError("课程蓝本尚未确认，不能建立课程语义模型。")
    resource_items = store.rows(
        "SELECT id,file_path,title,resource_type,project_hint,extraction_status FROM resource_items WHERE offering_id=?",
        (offering_id,),
    )
    facts = store.rows(
        "SELECT rf.*,ri.file_path FROM resource_facts rf JOIN resource_items ri ON ri.id=rf.resource_item_id WHERE rf.offering_id=?",
        (offering_id,),
    )
    projects = []
    knowledge = []
    tools = []
    standards = []
    processes = []
    methods = []
    assessment_scheme = get_scheme(offering_id)
    with store.connect() as db:
        db.execute("DELETE FROM course_evidence WHERE offering_id=?", (offering_id,))
        identity = {
            key: offering.get(key) for key in (
                "course_name", "course_code", "term", "major", "course_nature", "course_type", "credits",
                "total_hours", "weekly_hours", "teaching_class", "textbook_version",
                "prerequisite_courses", "followup_courses",
                "department", "teaching_mode", "assessment_type",
                "lecture_hours", "experiment_hours", "practice_hours",
            )
        }
        identity["offering_id"] = offering_id
        _evidence(db, offering_id, "course_identity", "fact", identity, "课程实例/教学安排表", f"offering:{offering_id}")
        for unit in units:
            skills = _parts(unit["source_skills"] or unit["revised_focus"])
            knowledge.extend(skill for skill in skills if skill not in knowledge)
            mapped = [
                {"resource_id": item["id"], "type": item["resource_type"], "path": item["file_path"]}
                for item in resource_items
                if item.get("project_hint") == unit["project_title"]
                or (unit.get("source_file") and item["file_path"].lower() == unit["source_file"].lower())
            ]
            project = {
                "seq": unit["seq"], "title": unit["project_title"], "hours": unit["suggested_hours"],
                "knowledge_skills": skills, "expected_outcome": f"完成“{unit['project_title']}”项目成果并进行检查、改进和提交",
                "resources": mapped,
            }
            projects.append(project)
            _evidence(db, offering_id, f"project:{unit['seq']}", "project", project, "课程蓝本/教材资源", f"curriculum_unit:{unit['id']}")
            for field, bucket in (("new_standards", standards), ("new_technology", tools), ("new_process", processes), ("new_methods", methods)):
                for value in _parts(unit[field]):
                    if value not in bucket:
                        bucket.append(value)
        for fact in facts:
            if fact["fact_key"] == "imports":
                for value in _parts(fact["fact_value"]):
                    if value not in tools:
                        tools.append(value)
        for name, values, source in (
            ("knowledge_system", knowledge, "教材PPT/课程蓝本"),
            ("tools_technology", tools, "课程蓝本/源码结构"),
            ("standards", standards, "课程蓝本"),
            ("work_process", processes, "课程蓝本"),
            ("teaching_methods", methods, "课程蓝本"),
        ):
            _evidence(db, offering_id, name, "derived_collection", values, source, name, 0.9)
        missing = []
        source_types = {item["source_type"] for item in store.rows("SELECT source_type FROM source_files WHERE offering_id=?", (offering_id,))}
        if not ({"行业标准", "职业技能标准"} & source_types):
            missing.append("尚未提供行业标准或职业技能标准原文，四新内容只能依据教材和课程蓝本分析")
        missing.append("原始资料未提供学生基础、学习困难和班级差异，学情分析只能依据课程先后关系形成待确认草案")
        plan = talent_plan.get_plan_for_major(offering.get("major") or "")
        talent_plan_data = None
        plan_positions = []
        if plan:
            course_info = talent_plan.find_course(plan, offering.get("course_name") or "")
            orientation = plan.get("orientation") or {}
            plan_positions = orientation.get("job_positions") or []
            talent_plan_data = {
                "major": plan.get("major"),
                "cohort": plan.get("cohort"),
                "goals": plan.get("goals", ""),
                "specs": plan.get("specs", ""),
                "orientation": orientation,
                "course_info": course_info,
            }
            _evidence(
                db, offering_id, "talent_plan", "fact", talent_plan_data,
                f"人才培养方案（{plan.get('cohort') or '版本未标注'}）", plan.get("source_path", ""),
            )
            if not course_info:
                missing.append(f"{plan.get('cohort') or ''}人才培养方案课程设置表中未找到《{offering.get('course_name')}》，课程类型与定位按现有资料推导")
            if plan_positions:
                missing.append(f"岗位方向取自{plan.get('cohort') or ''}人才培养方案“职业面向”表（{plan.get('major')}专业全部岗位群），请结合本课程筛选适用岗位后确认")
            else:
                missing.append("人才培养方案职业面向表中未提取到岗位群，岗位方向只能依据专业名称和项目成果分析形成待确认草案")
        else:
            missing.append(f"未提供{offering.get('major')}专业人才培养方案，课程性质中的培养目标与培养规格只能依据教材内容推导")
            missing.append("原始资料未提供正式岗位说明，岗位方向只能依据专业名称和项目成果分析形成待确认草案")
        model = {
            "schema_version": "semantic-course-v1",
            "identity": identity,
            "projects": projects,
            "knowledge_system": knowledge,
            "ability_outcomes": plan_positions,
            "tools_technology": tools,
            "standards": standards,
            "work_process": processes,
            "teaching_methods": methods,
            "course_links": {"prerequisite": offering.get("prerequisite_courses", ""), "followup": offering.get("followup_courses", "")},
            "assessment_evidence": ["项目过程记录", "阶段作品", "源文件或源代码", "成果展示与评价量表", "课程综合作品"],
            "assessment_scheme": assessment_scheme,
            "missing_evidence": missing,
            "resource_summary": {
                "total": len(resource_items),
                "parsed": sum(1 for item in resource_items if item.get("extraction_status") == "已解析"),
                "facts": len(facts),
            },
        }
        if talent_plan_data:
            model["talent_plan"] = talent_plan_data
        signature = hashlib.sha256(json.dumps(model, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()
        db.execute(
            "INSERT INTO course_content_models (offering_id,model_json,source_signature,generation_status,review_status,generated_at) VALUES (?,?,?,'语义模型已生成','待检查',CURRENT_TIMESTAMP) "
            "ON CONFLICT(offering_id) DO UPDATE SET model_json=excluded.model_json,source_signature=excluded.source_signature,generation_status=excluded.generation_status,review_status='待检查',generated_at=CURRENT_TIMESTAMP",
            (offering_id, json.dumps(model, ensure_ascii=False), signature),
        )
        db.commit()
    return model
