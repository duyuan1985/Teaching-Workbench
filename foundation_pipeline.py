import store
from content_author import author_course_content
from quality_checker import run_foundation_quality_check
from resource_analyzer import analyze_offering_resources
from semantic_model import build_semantic_model
from task_enricher import enrich_tasks_from_evidence
from template_analyzer import analyze_template


def _run_stage(offering_id, stage, fn):
    store.set_pipeline_stage(offering_id, stage, "running")
    try:
        result = fn()
        store.set_pipeline_stage(offering_id, stage, "done")
        return result
    except Exception as e:
        store.set_pipeline_stage(offering_id, stage, "failed", str(e))
        raise ValueError(f"{stage}阶段失败：{e}")


def rebuild_generation_foundation(offering_id):
    prev = store.get_pipeline_status(offering_id)
    has_failed = any(s["status"] == "failed" for s in prev.values())
    if has_failed:
        store.clear_pipeline_status(offering_id)
        prev = {}

    # 阶段1：资源解析
    status = prev.get("resources", {}).get("status")
    if status != "done":
        resource_result = _run_stage(offering_id, "resources", lambda: _check_or_analyze_resources(offering_id))
    else:
        fact_count = store.rows("SELECT COUNT(*) count FROM resource_facts WHERE offering_id=?", (offering_id,))[0]["count"]
        resource_result = {"facts": fact_count}

    # 阶段2：模板分析
    status = prev.get("templates", {}).get("status")
    if status != "done":
        template_result = _run_stage(offering_id, "templates", lambda: _analyze_templates(offering_id))
    else:
        template_result = {}

    # 阶段3：任务增强
    status = prev.get("tasks", {}).get("status")
    if status != "done":
        enriched_tasks = _run_stage(offering_id, "tasks", lambda: enrich_tasks_from_evidence(offering_id))
    else:
        enriched_tasks = 0

    # 阶段4：语义模型
    status = prev.get("model", {}).get("status")
    if status != "done":
        model = _run_stage(offering_id, "model", lambda: build_semantic_model(offering_id))
    else:
        model = store.rows("SELECT model_json FROM course_content_models WHERE offering_id=?", (offering_id,))[0]
        import json
        parsed = json.loads(model["model_json"])
        if not isinstance(parsed, dict):
            parsed = {}
        model = {"model_json": model["model_json"], "projects": parsed.get("projects", []),
                 "knowledge_system": parsed.get("knowledge_system", [])}

    # 阶段5：内容生成
    status = prev.get("content", {}).get("status")
    if status != "done":
        authored_sections = _run_stage(offering_id, "content", lambda: author_course_content(offering_id))
    else:
        authored_sections = store.rows("SELECT COUNT(*) count FROM authored_sections WHERE offering_id=?", (offering_id,))[0]["count"]

    # 阶段6：质量检查
    status = prev.get("quality", {}).get("status")
    if status != "done":
        issues = _run_stage(offering_id, "quality", lambda: run_foundation_quality_check(offering_id))
    else:
        issues = store.rows("SELECT COUNT(*) count FROM quality_issues WHERE offering_id=?", (offering_id,))[0]["count"]

    return {
        "resources": resource_result,
        "templates": template_result,
        "projects": len(model.get("projects", [])),
        "knowledge_items": len(model.get("knowledge_system", [])),
        "enriched_tasks": enriched_tasks,
        "authored_sections": authored_sections,
        "quality_issues": len(issues) if isinstance(issues, list) else issues,
    }


def _check_or_analyze_resources(offering_id):
    status = store.rows(
        "SELECT COUNT(*) total,SUM(CASE WHEN extraction_status='已解析' THEN 1 ELSE 0 END) analyzed,"
        "SUM(CASE WHEN extraction_status='解析失败' THEN 1 ELSE 0 END) failed "
        "FROM resource_items WHERE offering_id=?",
        (offering_id,),
    )[0]
    if not status["total"]:
        raise ValueError("资源索引为空，请先完成教材资源扫描。")
    if status["failed"] > 0:
        raise ValueError(f"存在{status['failed']}个解析失败的资源，请检查后再重建。")
    if status["analyzed"] != status["total"]:
        return analyze_offering_resources(offering_id)
    fact_count = store.rows("SELECT COUNT(*) count FROM resource_facts WHERE offering_id=?", (offering_id,))[0]["count"]
    return {"resources": status["total"], "analyzed": status["analyzed"], "failed": 0, "facts": fact_count}


def _analyze_templates(offering_id):
    templates = store.rows("SELECT id,document_type FROM template_files WHERE offering_id=? ORDER BY id", (offering_id,))
    template_result = {}
    for template in templates:
        rule_count = analyze_template(template["id"])
        slots = store.rows("SELECT COUNT(*) count FROM template_slots WHERE template_file_id=?", (template["id"],))[0]["count"]
        template_result[template["document_type"]] = {"rules": rule_count, "slots": slots}
    return template_result
