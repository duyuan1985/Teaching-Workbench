"""
AI驱动的课程蓝本审查模块

用智谱GLM分析PPT内容，自动识别：
- 项目分组
- 学习目标
- 技能点
- 建议学时

替代原有硬编码逻辑，适配任意课程教材。
"""

import json
import re
from pathlib import Path

import store
from ai.ai_router import ask_result

MAX_PPT_CHARS = 12000
MAX_GROUPS = 15


def _collect_ppt_summaries(offering_id):
    """收集PPT内容摘要，按文件分组"""
    ppts = store.rows(
        "SELECT * FROM resource_items WHERE offering_id=? "
        "AND resource_type IN ('PPT课件','旧版PPT课件') ORDER BY file_path",
        (offering_id,),
    )
    if not ppts:
        raise ValueError("资源索引中没有可读取的PPT课件。")

    main_ppts = [
        item for item in ppts
        if re.search(r"[\\/]0?\d+\s*课程\s*ppt[\\/]", item["file_path"], re.I)
    ]
    if main_ppts:
        ppts = main_ppts

    summaries = []
    for item in ppts:
        excerpt = item.get("content_excerpt", "") or ""
        if len(excerpt) > 2000:
            excerpt = excerpt[:2000]
        parent_name = Path(item["file_path"]).parent.name
        summaries.append({
            "file": Path(item["file_path"]).name,
            "parent_dir": parent_name,
            "content": excerpt,
        })
    return summaries


def _build_prompt(offering, ppt_summaries):
    """构建AI审查提示词"""
    course_name = offering.get("course_name", "")
    total_hours = offering.get("total_hours", 64)
    textbook = f"{offering.get('textbook_version', '')} {offering.get('textbook_path', '')}".strip()

    ppt_text = ""
    for i, s in enumerate(ppt_summaries, 1):
        ppt_text += f"\n--- PPT {i}: {s['file']} (目录: {s['parent_dir']}) ---\n{s['content']}\n"
        if len(ppt_text) > MAX_PPT_CHARS:
            ppt_text = ppt_text[:MAX_PPT_CHARS] + "\n...(后续PPT内容省略)"
            break

    return f"""你是一位资深职业教育课程专家。请分析以下课程PPT内容，输出课程蓝本审查结果。

## 课程信息
- 课程名称：{course_name}
- 教材：{textbook}
- 总学时：{total_hours}

## PPT内容
{ppt_text}

## 任务
请分析PPT内容，识别课程的项目结构，输出JSON格式的审查结果。

## 输出要求
输出一个JSON对象，格式如下：
{{
  "groups": [
    {{
      "title": "项目名称（如：HTML5基础认知）",
      "objectives": "学习目标1；学习目标2；学习目标3",
      "skills": "技能点1；技能点2；技能点3；技能点4",
      "suggested_hours": {total_hours // 6},
      "modernization": {{
        "standards": "相关质量/行业标准",
        "technology": "使用的工具/技术",
        "process": "工作流程（用—连接）",
        "methods": "教学方法"
      }}
    }}
  ]
}}

## 规则
1. 每个PPT文件或目录对应一个项目（如果多个PPT在同一目录下，合并为一个项目）
2. 项目数量不超过{MAX_GROUPS}个
3. 学时分配总和加"综合评价与课程总结"（占总学时的1/8左右）应等于总学时{total_hours}
4. 学习目标从PPT前几页的"学习目标"或"教学目标"页面提取
5. 技能点从PPT标题页和内容中提取，用分号分隔
6. 如果PPT内容不足以判断，根据课程名和教材信息合理推断
7. 现代化标准要与课程内容匹配

请只输出JSON，不要输出其他文字。"""


def ai_review_curriculum(offering_id):
    """AI驱动的课程蓝本审查"""
    offering = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))
    if not offering:
        raise ValueError("课程实例不存在")
    offering = offering[0]

    total_hours = int(offering.get("total_hours") or 64)

    ppt_summaries = _collect_ppt_summaries(offering_id)
    prompt = _build_prompt(offering, ppt_summaries)

    try:
        result = ask_result(
            prompt=prompt,
            system="你是一位资深职业教育课程专家，擅长分析课程结构和教学内容。",
            prefer_local=False,
            force_online=True,
        )
    except Exception as e:
        raise ValueError(f"AI调用失败: {e}")

    if not result.get("success"):
        raise ValueError(f"AI调用失败: {result.get('error', '未知错误')}")

    response = result.get("content", "")
    if not response or not response.strip():
        raise ValueError("AI返回空结果")

    text = response.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)

    try:
        result = json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start >= 0 and end > start:
            result = json.loads(text[start:end + 1])
        else:
            raise ValueError(f"AI返回的内容不是有效的JSON: {text[:200]}")

    if not isinstance(result, dict):
        raise ValueError(f"AI返回的JSON顶层不是对象: {type(result).__name__}")

    groups = result.get("groups", [])
    if not groups:
        raise ValueError("AI未返回任何项目分组")

    if len(groups) > MAX_GROUPS:
        groups = groups[:MAX_GROUPS]

    validated_groups = []
    for g in groups:
        if not isinstance(g, dict):
            continue
        title = g.get("title", "").strip()
        if not title:
            continue
        objectives = g.get("objectives", "").strip()
        skills = g.get("skills", "").strip()
        hours = g.get("suggested_hours", 0)
        try:
            hours = int(hours)
        except (ValueError, TypeError):
            hours = 0
        if hours <= 0 or hours > total_hours:
            hours = max(2, total_hours // max(len(groups), 1))

        mod = g.get("modernization", {})
        if not isinstance(mod, dict):
            mod = {}
        validated_groups.append({
            "title": title,
            "objectives": objectives,
            "skills": skills,
            "suggested_hours": hours,
            "standards": mod.get("standards", "").strip(),
            "technology": mod.get("technology", "").strip(),
            "process": mod.get("process", "").strip(),
            "methods": mod.get("methods", "").strip(),
        })

    if not validated_groups:
        raise ValueError("AI返回的项目分组数据无效（缺少标题）")

    assessment_hours = min(
        int(offering.get("weekly_hours") or 4),
        int(offering["total_hours"]),
    )

    units = []
    for group in validated_groups:
        title = group["title"]
        objectives = group["objectives"]
        skills = group["skills"]
        hours = group["suggested_hours"]

        standards = group["standards"]
        technology = group["technology"]
        process = group["process"]
        methods = group["methods"]

        units.append({
            "title": title,
            "source_file": "",
            "objectives": objectives,
            "skills": skills,
            "revised_focus": objectives,
            "rationale": f"AI分析：基于PPT内容识别的{title}项目",
            "standards": standards,
            "technology": technology,
            "process": process,
            "methods": methods,
            "hours": int(hours),
        })

    units.append({
        "title": "综合评价与课程总结",
        "source_file": "",
        "objectives": "课程成果提交、综合评价与学习总结",
        "skills": "课程成果汇报与评价；课程总结与复习",
        "revised_focus": "课程成果汇报与评价；课程总结与复习",
        "rationale": "完成课程成果验收、问题复盘与后续学习规划。",
        "standards": "课程质量标准与成果评价规范",
        "technology": "成果展示、文档整理与评价记录工具",
        "process": "成果整理—展示汇报—多元评价—复盘改进",
        "methods": "成果答辩、同伴互评、自我评价和总结反思",
        "hours": assessment_hours,
    })

    total_allocated = sum(u["hours"] for u in units)
    target = int(offering["total_hours"])
    if total_allocated != target:
        diff = target - total_allocated + assessment_hours
        if units:
            units[-2]["hours"] += diff

    with store.connect() as db:
        db.execute("DELETE FROM curriculum_units WHERE offering_id=?", (offering_id,))
        for index, unit in enumerate(units, 1):
            db.execute(
                """INSERT INTO curriculum_units
                (offering_id,seq,project_title,source_file,source_objectives,source_skills,review_action,
                 revised_focus,rationale,new_standards,new_technology,new_process,new_methods,suggested_hours,approval_status)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (offering_id, index, unit["title"], unit["source_file"],
                 unit["objectives"], unit["skills"], "AI审查",
                 unit["revised_focus"], unit["rationale"], unit["standards"],
                 unit["technology"], unit["process"], unit["methods"],
                 unit["hours"], "待确认"),
            )
        db.commit()

    return len(units), [u["hours"] for u in units]
