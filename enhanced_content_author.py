"""
增强型内容生成器 - 教材内容 + 模板要求双驱动

在原有 content_author 的基础上，增加：
1. 智能资源检索：从教材PPT、源码中提取详细内容
2. 模板要求注入：把模板中的填写说明作为写作约束
3. 质量提升：AI基于真实教材内容生成，避免空泛套话
"""

import json
import os
import re
import time

import store
import resource_retriever
from ai.ai_router import ask_result
from content_author import _ai_activity


def _dump_debug(raw_text, label="parse_fail"):
    """解析失败时保存原始AI返回，便于诊断。"""
    try:
        debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "debug")
        os.makedirs(debug_dir, exist_ok=True)
        path = os.path.join(debug_dir, f"{label}_{time.strftime('%Y%m%d_%H%M%S')}_{int(time.time()*1000)%1000:03d}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write(raw_text or "")
        return path
    except Exception:
        return None


# 增强版系统提示词：强调基于教材内容、遵循模板要求
ENHANCED_SYSTEM_PROMPT = """你是高职院校课程教学设计专家。
你的任务是基于提供的教材内容和模板填写要求，生成高质量的教学文档内容。

重要原则：
1. **以教材为依据**：所有知识点、技术细节、操作步骤必须来自提供的教材资源，不得编造
2. **遵循模板要求**：严格按照模板中的填写说明、结构要求、篇幅要求来生成内容
3. **内容要具体**：要有具体的知识点、代码示例、操作步骤，不能是空泛的套话
4. **贴合任务**：内容必须紧扣当前教学任务，不能写通用模板话

只返回符合指定结构的JSON，不要返回Markdown、解释或代码围栏。"""


def _repair_json(text):
    """尝试修复AI返回的常见JSON语法错误。"""
    repaired = text
    # 把误写成JSON数组的图片/源码标记转为字符串: "图片": [[图片:a.png]] → "[[图片:a.png]]"
    repaired = re.sub(r':\s*\[\[([^\[\]"]*)\]\]', r': "[[\1]]"', repaired)
    # 移除尾随逗号（},] 或 ],  前的逗号）
    repaired = re.sub(r",\s*([}\]])", r"\1", repaired)
    # 移除JavaScript风格注释 //...
    repaired = re.sub(r"//[^\n]*", "", repaired)
    # 将单引号替换为双引号（仅在键/值位置）
    repaired = repaired.replace("'", '"')
    # 移除值前多余的冒号空格导致的空值（如 "key": ,）
    repaired = re.sub(r':\s*,', ': "",', repaired)
    # 修复最后一个键值对后缺少逗号的情况（仅匹配行尾引号后紧跟换行和引号）
    repaired = re.sub(r'"\s*\n\s*"', '",\n"', repaired)
    return repaired


def _parse_json(text):
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", str(text or "").strip(), flags=re.IGNORECASE)
    try:
        result = json.loads(cleaned)
    except json.JSONDecodeError as e1:
        candidate = re.search(r"\{[\s\S]*\}", cleaned)
        if not candidate:
            raise ValueError(f"AI未返回JSON对象: {e1}")
        snippet = candidate.group(0)
        try:
            result = json.loads(snippet)
        except json.JSONDecodeError:
            try:
                result = json.loads(_repair_json(snippet))
            except json.JSONDecodeError as e2:
                raise ValueError(f"AI返回的JSON无法解析: {e2}; 原文前200字: {snippet[:200]}") from e2
    return result


def _normalize_activity(data):
    """
    标准化教学活动数据格式。
    AI可能返回三种格式：
    1. 纯字符串: {"教学导入": "文本内容..."}
    2. 结构化对象: {"教学导入": {"提问": "...", "教师活动": "..."}}
    3. 数组格式: [{"教学导入": {...}}, {"任务1": {...}}]
    
    Returns:
        (flat_dict, structured_dict): 文本版和结构化版
    """
    # 处理数组格式：合并为单个对象
    if isinstance(data, list):
        merged = {}
        for item in data:
            if isinstance(item, dict):
                merged.update(item)
        data = merged
    
    if not isinstance(data, dict):
        return {}, {}
    
    flat = {}
    structured = {}
    
    for key, value in data.items():
        structured[key] = value
        if isinstance(value, str):
            flat[key] = value
        elif isinstance(value, dict):
            # 结构化转文本段落
            parts = []
            for sub_key, sub_val in value.items():
                if sub_val and str(sub_val).strip():
                    parts.append(f"【{sub_key}】{sub_val}")
            flat[key] = "\n".join(parts)
        else:
            flat[key] = str(value)
    
    return flat, structured


def _validate_activity(data):
    """验证教学活动数据完整性"""
    activity_keys = (
        "教学导入", "任务1", "任务2", "任务3",
        "课堂小结", "课后作业", "教学反思",
    )
    for key in activity_keys:
        value = data.get(key)
        if not isinstance(value, str) or len(value.strip()) < 20:
            raise ValueError(f"AI教学活动字段不完整：{key}")


def _get_template_requirements(template_file_id, section_keywords=None):
    """
    获取模板中与指定章节相关的填写要求。
    
    Args:
        template_file_id: 模板文件ID
        section_keywords: 章节关键词列表，用于匹配相关规则
    
    Returns:
        list of dict: 相关的模板规则
    """
    rules = store.rows(
        "SELECT section_title, instruction_text, content_requirements, data_sources "
        "FROM template_rules WHERE template_file_id=? ORDER BY seq",
        (template_file_id,)
    )
    
    if not section_keywords:
        return rules
    
    # 按关键词匹配数排序
    scored = []
    for rule in rules:
        text = f"{rule['section_title']} {rule['instruction_text']} {rule['content_requirements']} {rule.get('data_sources', '')}"
        score = sum(1 for kw in section_keywords if kw and kw in text)
        if score > 0:
            scored.append((score, rule))

    scored.sort(key=lambda x: x[0], reverse=True)

    # 去重：同一 section_title 只保留得分最高的一条
    seen_titles = set()
    result = []
    for score, rule in scored:
        title = rule.get("section_title", "")
        if title not in seen_titles:
            seen_titles.add(title)
            result.append(rule)

    return result[:5] if result else rules[:3]


def _format_template_requirements(rules):
    """将模板规则格式化为提示词文本"""
    if not rules:
        return "（未找到模板填写要求）"
    
    lines = ["【模板填写要求】"]
    for i, rule in enumerate(rules[:5]):  # 最多5条
        if rule.get("section_title"):
            lines.append(f"  {i+1}. 区域：{rule['section_title']}")
        if rule.get("content_requirements"):
            req = rule["content_requirements"]
            if len(req) > 200:
                req = req[:200] + "..."
            lines.append(f"     内容要求：{req}")
        elif rule.get("instruction_text"):
            instr = rule["instruction_text"]
            if len(instr) > 200:
                instr = instr[:200] + "..."
            lines.append(f"     填写说明：{instr}")
        if rule.get("data_sources"):
            lines.append(f"     数据来源：{rule['data_sources'][:100]}")
    
    return "\n".join(lines)


def generate_enhanced_activity(task_id, offering_id, template_file_id=None, force_online=True):
    """
    生成增强版教学活动设计（基于教材内容+模板要求）。
    
    Args:
        task_id: 任务ID
        offering_id: 课程实例ID
        template_file_id: 教学设计模板文件ID（用于获取模板要求）
        force_online: 是否强制使用在线AI
    
    Returns:
        dict: 生成的教学活动内容
    """
    # 获取任务信息
    task_rows = store.rows("SELECT * FROM tasks WHERE id=?", (task_id,))
    if not task_rows:
        raise ValueError("任务不存在")
    task = task_rows[0]
    
    # 获取课程基本信息
    offering = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))[0]
    identity = {
        "course_name": offering["course_name"],
        "major": offering["major"],
        "term": offering["term"],
        "course_nature": offering.get("course_nature", ""),
        "assessment_type": offering.get("assessment_type", ""),
        "department": offering.get("department", ""),
        "teaching_mode": offering.get("teaching_mode", ""),
    }
    
    # 1. 智能检索教材资源
    resources = resource_retriever.retrieve_for_task(
        task=task,
        offering_id=offering_id,
        max_results=15,
        max_chars=5000,
    )
    
    # 1b. 获取可用的图片和源码文件列表（供AI选择插入）
    resource_files = resource_retriever.get_task_resource_files(
        task=task,
        offering_id=offering_id,
        max_images=12,
        max_code_files=10,
    )
    
    # 1c. 获取已采纳的内容更新（AI生成时融入最新内容）
    from content_updater import format_updates_for_prompt
    accepted_updates = store.get_accepted_updates(offering_id)
    update_text = format_updates_for_prompt(accepted_updates, max_chars=1500)
    
    # 2. 获取模板填写要求
    template_requirements = []
    if template_file_id:
        # 用任务标题和知识点作为关键词匹配模板规则
        keywords = [
            task.get("chapter", ""),
            "教学活动",
            "任务",
            "教学设计",
        ]
        # 加上知识点关键词
        keywords.extend(resource_retriever._extract_keywords(task.get("knowledge_goal", ""))[:5])
        template_requirements = _get_template_requirements(template_file_id, keywords)
    
    # 3. 构建增强型提示词
    resource_text = resource_retriever.format_resources_for_prompt(resources)
    req_text = _format_template_requirements(template_requirements)
    
    # 构建可用资源文件列表（图片和源码）
    image_list_text = ""
    if resource_files["images"]:
        image_list_text = "【可用图片素材（可选择插入）】\n"
        for img in resource_files["images"]:
            image_list_text += f"  - [[图片:{img['name']}]]  {img['title']}\n"
    
    code_list_text = ""
    if resource_files["code_files"]:
        code_list_text = "\n【可用源码文件（可选择插入）】\n"
        for cf in resource_files["code_files"]:
            code_list_text += f"  - [[源码:{cf['name']}]]  {cf['title']}\n"
    
    activity_keys = (
        "教学导入", "任务1", "任务2", "任务3",
        "课堂小结", "课后作业", "教学反思",
    )
    
    payload = {
        "课程": identity,
        "任务": {
            "序号": task.get("seq"),
            "项目": task.get("chapter"),
            "任务标题": task.get("title"),
            "学时": task.get("hours"),
            "知识目标": task.get("knowledge_goal"),
            "能力目标": task.get("ability_goal"),
            "思政目标": task.get("ideological_goal"),
            "素质目标": task.get("quality_goal"),
        },
    }
    
    prompt = f"""请为这一次课生成可直接套入教案的教学组织。

【重要要求】
1. 必须基于下面提供的教材内容进行设计，知识点、技术细节要准确
2. 必须遵循模板填写要求
3. 三个课堂任务必须体现递进关系（认知→练习→应用）
4. 每个环节要明确：提问、知识或技术分析、教师活动、学生活动、操作练习、检查标准
5. 要融入德育渗透和板书设计
6. 不得把目标原句简单改写为教学过程，必须有具体的教学活动设计
7. **可以在任务内容中插入图片和源码示例**，使教学设计更直观充实：
   - 图片：在合适位置插入 [[图片:文件名]] 标记，用于展示课件截图、界面效果、操作步骤等
   - 源码：在合适位置插入 [[源码:文件名]] 标记，用于展示关键代码、完整示例等
   - 只在有实际教学价值时插入，不需要硬加；一次课插入 1-3 张图片和 0-2 段源码即可
   - 必须从下面"可用图片素材"和"可用源码文件"列表中选择，文件名必须完全一致
   - 标记要独立占一行，放在相关教学内容描述之后
   - **标记是纯文本，只能写在"内容"等字段的字符串值内部**；禁止单独输出"图片"、"源码"字段，禁止把标记写成JSON数组值（如 "图片": [[图片:a.png]] 是错误写法）
8. **融入课程内容更新**：如果下面有"课程内容更新"部分，说明教材部分内容已过时或需要补充，请将更新内容自然融入教学过程中：
   - 技术更新类：用新方法/新API替代旧内容，同时简要说明新旧差异
   - 内容补充类：作为拓展知识或进阶内容融入相关任务
   - 废弃警告类：明确指出教材中哪些内容已不推荐使用，并给出替代方案
   - 更新内容要有机融入，不要生硬堆砌；与本次课无关的更新不必提及

{req_text}

{resource_text}

{update_text}
{image_list_text}{code_list_text}

【输出格式——严格遵守】
必须输出一个JSON对象，恰好{len(activity_keys)}个键，每个键的值都是字符串，
禁止嵌套对象、禁止数组、禁止输出多个JSON对象、禁止单独输出"图片"/"源码"字段。
骨架如下（把每个占位文本替换为完整内容）：
{json.dumps({key: "（完整内容）" for key in activity_keys}, ensure_ascii=False, indent=2)}

任务基本信息：
{json.dumps(payload, ensure_ascii=False)}

请直接输出JSON。"""
    
    # 调用AI（解析或校验失败时自动重试，AI偶发输出格式错误的JSON）
    data = None
    structured = None
    last_error = None
    retry_prompt = prompt
    for attempt in range(3):
        result = ask_result(retry_prompt, system=ENHANCED_SYSTEM_PROMPT, force_online=force_online, show_details=True)
        if not result.get("success"):
            raise RuntimeError(f"AI生成失败: {result.get('error')}")
        raw_content = result.get("content", "")
        try:
            raw_data = _parse_json(raw_content)
            data, structured = _normalize_activity(raw_data)
            _validate_activity(data)
            break
        except ValueError as e:
            last_error = e
            dumped = _dump_debug(raw_content)
            if dumped:
                print(f"    [调试] 原始AI返回已保存: {dumped}")
            retry_prompt = prompt + f"\n上次返回未通过结构校验：{e}。请补全缺失字段、确保每个字段内容完整后重新输出完整JSON。"
            if attempt < 2:
                print(f"    [重试] AI输出异常，重新请求一次...")
    if data is None:
        raise last_error

    source = f"增强AI/{result.get('source', '未知来源')}/{result.get('model', '未知模型')}"
    
    return {
        "content": data,
        "structured": structured,
        "source": source,
        "resources_used": len(resources),
        "template_rules_used": len(template_requirements),
    }


def generate_enhanced_overview(offering_id, template_file_id=None, force_online=True):
    """
    生成增强版课程概述（课程性质、课程目标、课程设计）。
    
    Args:
        offering_id: 课程实例ID
        template_file_id: 课程标准模板文件ID
        force_online: 是否强制使用在线AI
    
    Returns:
        dict: 生成的课程概述内容
    """
    offering = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))[0]
    identity = {
        "course_name": offering["course_name"],
        "major": offering["major"],
        "term": offering["term"],
        "course_type": offering.get("course_type", ""),
        "course_nature": offering.get("course_nature", ""),
        "total_hours": offering.get("total_hours", ""),
        "lecture_hours": offering.get("lecture_hours", 0),
        "practice_hours": offering.get("practice_hours", 0),
        "credits": offering.get("credits", 0),
        "assessment_type": offering.get("assessment_type", ""),
        "department": offering.get("department", ""),
        "teaching_mode": offering.get("teaching_mode", ""),
    }
    
    # 获取课程语义模型
    models = store.rows(
        "SELECT model_json FROM course_content_models WHERE offering_id=?",
        (offering_id,)
    )
    if not models:
        raise ValueError("尚未生成课程语义模型")
    
    model = json.loads(models[0]["model_json"])
    
    # 1. 检索课程级别的资源
    resources = resource_retriever.retrieve_for_section(
        offering_id=offering_id,
        section_key="course_nature",
        max_results=10,
        max_chars=3000,
    )
    
    # 2. 获取模板要求
    template_requirements = []
    if template_file_id:
        keywords = ["课程性质", "课程目标", "课程设计", "设计思路", "地位"]
        template_requirements = _get_template_requirements(template_file_id, keywords)
    
    # 3. 构建提示词
    resource_text = resource_retriever.format_resources_for_prompt(resources)
    req_text = _format_template_requirements(template_requirements)
    
    # 获取已采纳的内容更新
    from content_updater import format_updates_for_prompt
    accepted_updates = store.get_accepted_updates(offering_id)
    update_text = format_updates_for_prompt(accepted_updates, max_chars=2000)
    
    payload = {
        "课程基本信息": identity,
        "教材项目": model.get("projects", [])[:8],
        "知识体系": model.get("knowledge_system", [])[:20],
        "课程标准": model.get("standards", []),
        "技术工具": model.get("tools_technology", []),
        "教学方法": model.get("teaching_methods", []),
        "工作过程": model.get("work_process", []),
        "课程联系": model.get("course_links", {}),
    }
    
    prompt = f"""请生成课程标准中的课程性质、课程目标和课程设计总体思路。

【重要要求】
1. 内容必须结合具体专业、教材项目、知识技能和资源证据
2. 严格遵循模板中的填写要求和结构
3. 避免可以套用于任意课程的空泛表述
4. 要有具体的课程名称、项目名称、知识点名称
5. 课程性质要说明课程定位、地位作用、与前后课程关系
6. 课程目标要分知识目标、能力目标、思政目标、素质目标
7. 课程设计要说明设计思路、内容组织、教学方法
8. 必须依据课程基本信息中的专业、课程类别(必修/选修)、考核方式、学时构成、授课方式来生成内容，体现本课程在本专业中的具体定位
9. 学时数据以课程基本信息中的total_hours/lecture_hours/practice_hours为准，不得使用教材中的通用学时
10. **融入课程内容更新**：如果下面有"课程内容更新"部分，请将已确认的内容更新体现到课程标准中：
    - 技术更新/废弃警告类：在课程内容设计中体现新方法、新标准，说明课程会与时俱进
    - 内容补充类：作为课程拓展内容或选学模块纳入整体设计
    - 更新内容要自然融入，不要单独列出来

{req_text}

{resource_text}

{update_text}
JSON结构必须为：
{{
  "course_nature": ["三段正文"],
  "course_goals": {{
    "知识目标": ["至少3条"],
    "能力目标": ["至少3条"],
    "思政目标": ["至少3条"],
    "素质目标": ["至少3条"]
  }},
  "course_design": ["三段正文"]
}}

输入资料：
{json.dumps(payload, ensure_ascii=False)}

请直接输出JSON。"""
    
    result = ask_result(prompt, system=ENHANCED_SYSTEM_PROMPT, force_online=force_online, show_details=True)
    if not result.get("success"):
        raise RuntimeError(f"AI生成失败: {result.get('error')}")
    
    data = _parse_json(result.get("content", ""))
    
    # 验证
    for key in ("course_nature", "course_design"):
        values = data.get(key)
        if not isinstance(values, list) or len(values) < 3 or not all(str(item).strip() for item in values):
            raise ValueError(f"AI课程概述字段不完整：{key}")
    
    goals = data.get("course_goals")
    if not isinstance(goals, dict):
        raise ValueError("AI课程目标字段不是对象。")
    for key in ("知识目标", "能力目标", "思政目标", "素质目标"):
        values = goals.get(key)
        if not isinstance(values, list) or len(values) < 3 or not all(str(item).strip() for item in values):
            raise ValueError(f"AI课程目标字段不完整：{key}")
    
    source = f"增强AI/{result.get('source', '未知来源')}/{result.get('model', '未知模型')}"
    
    return {
        "content": data,
        "source": source,
        "resources_used": len(resources),
        "template_rules_used": len(template_requirements),
    }


def generate_and_save_unit_designs(offering_id, task_ids=None, force_regenerate=False):
    """
    批量生成增强版单元教学设计，并保存到 authored_sections 表。
    
    Args:
        offering_id: 课程实例ID
        task_ids: 指定任务ID列表，None表示所有任务
        force_regenerate: 是否强制重新生成（即使已有增强生成结果）
    
    Returns:
        dict: {task_id: result} 每个任务的生成结果
    """
    # 获取模板文件ID
    templates = store.rows(
        "SELECT id FROM template_files WHERE offering_id=? AND document_type=?",
        (offering_id, '教学设计')
    )
    template_file_id = templates[0]['id'] if templates else None
    
    # 获取任务列表
    if task_ids:
        placeholders = ",".join("?" * len(task_ids))
        tasks = store.rows(
            f"SELECT * FROM tasks WHERE offering_id=? AND id IN ({placeholders}) ORDER BY seq",
            (offering_id,) + tuple(task_ids)
        )
    else:
        tasks = store.rows(
            "SELECT * FROM tasks WHERE offering_id=? ORDER BY seq",
            (offering_id,)
        )
    
    results = {}
    total = len(tasks)
    
    print(f"\n[增强生成] 开始生成 {total} 个任务的单元教学设计...")
    print(f"  模板文件ID: {template_file_id}")
    
    for idx, task in enumerate(tasks, 1):
        task_id = task['id']
        seq = task['seq']
        chapter = task['chapter']
        
        print(f"\n  [{idx}/{total}] 任务{seq}: {chapter[:30]}...")
        
        # 检查是否已有增强生成结果
        if not force_regenerate:
            existing = store.rows(
                "SELECT authoring_status FROM authored_sections "
                "WHERE offering_id=? AND document_type='教学设计' AND section_key='unit_design' AND repeat_key=?",
                (offering_id, str(seq))
            )
            if existing and existing[0]['authoring_status'] == '增强AI生成':
                print(f"    已存在增强生成结果，跳过")
                results[task_id] = {"status": "skipped", "reason": "already_enhanced"}
                continue
        
        try:
            result = generate_enhanced_activity(
                task_id=task_id,
                offering_id=offering_id,
                template_file_id=template_file_id,
            )
            
            # 保存到 authored_sections
            _save_unit_design(offering_id, task, result, template_file_id)
            
            results[task_id] = {
                "status": "success",
                "resources_used": result["resources_used"],
                "template_rules_used": result["template_rules_used"],
            }
            print(f"    ✓ 成功（使用 {result['resources_used']} 条资源, {result['template_rules_used']} 条模板规则）")
            
        except Exception as e:
            print(f"    增强生成失败: {e}")
            try:
                model_row = store.rows("SELECT model_json FROM course_content_models WHERE offering_id=?", (offering_id,))
                if model_row:
                    model = json.loads(model_row[0]["model_json"])
                    identity = model["identity"]
                    facts = store.rows("SELECT project_hint,fact_type,fact_value,locator FROM resource_facts WHERE offering_id=?", (offering_id,))
                    activity, activity_status = _ai_activity(task, identity, facts)
                    content_json = {
                        "基本信息": {"周次": task.get("week_no"), "日期": task.get("lesson_date"), "项目": task.get("chapter"), "任务": task.get("title"), "学时": task.get("hours")},
                        "教学目标": {"知识目标": task.get("knowledge_goal", ""), "能力目标": task.get("ability_goal", ""), "思政目标": task.get("ideological_goal", ""), "素质目标": task.get("quality_goal", "")},
                        "教学组织": activity,
                    }
                    title = f"任务{task['seq']}：{task.get('chapter', '')}"
                    with store.connect() as db:
                        db.execute(
                            "INSERT INTO authored_sections "
                            "(offering_id,document_type,section_key,repeat_key,title,content_json,evidence_json,authoring_status,review_status,generated_at) "
                            "VALUES (?,?,?,?,?,?,?,?,'待检查',CURRENT_TIMESTAMP) "
                            "ON CONFLICT(offering_id,document_type,section_key,repeat_key) DO UPDATE SET "
                            "title=excluded.title,content_json=excluded.content_json,evidence_json=excluded.evidence_json,"
                            "authoring_status=excluded.authoring_status,review_status='待检查',generated_at=CURRENT_TIMESTAMP",
                            (offering_id, '教学设计', 'unit_design', str(task['seq']), title,
                             json.dumps(content_json, ensure_ascii=False),
                             json.dumps({"生成方式": "基础AI生成(增强失败回退)"}, ensure_ascii=False),
                             activity_status),
                        )
                        db.commit()
                    results[task_id] = {"status": "fallback", "error": str(e)}
                    print(f"    已回退到基础AI生成")
                else:
                    results[task_id] = {"status": "error", "error": str(e)}
            except Exception as fallback_err:
                results[task_id] = {"status": "error", "error": f"增强: {e}; 回退: {fallback_err}"}
                print(f"    回退也失败: {fallback_err}")
    
    # 统计
    success = sum(1 for r in results.values() if r["status"] == "success")
    skipped = sum(1 for r in results.values() if r["status"] == "skipped")
    fallback = sum(1 for r in results.values() if r["status"] == "fallback")
    failed = sum(1 for r in results.values() if r["status"] == "error")

    print(f"\n[增强生成] 完成：成功 {success}，跳过 {skipped}，回退 {fallback}，失败 {failed}")
    return results


def _save_unit_design(offering_id, task, result, template_file_id=None):
    """
    将增强生成的单元教学设计保存到 authored_sections 表。
    
    保存格式与原有 content_author 兼容：
    - section_key: 'unit_design'
    - repeat_key: str(task_seq)
    - content_json: {"教学组织": {...}, "结构化": {...}}
    - authoring_status: '增强AI生成'
    """
    seq = task['seq']
    content = result['content']
    structured = result.get('structured', {})
    
    # 构建兼容格式：教学组织字段直接包含内容字典
    # 原有格式: {"教学组织": {"教学导入": "文本", "任务1": "文本", ...}}
    content_json = {
        "教学组织": content,
        "结构化版本": structured,
    }
    
    # 证据：使用的资源
    resources_used = result.get('resources_used', 0)
    evidence = {
        "生成方式": "增强AI生成",
        "模型来源": result.get('source', ''),
        "使用资源数": resources_used,
        "使用模板规则数": result.get('template_rules_used', 0),
        "教材项目": task.get('chapter', ''),
    }
    
    title = f"任务{seq}：{task.get('chapter', '')}"
    
    with store.connect() as db:
        db.execute(
            "INSERT INTO authored_sections "
            "(offering_id,document_type,section_key,repeat_key,title,content_json,evidence_json,authoring_status,review_status,generated_at) "
            "VALUES (?,?,?,?,?,?,?,?,'待检查',CURRENT_TIMESTAMP) "
            "ON CONFLICT(offering_id,document_type,section_key,repeat_key) DO UPDATE SET "
            "title=excluded.title,content_json=excluded.content_json,evidence_json=excluded.evidence_json,"
            "authoring_status=excluded.authoring_status,review_status='待检查',generated_at=CURRENT_TIMESTAMP",
            (
                offering_id, '教学设计', 'unit_design', str(seq), title,
                json.dumps(content_json, ensure_ascii=False),
                json.dumps(evidence, ensure_ascii=False),
                '增强AI生成',
            ),
        )
        db.commit()


def is_enhanced_mode_enabled():
    """检查是否启用了增强生成模式"""
    val = store.get_setting("enhanced_generation", "0")
    return val == "1"


def enable_enhanced_mode(enable=True):
    """启用或禁用增强生成模式"""
    store.set_setting("enhanced_generation", "1" if enable else "0")


if __name__ == "__main__":
    # 测试：为任务1生成增强版教学活动
    offering_id = 27
    
    # 找教学设计模板
    templates = store.rows(
        "SELECT id, document_type FROM template_files WHERE offering_id=? AND document_type=?",
        (offering_id, '教学设计')
    )
    design_tpl_id = templates[0]['id'] if templates else None
    
    # 找第一个任务
    tasks = store.rows(
        "SELECT id, seq, chapter FROM tasks WHERE offering_id=? ORDER BY seq LIMIT 1",
        (offering_id,)
    )
    
    if tasks:
        task = tasks[0]
        print(f"测试生成: 任务{task['seq']} - {task['chapter']}")
        print(f"模板文件ID: {design_tpl_id}")
        print(f"\n正在调用增强型AI生成...")
        
        try:
            result = generate_enhanced_activity(
                task_id=task['id'],
                offering_id=offering_id,
                template_file_id=design_tpl_id,
            )
            
            print(f"\n✓ 生成成功!")
            print(f"  使用资源数: {result['resources_used']}")
            print(f"  使用模板规则数: {result['template_rules_used']}")
            print(f"  内容来源: {result['source']}")
            
            content = result['content']
            print(f"\n--- 生成内容预览 ---")
            for key in ["教学导入", "任务1", "任务2", "任务3", "课堂小结"]:
                val = content.get(key, "")
                preview = val[:100] + "..." if len(val) > 100 else val
                print(f"\n【{key}】({len(val)}字)")
                print(f"  {preview}")
                
        except Exception as e:
            print(f"\n✗ 生成失败: {e}")
            import traceback
            traceback.print_exc()
