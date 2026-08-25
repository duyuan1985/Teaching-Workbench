"""
资源检索引擎 - 基于任务智能检索教材内容

根据任务信息（项目名称、知识点、技能要求）从资源事实库中
检索相关的PPT内容、源码示例、文档片段，为AI内容生成提供
详细的素材支持。

检索策略：
1. 优先从任务已关联的资源中搜索
2. 按项目名称（project_hint）匹配
3. 按关键词匹配内容
4. 按相关性排序，取Top N
5. 控制总内容量，避免超出AI上下文
"""

import json
import re
from pathlib import Path
from collections import defaultdict

import store


def _extract_keywords(text):
    """从文本中提取关键词（简单的中文分词+去停用词）"""
    if not text:
        return []
    
    # 移除标点和空白
    text = re.sub(r'''[，。；！？、：：（）()【】\[\]《》""''\s\d]''', " ", text)
    
    # 常用停用词
    stop_words = {
        "的", "了", "和", "是", "在", "有", "与", "及", "等", "将", "对", "为",
        "以", "到", "从", "由", "把", "被", "给", "让", "向", "往", "从",
        "本", "该", "此", "其", "之", "于", "则", "而", "但", "并", "也",
        "就", "都", "还", "又", "再", "更", "最", "很", "可", "能", "会",
        "要", "不", "没", "请", "你", "我", "他", "她", "它", "们",
        "一个", "一种", "一些", "可以", "能够", "进行", "通过", "根据",
        "以及", "或者", "如果", "因为", "所以", "但是", "而且", "然后",
        "学习", "掌握", "了解", "理解", "熟悉", "知道", "认识",
    }
    
    # 简单分词：按2-4字滑窗 + 已知技术术语
    keywords = set()
    
    # 提取常见技术术语（HTML/CSS/JS相关）
    tech_terms = [
        "HTML", "HTML5", "CSS", "CSS3", "JavaScript", "JS", "Vue", "Vue.js",
        "div", "span", "p", "a", "img", "ul", "li", "ol", "table", "form",
        "input", "button", "header", "footer", "nav", "section", "article",
        "选择器", "类选择器", "ID选择器", "标签选择器", "后代选择器",
        "盒子模型", "margin", "padding", "border", "float", "position",
        "flex", "grid", "响应式", "动画", "transition", "transform",
        "函数", "变量", "数组", "对象", "循环", "条件", "事件",
        "DOM", "BOM", "AJAX", "JSON", "API",
    ]
    
    text_upper = text.upper()
    for term in tech_terms:
        if term.upper() in text_upper:
            keywords.add(term)
    
    # 简单的2字词和3字词提取
    clean = re.sub(r"\s+", "", text)
    for i in range(len(clean) - 1):
        word = clean[i:i+2]
        if word not in stop_words and len(word) == 2 and '\u4e00' <= word[0] <= '\u9fff':
            keywords.add(word)
    
    for i in range(len(clean) - 2):
        word = clean[i:i+3]
        if '\u4e00' <= word[0] <= '\u9fff':
            keywords.add(word)
    
    return list(keywords)


def _calc_relevance(fact, task_keywords, project_name, resource_type_weight=None):
    """
    计算资源事实与任务的相关性得分。
    
    得分因素：
    - 项目名称完全匹配: +10
    - project_hint 包含项目名: +5
    - 关键词命中数: 每个+1
    - 资源类型权重: PPT内容权重更高
    - 内容长度: 适中长度的内容权重更高
    """
    score = 0.0
    
    fact_value = fact.get("fact_value", "")
    fact_key = fact.get("fact_key", "")
    project_hint = fact.get("project_hint", "")
    fact_type = fact.get("fact_type", "")
    
    # 项目匹配
    if project_name and project_hint:
        if project_hint == project_name:
            score += 10
        elif project_name in project_hint or project_hint in project_name:
            score += 5
    
    # 关键词匹配
    content = f"{fact_key} {fact_value}"
    hit_count = 0
    for kw in task_keywords:
        if kw.lower() in content.lower():
            hit_count += 1
            # 技术术语权重更高
            if len(kw) <= 4 and not '\u4e00' <= kw[0] <= '\u9fff':
                score += 2  # 英文技术术语
            else:
                score += 1
    
    # 资源类型权重
    type_weights = {
        "ppt_slide": 1.5,      # PPT内容最相关
        "source_excerpt": 1.2,  # 源码示例次之
        "source_structure": 0.8, # 结构分析参考
        "image_metadata": 0.3,  # 图片信息辅助
    }
    weight = type_weights.get(fact_type, 1.0)
    score *= weight
    
    # 内容长度惩罚（过短信息量不足，过长可能是噪音）
    content_len = len(fact_value)
    if content_len < 20:
        score *= 0.5
    elif content_len > 2000:
        score *= 0.7
    
    return score


def retrieve_for_task(task_id=None, task=None, offering_id=None, max_results=15, max_chars=4000):
    """
    为指定任务检索相关资源内容。
    
    Args:
        task_id: 任务ID（可选，提供后自动查询任务信息）
        task: 任务字典（可选，直接提供任务信息）
        offering_id: 课程实例ID（task_id 或 task 未提供时需要）
        max_results: 最多返回多少条结果
        max_chars: 总内容字符数上限
    
    Returns:
        list of dict: 相关资源内容列表，按相关性排序
            每项包含: fact_type, title, fact_key, fact_value, 
                    project_hint, locator, relevance, resource_type
    """
    # 获取任务信息
    if task is None and task_id:
        task_rows = store.rows("SELECT * FROM tasks WHERE id=?", (task_id,))
        if not task_rows:
            return []
        task = task_rows[0]
    
    if task is None:
        return []
    
    if offering_id is None:
        offering_id = task.get("offering_id")
    
    project_name = task.get("chapter", "")
    task_title = task.get("title", "")
    
    # 提取关键词
    keywords = _extract_keywords(f"{project_name} {task_title}")
    # 加上任务目标中的关键词
    for goal_field in ["knowledge_goal", "ability_goal"]:
        goal_text = task.get(goal_field, "")
        if goal_text:
            keywords.extend(_extract_keywords(goal_text))
    
    # 去重
    keywords = list(dict.fromkeys(keywords))[:30]  # 最多30个关键词
    
    # 获取任务关联的资源ID
    task_resource_ids = set()
    refs_json = task.get("resource_refs", "[]")
    if refs_json:
        try:
            refs = json.loads(refs_json)
            for ref in refs:
                if isinstance(ref, dict):
                    ref_path = ref.get("path", ref.get("file_path", ""))
                else:
                    ref_path = str(ref)
                if ref_path:
                    # 通过路径查找资源ID
                    ritems = store.rows(
                        "SELECT id FROM resource_items WHERE offering_id=? AND file_path=?",
                        (offering_id, ref_path)
                    )
                    for r in ritems:
                        task_resource_ids.add(r["id"])
        except (json.JSONDecodeError, TypeError):
            pass
    
    # 检索资源事实
    if task_resource_ids:
        # 优先从关联资源中检索
        placeholders = ",".join("?" * len(task_resource_ids))
        facts = store.rows(
            f"SELECT rf.*, ri.title, ri.resource_type, ri.file_path "
            f"FROM resource_facts rf JOIN resource_items ri ON rf.resource_item_id = ri.id "
            f"WHERE ri.offering_id=? AND ri.id IN ({placeholders})",
            (offering_id,) + tuple(task_resource_ids)
        )
    else:
        # 没有关联资源，从全部资源中按项目名检索
        facts = store.rows(
            "SELECT rf.*, ri.title, ri.resource_type, ri.file_path "
            "FROM resource_facts rf JOIN resource_items ri ON rf.resource_item_id = ri.id "
            "WHERE ri.offering_id=?",
            (offering_id,)
        )
    
    if not facts:
        return []
    
    # 计算相关性并排序
    scored = []
    for fact in facts:
        score = _calc_relevance(fact, keywords, project_name)
        if score > 0:  # 只保留有相关性的
            scored.append((score, fact))
    
    scored.sort(key=lambda x: x[0], reverse=True)
    
    # 取Top N，并控制总字符数
    results = []
    total_chars = 0
    
    for score, fact in scored[:max_results]:
        fact_value = fact["fact_value"]
        
        # 截断过长的内容
        if len(fact_value) > 800:
            fact_value = fact_value[:800] + "..."
        
        # 检查字符数限制
        if total_chars + len(fact_value) > max_chars and results:
            break
        
        results.append({
            "fact_type": fact["fact_type"],
            "title": fact["title"],
            "resource_type": fact["resource_type"],
            "fact_key": fact["fact_key"],
            "fact_value": fact_value,
            "project_hint": fact.get("project_hint", ""),
            "locator": fact.get("locator", ""),
            "relevance": round(score, 2),
        })
        total_chars += len(fact_value)
    
    return results


def retrieve_for_section(offering_id, section_key, max_results=20, max_chars=6000):
    """
    为课程标准/教学设计的某个章节检索相关资源。
    
    Args:
        offering_id: 课程实例ID
        section_key: 章节标识（如 "course_nature", "course_design"）
        max_results: 最多返回多少条结果
        max_chars: 总内容字符数上限
    
    Returns:
        list of dict: 相关资源内容
    """
    # 根据章节类型确定检索关键词
    section_keywords = {
        "course_nature": ["课程性质", "课程定位", "专业", "人才培养", "地位", "作用"],
        "course_design": ["课程设计", "设计思路", "教学方法", "项目", "任务", "理实一体"],
        "course_goals": ["目标", "知识", "能力", "素质", "思政"],
        "assessment": ["考核", "评价", "成绩", "过程性", "终结性"],
        "teacher_requirements": ["教师", "能力", "素质", "要求"],
        "course_resources": ["教材", "资源", "参考", "资料"],
    }
    
    keywords = section_keywords.get(section_key, [])
    
    # 也从课程语义模型中提取关键词
    models = store.rows(
        "SELECT model_json FROM course_content_models WHERE offering_id=?",
        (offering_id,)
    )
    if models:
        try:
            model = json.loads(models[0]["model_json"])
            keywords.extend(model.get("knowledge_system", [])[:10])
            projects = model.get("projects", [])
            for p in projects[:3]:
                keywords.append(p.get("title", ""))
        except (json.JSONDecodeError, TypeError):
            pass
    
    # 检索所有资源事实
    facts = store.rows(
        "SELECT rf.*, ri.title, ri.resource_type, ri.file_path "
        "FROM resource_facts rf JOIN resource_items ri ON rf.resource_item_id = ri.id "
        "WHERE ri.offering_id=?",
        (offering_id,)
    )
    
    if not facts:
        return []
    
    # 计算相关性
    scored = []
    for fact in facts:
        score = _calc_relevance(fact, keywords, "")
        if score > 0:
            scored.append((score, fact))
    
    scored.sort(key=lambda x: x[0], reverse=True)
    
    results = []
    total_chars = 0
    for score, fact in scored[:max_results]:
        fact_value = fact["fact_value"]
        if len(fact_value) > 600:
            fact_value = fact_value[:600] + "..."
        if total_chars + len(fact_value) > max_chars and results:
            break
        results.append({
            "fact_type": fact["fact_type"],
            "title": fact["title"],
            "resource_type": fact["resource_type"],
            "fact_key": fact["fact_key"],
            "fact_value": fact_value,
            "locator": fact.get("locator", ""),
            "relevance": round(score, 2),
        })
        total_chars += len(fact_value)
    
    return results


def format_resources_for_prompt(resources):
    """
    将检索到的资源格式化为适合AI提示词的文本。
    
    Args:
        resources: retrieve_for_task 返回的结果列表
    
    Returns:
        str: 格式化后的资源文本
    """
    if not resources:
        return "（无相关资源）"
    
    # 按类型分组
    by_type = defaultdict(list)
    for r in resources:
        by_type[r["fact_type"]].append(r)
    
    lines = []
    
    # PPT内容
    if "ppt_slide" in by_type:
        lines.append("【相关PPT内容】")
        # 按文件分组
        by_file = defaultdict(list)
        for r in by_type["ppt_slide"]:
            by_file[r["title"]].append(r)
        
        for file_name, slides in by_file.items():
            lines.append(f"  课件：{file_name}")
            for s in slides[:5]:  # 每个文件最多5页
                lines.append(f"    - {s['fact_key']}: {s['fact_value'][:120]}")
    
    # 文档内容
    doc_types = ["document_paragraph", "document_table_row", "document_heading"]
    for dtype in doc_types:
        if dtype in by_type:
            label = {"document_paragraph": "文档段落", "document_table_row": "文档表格", "document_heading": "文档章节"}[dtype]
            lines.append(f"【相关{label}】")
            for r in by_type[dtype][:8]:
                lines.append(f"  - [{r['title']}] {r['fact_value'][:100]}")
    
    # 源码内容
    code_types = ["code_function", "code_class", "code_snippet"]
    found_code = [r for r in resources if r["fact_type"] in code_types]
    if found_code:
        lines.append("【相关源码片段】")
        for r in found_code[:8]:
            lines.append(f"  - {r['title']} / {r['fact_key']}: {r['fact_value'][:120]}")
    
    return "\n".join(lines)


def get_task_resource_files(task_id=None, task=None, offering_id=None, max_images=10, max_code_files=8):
    """
    获取任务关联的图片和源码文件列表（供AI选择插入哪些文件）。
    
    Returns:
        dict: {"images": [{"path": "...", "name": "...", "title": "..."}], 
               "code_files": [{"path": "...", "name": "...", "title": "..."}]}
    """
    if task is None and task_id:
        task_rows = store.rows("SELECT * FROM tasks WHERE id=?", (task_id,))
        if not task_rows:
            return {"images": [], "code_files": []}
        task = task_rows[0]
    
    if task is None:
        return {"images": [], "code_files": []}
    
    if offering_id is None:
        offering_id = task.get("offering_id")
    
    # 从任务关联的资源中找图片和源码
    refs_json = task.get("resource_refs", "[]")
    ref_paths = set()
    if refs_json:
        try:
            refs = json.loads(refs_json)
            for ref in refs:
                if isinstance(ref, dict):
                    ref_paths.add(ref.get("path", ref.get("file_path", "")))
                else:
                    ref_paths.add(str(ref))
        except (json.JSONDecodeError, TypeError):
            pass
    
    # 获取资源项
    images = []
    code_files = []
    
    if ref_paths:
        placeholders = ",".join("?" * len(ref_paths))
        items = store.rows(
            f"SELECT * FROM resource_items WHERE offering_id=? AND file_path IN ({placeholders})",
            [offering_id] + list(ref_paths),
        )
    else:
        # 没有关联资源，从任务章节找相似的资源
        project_name = task.get("chapter", "")
        items = store.rows(
            "SELECT * FROM resource_items WHERE offering_id=? AND (title LIKE ? OR file_path LIKE ?) LIMIT 50",
            (offering_id, f"%{project_name}%", f"%{project_name}%"),
        )
    
    for item in items:
        rtype = item["resource_type"]
        fpath = item["file_path"]
        name = Path(fpath).name
        
        if rtype == "图片素材" and len(images) < max_images:
            images.append({
                "path": fpath,
                "name": name,
                "title": item.get("title", name),
            })
        elif rtype in ("实训源码/数据",) and name.endswith((".py", ".html", ".css", ".js", ".vue", ".json")) and len(code_files) < max_code_files:
            code_files.append({
                "path": fpath,
                "name": name,
                "title": item.get("title", name),
            })
    
    return {"images": images, "code_files": code_files}


if __name__ == "__main__":
    # 测试：为任务1检索资源
    offering_id = 27
    tasks = store.rows(
        "SELECT * FROM tasks WHERE offering_id=? ORDER BY seq LIMIT 1",
        (offering_id,)
    )
    
    if tasks:
        task = tasks[0]
        print(f"任务: {task['chapter']} - {task['title']}")
        print(f"知识目标: {task['knowledge_goal']}")
        print(f"能力目标: {task['ability_goal']}")
        
        resources = retrieve_for_task(task=task, max_results=12, max_chars=3000)
        print(f"\n检索到 {len(resources)} 条相关资源:")
        for i, r in enumerate(resources):
            print(f"\n{i+1}. [{r['fact_type']}] {r['title']} / {r['fact_key']}")
            print(f"   相关性: {r['relevance']}")
            val = r['fact_value']
            if len(val) > 120:
                val = val[:120] + "..."
            print(f"   内容: {val}")
        
        print(f"\n{'='*50}")
        print("格式化输出（喂给AI的格式）:")
        print('='*50)
        print(format_resources_for_prompt(resources)[:1500])
