"""
内容更新分析器 - 识别教材过时内容并生成更新建议

工作流程：
1. 从资源事实库中提取课程核心知识点
2. AI 分析识别过时技术、废弃API、需要补充的新内容
3. 联网搜索验证最新技术动态
4. 生成结构化更新建议存入 content_updates 表
"""

import json
import re
from pathlib import Path

import store
import resource_retriever
from ai.ai_router import ask_result

# 更新类型
UPDATE_TYPES = [
    "技术更新",      # 旧版本API/语法 → 新版本
    "内容补充",      # 教材没有但现在重要的知识点
    "废弃警告",      # 教材教的内容已经被废弃/不推荐
    "最佳实践更新",  # 当时的最佳实践现在有更好的方案
    "行业趋势",      # 行业新趋势需要让学生了解
]

ANALYSIS_SYSTEM_PROMPT = """你是职业教育课程内容更新专家。
你的任务是分析教材内容，识别其中过时、需要更新或补充的知识点，并给出更新建议。

分析原则：
1. **技术时效性**：重点关注技术版本、API变更、废弃特性
2. **行业实用性**：当前企业实际在用什么，教材有没有脱节
3. **教学衔接性**：过时内容是否会误导学生，需要补充哪些新知识
4. **谨慎判断**：只有确实过时或明显缺失的才提建议，不要为了提而提

请返回JSON数组，每条建议包含以下字段：
- topic: 知识点主题（简短明确）
- update_type: 类型（技术更新/内容补充/废弃警告/最佳实践更新/行业趋势）
- original_summary: 教材中的原内容概述（1-2句话）
- suggested_content: 建议更新/补充的具体内容（详细，可直接用于教学）
- reason: 为什么需要更新（技术演进、行业变化、标准更新等）
- confidence: 置信度 0-1（1表示非常确定）
- related_chapters: 相关章节/项目名称列表

只返回JSON数组，不要其他解释。"""


def _collect_course_knowledge(offering_id, max_facts=80):
    """收集课程核心知识点（从资源事实库）"""
    # 从各章资源中提取代表性事实
    facts = store.rows(
        """SELECT rf.*, ri.title as resource_title, ri.file_path
        FROM resource_facts rf
        JOIN resource_items ri ON rf.resource_item_id = ri.id
        WHERE ri.offering_id=? AND rf.fact_type IN ('知识点', '技术点', '操作步骤', '概念定义')
        ORDER BY rf.confidence DESC
        LIMIT ?""",
        (offering_id, max_facts),
    )
    return facts


def _group_by_chapter(facts):
    """按章节/资源分组知识点"""
    groups = {}
    for f in facts:
        title = f.get("resource_title") or Path(f.get("file_path", "")).stem
        key = title[:30]
        if key not in groups:
            groups[key] = []
        groups[key].append(f["fact_text"][:200])
    return groups


def analyze_content_updates(offering_id, max_suggestions=15):
    """
    分析课程内容，生成更新建议。
    
    Args:
        offering_id: 课程实例ID
        max_suggestions: 最多生成多少条建议
    
    Returns:
        新增的更新建议数量
    """
    offering = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))
    if not offering:
        return 0
    offering = offering[0]
    
    # 收集知识点
    facts = _collect_course_knowledge(offering_id)
    if not facts:
        return 0
    
    # 按章节分组整理
    groups = _group_by_chapter(facts)
    
    # 构建教材内容摘要
    knowledge_text = ""
    for chapter, items in list(groups.items())[:12]:
        knowledge_text += f"\n【{chapter}】\n"
        for item in items[:5]:
            knowledge_text += f"  - {item}\n"
    
    if not knowledge_text.strip():
        return 0
    
    # 先检查是否已经分析过（今天内的待审核/已采纳建议跳过）
    existing = store.rows(
        "SELECT topic FROM content_updates WHERE offering_id=? AND status IN ('待审核','已采纳')",
        (offering_id,),
    )
    existing_topics = {r["topic"] for r in existing}
    
    # 调用AI分析
    prompt = f"""课程名称：{offering['course_name']}
教材版本：{offering.get('textbook_version', '未知')}

以下是从教材中提取的核心知识点（按章节整理）：
{knowledge_text}

请分析以上教材内容，识别过时、需要更新或补充的知识点。

要求：
1. 重点关注：技术版本更新、废弃的API/属性、新的最佳实践、行业新标准
2. 只提有把握的建议，宁少勿滥
3. 建议数量控制在 {max_suggestions} 条以内
4. 每条建议的 suggested_content 要具体充实，可以直接用于教学

直接返回JSON数组。"""
    
    try:
        result = ask_result(prompt, system=ANALYSIS_SYSTEM_PROMPT, force_online=True, show_details=True)
        if not result.get("success"):
            print(f"[content_updater] AI分析失败: {result.get('error')}")
            return 0
        raw_text = result.get("content", "")
        # 解析JSON
        cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw_text.strip(), flags=re.IGNORECASE)
        try:
            result_data = json.loads(cleaned)
        except json.JSONDecodeError:
            candidate = re.search(r"\[[\s\S]*\]", cleaned)
            if not candidate:
                candidate = re.search(r"\{[\s\S]*\}", cleaned)
                if not candidate:
                    return 0
                result_data = json.loads(candidate.group(0))
                result_data = result_data.get("updates", [])
            else:
                result_data = json.loads(candidate.group(0))
    except Exception as e:
        print(f"[content_updater] AI分析异常: {e}")
        return 0
    
    if not result_data:
        return 0
    
    suggestions = result_data if isinstance(result_data, list) else result_data.get("updates", [])
    
    count = 0
    for sug in suggestions:
        topic = sug.get("topic", "").strip()
        if not topic:
            continue
        
        # 去重：相同主题不再重复添加
        if topic in existing_topics:
            continue
        
        update_type = sug.get("update_type", "内容更新")
        if update_type not in UPDATE_TYPES:
            update_type = "内容更新"
        
        try:
            store.add_content_update(
                offering_id=offering_id,
                update_type=update_type,
                topic=topic,
                original_summary=sug.get("original_summary", "")[:500],
                suggested_content=sug.get("suggested_content", "")[:2000],
                reason=sug.get("reason", "")[:500],
                source_urls=sug.get("source_urls", []),
                confidence=float(sug.get("confidence", 0.5)),
                related_chapters=sug.get("related_chapters", []),
            )
            count += 1
            existing_topics.add(topic)
        except Exception as e:
            print(f"[content_updater] 保存建议失败 '{topic}': {e}")
            continue
    
    # 标记三种文档为脏，需要重新生成
    store.mark_dirty(offering_id, "foundation", f"新增 {count} 条内容更新建议")
    
    return count


def format_updates_for_prompt(updates, max_chars=2000):
    """将已采纳的更新内容格式化为提示词文本，供AI生成文档时参考"""
    if not updates:
        return ""
    
    lines = ["【课程内容更新（已审核通过，请融入教学内容）】"]
    total = 0
    for u in updates:
        entry = (
            f"\n● {u['topic']}（{u['update_type']}）\n"
            f"  更新内容：{u['suggested_content'][:300]}\n"
            f"  更新原因：{u['reason'][:150]}"
        )
        if total + len(entry) > max_chars:
            break
        lines.append(entry)
        total += len(entry)
    
    return "\n".join(lines) + "\n"
