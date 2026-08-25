import json
import re

import store


def _parts(text):
    return [item.strip() for item in re.split(r"[、；;]+", str(text or "")) if item.strip()]


def _numbered(items):
    return "\n".join(f"{index}. {item}" for index, item in enumerate(items, 1))


def _knowledge_statement(skill):
    if skill.endswith(("作用", "概述", "介绍", "概念")):
        return f"理解{skill}、基本规则和适用场景"
    return f"理解{skill}的作用、基本规则和适用场景"


def _representative_resources(resources):
    selected = []
    for resource_type in ("PPT课件", "教材/实训文档", "实训源码/数据", "图片素材", "音视频资源"):
        item = next((row for row in resources if row["resource_type"] == resource_type), None)
        if item:
            selected.append(item["file_path"])
        if len(selected) == 4:
            break
    return selected


def _ideological_goal(text, major, context):
    if "标注" in text:
        return _numbered([f"在“{context}”中严格执行标注规范，保护原始数据和个人信息，不擅自复制、传播任务数据", "坚持客观、准确地标注和复核，不隐瞒漏标错标，如实记录个人完成情况"])
    if any(word in text for word in ("表单", "注册", "数据", "用户")):
        return _numbered([f"在“{context}”任务中遵守个人信息保护、数据安全和网络安全要求，按必要范围处理业务数据", "坚持真实分析和规范验证，不篡改、不选择性隐瞒数据与分析结果"])
    if any(word in text for word in ("图像", "音频", "视频", "素材")):
        return _numbered([f"在“{context}”任务中尊重数字作品著作权，规范使用并标注素材来源", "坚持健康审美和正确价值导向，提升数字内容的社会责任意识"])
    if any(word in text for word in ("企业", "集团", "网站", "首页", "营销")):
        return _numbered([f"在“{context}”任务中坚持真实、准确、合规地表达组织和产品信息", "理解数字技术服务行业发展和社会需求的责任，形成用户意识与质量意识"])
    if any(word in text for word in ("Canvas", "绘制", "动画", "变形")):
        return _numbered([f"在“{context}”的反复调试和细节完善中践行严谨求实、精益求精的职业精神", "尊重技术规律和评价标准，形成以可靠成果证明能力的诚信意识"])
    second = "关注数字技术服务乡村振兴和农村电子商务发展的实际价值" if "农村电子商务" in major else "理解专业技术服务经济社会发展的责任"
    return _numbered([f"在“{context}”任务中遵守技术规范和知识产权要求，形成认真负责的职业态度", second])


def _quality_goal(text, implementation, context):
    if implementation:
        return _numbered([f"完成“{context}”时形成按步骤实施、及时测试、记录问题和持续改进的习惯", "提升任务分工、沟通反馈、成果检查和按时交付能力"])
    if any(word in text for word in ("布局", "样式", "图像", "界面")):
        return _numbered([f"在“{context}”中培养版式审美、细节检查和视觉一致性意识", "能够自主查阅资料、比较方案并说明设计选择"])
    return _numbered([f"完成“{context}”时形成规范操作、主动学习和依据标准检查成果的习惯", "提升独立分析、表达技术思路和接受反馈的能力"])


def enrich_tasks_from_evidence(offering_id):
    offering = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))[0]
    tasks = store.rows("SELECT * FROM tasks WHERE offering_id=? ORDER BY seq", (offering_id,))
    resources = store.rows(
        "SELECT file_path,resource_type,project_hint FROM resource_items WHERE offering_id=? AND extraction_status='已解析' ORDER BY id",
        (offering_id,),
    )
    chapter_counts = {}
    updated = 0
    with store.connect() as db:
        for task in tasks:
            chapter_counts[task["chapter"]] = chapter_counts.get(task["chapter"], 0) + 1
            detail = task["title"].split("：", 1)[-1]
            skills = _parts(detail)[:4]
            implementation = chapter_counts[task["chapter"]] > 1 or "任务实施" in task["title"]
            if task["chapter"] == "综合评价与课程总结":
                refs = ["课程项目成果", "作品提交要求", "课程评价量表", "成果汇报与答辩记录"]
                knowledge = _numbered(["掌握课程成果整理、提交、展示和评价要求", "归纳课程项目涉及的核心知识、技术路线和质量标准"])
                ability = _numbered(["能够提交可运行、结构完整且说明清晰的课程综合作品", "能够进行成果陈述、问题说明、自我评价并依据意见修改完善"])
                ideological = _numbered(["坚持诚信提交，能够如实说明个人贡献、素材来源和作品完成过程", "尊重评价规则和他人成果，以认真负责的态度完成课程总结"])
                quality = _numbered(["提升成果表达、沟通答辩和接受评价的能力", "形成复盘问题、制定改进计划和持续学习的习惯"])
            else:
                refs = _representative_resources([item for item in resources if item["project_hint"] == task["chapter"]])
                knowledge_items = [_knowledge_statement(skill) for skill in skills[:2]]
                if len(skills) > 2:
                    knowledge_items.append(f"掌握{'、'.join(skills[2:])}之间的配合关系和使用要求")
                knowledge = _numbered(knowledge_items or [f"掌握{task['chapter']}所需的核心知识和质量要求"])
                ability = _numbered([
                    f"能够运用{'、'.join(skills) or '本次知识技能'}完成“{task['chapter']}”对应页面或功能",
                    "能够借助配套示例和开发工具运行、检查并定位常见问题",
                    "能够依据任务要求整理代码、素材和运行结果并进行阶段提交",
                ])
                context = task["title"].replace("：", " - ", 1)
                ideological = _ideological_goal(f"{task['chapter']} {detail}", offering["major"], context)
                quality = _quality_goal(f"{task['chapter']} {detail}", implementation, context)
            db.execute(
                "UPDATE tasks SET resource_refs=?,knowledge_goal=?,ability_goal=?,ideological_goal=?,quality_goal=? WHERE id=?",
                (json.dumps(refs, ensure_ascii=False), knowledge, ability, ideological, quality, task["id"]),
            )
            updated += 1
        db.commit()
    return updated
