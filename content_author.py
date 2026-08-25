import json
import re
from pathlib import Path

import store
from ai.ai_router import ask_result


ACTIVITY_KEYS = (
    "教学导入", "任务1", "任务2", "任务3",
    "课堂小结", "课后作业", "教学反思",
)
AI_SYSTEM_PROMPT = (
    "你是高职院校课程标准和教学设计专家。只能使用输入资料中的事实，"
    "不得编造课程、项目、软件、数据或教材内容。只返回符合指定结构的JSON，"
    "不要返回Markdown、解释或代码围栏。"
)


def _list(text):
    return [line.strip().lstrip("0123456789.、 ") for line in str(text or "").splitlines() if line.strip()]


def _join(items, limit=8):
    return "、".join(list(dict.fromkeys(item for item in items if item))[:limit]) if items else ""


def _repair_json(text):
    """尝试修复AI返回的常见JSON语法错误。"""
    repaired = text
    repaired = re.sub(r':\s*\[\[([^\[\]"]*)\]\]', r': "[[\1]]"', repaired)
    repaired = re.sub(r",\s*([}\]])", r"\1", repaired)
    repaired = re.sub(r"//[^\n]*", "", repaired)
    repaired = repaired.replace("'", '"')
    repaired = re.sub(r':\s*,', ': "",', repaired)
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
    if not isinstance(result, dict):
        raise ValueError("AI返回的JSON顶层必须是对象。")
    return result


def _validate_overview(data):
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


def _validate_activity(data):
    for key in ACTIVITY_KEYS:
        value = data.get(key)
        if not isinstance(value, str) or len(value.strip()) < 20:
            raise ValueError(f"AI教学活动字段不完整：{key}")


def _ask_json(prompt, validator, attempts=3):
    errors = []
    for attempt in range(attempts):
        suffix = ""
        if attempt > 0 and errors:
            suffix = f"\n上次返回未通过结构校验：{errors[-1]}。请补全缺失字段、确保每个字段内容完整后重新输出完整JSON。"
        result = ask_result(prompt + suffix, system=AI_SYSTEM_PROMPT, show_details=False)
        if not result.get("success"):
            errors.append(result.get("error", "未知错误"))
            continue
        try:
            data = _parse_json(result.get("content", ""))
            validator(data)
            source = f"AI草稿/{result.get('source', '未知来源')}/{result.get('model', '未知模型')}"
            return data, source
        except (ValueError, json.JSONDecodeError) as error:
            errors.append(str(error))
    raise RuntimeError(f"AI内容生成失败：{'；'.join(errors)}")


def _ai_overview(identity, model, facts):
    evidence = [
        {"类型": fact["fact_type"], "内容": fact["fact_value"][:300], "位置": fact.get("locator", "")}
        for fact in facts[:16] if fact.get("fact_value")
    ]
    talent = model.get("talent_plan") or {}
    payload = {
        "课程": identity,
        "教材项目": model.get("projects", []),
        "知识体系": model.get("knowledge_system", []),
        "课程标准": model.get("standards", []),
        "技术工具": model.get("tools_technology", []),
        "人才培养方案": {
            "方案专业": talent.get("major"), "年级": talent.get("cohort"),
            "培养目标": (talent.get("goals") or "")[:400],
            "职业面向岗位群": (talent.get("orientation") or {}).get("job_positions", []),
            "本课程设置": talent.get("course_info"),
        },
        "资源证据": evidence,
    }
    prompt = """
请生成课程标准中的课程性质、课程目标和课程设计总体思路。内容必须结合具体专业、
教材项目、知识技能、人才培养方案（培养目标、培养规格、职业面向岗位群）和资源证据，
避免可以套用于任意课程的空泛表述。课程性质中应说明本课程在专业岗位群中的定位。
JSON结构必须为：
{
  "course_nature": ["三段正文"],
  "course_goals": {
    "知识目标": ["至少3条"], "能力目标": ["至少3条"],
    "思政目标": ["至少3条"], "素质目标": ["至少3条"]
  },
  "course_design": ["三段正文"]
}
输入资料：
""" + json.dumps(payload, ensure_ascii=False)
    return _ask_json(prompt, _validate_overview)


def _ai_activity(task, identity, facts):
    related = [
        {"类型": fact["fact_type"], "内容": fact["fact_value"][:400], "位置": fact.get("locator", "")}
        for fact in facts if fact.get("project_hint") == task["chapter"] and fact.get("fact_value")
    ][:8]
    refs = json.loads(task["resource_refs"] or "[]")
    payload = {
        "课程": identity,
        "任务": {key: task[key] for key in (
            "seq", "chapter", "title", "hours", "knowledge_goal",
            "ability_goal", "ideological_goal", "quality_goal",
        )},
        "关联资源": refs,
        "资源证据": related,
    }
    skeleton = json.dumps({key: "（不少于60字的完整内容）" for key in ACTIVITY_KEYS},
                          ensure_ascii=False, indent=2)
    prompt = f"""
请为这一次课生成可直接套入教案的教学组织。三个课堂任务必须体现递进关系，
并明确提问、知识或技术分析、教师活动、学生活动、操作练习、检查标准、结论。
不得把目标原句简单改写为教学过程。

【输出格式——严格遵守】
必须输出一个JSON对象，恰好{len(ACTIVITY_KEYS)}个键，每个键的值都是字符串，
禁止嵌套对象、禁止数组、禁止输出多个JSON对象。
骨架如下（把每个占位文本替换为完整内容）：
{skeleton}

【内容要求】
任务1/任务2/任务3的字符串中需依次写明：提问、知识或技术分析、教师活动、学生活动、操作练习、检查标准、结论。

输入资料：
""" + json.dumps(payload, ensure_ascii=False)
    return _ask_json(prompt, _validate_activity)


def _put(db, offering_id, document_type, section_key, title, content, evidence,
         repeat_key="", authoring_status="结构化生成"):
    db.execute(
        "INSERT INTO authored_sections (offering_id,document_type,section_key,repeat_key,title,content_json,evidence_json,authoring_status,review_status,generated_at) "
        "VALUES (?,?,?,?,?,?,?,?,'待检查',CURRENT_TIMESTAMP) "
        "ON CONFLICT(offering_id,document_type,section_key,repeat_key) DO UPDATE SET title=excluded.title,content_json=excluded.content_json,"
        "evidence_json=excluded.evidence_json,authoring_status=excluded.authoring_status,review_status='待检查',generated_at=CURRENT_TIMESTAMP",
        (offering_id, document_type, section_key, repeat_key, title,
         json.dumps(content, ensure_ascii=False), json.dumps(evidence, ensure_ascii=False), authoring_status),
    )


_DOMAIN_KEYWORDS = (
    ("数据分析", "数据采集、分析与决策支持"),
    ("H5", "网页设计与前端制作"),
    ("网页", "网页设计与前端制作"),
    ("Python", "程序设计与数据处理"),
    ("图形图像", "图像处理与视觉设计"),
    ("Photoshop", "图像处理与视觉设计"),
    ("标注", "数据整理与标注"),
    ("新媒体", "新媒体内容制作与运营推广"),
)


def _course_domain(identity):
    name = identity.get("course_name") or ""
    for keyword, domain in _DOMAIN_KEYWORDS:
        if keyword.lower() in name.lower():
            return domain
    return "本课程所面向业务"


def _plan_paragraphs(identity, model):
    """依据人才培养方案生成课程定位段落（培养目标/培养规格/课程归属）。"""
    plan = model.get("talent_plan") or {}
    goals = (plan.get("goals") or "").strip()
    if not goals:
        return []
    cohort = plan.get("cohort") or ""
    course_info = plan.get("course_info") or {}
    category = course_info.get("category") or identity.get("course_type") or ""
    major = identity.get("major") or plan.get("major") or ""
    plan_major = plan.get("major") or major
    if plan_major != major:
        plan_ref = f"{plan_major}专业{cohort}人才培养方案（本专业参照执行）"
    else:
        plan_ref = f"{major}专业{cohort}人才培养方案"
    first_goal = goals.split("。")[0] + "。"
    if len(first_goal) > 220:
        first_goal = first_goal[:220] + "……。"
    hours_desc = ""
    if course_info.get("hours"):
        credit_part = f"、{course_info['credits']}学分" if course_info.get("credits") else ""
        hours_desc = f"方案规定本课程{course_info['hours']}学时{credit_part}，"
    paragraphs = [
        f"{plan_ref}确定的培养目标是：{first_goal}本课程作为{category or '专业课程'}列入专业课程体系，{hours_desc}"
        f"承担着将专业培养目标中与本课程相关的知识、技能和素养要求转化为具体教学内容的任务。"
    ]
    positions = (plan.get("orientation") or {}).get("job_positions") or []
    if positions:
        paragraphs.append(
            f"方案确定的职业面向为{_join(positions, 6)}等岗位（群），本课程教学内容与其中涉及"
            f"{_course_domain(identity)}工作的岗位能力要求直接对应，为学生胜任相关岗位奠定基础。"
        )
    specs = (plan.get("specs") or "").strip()
    ability_match = re.search(r"3[.、．]\s*能力(.+?)(?:$|4[.、．])", specs, re.S)
    if ability_match:
        items = re.findall(r"（\d+）([^；;]{10,80})", ability_match.group(1))
        picked = [re.sub(r"\s+", "", item) for item in items[:2]]
        if picked:
            knowledge = model.get("knowledge_system") or []
            focus = _join(knowledge[:6]) if knowledge else "本课程核心内容"
            paragraphs.append(
                f"对照培养规格中“{picked[0]}”"
                + (f"“{picked[1]}”" if len(picked) > 1 else "")
                + f"等能力要求，本课程围绕{focus}组织教学，使学生在完成项目任务的过程中形成规范操作、数据思维和成果交付能力，落实专业培养规格。"
            )
    return paragraphs


def _nature(identity, model):
    name, major, course_type = identity["course_name"], identity["major"], identity["course_type"]
    nature = (identity.get("course_nature") or "必修课").removesuffix("课")
    department = identity.get("department", "")
    assessment = identity.get("assessment_type", "")
    teaching_mode = identity.get("teaching_mode", "")
    lecture_hours = identity.get("lecture_hours", 0)
    practice_hours = identity.get("practice_hours", 0)
    projects = [p["title"] for p in model["projects"] if p["title"] != "综合评价与课程总结"]
    knowledge = model["knowledge_system"]
    prerequisite = model["course_links"].get("prerequisite")
    followup = model["course_links"].get("followup")
    dept_prefix = f"{department}" if department else ""
    hours_desc = ""
    if lecture_hours and practice_hours:
        hours_desc = f"其中理论{lecture_hours}学时、实践{practice_hours}学时，"
    elif lecture_hours:
        hours_desc = f"其中理论{lecture_hours}学时，"
    assessment_desc = f"考核方式为{assessment}。" if assessment and assessment != "期末考核" else ""
    mode_desc = f"采用{teaching_mode}方式进行教学。" if teaching_mode else ""
    paragraphs = [
        f"《{name}》是{dept_prefix}{major}专业开设的{nature}{course_type}。{hours_desc}课程以{_join(projects, 5)}等教材项目为载体，组织学生完成从知识学习、技术练习到项目成果交付的递进训练，具有较强的实践性、综合性和职业性。{assessment_desc}{mode_desc}",
    ]
    paragraphs.extend(_plan_paragraphs(identity, model))
    paragraphs.append(
        f"课程围绕{_join(knowledge, 10)}等内容展开，使学生能够分析项目需求，选择适当技术完成制作、运行检查、问题修改和规范提交；同时将技术规范、用户体验、网络安全、个人信息保护、数字版权和诚信交付落实到项目过程与成果评价中。"
    )
    if prerequisite and followup:
        paragraphs.append(f"本课程以《{prerequisite}》形成的基础知识与操作能力为学习起点，并为《{followup}》中的综合项目实施奠定知识、技术和项目协作基础。")
    elif prerequisite:
        paragraphs.append(f"本课程以《{prerequisite}》形成的基础知识与操作能力为学习起点；后续课程关系尚待教师确认。")
    elif followup:
        paragraphs.append(f"本课程为《{followup}》中的综合项目实施奠定知识、技术和项目协作基础；先导课程关系尚待教师确认。")
    else:
        paragraphs.append("本课程的先导、后续课程关系尚未提供，不在正文中推测具体课程名称。")
    return paragraphs


def _course_goals(model):
    knowledge = model["knowledge_system"]
    abilities = []
    for project in model["projects"][:6]:
        if project["title"] != "综合评价与课程总结":
            abilities.append(f"能够运用{_join(project['knowledge_skills'], 4)}完成“{project['title']}”并检查、修改和提交成果")
    return {
        "知识目标": [f"了解{_join(knowledge[:4])}的基本概念、作用和应用场景", f"掌握{_join(knowledge[4:12])}的规则、使用方法及配合关系", f"熟悉{_join(model['standards'])}以及项目测试、优化和成果交付要求"],
        "能力目标": abilities[:4],
        "思政目标": ["在数字作品制作中坚持正确价值导向，遵守网络安全、个人信息保护和数字版权要求", "在代码、素材和成果提交中坚持诚实守信，能够如实说明素材来源与个人贡献", "理解数字技术服务专业实践和社会需求的责任，形成用户意识与质量意识"],
        "素质目标": ["形成依据规范编码、及时测试、记录问题和持续改进的职业习惯", "提升项目分工、沟通反馈、成果展示和按时交付能力", "形成主动查阅资料、比较技术方案和持续学习新技术的意识"],
    }


def _course_design(identity, model):
    plan = model.get("talent_plan") or {}
    major = identity.get("major") or ""
    if plan.get("goals"):
        plan_major = plan.get("major") or major
        cohort = plan.get("cohort") or ""
        if plan_major and plan_major != major:
            plan_basis = f"{cohort}{plan_major}专业人才培养方案（本专业参照执行）"
        else:
            plan_basis = f"{major}专业{cohort}人才培养方案"
        basis_text = f"{plan_basis}确定的培养目标与培养规格"
    else:
        basis_text = "专业人才培养的技术技能要求"
    return [
        f"课程依据{basis_text}、{identity['term']}教学安排和本学期指定教材资源进行设计，以{_join([p['title'] for p in model['projects'][:6]])}等完整项目贯穿教学。",
        f"内容组织体现新标准、新技术、新工艺和新方法：标准要求包括{_join(model['standards'])}；技术内容包括{_join(model['tools_technology'])}；项目按照{_join(model['work_process'])}组织；教学采用{_join(model['teaching_methods'])}。",
        "课程遵循由单项知识技能到综合项目实施、由教师示范到学生独立完成的递进规律，将知识目标、能力目标、思政目标和素质目标落实到任务要求、课堂活动、成果检查和过程性评价中。",
    ]


def _implementation(identity, model):
    knowledge = model["knowledge_system"]
    tools = model["tools_technology"]
    return {
        "教师知识能力要求": [
            "任课教师应坚持正确政治方向，强化课程思政意识，将职业道德、网络文明、社会责任和诚信评价融入专业教学全过程。",
            f"任课教师应具备扎实的{identity['course_name']}专业基础，熟悉{_join(knowledge[:12])}，能够依据教材项目解释原理、示范操作并处理学生实践中的典型问题。",
            f"任课教师应能够熟练使用{_join(tools)}开展项目实践、运行测试和成果检查，持续关注课程涉及的标准、技术和工具更新。",
            f"任课教师应能按照{_join(model['work_process'])}把教材项目转化为教学任务，实施{_join(model['teaching_methods'])}，并兼顾网络安全、数据安全、知识产权、创新意识和分层指导。",
        ],
        "教材与课程资源": {
            "教材": identity["textbook_version"],
            "教学资料": "使用本学期指定教材包中的课程PPT、实训文档、项目源码、图片与音视频素材；教学时按项目映射调用，不以历史课程文档作为内容来源。",
            "开发利用": "围绕各项目建设课件要点、操作演示、源码片段、任务单、评价清单和优秀作品库；资源更新应记录来源、适用项目和版本，并检查版权、个人信息和运行安全。",
        },
    }


def _assessment(model):
    projects = [p["title"] for p in model["projects"] if p["title"] != "综合评价与课程总结"]
    scheme = model.get("assessment_scheme", {})
    process_total = float(scheme.get("process_total", 40))
    final_total = float(scheme.get("final_total", 60))
    component_names = [f"{item['component_name']}{float(item['weight']):g}分" for item in scheme.get("components", [])]
    return {
        "原则": "采用过程性评价与终结性评价相结合的方式，同时评价知识理解、技术操作、项目成果、规范意识、诚信与协作表现。",
        "过程性评价": {"权重": f"{process_total:g}%", "内容": component_names or ["签到与课堂参与", "技术练习与阶段作业"], "抽取项目": projects[:4]},
        "终结性评价": {"权重": f"{final_total:g}%", "内容": ["课程综合作品", "源文件或源代码与素材说明", "成果展示与答辩", "问题修改记录和课程总结"]},
    }


def _domain_context(identity, focus):
    course = identity["course_name"]
    if course == "Python程序设计":
        return {
            "artifact": "Python程序及运行结果", "tool": "Python解释器、IDLE或当前主流Python开发环境",
            "operation": "分析问题、设计算法，编写并运行Python代码，利用报错信息、断点或输出结果定位问题",
            "issues": "语法、数据类型、控制流程、函数调用、文件路径和运行异常",
            "quality": "程序可运行、结果正确、代码结构清晰、命名与注释规范，并具有必要的异常处理",
            "action": "编程实现、测试调试和程序交付", "homework": "源代码、运行结果截图、测试数据和问题修改记录",
        }
    if course in ("商务数据分析", "商务分析与决策", "商务数据分析与应用"):
        return {
            "artifact": "数据分析结果、图表和业务结论", "tool": "教材指定的数据分析软件、电子表格或Python分析环境",
            "operation": "明确业务问题，整理数据，选择指标与分析方法，完成计算或可视化并解释结果",
            "issues": "数据缺失、类型错误、指标口径、公式或代码、图表选择和结论证据不足",
            "quality": "数据处理可追溯、指标口径一致、计算结果正确、图表清晰、结论有数据依据",
            "action": "数据处理、分析计算、结果解释和决策呈现", "homework": "数据文件、分析过程、图表、业务结论和问题修改记录",
        }
    if course == "图形图像设计":
        return {
            "artifact": "图像设计作品及可编辑源文件", "tool": "Adobe Photoshop及教材配套素材",
            "operation": "分析视觉需求，选择工具与图层组织方式，完成选区、调整、合成或特效制作并检查细节",
            "issues": "分辨率与色彩模式、选区边缘、图层顺序、蒙版、参数设置、素材版权和导出格式",
            "quality": "构图与色彩协调、边缘处理细致、图层命名清楚、源文件可编辑、导出规格正确",
            "action": "图像处理、视觉设计、细节修整和规范输出", "homework": "PSD源文件、导出作品、素材来源说明和修改前后对比",
        }
    if course == "数据标注":
        return {
            "artifact": "符合规范的标注数据集和质量记录", "tool": "教材指定的标注平台或标注工具",
            "operation": "解读标注规范，配置任务，完成样本标注、复核、纠错和数据导出",
            "issues": "标签定义、边界精度、漏标错标、一致性、个人信息与数据安全、文件格式",
            "quality": "标签使用正确、边界准确、漏标错标率符合要求、记录完整、数据安全合规",
            "action": "样本识别、规范标注、质量复核和数据交付", "homework": "标注文件、质量检查表、典型问题截图和纠错记录",
        }
    if course == "新媒体平台运营与推广":
        return {
            "artifact": "新媒体运营方案、内容作品或数据复盘报告", "tool": "教材案例、主流新媒体平台和数据分析工具",
            "operation": "分析用户与传播目标，策划内容和渠道，设计执行方案，依据平台数据开展效果复盘",
            "issues": "用户定位、内容价值、平台规则、广告合规、版权、互动转化和数据指标解释",
            "quality": "定位明确、内容合规、方案可执行、素材来源清楚、指标选择合理、复盘建议有依据",
            "action": "用户分析、内容策划、运营执行和效果复盘", "homework": "运营方案、内容样稿、发布或模拟数据、复盘结论和修改记录",
        }
    return {
        "artifact": "H5页面、代码及运行效果", "tool": "代码编辑器和浏览器开发者工具",
        "operation": "编写页面结构、设置样式与交互，运行页面并使用浏览器开发者工具检查结果",
        "issues": "HTML语义结构、CSS选择器与布局、资源路径、兼容性、响应式表现和代码规范",
        "quality": "页面可运行、结构语义清晰、视觉效果符合要求、不同视口显示正常、代码与素材规范",
        "action": "页面设计、编码实现、调试测试和成果交付", "homework": "源代码、页面截图、素材说明和问题修改记录",
    }


def _activity(task, resources, facts, identity):
    skills = [item.strip() for item in re.split(r"[、；]", task["title"].split("：", 1)[-1]) if item.strip()]
    ability = _list(task["ability_goal"])
    refs = json.loads(task["resource_refs"] or "[]")
    names = [Path(item).name for item in refs if ":" in item or "\\" in item]
    fact_text = [f["fact_value"] for f in facts if f.get("project_hint") == task["chapter"] and f["fact_type"] in ("ppt_slide", "source_structure", "source_excerpt")]
    source_focus = next((
        text for text in fact_text
        if 25 <= len(text) <= 500
        and "企业级卓越人才培养" not in text
        and "……" not in text and "等等" not in text
        and "口令：" not in text and "滨海迅腾科技集团" not in text
        and "___PPT" not in text
    ), "")
    if task["chapter"] == "综合评价与课程总结":
        return {
            "教学导入": "公布课程综合作品提交清单、展示顺序和评价量表，说明作品可运行性、完整性、规范性、创新性以及汇报答辩要求；学生对照清单完成提交前自查。",
            "任务1": "教学活动名称：成果验收与规范检查。教师说明文件结构、运行效果、素材来源、代码注释和过程记录的验收方法；学生逐项检查并修复影响运行与提交的问题。结论：综合作品能够正常运行，提交材料完整且来源说明清楚。德育渗透：强调原创、诚信和数字版权。板书：必交材料｜运行检查｜版权说明｜一票否决项。",
            "任务2": "教学活动名称：作品展示与答辩。学生演示核心页面或功能，说明需求、技术路线、关键实现、调试过程和改进之处；教师与同伴依据量表提问、评分并提出修改意见。结论：学生能够用成果和过程证据说明个人贡献与目标达成情况。德育渗透：尊重他人成果，客观评价，如实陈述。板书：展示结构｜答辩问题｜评分依据。",
            "任务3": "教学活动名称：课程复盘与迭代提交。按知识体系和项目流程梳理共性难点，学生根据评价意见完成最后修改，提交最终版本并形成个人课程总结。结论：学生能够归纳技术路线、分析问题成因并提出后续学习计划。德育渗透：培养责任担当、持续改进和终身学习意识。板书：知识图谱｜典型问题｜修改清单｜学习计划。",
            "课堂小结": "汇总综合作品达成情况和高频问题，明确最终提交要求；学生完成自评、互评与个人贡献说明，教师反馈课程目标达成情况。",
            "课后作业": "在规定时间内提交修改后的课程综合作品、运行截图、素材来源说明、问题修改记录和个人课程总结。",
            "教学反思": "依据作品通过率、量表得分分布、答辩表现和诚信提交情况，分析课程目标与各项目衔接效果；记录需调整的学时、示例、评价指标和分层任务，为下一轮课程改进提供依据。",
        }
    focus = _join(skills, 4)
    domain = _domain_context(identity, focus)
    content = {
        "教学导入": f"展示“{task['chapter']}”对应的{domain['artifact']}或教材案例，联系前一任务提出本次需要解决的实际问题；学生观察、比较并说明初步方案，教师归纳教学目标和验收要求。",
        "任务1": f"教学活动名称：知识原理与方案分析。\n提问：围绕{focus}提出“它解决什么问题、适用于什么场景、怎样判断结果正确”，学生结合教材案例说明初步判断。\n概念与技术分析：依据教材项目和对应PPT，分析核心概念、操作规则、方法之间的配合关系、适用条件及常见错误。" + (f"课件要点：{source_focus}\n" if source_focus else "") + f"练习技术：学生辨析正误示例，标注{domain['issues']}方面的问题并说明修改依据。\n结论：能够解释本次关键知识，并依据任务需求选择适当方法。\n适时德育渗透：强调遵守职业规范、保护数据与知识产权、如实记录过程和结果。\n板书设计：任务目标｜核心规则｜适用场景｜易错点｜验收标准。",
        "任务2": f"教学活动名称：资源研读、操作演示与技术练习。\n提问：对照教材示例的预期结果与实际结果，判断{domain['issues']}中可能存在的问题。\n技术分析：结合{_join(names[:3]) or '教材配套PPT、实训资料和项目文件'}定位关键步骤、参数、代码或操作，说明其与成果之间的对应关系。\n实验/操作演示：教师使用{domain['tool']}分步演示如何{domain['operation']}，完整呈现保存、检查、定位问题和修改验证的过程。\n练习技术：学生同步操作、观察和修改，保留关键过程、结果截图与错误修改记录。\n结论：学生能够{_join([item.removeprefix('能够') for item in ability[:2]], 2)}。\n适时德育渗透：强调规范操作、数据与素材安全、严谨验证和独立完成。\n板书设计：文件与资源｜关键步骤/参数｜结果现象｜检查路径｜修改结果。",
        "任务3": f"教学活动名称：阶段任务实施、检查与评价。\n提问：对照“{task['chapter']}”要求，明确本阶段必须完成的{domain['artifact']}及质量标准。\n技术分析：教师按照{domain['action']}的顺序梳理实施步骤，指出与前序任务的衔接点及{domain['issues']}方面的风险。\n练习技术/项目实施：学生独立或协作完成阶段成果，使用{domain['tool']}开展检查，依据评价清单进行自评、互评和迭代修改；教师针对共性问题作现场演示。\n结论：{domain['quality']}。\n适时德育渗透：落实分工责任、诚信提交、用户意识、质量意识和精益求精的职业精神。\n板书设计：实施步骤｜质量检查点｜典型问题｜修改清单｜提交要求。",
        "课堂小结": f"围绕“{task['chapter']}”回顾知识原理、操作步骤、调试方法和质量标准。学生展示代表性成果并说明问题解决过程，教师归纳共性问题，形成修改清单。",
        "课后作业": f"继续完善“{task['chapter']}”本次成果，提交{domain['homework']}；选择一个关键细节自主优化并说明修改依据。",
        "教学反思": "依据任务完成率、典型错误、课堂参与和成果评价记录，分析目标达成情况；重点检查任务分解、演示节奏、分层指导和思政融入是否有效，并据此补充示例、调整练习和安排针对性复习。",
    }
    return content


def author_course_content(offering_id):
    row = store.rows("SELECT model_json FROM course_content_models WHERE offering_id=?", (offering_id,))
    if not row:
        raise ValueError("尚未建立课程语义模型。")
    model = json.loads(row[0]["model_json"])
    identity = model["identity"]
    tasks = store.rows("SELECT * FROM tasks WHERE offering_id=? ORDER BY seq", (offering_id,))
    facts = store.rows("SELECT project_hint,fact_type,fact_value,locator FROM resource_facts WHERE offering_id=?", (offering_id,))
    evidence = ["course_identity", "knowledge_system", "projects", "standards", "tools_technology", "work_process"]
    try:
        overview, overview_status = _ai_overview(identity, model, facts)
    except (RuntimeError, ValueError) as error:
        print(f"  [兜底] 课程概述AI生成失败，改用规则生成：{error}")
        overview = {
            "course_nature": _nature(identity, model),
            "course_goals": _course_goals(model),
            "course_design": _course_design(identity, model),
        }
        overview_status = "结构化生成"
    activities = {}
    for task in tasks:
        try:
            activities[str(task["seq"])] = _ai_activity(task, identity, facts)
        except (RuntimeError, ValueError) as error:
            print(f"  [兜底] 第{task['seq']}次课教学组织AI生成失败，改用规则生成：{error}")
            activities[str(task["seq"])] = (_activity(task, [], facts, identity), "结构化生成")
    with store.connect() as db:
        db.execute("DELETE FROM authored_sections WHERE offering_id=?", (offering_id,))
        _put(db, offering_id, "课程标准", "course_nature", "一、课程性质", overview["course_nature"], evidence, authoring_status=overview_status)
        _put(db, offering_id, "课程标准", "course_goals", "二、课程目标", overview["course_goals"], evidence, authoring_status=overview_status)
        _put(db, offering_id, "课程标准", "course_design", "课程设计总体思路", overview["course_design"], evidence, authoring_status=overview_status)
        _put(db, offering_id, "课程标准", "content_hours", "课程内容划分及课时分配", model["projects"], ["projects", "tasks"])
        _put(db, offering_id, "课程标准", "learning_scenarios", "四、学习情境描述", [{"项目": p["title"], "学时": p["hours"], "知识技能": p["knowledge_skills"], "预期成果": p["expected_outcome"]} for p in model["projects"]], ["projects", "tasks"])
        _put(db, offering_id, "课程标准", "assessment", "课程考核评价", _assessment(model), ["assessment_evidence", "projects", "tasks"])
        implementation = _implementation(identity, model)
        _put(db, offering_id, "课程标准", "teacher_requirements", "五、课程实施", implementation["教师知识能力要求"], ["knowledge_system", "tools_technology", "standards", "work_process"])
        _put(db, offering_id, "课程标准", "course_resources", "教材编写选用与课程资源", implementation["教材与课程资源"], ["course_identity", "resource_summary", "projects"])
        _put(db, offering_id, "授课计划", "schedule_rows", "授课计划明细", [{k: task[k] for k in ("seq", "chapter", "title", "hours", "theory_hours", "practice_hours", "week_no", "lesson_date")} for task in tasks], ["tasks", "sessions"])
        _put(db, offering_id, "教学设计", "course_goals", "课程目标设计", overview["course_goals"], evidence, authoring_status=overview_status)
        _put(db, offering_id, "教学设计", "implementation_conditions", "课程教学实施条件与教学资源", implementation, ["course_identity", "resource_summary", "tools_technology"])
        first = tasks[0] if tasks else None
        if first:
            first_activity, first_status = activities[str(first["seq"])]
            _put(db, offering_id, "教学设计", "first_lesson_outline", "第一节课设计梗概", first_activity, [f"task:{first['id']}", *json.loads(first["resource_refs"] or "[]")], authoring_status=first_status)
        for task in tasks:
            refs = json.loads(task["resource_refs"] or "[]")
            activity, activity_status = activities[str(task["seq"])]
            content = {
                "基本信息": {"周次": task["week_no"], "日期": task["lesson_date"], "项目": task["chapter"], "任务": task["title"], "学时": task["hours"]},
                "教学目标": {"知识目标": _list(task["knowledge_goal"]), "能力目标": _list(task["ability_goal"]), "思政目标": _list(task["ideological_goal"]), "素质目标": _list(task["quality_goal"])},
                "教学组织": activity,
            }
            _put(db, offering_id, "教学设计", "unit_design", f"第{task['seq']}次教学设计", content, [f"task:{task['id']}", *refs], str(task["seq"]), activity_status)
        db.commit()
    return len(tasks) + 12
