import json
import math
import hashlib
from pathlib import Path

import store
from assessment_scheme import get_scheme


CHINESE_NUMBERS = "零一二三四五六七八九十"


def _unique_parts(text):
    result = []
    for raw in str(text or "").replace(";", "；").split("；"):
        part = raw.strip()
        if part and part not in result:
            result.append(part)
    return result


def _group(parts, count):
    if not parts or count <= 0:
        return []
    groups = []
    for index in range(count):
        start = math.floor(index * len(parts) / count)
        end = math.floor((index + 1) * len(parts) / count)
        selected = parts[start:end] or [parts[min(start, len(parts) - 1)]]
        groups.append("、".join(selected))
    return groups


def _subscenarios(unit):
    hours = int(unit["suggested_hours"])
    skills = _unique_parts(unit["source_skills"])
    if unit["project_title"] == "综合评价与课程总结":
        groups = _group(skills, max(1, hours // 2))
        return [
            {"seq": index + 1, "title": title, "hours": 2, "kind": "总结评价"}
            for index, title in enumerate(groups)
        ]
    implementation_hours = max(2, hours // 2)
    knowledge_hours = hours - implementation_hours
    group_count = min(len(skills), knowledge_hours) if skills else knowledge_hours
    groups = _group(skills, group_count)
    group_hours = [1] * len(groups)
    extra_hours = knowledge_hours - sum(group_hours)
    allocation_order = sorted(range(len(groups)), key=lambda index: (abs(index - (len(groups) - 1) / 2), index))
    for index in range(extra_hours):
        group_hours[allocation_order[index % len(allocation_order)]] += 1
    items = [
        {"seq": index + 1, "title": title, "hours": group_hours[index], "kind": "知识技能"}
        for index, title in enumerate(groups)
    ]
    items.append({"seq": len(items) + 1, "title": "任务实施", "hours": implementation_hours, "kind": "任务实施"})
    return items


def _course_nature(offering, units):
    titles = "、".join(unit["project_title"] for unit in units[:6])
    prerequisite = offering.get("prerequisite_courses") or "相关专业基础课程"
    followup = offering.get("followup_courses") or "相关综合实践课程"
    nature = (offering.get("course_nature") or "必修课").removesuffix("课")
    return [
        f"《{offering['course_name']}》是{offering['major']}专业的一门{nature}{offering['course_type']}，课程面向与本专业相关的技术应用、业务实施和项目协作岗位，具有实践性、职业性和综合性。",
        f"课程以本学期指定教材及配套资源为内容依据，围绕{titles}等典型学习项目组织教学，使学生掌握完成课程项目所需的基础知识、操作方法、质量规范和问题解决方法，培养规范实施、检查改进与成果交付能力。",
        f"本课程的先导课程为《{prerequisite}》，为课程学习提供必要基础；后续课程为《{followup}》，进一步培养学生综合运用本课程知识和技能解决实际问题的能力。",
    ]


def _course_goals(offering, units):
    focuses = []
    for unit in units:
        focuses.extend(_unique_parts(unit.get("source_skills")))
    focus_text = "、".join(focuses[:8]) or "课程核心知识与操作规范"
    return {
        "knowledge": [
            f"了解《{offering['course_name']}》相关行业的发展、岗位任务及应用场景。",
            f"掌握{focus_text}等核心知识。",
            "熟悉课程项目的工作流程、质量标准、安全要求和成果交付规范。",
            "了解本领域新标准、新技术、新工艺和新方法的基本应用要求。",
        ],
        "ability": [
            "具有分析项目需求、分解工作任务并制定实施方案的能力。",
            "具有运用课程知识和工具完成典型项目任务的能力。",
            "具有检查成果、定位问题、调试改进和规范交付的能力。",
            "具有查阅技术资料、学习新技术并迁移应用的能力。",
        ],
        "ideological": [
            "坚定职业理想，理解专业技术服务经济社会发展的责任。",
            "践行社会主义核心价值观，形成诚实守信、认真负责的职业态度。",
            "增强文化自信，在项目内容与成果表达中坚持正确价值导向。",
            "强化安全、法治、知识产权和职业道德意识，自觉遵守行业规范。",
        ],
        "quality": [
            "形成创新意识和持续学习意识，主动关注行业发展与技术更新。",
            "形成规范操作、安全实施、细节检查和精益求精的质量意识。",
            "形成团队协作、沟通表达、任务管理和按期交付的职业素养。",
            "形成独立分析、主动实践、自我评价和持续改进的学习习惯。",
        ],
    }


def _course_design(offering, units):
    standards = "；".join(filter(None, (unit.get("new_standards") for unit in units)))
    technology = "；".join(filter(None, (unit.get("new_technology") for unit in units)))
    process = "；".join(filter(None, (unit.get("new_process") for unit in units)))
    methods = "；".join(filter(None, (unit.get("new_methods") for unit in units)))
    return [
        f"课程设计面向{offering['major']}专业相关岗位的典型工作任务，以本学期教材项目为载体。新标准方面融入{standards or '现行职业规范、质量标准、安全要求和知识产权要求'}；新技术方面融入{technology or '教材配套工具及本领域主流技术'}。",
        f"课程按{process or '需求分析、任务分解、操作实施、检查测试、改进优化和成果交付'}的工作过程组织教学，采用{methods or '项目教学、任务驱动、案例分析、分层指导、合作学习和过程性评价'}等方法，实现知识、技能、思政和职业素质协同培养。",
        f"课程设置以{offering['term']}教学安排表确定的课程类型、学时学分和开课班级为执行依据，以指定教材、PPT、实训及源码等资源为内容依据。课程与先导课程《{offering.get('prerequisite_courses') or '相关专业基础课程'}》衔接，为后续课程《{offering.get('followup_courses') or '相关综合实践课程'}》奠定基础。",
    ]


def _teacher_requirements(offering, units):
    focus = "、".join(
        part for unit in units[:5] for part in _unique_parts(unit.get("source_skills"))[:2]
    ) or "课程核心知识与技能"
    return [
        "任课教师应坚持正确政治方向，强化课程思政意识，将价值引领、职业道德和社会责任融入专业教学全过程。",
        f"任课教师应具有扎实的《{offering['course_name']}》专业基础，熟悉{focus}，持续学习行业新标准、新技术、新工艺和新方法，并及时更新教学内容。",
        "任课教师应具有相应项目实践和问题解决能力，能够把真实或仿真岗位任务转化为教学项目，指导学生完成实施、检查、优化和交付。",
        "任课教师应具备项目教学、任务驱动、分层指导和多元评价能力，并加强安全、创新、知识产权和终身学习等方面的认知。",
    ]


def _signature(offering, units):
    payload = {"offering": {key: offering.get(key) for key in (
        "course_name", "term", "major", "course_nature", "course_type", "prerequisite_courses", "followup_courses",
        "textbook_version", "textbook_path", "total_hours")}, "units": units}
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def build_course_standard_model(offering_id):
    offering = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))[0]
    units = store.rows(
        "SELECT * FROM curriculum_units WHERE offering_id=? AND approval_status='已确认' "
        "AND review_action<>'删除' ORDER BY seq", (offering_id,),
    )
    scenarios = []
    for unit in units:
        subs = _subscenarios(unit)
        scenarios.append({
            "seq": unit["seq"],
            "number": CHINESE_NUMBERS[unit["seq"]] if unit["seq"] < len(CHINESE_NUMBERS) else str(unit["seq"]),
            "title": unit["project_title"],
            "hours": unit["suggested_hours"],
            "theory_hours": int(unit["suggested_hours"]) // 2,
            "practice_hours": int(unit["suggested_hours"]) - int(unit["suggested_hours"]) // 2,
            "subscenarios": subs,
            "standards": unit["new_standards"],
            "technology": unit["new_technology"],
            "process": unit["new_process"],
            "methods": unit["new_methods"],
        })
    course_nature = _course_nature(offering, units)
    course_goals = _course_goals(offering, units)
    course_design = _course_design(offering, units)
    if not course_nature or not all(course_goals.values()) or not course_design:
        raise ValueError(f"《{offering['course_name']}》尚未建立课程内容模型，不能套用其他课程正文。")
    scheme = get_scheme(offering_id)
    model = {
        "offering": offering,
        "course_nature": course_nature,
        "course_goals": course_goals,
        "course_design": course_design,
        "job_direction": f"面向{offering['major']}专业相关的技术应用、业务实施、项目协作与成果交付岗位。",
        "teacher_requirements": _teacher_requirements(offering, units),
        "scenarios": scenarios,
        "assessment": {
            "process_percent": scheme["process_total"],
            "final_percent": scheme["final_total"],
            "components": scheme["components"],
            "projects": [
                {
                    "title": item["title"], "project_weight": 25,
                    "items": [
                        {"task": "知识技能", "attendance": 2, "performance": 3, "assignment": 6},
                        {"task": "任务实施", "attendance": 2, "performance": 3, "assignment": 9},
                    ],
                }
                for item in scenarios[:4]
            ],
        },
    }
    signature = _signature(offering, units)
    store.execute(
        "INSERT INTO course_content_models (offering_id,model_json,source_signature,generation_status,review_status,generated_at) "
        "VALUES (?,?,?,'已生成','待检查',CURRENT_TIMESTAMP) "
        "ON CONFLICT(offering_id) DO UPDATE SET model_json=excluded.model_json,source_signature=excluded.source_signature,"
        "generation_status='已生成',review_status='待检查',generated_at=CURRENT_TIMESTAMP",
        (offering_id, json.dumps(model, ensure_ascii=False), signature),
    )
    return model


def write_course_standard_model(offering_id, output_path):
    model = build_course_standard_model(offering_id)
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(model, ensure_ascii=False, indent=2), encoding="utf-8")
    return path
