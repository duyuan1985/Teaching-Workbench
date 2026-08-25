import json

import store


def _without_ppt_footer(text):
    parts = []
    for raw in str(text or "").replace(";", "；").split("；"):
        part = raw.strip()
        if not part or "企业级卓越人才培养" in part or part.replace(" ", "") == "任务技能":
            continue
        if any(noise in part for noise in ("建设制造强国", "判断标准如表", "___PPT", "口令：", "滨海迅腾科技集团")):
            continue
        if part not in parts:
            parts.append(part)
    return "；".join(parts)


def _quality_goal(unit):
    text = f"{unit['project_title']} {unit['source_skills']}"
    goals = ["形成规范操作、自主学习和持续改进的职业素养"]
    if any(word in text for word in ("表单", "注册", "数据")):
        goals.append("增强个人信息保护与数据安全意识")
    if any(word in text for word in ("图像", "音频", "视频", "素材")):
        goals.append("养成尊重知识产权和规范使用数字素材的习惯")
    if any(word in text for word in ("CSS", "布局", "样式", "动画")):
        goals.append("培养审美意识、细节意识和精益求精的职业态度")
    if any(word in text for word in ("项目", "网页", "页面")):
        goals.append("提升团队协作、沟通表达和成果交付意识")
    return "；".join(goals) + "。"


def _ideological_goal(unit, offering):
    text = f"{unit['project_title']} {unit['source_skills']}"
    goals = ["树立诚实守信、遵纪守法和认真负责的职业价值观"]
    if any(word in text for word in ("表单", "注册", "数据", "网络")):
        goals.append("强化个人信息保护、网络安全和依法用网意识")
    if any(word in text for word in ("图像", "音频", "视频", "素材")):
        goals.append("增强知识产权保护意识，自觉抵制侵权和不规范使用数字资源的行为")
    if any(word in text for word in ("CSS", "布局", "样式", "动画", "页面")):
        goals.append("在页面设计中践行健康审美、文化自信和精益求精的工匠精神")
    if any(word in text for word in ("项目", "网站", "网页")):
        goals.append("理解数字技术服务社会与行业发展的责任，形成用户意识和质量意识")
    if "农村电子商务" in offering.get("major", ""):
        goals.append("关注数字技术赋能乡村振兴和农村电子商务发展的实际价值")
    return "；".join(goals) + "。"


def build_tasks(offering, replace=False):
    existing = store.rows("SELECT id FROM tasks WHERE offering_id=?", (offering["id"],))
    if existing and not replace:
        raise ValueError("课程主表中已有任务，请先核对现有内容，系统不会自动覆盖。")
    units = store.rows(
        "SELECT * FROM curriculum_units WHERE offering_id=? AND approval_status='已确认' AND review_action<>'删除' ORDER BY seq",
        (offering["id"],),
    )
    if not units:
        raise ValueError("课程蓝本尚未确认，不能生成课程任务。")
    hours = sum(unit["suggested_hours"] for unit in units)
    if hours != offering["total_hours"]:
        raise ValueError(f"已确认项目合计{hours}学时，与课程总学时{offering['total_hours']}不一致。")
    sessions = store.rows(
        "SELECT * FROM sessions WHERE offering_id=? AND status='已确认' AND session_type<>'停课' ORDER BY COALESCE(lesson_date,''),COALESCE(week_no,999),id",
        (offering["id"],),
    )
    class_names = [item.strip() for item in str(offering.get("teaching_class", "")).replace("；", ";").split(";") if item.strip()]
    if class_names:
        primary = class_names[0]
        sessions = [item for item in sessions if not item.get("class_name") or item.get("class_name") == primary]
    session_index = 0
    session_remaining = sessions[0]["hours"] if sessions else 4
    seq = 1
    with store.connect() as db:
        if replace:
            db.execute("DELETE FROM tasks WHERE offering_id=?", (offering["id"],))
        for unit in units:
            skills = [part.strip() for part in _without_ppt_footer(unit["source_skills"]).replace("；", ";").split(";") if part.strip()]
            unit_remaining = unit["suggested_hours"]
            chunks = []
            while unit_remaining > 0:
                session = sessions[session_index] if session_index < len(sessions) else {}
                chunk = min(unit_remaining, session_remaining)
                chunks.append((chunk, session))
                unit_remaining -= chunk
                session_remaining -= chunk
                if session_remaining == 0:
                    session_index += 1
                    session_remaining = sessions[session_index]["hours"] if session_index < len(sessions) else 4
            content_count = len(chunks) - 1 if len(chunks) > 1 else 1
            groups = [[] for _ in range(content_count)]
            for skill_index, skill in enumerate(skills):
                groups[skill_index % content_count].append(skill)
            for part_index, (chunk, session) in enumerate(chunks):
                implementation = len(chunks) > 1 and part_index == len(chunks) - 1
                selected = groups[part_index] if part_index < content_count else []
                if implementation:
                    detail = f"任务实施：{unit['project_title']}成果制作、检查与优化"
                else:
                    detail = "、".join(selected[:4]) or unit["revised_focus"] or unit["project_title"]
                title = f"{unit['project_title']}：{detail}"
                theory = chunk // 2
                practice = chunk - theory
                refs = [unit["source_file"]] if unit["source_file"] else []
                knowledge = _without_ppt_footer(unit["revised_focus"] or unit["source_objectives"])
                ability = f"能够完成{unit['project_title']}相关任务并按规范调试、验收和提交成果。"
                ideological = _ideological_goal(unit, offering)
                quality = _quality_goal(unit)
                if unit["project_title"] == "综合评价与课程总结":
                    refs = ["课程项目成果", "作品提交要求", "课程评价量表", "成果汇报与答辩记录"]
                    knowledge = "掌握课程项目成果整理、提交、汇报和评价的基本要求。"
                    ability = "能够完成作品提交、成果展示、问题说明和自我总结，并依据评价量表进行修改完善。"
                    ideological = "树立诚实守信、尊重规则、认真负责的职业价值观，增强成果质量意识和社会责任意识。"
                    quality = "形成自主总结、沟通表达、团队协作、接受评价和持续改进的职业素养。"
                db.execute(
                    """INSERT INTO tasks
                    (offering_id,seq,chapter,title,hours,theory_hours,practice_hours,week_no,lesson_date,
                     resource_refs,knowledge_goal,ability_goal,ideological_goal,quality_goal)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        offering["id"], seq, unit["project_title"], title, chunk, theory, practice,
                        session.get("week_no"), session.get("lesson_date", ""), json.dumps(refs, ensure_ascii=False),
                        knowledge, ability, ideological, quality,
                    ),
                )
                seq += 1
        db.commit()
    return seq - 1
