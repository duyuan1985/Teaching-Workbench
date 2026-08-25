import store


PHRASES = ("企业级卓越人才培养（信息类专业集群）", "企业级卓越人才培养")
NOISE = ("天津滨海迅腾科技集团", "迅腾科技集团", "http://", "https://", "网页设计与制作—HTML5+CSS3项目实战")


def clean(value):
    parts = [part.strip() for part in str(value or "").replace(";", "；").split("；") if part.strip()]
    return "；".join(
        part for part in parts
        if part != "任务技能"
        and not any(phrase in part for phrase in PHRASES)
        and not any(noise in part for noise in NOISE)
    )


def main():
    curriculum_fields = ("source_objectives", "source_skills", "revised_focus")
    task_fields = ("title", "knowledge_goal", "ability_goal", "ideological_goal", "quality_goal")
    with store.connect() as db:
        for row in db.execute("SELECT * FROM curriculum_units").fetchall():
            values = [clean(row[field]) for field in curriculum_fields]
            db.execute(
                f"UPDATE curriculum_units SET {','.join(field + '=?' for field in curriculum_fields)} WHERE id=?",
                (*values, row["id"]),
            )
        for row in db.execute("SELECT * FROM tasks").fetchall():
            values = [clean(row[field]) for field in task_fields]
            db.execute(
                f"UPDATE tasks SET {','.join(field + '=?' for field in task_fields)} WHERE id=?",
                (*values, row["id"]),
            )
        db.commit()


if __name__ == "__main__":
    main()
