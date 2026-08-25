import store
from task_builder import _ideological_goal, _quality_goal


def main():
    offerings = store.rows("SELECT * FROM offerings")
    for offering in offerings:
        units = store.rows(
            "SELECT * FROM curriculum_units WHERE offering_id=?",
            (offering["id"],),
        )
        by_title = {unit["project_title"]: unit for unit in units}
        for task in store.rows("SELECT * FROM tasks WHERE offering_id=?", (offering["id"],)):
            unit = by_title.get(task["chapter"])
            if unit:
                store.execute(
                    "UPDATE tasks SET quality_goal=?,ideological_goal=? WHERE id=?",
                    (_quality_goal(unit), _ideological_goal(unit, offering), task["id"]),
                )


if __name__ == "__main__":
    main()
