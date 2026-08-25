import store
from catalog import catalog_from_arrangement
from importers import rebuild_schedule
from term_sources import term_source_paths


def main():
    store.initialize()
    arrangement_path = store.get_setting("teaching_arrangement_path")
    for offering in store.rows("SELECT * FROM offerings"):
        offering_id = offering["id"]
        if not offering.get("teaching_class"):
            candidates = [
                item for item in catalog_from_arrangement(arrangement_path, offering["term"])
                if item["course_code"] == offering["course_code"] and item["major"] == offering["major"]
            ]
            if len(candidates) == 1:
                store.execute(
                    "UPDATE offerings SET teaching_class=? WHERE id=?",
                    (candidates[0]["class_name"], offering_id),
                )
        arrangement_sources = store.rows(
            "SELECT id FROM source_files WHERE offering_id=? AND source_type='教学安排表'",
            (offering_id,),
        )
        if arrangement_sources:
            store.execute(
                "UPDATE source_files SET source_path=? WHERE id=?",
                (arrangement_path, arrangement_sources[0]["id"]),
            )
        else:
            store.create_source_file(
                offering_id,
                {"source_type": "教学安排表", "source_path": arrangement_path, "required": "1", "notes": "教学安排总表"},
            )
        sources = term_source_paths(offering["term"])
        if sources["progress_path"]:
            progress_sources = store.rows(
                "SELECT id FROM source_files WHERE offering_id=? AND source_type='学期进程表'",
                (offering_id,),
            )
            if progress_sources:
                store.execute(
                    "UPDATE source_files SET source_path=? WHERE id=?",
                    (sources["progress_path"], progress_sources[0]["id"]),
                )
            else:
                store.create_source_file(
                    offering_id,
                    {"source_type": "学期进程表", "source_path": sources["progress_path"], "required": "1", "notes": "按课程实例学期自动匹配"},
                )
            store.execute(
                "UPDATE offerings SET schedule_path=? WHERE id=?",
                (sources["progress_path"], offering_id),
            )
        if sources["calendar_path"]:
            existing = store.rows(
                "SELECT id FROM source_files WHERE offering_id=? AND source_type='学校校历'",
                (offering_id,),
            )
            if existing:
                store.execute(
                    "UPDATE source_files SET source_path=? WHERE id=?",
                    (sources["calendar_path"], existing[0]["id"]),
                )
            else:
                store.create_source_file(
                    offering_id,
                    {
                        "source_type": "学校校历",
                        "source_path": sources["calendar_path"],
                        "required": "1",
                        "notes": "按课程实例学期自动匹配",
                    },
                )
        refreshed = store.rows("SELECT * FROM offerings WHERE id=?", (offering_id,))[0]
        try:
            rebuild_schedule(refreshed)
        except Exception as error:
            print(f"{refreshed['term']} {refreshed['course_name']}: {error}")


if __name__ == "__main__":
    main()
