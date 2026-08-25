import store


DEFAULT_COMPONENTS = (("出勤", 10, "考勤"), ("课堂表现", 10, "课堂"), ("平时作业", 20, "作业"))


def ensure_scheme(offering_id):
    if not store.rows("SELECT id FROM grade_components WHERE offering_id=? LIMIT 1", (offering_id,)):
        for order, (name, score, source) in enumerate(DEFAULT_COMPONENTS, 1):
            store.execute(
                "INSERT INTO grade_components(offering_id,component_name,weight,source_type,sort_order) VALUES (?,?,?,?,?)",
                (offering_id, name, score, source, order),
            )
    store.execute("INSERT OR IGNORE INTO grade_scheme_meta(offering_id) VALUES (?)", (offering_id,))


def get_scheme(offering_id):
    ensure_scheme(offering_id)
    components = store.rows(
        "SELECT id,component_name,weight,source_type,sort_order FROM grade_components WHERE offering_id=? ORDER BY sort_order,id",
        (offering_id,),
    )
    process_total = sum(float(item["weight"]) for item in components)
    meta = store.rows("SELECT * FROM grade_scheme_meta WHERE offering_id=?", (offering_id,))[0]
    return {
        "components": components,
        "process_total": process_total,
        "final_total": max(0, 100 - process_total),
        "source_label": meta["source_label"],
        "review_status": meta["review_status"],
    }


def component_text(offering_id, separator="、"):
    scheme = get_scheme(offering_id)
    return separator.join(f"{item['component_name']}{float(item['weight']):g}分" for item in scheme["components"])
