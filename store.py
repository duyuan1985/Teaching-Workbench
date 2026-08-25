import json
import os
import sqlite3
from contextlib import contextmanager
from pathlib import Path

from config import DEFAULTS, DB_PATH as _DEFAULT_DB_PATH


DB_PATH = Path(os.environ.get("WORKBENCH_DB_PATH", str(_DEFAULT_DB_PATH)))


def connect():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA journal_mode = WAL")
    db.execute("PRAGMA busy_timeout = 5000")
    db.execute("PRAGMA foreign_keys = ON")
    return db


@contextmanager
def db_transaction():
    """统一事务上下文管理器，确保批量操作的原子性。"""
    db = connect()
    try:
        yield db
        db.commit()
    except Exception:
        db.rollback()
        raise
    finally:
        db.close()


def initialize():
    with connect() as db:
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS talent_plans (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              major TEXT NOT NULL,
              cohort TEXT NOT NULL DEFAULT '',
              source_path TEXT NOT NULL,
              plan_json TEXT NOT NULL,
              created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              UNIQUE(major, cohort)
            );
            CREATE TABLE IF NOT EXISTS offerings (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              course_name TEXT NOT NULL,
              term TEXT NOT NULL,
              major TEXT NOT NULL DEFAULT '',
              course_code TEXT NOT NULL DEFAULT '',
              course_nature TEXT NOT NULL DEFAULT '',
              course_type TEXT NOT NULL DEFAULT '',
              assessment_type TEXT NOT NULL DEFAULT '期末考核',
              assessment_method TEXT NOT NULL DEFAULT '实操',
              credits REAL NOT NULL DEFAULT 0,
              total_hours INTEGER NOT NULL,
              weekly_hours INTEGER NOT NULL,
              template_version TEXT NOT NULL,
              textbook_version TEXT NOT NULL,
              textbook_path TEXT NOT NULL DEFAULT '',
              template_path TEXT NOT NULL DEFAULT '',
              schedule_path TEXT NOT NULL DEFAULT '',
              notes TEXT NOT NULL DEFAULT '',
              created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              UNIQUE(course_name, term, major, textbook_version, template_version)
            );
            CREATE TABLE IF NOT EXISTS tasks (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              seq INTEGER NOT NULL,
              chapter TEXT NOT NULL,
              title TEXT NOT NULL,
              hours INTEGER NOT NULL,
              theory_hours INTEGER NOT NULL DEFAULT 0,
              practice_hours INTEGER NOT NULL DEFAULT 0,
              week_no INTEGER,
              lesson_date TEXT NOT NULL DEFAULT '',
              resource_refs TEXT NOT NULL DEFAULT '[]',
              knowledge_goal TEXT NOT NULL DEFAULT '',
              ability_goal TEXT NOT NULL DEFAULT '',
              ideological_goal TEXT NOT NULL DEFAULT '',
              quality_goal TEXT NOT NULL DEFAULT '',
              UNIQUE(offering_id, seq)
            );
            CREATE TABLE IF NOT EXISTS sessions (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              session_type TEXT NOT NULL CHECK(session_type IN ('正常排课','补课','调课','停课')),
              week_no INTEGER,
              lesson_date TEXT NOT NULL DEFAULT '',
              weekday TEXT NOT NULL DEFAULT '',
              periods TEXT NOT NULL DEFAULT '',
              hours INTEGER NOT NULL,
              classroom TEXT NOT NULL DEFAULT '',
              status TEXT NOT NULL DEFAULT '已确认' CHECK(status IN ('待确认','已确认','已取消')),
              source_note TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS calendar_events (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              event_name TEXT NOT NULL,
              start_date TEXT NOT NULL,
              end_date TEXT NOT NULL,
              suspends_classes INTEGER NOT NULL DEFAULT 1,
              replacement_date TEXT NOT NULL DEFAULT '',
              target_week_no INTEGER,
              target_weekday TEXT NOT NULL DEFAULT '',
              source_note TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS template_files (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              document_type TEXT NOT NULL,
              template_name TEXT NOT NULL,
              template_path TEXT NOT NULL,
              required INTEGER NOT NULL DEFAULT 1,
              output_name_pattern TEXT NOT NULL DEFAULT '',
              notes TEXT NOT NULL DEFAULT '',
              UNIQUE(offering_id, document_type, template_path)
            );
            CREATE TABLE IF NOT EXISTS source_files (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              source_type TEXT NOT NULL,
              source_path TEXT NOT NULL,
              required INTEGER NOT NULL DEFAULT 1,
              notes TEXT NOT NULL DEFAULT '',
              UNIQUE(offering_id, source_type, source_path)
            );
            CREATE TABLE IF NOT EXISTS course_types (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              name TEXT NOT NULL UNIQUE,
              sort_order INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS settings (
              setting_key TEXT PRIMARY KEY,
              setting_value TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS template_rules (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              template_file_id INTEGER NOT NULL REFERENCES template_files(id) ON DELETE CASCADE,
              seq INTEGER NOT NULL,
              location_type TEXT NOT NULL,
              location_ref TEXT NOT NULL,
              section_title TEXT NOT NULL DEFAULT '',
              instruction_text TEXT NOT NULL DEFAULT '',
              content_requirements TEXT NOT NULL DEFAULT '',
              data_sources TEXT NOT NULL DEFAULT '',
              format_json TEXT NOT NULL DEFAULT '{}',
              approval_status TEXT NOT NULL DEFAULT '待确认' CHECK(approval_status IN ('待确认','已确认','需修改')),
              UNIQUE(template_file_id,seq)
            );
            CREATE TABLE IF NOT EXISTS template_analyses (
              template_file_id INTEGER PRIMARY KEY REFERENCES template_files(id) ON DELETE CASCADE,
              source_hash TEXT NOT NULL,
              rule_count INTEGER NOT NULL DEFAULT 0,
              contract_json TEXT NOT NULL DEFAULT '{}',
              analysis_status TEXT NOT NULL DEFAULT '待确认',
              analyzed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS template_slots (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              template_file_id INTEGER NOT NULL REFERENCES template_files(id) ON DELETE CASCADE,
              slot_key TEXT NOT NULL,
              locator TEXT NOT NULL,
              section_title TEXT NOT NULL DEFAULT '',
              field_name TEXT NOT NULL,
              content_kind TEXT NOT NULL DEFAULT '事实填写',
              repeat_scope TEXT NOT NULL DEFAULT '单次',
              source_priority TEXT NOT NULL DEFAULT '',
              instruction_text TEXT NOT NULL DEFAULT '',
              format_json TEXT NOT NULL DEFAULT '{}',
              required INTEGER NOT NULL DEFAULT 1,
              approval_status TEXT NOT NULL DEFAULT '待确认',
              UNIQUE(template_file_id,slot_key)
            );
            CREATE TABLE IF NOT EXISTS generated_documents (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              template_file_id INTEGER NOT NULL REFERENCES template_files(id) ON DELETE CASCADE,
              document_type TEXT NOT NULL,
              output_path TEXT NOT NULL,
              generation_status TEXT NOT NULL DEFAULT '草稿',
              structural_check TEXT NOT NULL DEFAULT '待检查',
              visual_check TEXT NOT NULL DEFAULT '待检查',
              notes TEXT NOT NULL DEFAULT '',
              generated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              UNIQUE(offering_id, document_type)
            );
            CREATE TABLE IF NOT EXISTS resource_items (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              resource_type TEXT NOT NULL,
              file_path TEXT NOT NULL,
              title TEXT NOT NULL DEFAULT '',
              content_excerpt TEXT NOT NULL DEFAULT '',
              source_hash TEXT NOT NULL DEFAULT '',
              UNIQUE(offering_id, file_path)
            );
            CREATE TABLE IF NOT EXISTS resource_facts (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              resource_item_id INTEGER NOT NULL REFERENCES resource_items(id) ON DELETE CASCADE,
              project_hint TEXT NOT NULL DEFAULT '',
              fact_type TEXT NOT NULL,
              fact_key TEXT NOT NULL DEFAULT '',
              fact_value TEXT NOT NULL,
              locator TEXT NOT NULL DEFAULT '',
              confidence REAL NOT NULL DEFAULT 1.0
            );
            CREATE TABLE IF NOT EXISTS curriculum_units (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              seq INTEGER NOT NULL,
              project_title TEXT NOT NULL,
              source_file TEXT NOT NULL,
              source_objectives TEXT NOT NULL DEFAULT '',
              source_skills TEXT NOT NULL DEFAULT '',
              review_action TEXT NOT NULL DEFAULT '更新' CHECK(review_action IN ('保留','更新','补充','删除')),
              revised_focus TEXT NOT NULL DEFAULT '',
              rationale TEXT NOT NULL DEFAULT '',
              new_standards TEXT NOT NULL DEFAULT '',
              new_technology TEXT NOT NULL DEFAULT '',
              new_process TEXT NOT NULL DEFAULT '',
              new_methods TEXT NOT NULL DEFAULT '',
              suggested_hours INTEGER NOT NULL DEFAULT 0,
              approval_status TEXT NOT NULL DEFAULT '待确认' CHECK(approval_status IN ('待确认','已确认','退回修改')),
              UNIQUE(offering_id, seq)
            );
            CREATE TABLE IF NOT EXISTS course_content_models (
              offering_id INTEGER PRIMARY KEY REFERENCES offerings(id) ON DELETE CASCADE,
              model_json TEXT NOT NULL DEFAULT '{}',
              source_signature TEXT NOT NULL DEFAULT '',
              generation_status TEXT NOT NULL DEFAULT '待生成',
              review_status TEXT NOT NULL DEFAULT '待检查',
              generated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS model_cache (
              cache_key TEXT PRIMARY KEY,
              provider TEXT NOT NULL DEFAULT 'local',
              model_name TEXT NOT NULL DEFAULT '',
              prompt_hash TEXT NOT NULL DEFAULT '',
              response_json TEXT NOT NULL DEFAULT '{}',
              created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS students (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              class_name TEXT NOT NULL,
              student_no TEXT NOT NULL,
              student_name TEXT NOT NULL,
              gender TEXT NOT NULL DEFAULT '',
              notes TEXT NOT NULL DEFAULT '',
              active INTEGER NOT NULL DEFAULT 1,
              UNIQUE(class_name,student_no)
            );
            CREATE TABLE IF NOT EXISTS attendance_records (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
              lesson_date TEXT NOT NULL,
              status TEXT NOT NULL DEFAULT '出勤' CHECK(status IN ('出勤','迟到','早退','请假','旷课')),
              score REAL NOT NULL DEFAULT 0,
              notes TEXT NOT NULL DEFAULT '',
              updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              UNIQUE(offering_id,student_id,lesson_date)
            );
            CREATE TABLE IF NOT EXISTS grade_components (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              component_name TEXT NOT NULL,
              weight REAL NOT NULL DEFAULT 0,
              source_type TEXT NOT NULL DEFAULT '手工',
              sort_order INTEGER NOT NULL DEFAULT 0,
              UNIQUE(offering_id,component_name)
            );
            CREATE TABLE IF NOT EXISTS attendance_rules (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              status TEXT NOT NULL,
              deduction REAL NOT NULL DEFAULT 0,
              sort_order INTEGER NOT NULL DEFAULT 0,
              UNIQUE(offering_id,status)
            );
            CREATE TABLE IF NOT EXISTS student_leave_periods (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
              start_date TEXT NOT NULL,
              end_date TEXT NOT NULL,
              reason TEXT NOT NULL,
              created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              UNIQUE(offering_id,student_id,start_date,end_date,reason)
            );
            CREATE TABLE IF NOT EXISTS assignments (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              assignment_name TEXT NOT NULL,
              max_score REAL NOT NULL DEFAULT 5,
              sort_order INTEGER NOT NULL DEFAULT 0,
              UNIQUE(offering_id,assignment_name)
            );
            CREATE TABLE IF NOT EXISTS assignment_scores (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              assignment_id INTEGER NOT NULL REFERENCES assignments(id) ON DELETE CASCADE,
              student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
              score REAL NOT NULL DEFAULT 0,
              source TEXT NOT NULL DEFAULT '手工录入',
              updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              UNIQUE(assignment_id,student_id)
            );
            CREATE TABLE IF NOT EXISTS assignment_import_issues (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              assignment_id INTEGER NOT NULL REFERENCES assignments(id) ON DELETE CASCADE,
              raw_student_no TEXT NOT NULL DEFAULT '',
              raw_student_name TEXT NOT NULL DEFAULT '',
              raw_score TEXT NOT NULL DEFAULT '',
              issue TEXT NOT NULL,
              created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS classroom_performance_records (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
              lesson_date TEXT NOT NULL,
              behavior TEXT NOT NULL,
              deduction REAL NOT NULL DEFAULT 0,
              notes TEXT NOT NULL DEFAULT '',
              created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS grade_scheme_meta (
              offering_id INTEGER PRIMARY KEY REFERENCES offerings(id) ON DELETE CASCADE,
              source_label TEXT NOT NULL DEFAULT '系统默认（课程标准尚未提供结构化分项分值）',
              review_status TEXT NOT NULL DEFAULT '待确认',
              updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS grade_analysis_documents (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              class_name TEXT NOT NULL,
              source_filename TEXT NOT NULL,
              output_path TEXT NOT NULL,
              generation_status TEXT NOT NULL DEFAULT '已生成',
              generated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS training_documents (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              class_name TEXT NOT NULL DEFAULT '',
              source_dir TEXT NOT NULL,
              output_path TEXT NOT NULL,
              generation_status TEXT NOT NULL DEFAULT '已生成',
              generated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS offering_classes (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              class_name TEXT NOT NULL,
              enrollment_count INTEGER NOT NULL DEFAULT 0,
              source_note TEXT NOT NULL DEFAULT '',
              updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              UNIQUE(offering_id,class_name)
            );
            CREATE TABLE IF NOT EXISTS listening_records (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              session_id INTEGER NOT NULL REFERENCES sessions(id) ON DELETE CASCADE,
              template_path TEXT NOT NULL,
              output_path TEXT NOT NULL,
              generated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS student_scores (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              student_id INTEGER NOT NULL REFERENCES students(id) ON DELETE CASCADE,
              component_id INTEGER NOT NULL REFERENCES grade_components(id) ON DELETE CASCADE,
              score REAL NOT NULL DEFAULT 0,
              notes TEXT NOT NULL DEFAULT '',
              UNIQUE(offering_id,student_id,component_id)
            );
            CREATE TABLE IF NOT EXISTS course_evidence (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              evidence_key TEXT NOT NULL,
              evidence_type TEXT NOT NULL,
              value_json TEXT NOT NULL,
              source_type TEXT NOT NULL,
              source_locator TEXT NOT NULL DEFAULT '',
              confidence REAL NOT NULL DEFAULT 1.0
            );
            CREATE TABLE IF NOT EXISTS quality_issues (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              stage TEXT NOT NULL,
              severity TEXT NOT NULL,
              issue_code TEXT NOT NULL,
              location TEXT NOT NULL DEFAULT '',
              message TEXT NOT NULL,
              status TEXT NOT NULL DEFAULT '待处理'
            );
            CREATE TABLE IF NOT EXISTS authored_sections (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              document_type TEXT NOT NULL,
              section_key TEXT NOT NULL,
              repeat_key TEXT NOT NULL DEFAULT '',
              title TEXT NOT NULL DEFAULT '',
              content_json TEXT NOT NULL DEFAULT '{}',
              evidence_json TEXT NOT NULL DEFAULT '[]',
              authoring_status TEXT NOT NULL DEFAULT '草稿',
              review_status TEXT NOT NULL DEFAULT '待检查',
              generated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              UNIQUE(offering_id,document_type,section_key,repeat_key)
            );
            CREATE TABLE IF NOT EXISTS curriculum_review_rules (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              course_name TEXT NOT NULL,
              rule_name TEXT NOT NULL DEFAULT '',
              ppt_group_mode TEXT NOT NULL DEFAULT 'each',
              project_pattern TEXT NOT NULL DEFAULT '',
              outline_keyword TEXT NOT NULL DEFAULT '教学大纲',
              title_extraction TEXT NOT NULL DEFAULT 'first_slide',
              objective_keyword TEXT NOT NULL DEFAULT '学习目标',
              skill_keywords TEXT NOT NULL DEFAULT '',
              modernization_tags TEXT NOT NULL DEFAULT '',
              is_active INTEGER NOT NULL DEFAULT 1,
              created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              UNIQUE(course_name)
            );
            CREATE TABLE IF NOT EXISTS dirty_flags (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              flag TEXT NOT NULL,
              reason TEXT NOT NULL DEFAULT '',
              created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              UNIQUE(offering_id, flag)
            );
            CREATE TABLE IF NOT EXISTS content_updates (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              update_type TEXT NOT NULL DEFAULT '内容更新',
              topic TEXT NOT NULL DEFAULT '',
              original_summary TEXT NOT NULL DEFAULT '',
              suggested_content TEXT NOT NULL DEFAULT '',
              reason TEXT NOT NULL DEFAULT '',
              source_urls TEXT NOT NULL DEFAULT '[]',
              confidence REAL NOT NULL DEFAULT 0.5,
              status TEXT NOT NULL DEFAULT '待审核',
              related_chapters TEXT NOT NULL DEFAULT '[]',
              reviewed_by TEXT NOT NULL DEFAULT '',
              created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              reviewed_at TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_content_updates_offering ON content_updates(offering_id);
            CREATE TABLE IF NOT EXISTS pipeline_status (
              offering_id INTEGER NOT NULL REFERENCES offerings(id) ON DELETE CASCADE,
              stage TEXT NOT NULL,
              status TEXT NOT NULL DEFAULT 'pending',
              error_message TEXT NOT NULL DEFAULT '',
              started_at TEXT NOT NULL DEFAULT '',
              completed_at TEXT NOT NULL DEFAULT '',
              UNIQUE(offering_id, stage)
            );
            CREATE TABLE IF NOT EXISTS template_library (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              doc_type TEXT NOT NULL,
              name TEXT NOT NULL,
              version_label TEXT NOT NULL DEFAULT '',
              file_path TEXT NOT NULL UNIQUE,
              file_hash TEXT NOT NULL DEFAULT '',
              file_format TEXT NOT NULL DEFAULT '',
              status TEXT NOT NULL DEFAULT '待解析' CHECK(status IN ('待解析','解析中','已解析','解析失败','暂不支持')),
              notes TEXT NOT NULL DEFAULT '',
              created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
              updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS template_contracts (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              library_id INTEGER NOT NULL REFERENCES template_library(id) ON DELETE CASCADE,
              version INTEGER NOT NULL DEFAULT 1,
              status TEXT NOT NULL DEFAULT '草稿' CHECK(status IN ('草稿','已确认')),
              structural_json TEXT NOT NULL DEFAULT '{}',
              content_json TEXT NOT NULL DEFAULT '{}',
              slot_count INTEGER NOT NULL DEFAULT 0,
              parse_message TEXT NOT NULL DEFAULT '',
              parsed_at TEXT NOT NULL DEFAULT '',
              confirmed_at TEXT NOT NULL DEFAULT '',
              UNIQUE(library_id, version)
            );
            CREATE TABLE IF NOT EXISTS contract_slots (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              contract_id INTEGER NOT NULL REFERENCES template_contracts(id) ON DELETE CASCADE,
              slot_key TEXT NOT NULL,
              locator TEXT NOT NULL,
              section_title TEXT NOT NULL DEFAULT '',
              field_name TEXT NOT NULL,
              classification TEXT NOT NULL DEFAULT 'A' CHECK(classification IN ('A','B','C','人工')),
              structure_json TEXT NOT NULL DEFAULT '{}',
              format_json TEXT NOT NULL DEFAULT '{}',
              content_req TEXT NOT NULL DEFAULT '',
              confidence TEXT NOT NULL DEFAULT '中' CHECK(confidence IN ('高','中','低')),
              required INTEGER NOT NULL DEFAULT 1,
              manual_override INTEGER NOT NULL DEFAULT 0,
              approval_status TEXT NOT NULL DEFAULT '待确认',
              UNIQUE(contract_id, slot_key)
            );
            """
        )
        for index, name in enumerate(("专业核心课", "专业基础课", "专业拓展课"), 1):
            db.execute(
                "INSERT OR IGNORE INTO course_types (name,sort_order) VALUES (?,?)",
                (name, index),
            )
        offering_columns = {row[1] for row in db.execute("PRAGMA table_info(offerings)")}
        template_file_columns = {row[1] for row in db.execute("PRAGMA table_info(template_files)")}
        if "library_id" not in template_file_columns:
            db.execute("ALTER TABLE template_files ADD COLUMN library_id INTEGER REFERENCES template_library(id)")
        contract_slot_columns = {row[1] for row in db.execute("PRAGMA table_info(contract_slots)")}
        if "sort_order" not in contract_slot_columns:
            db.execute("ALTER TABLE contract_slots ADD COLUMN sort_order INTEGER NOT NULL DEFAULT 0")
        if "teaching_class" not in offering_columns:
            db.execute("ALTER TABLE offerings ADD COLUMN teaching_class TEXT NOT NULL DEFAULT ''")
        if "course_nature" not in offering_columns:
            db.execute("ALTER TABLE offerings ADD COLUMN course_nature TEXT NOT NULL DEFAULT ''")
        for column, default in (("assessment_type", "期末考核"), ("assessment_method", "实操")):
            if column not in offering_columns:
                db.execute(f"ALTER TABLE offerings ADD COLUMN {column} TEXT NOT NULL DEFAULT '{default}'")
        for column in ("prerequisite_courses", "followup_courses"):
            if column not in offering_columns:
                db.execute(f"ALTER TABLE offerings ADD COLUMN {column} TEXT NOT NULL DEFAULT ''")
        if "offering_kind" not in offering_columns:
            db.execute("ALTER TABLE offerings ADD COLUMN offering_kind TEXT NOT NULL DEFAULT '普通课程'")
        if "base_offering_id" not in offering_columns:
            db.execute("ALTER TABLE offerings ADD COLUMN base_offering_id INTEGER")
        if "teacher_name" not in offering_columns:
            db.execute("ALTER TABLE offerings ADD COLUMN teacher_name TEXT NOT NULL DEFAULT ''")
        if "department" not in offering_columns:
            db.execute("ALTER TABLE offerings ADD COLUMN department TEXT NOT NULL DEFAULT ''")
        if "teaching_mode" not in offering_columns:
            db.execute("ALTER TABLE offerings ADD COLUMN teaching_mode TEXT NOT NULL DEFAULT ''")
        if "lecture_hours" not in offering_columns:
            db.execute("ALTER TABLE offerings ADD COLUMN lecture_hours INTEGER NOT NULL DEFAULT 0")
        if "experiment_hours" not in offering_columns:
            db.execute("ALTER TABLE offerings ADD COLUMN experiment_hours INTEGER NOT NULL DEFAULT 0")
        if "practice_hours" not in offering_columns:
            db.execute("ALTER TABLE offerings ADD COLUMN practice_hours INTEGER NOT NULL DEFAULT 0")
        session_columns = {row[1] for row in db.execute("PRAGMA table_info(sessions)")}
        if "class_name" not in session_columns:
            db.execute("ALTER TABLE sessions ADD COLUMN class_name TEXT NOT NULL DEFAULT ''")
            db.execute(
                "UPDATE sessions SET class_name=TRIM(SUBSTR(source_note,INSTR(source_note,'班级:')+3,"
                "CASE WHEN INSTR(SUBSTR(source_note,INSTR(source_note,'班级:')+3),'；')>0 "
                "THEN INSTR(SUBSTR(source_note,INSTR(source_note,'班级:')+3),'；')-1 ELSE LENGTH(source_note) END)) "
                "WHERE source_note LIKE '%班级:%'"
            )
        db.execute(
            "UPDATE sessions SET class_name=(SELECT teaching_class FROM offerings WHERE offerings.id=sessions.offering_id) "
            "WHERE class_name='' AND (SELECT teaching_class FROM offerings WHERE offerings.id=sessions.offering_id)<>'' "
            "AND (SELECT teaching_class FROM offerings WHERE offerings.id=sessions.offering_id) NOT LIKE '%；%' "
            "AND (SELECT teaching_class FROM offerings WHERE offerings.id=sessions.offering_id) NOT LIKE '%;%'"
        )
        db.execute("UPDATE grade_components SET source_type='作业' WHERE component_name='平时作业' AND source_type='手工'")
        db.execute("UPDATE grade_components SET source_type='课堂' WHERE component_name='课堂表现' AND source_type='手工'")
        analysis_columns = {row[1] for row in db.execute("PRAGMA table_info(template_analyses)")}
        if "contract_json" not in analysis_columns:
            db.execute("ALTER TABLE template_analyses ADD COLUMN contract_json TEXT NOT NULL DEFAULT '{}'")
        # AI增强分析相关字段
        for column, definition in (
            ("ai_analysis_json", "TEXT NOT NULL DEFAULT '{}'"),
            ("ai_analysis_status", "TEXT NOT NULL DEFAULT '未分析'"),
        ):
            if column not in analysis_columns:
                db.execute(f"ALTER TABLE template_analyses ADD COLUMN {column} {definition}")
        resource_columns = {row[1] for row in db.execute("PRAGMA table_info(resource_items)")}
        for column, definition in (
            ("structured_json", "TEXT NOT NULL DEFAULT '{}'"),
            ("extraction_status", "TEXT NOT NULL DEFAULT '待解析'"),
            ("project_hint", "TEXT NOT NULL DEFAULT ''"),
        ):
            if column not in resource_columns:
                db.execute(f"ALTER TABLE resource_items ADD COLUMN {column} {definition}")
        for key, value in DEFAULTS.items():
            db.execute(
                "INSERT OR IGNORE INTO settings (setting_key,setting_value) VALUES (?,?)",
                (key, str(value)),
            )
        base = db.execute(
            "SELECT * FROM offerings WHERE course_name='商务数据分析' AND term='2025-2026-2' AND major='农村电子商务' ORDER BY id LIMIT 1"
        ).fetchone()
        if base and not db.execute(
            "SELECT id FROM offerings WHERE course_name='商务数据分析实训' AND term='2025-2026-2' AND major='农村电子商务'"
        ).fetchone():
            cursor = db.execute(
                """INSERT INTO offerings
                (course_name,term,major,course_code,teaching_class,course_nature,course_type,assessment_type,assessment_method,
                 prerequisite_courses,followup_courses,credits,total_hours,weekly_hours,template_version,textbook_version,
                 textbook_path,template_path,schedule_path,notes,offering_kind,base_offering_id)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    "商务数据分析实训", "2025-2026-2", "农村电子商务", "", "农商245702", "必修课",
                    base["course_type"], "实训考核", "成果评价", "商务数据分析", "", 0, 24, 24,
                    "由模板套件自动识别", base["textbook_version"], base["textbook_path"], "", "",
                    "集中实训课程；依据现有商务数据分析实训资料建立。", "实训课程", base["id"],
                ),
            )
            training_id = cursor.lastrowid
            db.execute(
                "INSERT OR IGNORE INTO offering_classes(offering_id,class_name,enrollment_count,source_note) VALUES (?,?,?,?)",
                (training_id, "农商245702", 36, "实训资料识别：商务数据分析实训总结"),
            )


def rows(sql, params=()):
    with connect() as db:
        return [dict(row) for row in db.execute(sql, params)]


def execute(sql, params=()):
    with connect() as db:
        cur = db.execute(sql, params)
        db.commit()
        return cur.lastrowid


def get_setting(key, default=""):
    items = rows("SELECT setting_value FROM settings WHERE setting_key=?", (key,))
    return items[0]["setting_value"] if items else default


def set_setting(key, value):
    execute(
        "INSERT INTO settings (setting_key,setting_value) VALUES (?,?) "
        "ON CONFLICT(setting_key) DO UPDATE SET setting_value=excluded.setting_value",
        (key, value),
    )


def create_offering(data):
    fields = [
        "course_name", "term", "major", "course_code", "teaching_class", "course_nature", "course_type", "assessment_type", "assessment_method", "prerequisite_courses", "followup_courses",
        "credits", "total_hours", "weekly_hours", "template_version",
        "textbook_version", "textbook_path", "template_path", "schedule_path", "notes"
    ]
    values = [data.get(field, "") for field in fields]
    values[fields.index("template_version")] = "由模板套件自动识别"
    return execute(
        f"INSERT INTO offerings ({','.join(fields)}) VALUES ({','.join('?' for _ in fields)})",
        values,
    )


def update_offering(offering_id, data):
    fields = [
        "course_name", "term", "major", "course_code", "teaching_class", "course_nature", "course_type", "assessment_type", "assessment_method", "prerequisite_courses", "followup_courses",
        "credits", "total_hours", "weekly_hours",
        "textbook_version", "notes"
    ]
    values = [data.get(field, "") for field in fields]
    values.append(offering_id)
    execute(
        f"UPDATE offerings SET {','.join(field + '=?' for field in fields)} WHERE id=?",
        values,
    )


FLAG_IMPACTS = {
    "basic_info": {"affects": ["documents"], "label": "课程基本信息修改", "hint": "需要重新生成文档", "action": "generate"},
    "schedule": {"affects": ["documents"], "label": "排课数据变更", "hint": "需要重新生成文档（排课影响授课计划/教学设计）", "action": "generate"},
    "resources": {"affects": ["review", "foundation", "documents"], "label": "教材资源变更", "hint": "需要重建蓝本→重建基础→重新生成文档", "action": "rebuild_review"},
    "review": {"affects": ["tasks", "foundation", "documents"], "label": "蓝本审查修改", "hint": "需要重新构建任务→重建基础→重新生成文档", "action": "build_tasks"},
    "tasks": {"affects": ["foundation", "documents"], "label": "教学任务变更", "hint": "需要重建基础→重新生成文档", "action": "rebuild_foundation"},
    "foundation": {"affects": ["documents"], "label": "内容基础变更", "hint": "需要重新生成文档", "action": "generate"},
    "templates": {"affects": ["foundation", "documents"], "label": "模板文件变更", "hint": "需要重建基础→重新生成文档", "action": "rebuild_foundation"},
    "teacher_name": {"affects": ["documents"], "label": "教师姓名修改", "hint": "只需重新生成文档", "action": "generate"},
}


def mark_dirty(offering_id, flag, reason=""):
    """标记某项数据变更，影响下游产物"""
    with connect() as db:
        db.execute(
            "INSERT INTO dirty_flags (offering_id, flag, reason) VALUES (?, ?, ?) "
            "ON CONFLICT(offering_id, flag) DO UPDATE SET reason=excluded.reason, created_at=CURRENT_TIMESTAMP",
            (offering_id, flag, reason),
        )
        db.commit()


def clear_dirty(offering_id, flag):
    """清除某项变更标记（对应操作完成后调用）"""
    with connect() as db:
        db.execute("DELETE FROM dirty_flags WHERE offering_id=? AND flag=?", (offering_id, flag))
        db.commit()


def clear_dirty_by_action(offering_id, action):
    """根据操作类型清除对应的变更标记"""
    flags_to_clear = [f for f, info in FLAG_IMPACTS.items() if info["action"] == action]
    if not flags_to_clear:
        return
    with connect() as db:
        placeholders = ",".join("?" for _ in flags_to_clear)
        db.execute(
            f"DELETE FROM dirty_flags WHERE offering_id=? AND flag IN ({placeholders})",
            [offering_id] + flags_to_clear,
        )
        db.commit()


# ==================== 管道状态 ====================

PIPELINE_STAGES = ("resources", "templates", "tasks", "model", "content", "quality")


def set_pipeline_stage(offering_id, stage, status, error=""):
    """更新管道阶段状态: pending / running / done / failed"""
    with connect() as db:
        if status == "running":
            db.execute(
                "INSERT INTO pipeline_status (offering_id, stage, status, started_at, error_message) "
                "VALUES (?,?,?,CURRENT_TIMESTAMP,'') "
                "ON CONFLICT(offering_id, stage) DO UPDATE SET status=excluded.status, "
                "started_at=excluded.started_at, error_message='', completed_at=''",
                (offering_id, stage, status),
            )
        else:
            db.execute(
                "INSERT INTO pipeline_status (offering_id, stage, status, completed_at, error_message) "
                "VALUES (?,?,?,CURRENT_TIMESTAMP,?) "
                "ON CONFLICT(offering_id, stage) DO UPDATE SET status=excluded.status, "
                "completed_at=excluded.completed_at, error_message=excluded.error_message",
                (offering_id, stage, status, error),
            )
        db.commit()


def get_pipeline_status(offering_id):
    """获取管道各阶段状态"""
    return {r["stage"]: dict(r) for r in rows(
        "SELECT stage, status, error_message, started_at, completed_at "
        "FROM pipeline_status WHERE offering_id=?", (offering_id,)
    )}


def clear_pipeline_status(offering_id):
    """清除管道状态（重新开始时调用）"""
    with connect() as db:
        db.execute("DELETE FROM pipeline_status WHERE offering_id=?", (offering_id,))
        db.commit()


def get_dirty_flags(offering_id):
    """获取当前未处理的变更标记"""
    return rows("SELECT flag, reason, created_at FROM dirty_flags WHERE offering_id=? ORDER BY created_at", (offering_id,))


def get_dirty_flags_with_meta(offering_id):
    """获取变更标记及其影响元数据，按处理优先级排序"""
    active = {r["flag"]: dict(r) for r in get_dirty_flags(offering_id)}
    result = []
    for flag, info in FLAG_IMPACTS.items():
        is_active = flag in active
        result.append({
            "flag": flag,
            "label": info["label"],
            "hint": info["hint"],
            "action": info["action"],
            "affects": info["affects"],
            "active": is_active,
            "reason": active.get(flag, {}).get("reason", ""),
            "created_at": active.get(flag, {}).get("created_at", ""),
        })
    return result


# ============================================================
# 内容更新建议
# ============================================================

def list_content_updates(offering_id, status=None):
    """列出课程的内容更新建议"""
    if status:
        return rows(
            "SELECT * FROM content_updates WHERE offering_id=? AND status=? ORDER BY confidence DESC, id DESC",
            (offering_id, status),
        )
    return rows(
        "SELECT * FROM content_updates WHERE offering_id=? ORDER BY status='待审核' DESC, confidence DESC, id DESC",
        (offering_id,),
    )


def add_content_update(offering_id, update_type, topic, original_summary,
                       suggested_content, reason, source_urls=None,
                       confidence=0.5, related_chapters=None):
    """新增一条内容更新建议"""
    import json as _json
    return execute(
        """INSERT INTO content_updates
        (offering_id, update_type, topic, original_summary, suggested_content,
         reason, source_urls, confidence, related_chapters, status)
        VALUES (?,?,?,?,?,?,?,?,?,'待审核')""",
        (
            offering_id, update_type, topic, original_summary, suggested_content,
            reason, _json.dumps(source_urls or [], ensure_ascii=False),
            confidence, _json.dumps(related_chapters or [], ensure_ascii=False),
        ),
    )


def update_content_update_status(update_id, status, reviewed_by=""):
    """审核更新建议：状态=已采纳/已忽略/待审核"""
    from datetime import datetime
    execute(
        "UPDATE content_updates SET status=?, reviewed_by=?, reviewed_at=? WHERE id=?",
        (status, reviewed_by, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), update_id),
    )


def get_accepted_updates(offering_id):
    """获取已采纳的更新内容（用于文档生成时融合）"""
    return rows(
        "SELECT * FROM content_updates WHERE offering_id=? AND status='已采纳' ORDER BY id",
        (offering_id,),
    )


def delete_content_update(update_id):
    """删除一条更新建议"""
    execute("DELETE FROM content_updates WHERE id=?", (update_id,))


def create_task(offering_id, data):
    refs = data.get("resource_refs", [])
    if isinstance(refs, str):
        refs = [part.strip() for part in refs.split(";") if part.strip()]
    if int(data.get("theory_hours", 0)) + int(data.get("practice_hours", 0)) != int(data["hours"]):
        raise ValueError("理论学时与实践学时之和必须等于任务学时。")
    fields = [
        "offering_id", "seq", "chapter", "title", "hours", "theory_hours",
        "practice_hours", "week_no", "lesson_date", "resource_refs",
        "knowledge_goal", "ability_goal", "ideological_goal", "quality_goal"
    ]
    values = [
        offering_id, int(data["seq"]), data["chapter"], data["title"], int(data["hours"]),
        int(data.get("theory_hours", 0)), int(data.get("practice_hours", 0)),
        int(data["week_no"]) if str(data.get("week_no", "")).strip() else None,
        data.get("lesson_date", ""), json.dumps(refs, ensure_ascii=False),
        data.get("knowledge_goal", ""), data.get("ability_goal", ""),
        data.get("ideological_goal", ""), data.get("quality_goal", "")
    ]
    return execute(
        f"INSERT INTO tasks ({','.join(fields)}) VALUES ({','.join('?' for _ in fields)})",
        values,
    )


def create_session(offering_id, data):
    fields = [
        "offering_id", "class_name", "session_type", "week_no", "lesson_date", "weekday",
        "periods", "hours", "classroom", "status", "source_note"
    ]
    values = [
        offering_id, data.get("class_name", ""), data["session_type"],
        int(data["week_no"]) if str(data.get("week_no", "")).strip() else None,
        data.get("lesson_date", ""), data.get("weekday", ""), data.get("periods", ""),
        int(data["hours"]), data.get("classroom", ""), data.get("status", "待确认"),
        data.get("source_note", "")
    ]
    return execute(
        f"INSERT INTO sessions ({','.join(fields)}) VALUES ({','.join('?' for _ in fields)})",
        values,
    )


def create_calendar_event(offering_id, data):
    fields = [
        "offering_id", "event_name", "start_date", "end_date", "suspends_classes",
        "replacement_date", "target_week_no", "target_weekday", "source_note"
    ]
    values = [
        offering_id, data["event_name"], data["start_date"], data.get("end_date") or data["start_date"],
        1 if data.get("suspends_classes", "1") == "1" else 0,
        data.get("replacement_date", ""),
        int(data["target_week_no"]) if str(data.get("target_week_no", "")).strip() else None,
        data.get("target_weekday", ""), data.get("source_note", "")
    ]
    return execute(
        f"INSERT INTO calendar_events ({','.join(fields)}) VALUES ({','.join('?' for _ in fields)})",
        values,
    )


def create_template_file(offering_id, data):
    fields = [
        "offering_id", "document_type", "template_name", "template_path",
        "required", "output_name_pattern", "notes"
    ]
    values = [
        offering_id, data["document_type"], data["template_name"], data["template_path"],
        1 if data.get("required", "1") == "1" else 0,
        data.get("output_name_pattern", ""), data.get("notes", "")
    ]
    return execute(
        f"INSERT INTO template_files ({','.join(fields)}) VALUES ({','.join('?' for _ in fields)})",
        values,
    )


def create_source_file(offering_id, data):
    fields = ["offering_id", "source_type", "source_path", "required", "notes"]
    values = [
        offering_id, data["source_type"], data["source_path"],
        1 if data.get("required", "1") == "1" else 0, data.get("notes", "")
    ]
    return execute(
        f"INSERT INTO source_files ({','.join(fields)}) VALUES ({','.join('?' for _ in fields)})",
        values,
    )


def update_source_file(source_id, data):
    execute(
        "UPDATE source_files SET source_type=?,source_path=?,required=?,notes=? WHERE id=?",
        (
            data["source_type"], data["source_path"],
            1 if data.get("required", "1") == "1" else 0,
            data.get("notes", ""), source_id,
        ),
    )


def update_curriculum_unit(unit_id, data):
    fields = [
        "review_action", "revised_focus", "rationale", "new_standards",
        "new_technology", "new_process", "new_methods", "suggested_hours", "approval_status"
    ]
    values = [data.get(field, "") for field in fields]
    values[7] = int(values[7])
    values.append(unit_id)
    execute(
        f"UPDATE curriculum_units SET {','.join(field + '=?' for field in fields)} WHERE id=?",
        values,
    )


def create_course_type(name):
    name = name.strip()
    if not name:
        raise ValueError("课程类型名称不能为空。")
    order = rows("SELECT COALESCE(MAX(sort_order),0)+1 AS next_order FROM course_types")[0]["next_order"]
    return execute("INSERT INTO course_types (name,sort_order) VALUES (?,?)", (name, order))


def update_course_type(type_id, name):
    name = name.strip()
    if not name:
        raise ValueError("课程类型名称不能为空。")
    current = rows("SELECT name FROM course_types WHERE id=?", (type_id,))
    if not current:
        raise ValueError("课程类型不存在。")
    old_name = current[0]["name"]
    with connect() as db:
        db.execute("UPDATE course_types SET name=? WHERE id=?", (name, type_id))
        db.execute("UPDATE offerings SET course_type=? WHERE course_type=?", (name, old_name))
        db.commit()


def delete_course_type(type_id):
    current = rows("SELECT name FROM course_types WHERE id=?", (type_id,))
    if not current:
        raise ValueError("课程类型不存在。")
    name = current[0]["name"]
    count = rows("SELECT COUNT(*) AS count FROM offerings WHERE course_type=?", (name,))[0]["count"]
    if count:
        raise ValueError(f"课程类型“{name}”正在被 {count} 个课程实例使用，不能删除。")
    execute("DELETE FROM course_types WHERE id=?", (type_id,))
