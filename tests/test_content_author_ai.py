import json
import sqlite3
import unittest
from unittest.mock import patch

import content_author


class ContentAuthorAITest(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(":memory:")
        self.db.row_factory = sqlite3.Row
        self.db.execute(
            """
            CREATE TABLE authored_sections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                offering_id INTEGER NOT NULL,
                document_type TEXT NOT NULL,
                section_key TEXT NOT NULL,
                repeat_key TEXT NOT NULL DEFAULT '',
                title TEXT NOT NULL,
                content_json TEXT NOT NULL,
                evidence_json TEXT NOT NULL,
                authoring_status TEXT NOT NULL,
                review_status TEXT NOT NULL,
                generated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(offering_id, document_type, section_key, repeat_key)
            )
            """
        )
        self.model = {
            "identity": {
                "course_name": "Python程序设计",
                "major": "电子商务",
                "course_type": "专业核心课",
                "course_nature": "必修课",
                "term": "2026-2027-1",
                "textbook_version": "Python程序设计项目化教程",
            },
            "projects": [
                {
                    "title": "Python基础",
                    "hours": 8,
                    "knowledge_skills": ["变量", "分支结构", "循环结构"],
                    "expected_outcome": "可运行的控制台程序",
                }
            ],
            "knowledge_system": ["变量", "数据类型", "分支结构", "循环结构"],
            "course_links": {"prerequisite": "计算机基础", "followup": "数据分析"},
            "standards": ["Python编码规范"],
            "tools_technology": ["Python 3", "IDLE"],
            "work_process": ["需求分析", "编码", "测试", "提交"],
            "teaching_methods": ["任务驱动", "演示练习"],
            "assessment_scheme": {
                "process_total": 40,
                "final_total": 60,
                "components": [{"component_name": "课堂任务", "weight": 40}],
            },
        }
        self.tasks = [self._task(1, "变量与输入输出"), self._task(2, "分支结构")]
        self.facts = [
            {
                "project_hint": "Python基础",
                "fact_type": "ppt_slide",
                "fact_value": "使用input读取数据，转换类型后输出计算结果。",
                "locator": "第12页",
            }
        ]

    def tearDown(self):
        self.db.close()

    @staticmethod
    def _task(seq, title):
        return {
            "id": seq,
            "seq": seq,
            "chapter": "Python基础",
            "title": title,
            "hours": 4,
            "theory_hours": 2,
            "practice_hours": 2,
            "week_no": seq,
            "lesson_date": f"2026-09-{seq:02d}",
            "knowledge_goal": "掌握变量、数据类型和程序结构",
            "ability_goal": "能够编写并调试Python程序",
            "ideological_goal": "形成诚信编程意识",
            "quality_goal": "形成规范编码习惯",
            "resource_refs": json.dumps(["教材/Python基础.pptx"], ensure_ascii=False),
        }

    def _rows(self, sql, params=()):
        if "course_content_models" in sql:
            return [{"model_json": json.dumps(self.model, ensure_ascii=False)}]
        if "FROM tasks" in sql:
            return self.tasks
        if "resource_facts" in sql:
            return self.facts
        raise AssertionError(f"未处理的查询：{sql}")

    @staticmethod
    def _overview():
        return {
            "course_nature": [
                "本课程面向电子商务专业的数据处理与程序应用能力培养。",
                "课程以Python基础项目为载体组织知识学习和技能训练。",
                "课程衔接计算机基础，并为后续数据分析课程提供支撑。",
            ],
            "course_goals": {
                "知识目标": ["掌握变量", "掌握数据类型", "掌握程序结构"],
                "能力目标": ["能够分析问题", "能够编写程序", "能够测试程序"],
                "思政目标": ["坚持诚信编程", "遵守数据规范", "尊重知识产权"],
                "素质目标": ["形成规范习惯", "形成协作意识", "形成质量意识"],
            },
            "course_design": [
                "按照需求分析、编码和测试的工作过程组织教学。",
                "使用教材项目和课件证据设计递进任务。",
                "通过过程检查和成果评价检验课程目标。",
            ],
        }

    @staticmethod
    def _activity():
        return {
            key: f"{key}：结合Python基础项目开展提问、教师演示、学生练习、结果检查和总结评价。"
            for key in content_author.ACTIVITY_KEYS
        }

    def _successful_ai(self, prompt, **kwargs):
        is_overview = "课程标准中的课程性质" in prompt
        payload = self._overview() if is_overview else self._activity()
        content = json.dumps(payload, ensure_ascii=False)
        if is_overview:
            content = f"```json\n{content}\n```"
        return {
            "success": True,
            "source": "MockAI",
            "model": "mock-model",
            "content": content,
        }

    def test_mock_ai_generates_valid_sections(self):
        with (
            patch.object(content_author.store, "rows", side_effect=self._rows),
            patch.object(content_author.store, "connect", return_value=self.db),
            patch.object(content_author, "ask_result", side_effect=self._successful_ai) as mocked_ai,
        ):
            count = content_author.author_course_content(1)

        self.assertEqual(count, 14)
        self.assertEqual(mocked_ai.call_count, 3)
        self.assertEqual(
            self.db.execute("SELECT COUNT(*) FROM authored_sections").fetchone()[0],
            14,
        )
        ai_rows = self.db.execute(
            "SELECT COUNT(*) FROM authored_sections WHERE authoring_status=?",
            ("AI草稿/MockAI/mock-model",),
        ).fetchone()[0]
        self.assertEqual(ai_rows, 7)
        goals = json.loads(self.db.execute(
            "SELECT content_json FROM authored_sections "
            "WHERE document_type='课程标准' AND section_key='course_goals'"
        ).fetchone()[0])
        self.assertIn("知识目标", goals)

    def test_invalid_ai_does_not_delete_existing_draft(self):
        self.db.execute(
            "INSERT INTO authored_sections "
            "(offering_id,document_type,section_key,repeat_key,title,content_json,"
            "evidence_json,authoring_status,review_status) VALUES (1,'课程标准',"
            "'course_nature','','旧稿','[\"旧内容\"]','[]','旧稿','已确认')"
        )
        self.db.commit()
        failed = {
            "success": True,
            "source": "MockAI",
            "model": "mock-model",
            "content": '{"course_nature":[]}',
        }

        with (
            patch.object(content_author.store, "rows", side_effect=self._rows),
            patch.object(content_author.store, "connect", return_value=self.db),
            patch.object(content_author, "ask_result", return_value=failed),
        ):
            with self.assertRaisesRegex(RuntimeError, "AI内容生成失败"):
                content_author.author_course_content(1)

        old = self.db.execute(
            "SELECT title,content_json FROM authored_sections"
        ).fetchone()
        self.assertEqual((old["title"], old["content_json"]), ("旧稿", '["旧内容"]'))


if __name__ == "__main__":
    unittest.main()
