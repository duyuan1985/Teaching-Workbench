"""
模板分析器单元测试

测试 template_analyzer.py 中的核心功能：
- 字段匹配（_match_field）
- 标签文本判断（_is_label_text）
- 表格角色识别（_table_role）
- 定位器解析
"""

import sys
import os
import unittest
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from template_analyzer import (
    _match_field,
    _is_label_text,
    _table_role,
    FIELD_SPECS,
    _identify_label_value_pairs,
    _cell_text,
)
from template_filler import _parse_locator, _get_slot_value


class TestLabelTextDetection(unittest.TestCase):
    """测试标签文本判断功能"""

    def test_short_label_with_colon(self):
        """带冒号的短标签应被识别"""
        self.assertTrue(_is_label_text("课程名称："))
        self.assertTrue(_is_label_text("课程类型:"))
        self.assertTrue(_is_label_text("学时学分："))

    def test_short_label_without_colon(self):
        """不带冒号的短文本也应被识别（表格左列）"""
        self.assertTrue(_is_label_text("课程名称"))
        self.assertTrue(_is_label_text("课程类型"))
        self.assertTrue(_is_label_text("周次"))

    def test_long_text_not_label(self):
        """长文本不应被识别为标签"""
        self.assertFalse(_is_label_text("课程名称及课程编号是课程的基本信息"))
        self.assertFalse(_is_label_text("本课程的教学目标是培养学生的实践能力"))

    def test_punctuation_not_label(self):
        """包含标点的较长文本不应被识别为标签"""
        self.assertFalse(_is_label_text("课程名称、课程类型和学时学分"))
        self.assertFalse(_is_label_text("注：请填写课程名称"))

    def test_label_with_spaces(self):
        """带空白的标签应被识别"""
        self.assertTrue(_is_label_text(" 课程名称 ： "))
        self.assertTrue(_is_label_text("授课教师  "))


class TestFieldMatching(unittest.TestCase):
    """测试字段匹配功能"""

    def test_label_mode_matching(self):
        """标签模式下应匹配短标签文本"""
        result = _match_field("课程名称：", mode="label")
        self.assertIsNotNone(result)
        field_name, kind, sources = result
        self.assertEqual(field_name, "课程基本信息")

        result = _match_field("课程类型：", mode="label")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "课程类型")

    def test_label_mode_rejects_long_text(self):
        """标签模式下不应匹配长文本"""
        result = _match_field("课程名称是课程的重要属性", mode="label")
        self.assertIsNone(result)

    def test_label_mode_rejects_notes(self):
        """标签模式下不应匹配注释说明"""
        result = _match_field("注：请填写课程名称", mode="label")
        self.assertIsNone(result)

        result = _match_field("（注：课程名称需准确）", mode="label")
        self.assertIsNone(result)

    def test_label_mode_rejects_placeholder(self):
        """标签模式下不应匹配占位词"""
        result = _match_field("课程名称填写处", mode="label")
        self.assertIsNone(result)

    def test_heading_mode_matching(self):
        """标题模式下应匹配短标题"""
        result = _match_field("一、课程性质", mode="heading")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "课程性质")

    def test_heading_mode_rejects_long(self):
        """标题模式下不应匹配过长文本"""
        result = _match_field("课程性质是本课程区别于其他课程的根本属性", mode="heading")
        self.assertIsNone(result)

    def test_course_name_and_code(self):
        """课程名称及课程编号应正确识别"""
        result = _match_field("课程名称及课程编号：", mode="label")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "课程基本信息")

    def test_class_field(self):
        """授课班级字段应正确识别"""
        result = _match_field("授课班级：", mode="label")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "授课班级")

    def test_teacher_field(self):
        """授课教师字段应正确识别"""
        result = _match_field("授课教师：", mode="label")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "授课教师")

        result = _match_field("主讲教师：", mode="label")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "授课教师")

    def test_week_field(self):
        """周次字段应正确识别"""
        result = _match_field("周次", mode="label")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "周次")

    def test_major_field(self):
        """适用专业字段应正确识别"""
        result = _match_field("适用专业：", mode="label")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "适用专业")

    def test_department_field(self):
        """所属系部字段应正确识别"""
        result = _match_field("所属系部：", mode="label")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "所属系部")


class TestTableRoleDetection(unittest.TestCase):
    """测试表格角色识别功能"""

    def test_basic_info_table(self):
        """课程基本信息表应被正确识别"""
        header = "课程名称 | 课程类型 | 学时学分"
        body = "网页设计 | 必修课 | 64学时"
        role = _table_role(header, body)
        self.assertEqual(role, "课程基本信息表")

    def test_course_goals_table(self):
        """课程目标表应被正确识别"""
        header = "知识目标 | 能力目标 | 思政目标 | 素质目标"
        body = "理解... | 掌握... | 培养... | 形成..."
        role = _table_role(header, body)
        self.assertEqual(role, "课程目标表")

    def test_schedule_detail_table(self):
        """授课计划明细表应被正确识别"""
        header = "周次 | 日期 | 课堂教学 | 课时"
        body = "1 | 2024-09-01 | 课程介绍 | 4"
        role = _table_role(header, body)
        self.assertEqual(role, "授课计划明细表")

    def test_content_structure_table(self):
        """课程内容结构表应被正确识别"""
        header = "序号 | 模块 | 学时"
        body = "1 | HTML基础 | 8"
        role = _table_role(header, body)
        self.assertEqual(role, "课程内容结构表")

    def test_assessment_table(self):
        """考核评价表应被正确识别"""
        header = "评价类型 | 评价内容 | 评价方式"
        body = "过程性考核 | 项目作品 | 教师评价"
        role = _table_role(header, body)
        self.assertEqual(role, "考核评价表")

    def test_design_basic_info_table(self):
        """教学设计基本信息表应被正确识别"""
        header = "周次 | 课时 | 授课班级"
        body = "1 | 4 | 营销2458"
        role = _table_role(header, body)
        self.assertEqual(role, "教学设计基本信息表")

    def test_unit_design_table(self):
        """单元教学设计表应被正确识别"""
        header = "教学步骤与内容 | 达成目标 | 教学方法"
        body = "导入... | 理解... | 讲授..."
        role = _table_role(header, body)
        self.assertEqual(role, "单元教学设计表")

    def test_generic_table(self):
        """普通表格应被识别为普通表格"""
        header = "列1 | 列2 | 列3"
        body = "a | b | c"
        role = _table_role(header, body)
        self.assertEqual(role, "普通表格")


class TestLocatorParsing(unittest.TestCase):
    """测试定位器解析功能"""

    def test_paragraph_locator(self):
        """段落定位器应正确解析"""
        loc_type, table_idx, row_idx, col_idx, direction = _parse_locator("paragraph:123")
        self.assertEqual(loc_type, "paragraph")
        self.assertEqual(table_idx, 123)  # paragraph 模式下 table_idx 存的是段落索引
        self.assertIsNone(row_idx)
        self.assertIsNone(col_idx)
        self.assertIsNone(direction)

    def test_table_horizontal_locator(self):
        """表格水平标签定位器应正确解析"""
        loc_type, table_idx, row_idx, col_idx, direction = _parse_locator(
            "table:0/row:2/col:3/horizontal"
        )
        self.assertEqual(loc_type, "table")
        self.assertEqual(table_idx, 0)
        self.assertEqual(row_idx, 2)
        self.assertEqual(col_idx, 3)
        self.assertEqual(direction, "horizontal")

    def test_table_vertical_locator(self):
        """表格垂直表头定位器应正确解析"""
        loc_type, table_idx, row_idx, col_idx, direction = _parse_locator(
            "table:1/col:2/vertical"
        )
        self.assertEqual(loc_type, "table")
        self.assertEqual(table_idx, 1)
        self.assertIsNone(row_idx)
        self.assertEqual(col_idx, 2)
        self.assertEqual(direction, "vertical")

    def test_table_header_locator(self):
        """表头定位器应正确解析"""
        loc_type, table_idx, row_idx, col_idx, direction = _parse_locator(
            "table:0/col:1/header"
        )
        self.assertEqual(loc_type, "table")
        self.assertEqual(table_idx, 0)
        self.assertEqual(col_idx, 1)
        self.assertEqual(direction, "header")

    def test_unknown_locator(self):
        """未知格式定位器应返回 unknown"""
        loc_type, *_ = _parse_locator("invalid:format")
        self.assertEqual(loc_type, "unknown")


class TestSlotValueMapping(unittest.TestCase):
    """测试槽位值映射功能"""

    def setUp(self):
        self.offering = {
            "course_name": "网页设计与制作",
            "course_type": "专业核心课",
            "course_nature": "必修课",
            "course_code": "KC001",
            "major": "市场营销",
            "term": "2024-2025-1",
            "teaching_class": "营销245801",
            "total_hours": 64,
            "credits": 4,
        }

    def test_basic_fields(self):
        """基本信息字段应正确映射"""
        self.assertEqual(
            _get_slot_value("课程名称", self.offering),
            "网页设计与制作"
        )
        self.assertEqual(
            _get_slot_value("课程类型", self.offering),
            "专业核心课"
        )
        self.assertEqual(
            _get_slot_value("课程代码", self.offering),
            "KC001"
        )
        self.assertEqual(
            _get_slot_value("适用专业", self.offering),
            "市场营销"
        )

    def test_credits_hours(self):
        """学时学分应正确格式化"""
        value = _get_slot_value("学时学分", self.offering)
        self.assertEqual(value, "64学时（4学分）")

    def test_floating_credits(self):
        """小数化学分应正确保留"""
        offering = dict(self.offering, credits=3.5)
        value = _get_slot_value("学时学分", offering)
        self.assertEqual(value, "64学时（3.5学分）")

    def test_teacher(self):
        """授课教师应正确返回"""
        self.assertEqual(_get_slot_value("授课教师", self.offering), "杜媛")

    def test_department(self):
        """所属系部应正确返回"""
        self.assertEqual(_get_slot_value("所属系部", self.offering), "经济贸易系")

    def test_complex_fields_return_none(self):
        """复杂字段应返回 None（由专用逻辑处理）"""
        self.assertIsNone(_get_slot_value("教学任务", self.offering))
        self.assertIsNone(_get_slot_value("知识目标", self.offering))
        self.assertIsNone(_get_slot_value("能力目标", self.offering))
        self.assertIsNone(_get_slot_value("课程设计总体思路", self.offering))
        self.assertIsNone(_get_slot_value("考核评价", self.offering))


class TestFieldSpecCompleteness(unittest.TestCase):
    """测试字段定义完整性"""

    def test_no_duplicate_keyword_groups(self):
        """同一关键词不应出现在多个字段定义中（避免歧义）"""
        keyword_to_field = {}
        for spec in FIELD_SPECS:
            keywords, field_name, kind, sources, match_mode = spec
            for kw in keywords:
                kw_lower = kw.strip()
                if kw_lower in keyword_to_field:
                    self.fail(
                        f"关键词 '{kw_lower}' 同时属于字段 "
                        f"'{keyword_to_field[kw_lower]}' 和 '{field_name}'"
                    )
                keyword_to_field[kw_lower] = field_name

    def test_all_fields_have_keywords(self):
        """所有字段定义都应有至少一个关键词"""
        for spec in FIELD_SPECS:
            keywords, field_name, kind, sources, match_mode = spec
            self.assertGreater(
                len(keywords), 0,
                f"字段 '{field_name}' 没有定义关键词"
            )

    def test_match_mode_valid(self):
        """匹配模式应为有效值"""
        valid_modes = {"label", "heading", "any"}
        for spec in FIELD_SPECS:
            keywords, field_name, kind, sources, match_mode = spec
            self.assertIn(
                match_mode, valid_modes,
                f"字段 '{field_name}' 的匹配模式 '{match_mode}' 无效"
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
