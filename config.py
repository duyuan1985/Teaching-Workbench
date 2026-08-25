"""
集中配置管理

所有默认值、常量、配置项统一定义在此文件。
其他模块通过 store.get_setting(key, DEFAULTS[key]) 引用。
"""

from pathlib import Path

PROJECT_ROOT = Path(__file__).parent

DEFAULTS = {
    "teacher_name": "杜媛",
    "department": "经济贸易系",
    "partner_company": "天津滨海迅腾科技集团有限公司",
    "output_root": str(PROJECT_ROOT / "生成结果"),
    "teaching_arrangement_path": "",
    "enhanced_generation": "0",
    "ai_curriculum_review": "0",
    "ai_model_preference": "online",
    "ollama_url": "http://127.0.0.1:11434",
    "ollama_model": "qwen3:8b",
}

ALLOWED_UPLOAD_EXTENSIONS = [".xlsx", ".xls", ".pdf", ".pptx", ".docx", ".doc"]
MAX_UPLOAD_SIZE = 50 * 1024 * 1024

DB_PATH = Path(__file__).parent / "data" / "workbench.db"

HTTP_HOST = "0.0.0.0"
HTTP_PORT = 8080

CORS_ORIGINS = [
    "http://127.0.0.1:5173",
    "http://127.0.0.1:8080",
    "http://localhost:5173",
    "http://localhost:8080",
]

SOURCE_ROOT = PROJECT_ROOT / "原始资料"
TEMPLATE_DIR = SOURCE_ROOT / "模板"
TEXTBOOK_ROOT = SOURCE_ROOT / "教材"

# 教材包目录根名 → 课程名（原始资料按课程建一级目录）
TEXTBOOK_COURSE_DIRS = {
    "H5设计与制作": "H5设计与制作",
    "Python程序设计": "Python程序设计",
    "商务数据分析": "商务数据分析",
    "图形图像设计": "图形图像设计",
    "数据标注": "数据标注",
    "新媒体平台运营与推广": "新媒体平台运营与推广",
}

# 教材选用规则：按 (课程名, 条件) 选择教材包目录与教材名称。
# 条件 term 精确匹配学期；major 为专业名匹配；无条件时为默认教材。
TEXTBOOK_RULES = [
    {"course": "H5设计与制作", "term": "2023-2024-2",
     "pack": "HTML5与CSS3项目实战", "name": "HTML5与CSS3项目实战（南开大学出版社）"},
    {"course": "H5设计与制作",
     "pack": "网页设计与制作—HTML5+CSS3项目实战(资料包)", "name": "网页设计与制作-HTML5与CSS3项目实战（天津大学出版社）"},
    {"course": "商务数据分析", "major": "农村电子商务",
     "pack": "大数据分析方法项目实战", "name": "大数据分析方法项目实战（天津大学出版社）"},
    {"course": "商务数据分析", "major": "市场营销",
     "pack": "《商务数据分析与决策》教学资料", "name": "商务数据分析与决策（南开大学出版社）"},
    {"course": "Python程序设计",
     "pack": "Python程序设计", "name": "Python程序设计（南开大学出版社）"},
    {"course": "数据标注",
     "pack": "数据标注任务式教程", "name": "数据标注任务式教程（天津大学出版社）"},
    {"course": "图形图像设计",
     "pack": "Adobe Photoshop CC 2018 案例化教程", "name": "Adobe Photoshop CC 2018 案例化教程（天津大学出版社）"},
    {"course": "新媒体平台运营与推广",
     "pack": "《新媒体营销综合案例教程》资料包", "name": "新媒体营销综合案例教程（天津大学出版社）"},
]

# 教学安排表课程名别名 → 标准课程名
COURSE_NAME_ALIASES = {"phthon程序设计": "Python程序设计"}

# 开课专业 → 可参照的人才培养方案专业（市场营销无单独方案，参照网络营销与直播电商）
MAJOR_PLAN_ALIASES = {"市场营销": "网络营销与直播电商"}

# 上课班级名称前缀 → 专业（用于推断开课专业，决定教材等）
MAJOR_BY_CLASS_PREFIX = [
    ("农商", "农村电子商务"),
    ("电商", "农村电子商务"),
    ("营销", "市场营销"),
    ("全媒体", "全媒体广告策划与营销"),
]


def resolve_textbook(course_name, term, major):
    """按课程/学期/专业匹配教材包目录与教材名称；返回 (pack_dir_path, textbook_name)。"""
    for rule in TEXTBOOK_RULES:
        if rule["course"] != course_name:
            continue
        if rule.get("term") and rule["term"] != term:
            continue
        if rule.get("major") and rule["major"] != major:
            continue
        course_dir = TEXTBOOK_COURSE_DIRS[course_name]
        return TEXTBOOK_ROOT / course_dir / rule["pack"], rule["name"]
    return "", ""


def resolve_major(class_name):
    for prefix, major in MAJOR_BY_CLASS_PREFIX:
        if prefix in str(class_name or ""):
            return major
    return ""
