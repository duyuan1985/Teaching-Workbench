import os

from generate import get_opening_semester, get_template_path, get_template_version


def test_opening_semester_from_long_class_name():
    assert get_opening_semester({"term": "2023-2024-2", "teaching_class": "2022电商教学班"}) == "第四学期"


def test_opening_semester_from_compact_class_code():
    assert get_opening_semester({"term": "2023-2024-2", "teaching_class": "235701班"}) == "第二学期"


def test_template_versions():
    assert get_template_version("2023-2024-2") == "2023-2024"
    assert get_template_version("2024-2025-1") == "2024-2025-1"
    assert get_template_version("2024-2025-2") == "2024-2025-2"
    assert get_template_version("2025-2026-2") == "2025-2026"


def test_all_declared_template_versions_exist():
    for term in ("2023-2024-1", "2023-2024-2", "2024-2025-1", "2024-2025-2", "2025-2026-1", "2025-2026-2"):
        offering = {"term": term}
        assert os.path.exists(get_template_path("standard", offering))
        assert os.path.exists(get_template_path("design", offering))
