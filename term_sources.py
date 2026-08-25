"""按学期自动匹配原始资料文件路径。"""

from config import SOURCE_ROOT


def term_source_paths(term):
    """根据学期匹配学期进程表和学校校历文件路径。"""
    result = {"progress_path": "", "calendar_path": ""}

    progress_dir = SOURCE_ROOT / "学期进程表"
    if progress_dir.exists():
        matches = sorted(
            path for path in progress_dir.iterdir()
            if path.is_file() and term in path.name and path.suffix.lower() in (".xls", ".xlsx")
        )
        if len(matches) == 1:
            result["progress_path"] = str(matches[0])

    parts = term.split("-")
    if len(parts) == 3:
        calendar_year = parts[0] if parts[2] == "1" else parts[1]
        calendar_dir = SOURCE_ROOT / "学校校历"
        if calendar_dir.exists():
            matches = sorted(
                path for path in calendar_dir.iterdir()
                if path.is_file() and calendar_year in path.name and path.suffix.lower() in (".xls", ".xlsx")
            )
            if len(matches) == 1:
                result["calendar_path"] = str(matches[0])
    return result
