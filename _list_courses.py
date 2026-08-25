import store

# 查询所有课程
courses = store.rows("""
    SELECT o.id as offering_id, o.course_id, c.name as course_name, 
           c.code as course_code, o.semester, o.major, o.class_name,
           o.teacher
    FROM offerings o
    JOIN courses c ON o.course_id = c.id
    ORDER BY o.semester, c.name
""")

print(f'共 {len(courses)} 个课程\n')
for c in courses:
    print(f"offering_id={c['offering_id']}: {c['course_name']} ({c['course_code']}) | 学期={c['semester']} | 专业={c['major']} | 班级={c['class_name']} | 教师={c['teacher']}")

# 检查已有生成结果
import os
result_dir = r'生成结果\精修版'
if os.path.exists(result_dir):
    files = [f for f in os.listdir(result_dir) if f.endswith('.docx')]
    print(f'\n已生成文档({len(files)}个):')
    for f in sorted(files):
        print(f'  {f}')
else:
    print('\n生成结果目录不存在')

# 检查原始资料中的课程信息
src_dir = r'原始资料'
if os.path.exists(src_dir):
    subdirs = [d for d in os.listdir(src_dir) if os.path.isdir(os.path.join(src_dir, d))]
    print(f'\n原始资料目录({len(subdirs)}个):')
    for d in sorted(subdirs):
        print(f'  {d}')
