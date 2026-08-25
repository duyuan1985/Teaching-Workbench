import store
import os

# 查询所有课程
offerings = store.rows("SELECT id, course_name, term, major, course_code, teaching_class, teacher_name, offering_kind FROM offerings ORDER BY id")
print(f'共 {len(offerings)} 个课程\n')
print(f'{"ID":>3} | {"课程名":<20} | {"学期":<12} | {"专业":<20} | {"班级":<20} | {"教师":<6} | {"类型":<6}')
print('-' * 120)
for o in offerings:
    print(f"{o['id']:>3} | {o['course_name']:<20} | {o['term']:<12} | {o['major']:<20} | {str(o['teaching_class'])[:20]:<20} | {str(o['teacher_name']):<6} | {str(o['offering_kind']):<6}")

# 检查已生成文档
print('\n=== 已生成文档 ===')
for subdir in ['生成结果/精修版', '生成结果']:
    if os.path.exists(subdir):
        files = [f for f in os.listdir(subdir) if f.endswith('.docx')]
        print(f'{subdir}: {len(files)}个文件')
        for f in sorted(files)[:10]:
            print(f'  {f}')
        if len(files) > 10:
            print(f'  ... 还有{len(files)-10}个')

# 检查哪些是培训课程（需要跳过）
training = [o for o in offerings if o.get('offering_kind') == '培训']
regular = [o for o in offerings if o.get('offering_kind') != '培训']
print(f'\n普通课程: {len(regular)}个')
print(f'培训课程: {len(training)}个 (跳过)')
