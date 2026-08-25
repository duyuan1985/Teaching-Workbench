import store

# 开启AI增强生成
with store.connect() as db:
    db.execute("UPDATE settings SET setting_value='1' WHERE setting_key='enhanced_generation'")
    db.commit()
print("enhanced_generation 已设置为 1")

# 检查course_content_models表
models = store.rows("SELECT offering_id FROM course_content_models")
model_oids = [m['offering_id'] for m in models]
print(f"\ncourse_content_models: {len(models)}条")

# 检查哪些课程有/没有模型
offerings = store.rows("SELECT id, course_name, term FROM offerings WHERE offering_kind != '实训课程' ORDER BY id")
print(f"\n=== 模型数据检查 ===")
for o in offerings:
    oid = o['id']
    has_model = oid in model_oids
    tasks = store.rows("SELECT COUNT(*) as cnt FROM tasks WHERE offering_id=?", [oid])
    task_count = tasks[0]['cnt'] if tasks else 0
    status = '✅' if has_model else '❌'
    print(f"  {status} ID={oid} {o['course_name']} ({o['term']}): model={'有' if has_model else '无'}, tasks={task_count}")

# 检查content_author.py的AI增强逻辑
import inspect, content_author
# 找到build_content或类似的入口函数
funcs = [name for name, obj in inspect.getmembers(content_author, inspect.isfunction)]
print(f"\ncontent_author函数: {funcs}")

# 检查build_content函数
if 'build_content' in funcs:
    src = inspect.getsource(content_author.build_content)
    print("\n=== build_content 函数 ===")
    print(src[:3000])
