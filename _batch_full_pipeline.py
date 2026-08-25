"""
批量执行 步骤4(语义模型) + 步骤5(AI内容) + 步骤6(生成文档)
对已有模型/AI内容的课程跳过，直接生成文档。
"""
import store
import semantic_model
import content_author
import generate
import time
import sys

# 获取所有非实训课程
offerings = store.rows(
    'SELECT * FROM offerings WHERE offering_kind != "实训课程" ORDER BY id'
)
total = len(offerings)
print(f"共 {total} 门课程需要处理")
print("=" * 70)
sys.stdout.flush()

results_log = []
overall_start = time.time()

for i, o in enumerate(offerings, 1):
    oid = o['id']
    name = o['course_name']
    term = o['term']

    print(f"\n[{i}/{total}] ID={oid} {name} ({term})")
    print("-" * 70)
    sys.stdout.flush()

    course_log = {'id': oid, 'name': name, 'term': term}

    # --- 步骤4: 构建语义模型 ---
    existing_model = store.rows(
        'SELECT COUNT(*) as cnt FROM course_content_models WHERE offering_id=?',
        [oid]
    )
    if existing_model[0]['cnt'] > 0:
        print(f"  [步骤4] 语义模型: 已存在，跳过")
        course_log['model'] = 'skip'
    else:
        try:
            model = semantic_model.build_semantic_model(oid)
            n_proj = len(model.get("projects", []))
            n_know = len(model.get("knowledge_system", []))
            print(f"  [步骤4] 语义模型: {n_proj}项目, {n_know}知识点 OK")
            course_log['model'] = f'ok({n_proj}p,{n_know}k)'
        except Exception as e:
            print(f"  [步骤4] 语义模型: 失败 - {e}")
            course_log['model'] = f'FAIL: {e}'
            results_log.append(course_log)
            continue
    sys.stdout.flush()

    # --- 步骤5: AI内容生成 ---
    existing_sections = store.rows(
        'SELECT COUNT(*) as cnt FROM authored_sections WHERE offering_id=?',
        [oid]
    )
    if existing_sections[0]['cnt'] > 0:
        print(f"  [步骤5] AI内容: 已有{existing_sections[0]['cnt']}个section，跳过")
        course_log['ai'] = 'skip'
    else:
        t0 = time.time()
        try:
            count = content_author.author_course_content(oid)
            elapsed = time.time() - t0
            print(f"  [步骤5] AI内容: {count}个section ({elapsed:.0f}秒) OK")
            course_log['ai'] = f'ok({count}s,{elapsed:.0f}s)'
        except Exception as e:
            print(f"  [步骤5] AI内容: 失败 - {e}")
            course_log['ai'] = f'FAIL: {e}'
            results_log.append(course_log)
            continue
    sys.stdout.flush()

    # --- 步骤6: 生成文档 ---
    t0 = time.time()
    try:
        results = generate.generate_all(oid)
        elapsed = time.time() - t0
        for dt, r in results.items():
            if 'path' in r:
                print(f"  [步骤6] {dt}: {r['path']} OK")
                course_log[f'doc_{dt}'] = r['path']
            elif 'error' in r:
                print(f"  [步骤6] {dt}: 错误 - {r['error']}")
                course_log[f'doc_{dt}'] = f"ERR: {r['error']}"
        course_log['gen_time'] = f'{elapsed:.0f}s'
    except Exception as e:
        print(f"  [步骤6] 生成文档: 失败 - {e}")
        course_log['docs'] = f'FAIL: {e}'

    results_log.append(course_log)
    sys.stdout.flush()

# --- 汇总 ---
total_time = time.time() - overall_start
print("\n" + "=" * 70)
print(f"全部完成! 总耗时: {total_time:.0f}秒 ({total_time/60:.1f}分钟)")
print("=" * 70)

print("\n=== 汇总 ===")
for r in results_log:
    parts = [f"模型:{r.get('model','?')}", f"AI:{r.get('ai','?')}"]
    for k in r:
        if k.startswith('doc_'):
            parts.append(f"{k[4:]}:{r[k][:50]}")
    if 'gen_time' in r:
        parts.append(f"耗时:{r['gen_time']}")
    print(f"ID={r['id']:2d} {r['name'][:12]:12s} | {' | '.join(parts)}")

print("\n=== 失败列表 ===")
failures = [r for r in results_log
            if 'FAIL' in str(r.get('model','')) or 'FAIL' in str(r.get('ai','')) or 'FAIL' in str(r.get('docs',''))]
if failures:
    for r in failures:
        print(f"  ID={r['id']} {r['name']}: {r}")
else:
    print("  无失败")
