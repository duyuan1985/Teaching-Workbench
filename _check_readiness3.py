import store

offerings = store.rows("""
    SELECT o.id, o.course_name, o.term, o.major, o.teaching_class, o.offering_kind
    FROM offerings o 
    WHERE o.offering_kind != '实训课程'
    ORDER BY o.id
""")

print(f'共 {len(offerings)} 个普通课程\n')

for o in offerings:
    oid = o['id']
    
    tasks = store.rows("SELECT COUNT(*) as cnt FROM tasks WHERE offering_id=?", [oid])
    task_count = tasks[0]['cnt'] if tasks else 0
    
    sessions = store.rows("SELECT COUNT(*) as cnt FROM sessions WHERE offering_id=?", [oid])
    session_count = sessions[0]['cnt'] if sessions else 0
    
    models = store.rows("SELECT review_status FROM course_content_models WHERE offering_id=?", [oid])
    model_status = models[0]['review_status'] if models else '(无)'
    
    issues = store.rows("SELECT COUNT(*) as cnt, SUM(CASE WHEN severity='错误' THEN 1 ELSE 0 END) as errors FROM quality_issues WHERE offering_id=?", [oid])
    issue_count = issues[0]['cnt'] if issues else 0
    error_count = issues[0]['errors'] if issues else 0
    
    docs = store.rows("SELECT COUNT(*) as cnt, GROUP_CONCAT(document_type) as types FROM generated_documents WHERE offering_id=?", [oid])
    doc_count = docs[0]['cnt'] if docs else 0
    doc_types = docs[0]['types'] if docs else ''
    
    ready = 'YES' if (task_count > 0 and session_count > 0 and model_status == '已确认' and error_count == 0) else 'NO'
    
    print(f'[{ready}] ID={oid}: {o["course_name"]} ({o["term"]}) | 任务={task_count} 课程={session_count} 模型={model_status} 质量错误={error_count} 已生成={doc_count}({doc_types})')
