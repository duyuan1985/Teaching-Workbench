import store

# 查询所有课程的状态
offerings = store.rows("""
    SELECT o.id, o.course_name, o.term, o.major, o.teaching_class, o.offering_kind,
           o.template_path, o.textbook_path
    FROM offerings o 
    WHERE o.offering_kind != '实训课程'
    ORDER BY o.id
""")

print(f'共 {len(offerings)} 个普通课程\n')

for o in offerings:
    oid = o['id']
    
    # 检查tasks
    tasks = store.rows("SELECT COUNT(*) as cnt FROM tasks WHERE offering_id=?", [oid])
    task_count = tasks[0]['cnt'] if tasks else 0
    
    # 检查sessions
    sessions = store.rows("SELECT COUNT(*) as cnt FROM sessions WHERE offering_id=?", [oid])
    session_count = sessions[0]['cnt'] if sessions else 0
    
    # 检查curriculum_units
    units = store.rows("SELECT COUNT(*) as cnt, SUM(CASE WHEN approval_status='已确认' THEN 1 ELSE 0 END) as confirmed FROM curriculum_units WHERE offering_id=?", [oid])
    unit_count = units[0]['cnt'] if units else 0
    unit_confirmed = units[0]['confirmed'] if units else 0
    
    # 检查content_model
    models = store.rows("SELECT review_status FROM course_content_models WHERE offering_id=?", [oid])
    model_status = models[0]['review_status'] if models else '(无)'
    
    # 检查quality_issues
    issues = store.rows("SELECT COUNT(*) as cnt, SUM(CASE WHEN severity='错误' THEN 1 ELSE 0 END) as errors FROM quality_issues WHERE offering_id=?", [oid])
    issue_count = issues[0]['cnt'] if issues else 0
    error_count = issues[0]['errors'] if issues else 0
    
    # 检查generated_documents
    docs = store.rows("SELECT COUNT(*) as cnt, GROUP_CONCAT(doc_type) as types FROM generated_documents WHERE offering_id=?", [oid])
    doc_count = docs[0]['cnt'] if docs else 0
    doc_types = docs[0]['types'] if docs else ''
    
    # 检查template_files
    tmpls = store.rows("SELECT COUNT(*) as cnt, SUM(CASE WHEN analysis_status='已确认' THEN 1 ELSE 0 END) as confirmed FROM template_files WHERE offering_id=?", [oid])
    tmpl_count = tmpls[0]['cnt'] if tmpls else 0
    tmpl_confirmed = tmpls[0]['confirmed'] if tmpls else 0
    
    # 判断就绪状态
    ready = '✅' if (task_count > 0 and session_count > 0 and model_status == '已确认' and error_count == 0) else '❌'
    
    print(f'{ready} ID={oid}: {o["course_name"]} ({o["term"]})')
    print(f'   任务={task_count} | 课程={session_count} | 单元={unit_count}(确认{unit_confirmed}) | 模型={model_status} | 质量={issue_count}(错误{error_count}) | 模板={tmpl_count}(确认{tmpl_confirmed}) | 已生成={doc_count}({doc_types})')
