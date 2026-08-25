"""
批量添加18周补课记录
- ID=18,19,22,32: 已有"待确认"补课，确认并补充
- ID=25,27: 无补课，新建18周补课
- ID=33,34: 全部sessions"待确认"，先确认全部，再补课
"""
import store

def get_confirmed_hours(oid):
    sessions = store.rows("SELECT * FROM sessions WHERE offering_id=?", [oid])
    return sum(s['hours'] for s in sessions if s['status'] == '已确认' and s.get('session_type') != '停课')

def confirm_existing_makeup(oid):
    """确认已有的待确认补课记录"""
    makeup = store.rows("SELECT * FROM sessions WHERE offering_id=? AND session_type='补课' AND status='待确认'", [oid])
    with store.connect() as db:
        for m in makeup:
            db.execute("UPDATE sessions SET status='已确认' WHERE id=?", [m['id']])
        db.commit()
    return len(makeup)

def add_makeup_sessions(oid, hours_needed, week=18):
    """在18周添加补课记录"""
    offering = store.rows("SELECT * FROM offerings WHERE id=?", [oid])[0]
    class_name = offering.get('teaching_class', '')
    # 取第一个班级
    if '；' in class_name:
        class_name = class_name.split('；')[0]
    elif ';' in class_name:
        class_name = class_name.split(';')[0]
    
    # 按每次4学时拆分，不足4的用2学时
    sessions_added = []
    remaining = hours_needed
    while remaining > 0:
        h = min(4, remaining)
        if remaining - h < 2 and h > 2:  # 避免剩余1学时
            h = remaining
        sessions_added.append(h)
        remaining -= h
    
    with store.connect() as db:
        for h in sessions_added:
            db.execute("""
                INSERT INTO sessions (offering_id, week_no, lesson_date, hours, status, session_type, class_name)
                VALUES (?, ?, '', ?, '已确认', '补课', ?)
            """, [oid, week, h, class_name])
        db.commit()
    return sessions_added

# ============================================================
# 处理8个课程
# ============================================================

courses = [
    # (oid, 描述)
    (18, 'H5设计与制作 2023-2024-2'),
    (19, 'Python程序设计 2023-2024-2'),
    (22, '图形图像设计 2024-2025-1'),
    (25, 'Python程序设计 2024-2025-2'),
    (27, 'H5设计与制作 2025-2026-1'),
    (32, '新媒体平台运营与推广 2025-2026-2'),
    (33, 'H5设计与制作 2026-2027-1'),
    (34, 'Python程序设计 2026-2027-1'),
]

for oid, desc in courses:
    offering = store.rows("SELECT * FROM offerings WHERE id=?", [oid])[0]
    total_hours = offering['total_hours']
    
    print(f"\n=== ID={oid} {desc} 总学时={total_hours} ===")
    
    # 步骤1: 确认已有的待确认补课
    confirmed_makeup = confirm_existing_makeup(oid)
    if confirmed_makeup:
        print(f"  确认已有补课: {confirmed_makeup}条")
    
    # 步骤2: 对ID=33,34，先确认全部sessions
    if oid in [33, 34]:
        with store.connect() as db:
            db.execute("UPDATE sessions SET status='已确认' WHERE offering_id=? AND status='待确认'", [oid])
            db.commit()
        print(f"  确认全部待确认sessions")
    
    # 步骤3: 检查是否还需要补课
    confirmed_hours = get_confirmed_hours(oid)
    deficit = total_hours - confirmed_hours
    
    if deficit > 0:
        # 添加18周补课
        added = add_makeup_sessions(oid, deficit)
        print(f"  差{deficit}学时 → 添加补课: {added} (共{sum(added)}学时)")
    elif deficit < 0:
        print(f"  超出{-deficit}学时（多班合并，需后续处理）")
    else:
        print(f"  学时已满，无需补课")
    
    # 最终验证
    final_hours = get_confirmed_hours(oid)
    print(f"  最终: 总学时={total_hours} 已确认={final_hours} 差={total_hours - final_hours}")

print("\n完成!")
