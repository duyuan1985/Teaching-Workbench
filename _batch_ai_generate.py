"""
批量AI生成流程：
1. build_semantic_model — 构建语义模型（不调AI）
2. author_course_content — 调用智谱GLM生成内容
3. generate_all — 生成文档

先只处理一个课程测试，确认AI生成正常后再批量执行
"""
import store
import semantic_model
import content_author
import generate
import time

# 先用ID=18测试
oid = 18
offering = store.rows("SELECT * FROM offerings WHERE id=?", [oid])[0]
print(f"=== 测试课程: ID={oid} {offering['course_name']} ({offering['term']}) ===")

# 步骤1: 构建语义模型
print("\n--- 步骤1: 构建语义模型 ---")
try:
    model = semantic_model.build_semantic_model(oid)
    print(f"  ✅ 语义模型构建成功")
    print(f"  项目数: {len(model.get('projects', []))}")
    print(f"  知识点: {len(model.get('knowledge_system', []))}")
    print(f"  工具技术: {len(model.get('tools_technology', []))}")
except Exception as e:
    print(f"  ❌ 失败: {e}")
    raise

# 步骤2: AI生成内容（调用智谱GLM）
print("\n--- 步骤2: AI生成内容（智谱GLM）---")
t0 = time.time()
try:
    count = content_author.author_course_content(oid)
    elapsed = time.time() - t0
    print(f"  ✅ AI内容生成完成，共{count}个section，耗时{elapsed:.0f}秒")
except Exception as e:
    elapsed = time.time() - t0
    print(f"  ❌ AI内容生成失败（耗时{elapsed:.0f}秒）: {e}")
    raise

# 检查生成的sections
sections = store.rows("SELECT document_type, section_key, authoring_status FROM authored_sections WHERE offering_id=?", [oid])
ai_count = sum(1 for s in sections if 'AI' in s.get('authoring_status', ''))
rule_count = sum(1 for s in sections if '结构化' in s.get('authoring_status', ''))
print(f"  生成sections: {len(sections)}个 (AI生成: {ai_count}, 规则兜底: {rule_count})")

print("\n=== 测试完成，可以批量执行 ===")
