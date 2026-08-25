import store

# 检查id=20的生成记录
docs = store.rows("SELECT * FROM generated_documents WHERE offering_id=20")
print("=== ID=20 生成记录 ===")
for d in docs:
    print(f"  doc_type={d['document_type']} status={d['generation_status']} path={d['output_path']}")
    print(f"    template_file_id={d.get('template_file_id','')} structural_check={d.get('structural_check','')} visual_check={d.get('visual_check','')}")

# 检查generate.py的main函数
import generate, inspect
src = inspect.getsource(generate.generate_all)
print(f"\n=== generate_all函数 ===")
print(src)
