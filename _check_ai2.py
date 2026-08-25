import store, inspect

# 检查settings表结构
cols = store.rows("PRAGMA table_info(settings)")
print("settings字段:", [c['name'] for c in cols])

settings = store.rows("SELECT * FROM settings")
print(f"\nsettings: {len(settings)}条")
for s in settings:
    print(f"  {dict(s)}")

# 检查content_author中的AI调用
print("\n=== content_author AI调用 ===")
import content_author
src = inspect.getsource(content_author)
for i, line in enumerate(src.split('\n')):
    if any(kw in line.lower() for kw in ['openai', 'claude', 'api_key', 'model', 'gpt', 'llm', 'ai_client', 'ask_result', 'enhanced_content', 'requests.post', 'http', 'url', 'endpoint']):
        print(f"  L{i}: {line.strip()}")

# 检查是否有ai_client或类似模块
import os
ai_files = [f for f in os.listdir('.') if 'ai' in f.lower() and f.endswith('.py')]
print(f"\nAI相关文件: {ai_files}")

# 检查gen_plan/gen_standard/gen_design中的AI调用
for mod_name in ['gen_plan', 'gen_standard', 'gen_design']:
    try:
        mod = __import__(mod_name)
        src = inspect.getsource(mod)
        ai_lines = []
        for i, line in enumerate(src.split('\n')):
            if any(kw in line.lower() for kw in ['openai', 'claude', 'api_key', 'model', 'gpt', 'llm', 'ai_client', 'ask_result', 'enhanced_content', 'requests.post', 'http']):
                ai_lines.append(f"  {mod_name} L{i}: {line.strip()}")
        if ai_lines:
            print(f"\n=== {mod_name} AI调用 ===")
            for l in ai_lines[:10]:
                print(l)
    except Exception as e:
        print(f"\n{mod_name}: {e}")
