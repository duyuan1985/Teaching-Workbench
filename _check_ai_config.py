import os, json

# 检查配置文件
config_paths = [
    'config.py', 'config.json', '.env', 'settings.json',
    'config/settings.py', 'config/config.json'
]

for cp in config_paths:
    if os.path.exists(cp):
        print(f"=== {cp} ===")
        with open(cp, 'r', encoding='utf-8') as f:
            content = f.read()
            # 隐藏API密钥
            import re
            content = re.sub(r'(sk-[a-zA-Z0-9]{10})[a-zA-Z0-9]+', r'\1***', content)
            content = re.sub(r'(api[_-]?key["\s:=]+["\s]?)([a-zA-Z0-9]{6})[a-zA-Z0-9]+', r'\1\2***', content, flags=re.IGNORECASE)
            print(content[:2000])
        print()

# 检查settings表
import store
settings = store.rows("SELECT * FROM settings")
print("=== 数据库settings表 ===")
for s in settings:
    val = str(s.get('value', ''))
    # 隐藏敏感信息
    if 'key' in s.get('key', '').lower() or 'secret' in s.get('key', '').lower() or 'token' in s.get('key', '').lower():
        val = val[:6] + '***'
    print(f"  {s['key']} = {val}")

# 检查content_author中的AI调用
import inspect
try:
    import content_author
    src = inspect.getsource(content_author)
    # 找AI相关调用
    for line in src.split('\n'):
        if any(kw in line.lower() for kw in ['openai', 'claude', 'api_key', 'model', 'gpt', 'llm', 'ai_client', 'ask_result', 'enhanced_content']):
            print(f"  content_author: {line.strip()}")
except Exception as e:
    print(f"content_author: {e}")

# 检查api.py中的AI路由
try:
    import api
    src = inspect.getsource(api)
    for line in src.split('\n'):
        if any(kw in line.lower() for kw in ['openai', 'claude', 'api_key', 'model', 'gpt', 'llm', 'ai_client', '/api/ai', 'generate-content', 'ask']):
            print(f"  api.py: {line.strip()}")
except Exception as e:
    print(f"api.py: {e}")
