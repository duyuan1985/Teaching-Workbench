import os
import store

# 检查环境变量
zhipu_key = os.getenv("ZHIPU_API_KEY", "")
print(f"ZHIPU_API_KEY 环境变量: {'已设置(' + zhipu_key[:6] + '...' + ')' if zhipu_key else '未设置'}")

# 检查数据库settings
db_key = store.get_setting("zhipu_api_key", "")
print(f"数据库 zhipu_api_key: {'已设置(' + db_key[:6] + '...' + ')' if db_key else '未设置'}")

# 检查.ai_env文件
for f in ['.ai_env', '.env', 'ai_config.json']:
    if os.path.exists(f):
        with open(f, 'r', encoding='utf-8') as fh:
            content = fh.read()
            if 'ZHIPU' in content or 'zhipu' in content.lower():
                print(f"\n{f} 文件中发现ZHIPU配置:")
                for line in content.split('\n'):
                    if 'ZHIPU' in line or 'zhipu' in line.lower():
                        # 隐藏key
                        import re
                        line = re.sub(r'(ZHIPU_API_KEY\s*=\s*["\']?)([a-zA-Z0-9]{6})[a-zA-Z0-9]+', r'\1\2***', line)
                        print(f"  {line}")

# 检查enhanced_generation
enhanced = store.get_setting("enhanced_generation", "0")
print(f"\nenhanced_generation: {enhanced}")
print(f"ai_model_preference: {store.get_setting('ai_model_preference', 'online')}")

if not zhipu_key and not db_key:
    print("\n❌ 智谱API Key未配置")
    print("请设置环境变量 ZHIPU_API_KEY")
else:
    print("\n✅ 智谱API Key已配置")
