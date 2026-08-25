import inspect
import os

# 查找build_model或类似函数
for fname in os.listdir('.'):
    if fname.endswith('.py') and not fname.startswith('_'):
        try:
            mod = __import__(fname.replace('.py', ''))
            src = inspect.getsource(mod)
            if 'course_content_models' in src and ('def build' in src or 'def extract' in src or 'def create_model' in src or 'model_json' in src):
                print(f"\n=== {fname} ===")
                for i, line in enumerate(src.split('\n')):
                    if 'course_content_model' in line or 'model_json' in line or ('def build' in line and 'model' in line.lower()):
                        print(f"  L{i}: {line.strip()}")
        except:
            pass

# 检查api.py中的模型构建路由
try:
    import api
    src = inspect.getsource(api)
    for i, line in enumerate(src.split('\n')):
        if 'course_content_model' in line or 'build_model' in line or 'extract_model' in line or 'create_model' in line:
            print(f"  api.py L{i}: {line.strip()}")
except:
    pass
