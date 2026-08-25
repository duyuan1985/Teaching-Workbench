import inspect, os

# 检查ai_router模块
try:
    from ai.ai_router import ask_result
    src = inspect.getsource(ask_result)
    print("=== ai_router.ask_result ===")
    print(src[:3000])
except Exception as e:
    print(f"ai_router: {e}")

# 检查ai目录
ai_dir = os.path.join(os.path.dirname(__file__), 'ai')
if os.path.exists(ai_dir):
    print(f"\n=== ai目录 ===")
    for f in os.listdir(ai_dir):
        print(f"  {f}")
