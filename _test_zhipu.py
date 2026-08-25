import store

# 测试智谱GLM连接
from ai.ai_router import ask_result

print("=== 测试智谱GLM连接 ===")
result = ask_result(
    "请回复'AI连接成功'，并说明你的模型名称。",
    system="你是教学辅助AI助手。",
    force_online=True,
    show_details=True
)

print(f"\n结果: success={result['success']}")
print(f"来源: {result.get('source', '')} / {result.get('model', '')}")
print(f"内容: {result.get('content', '')[:200]}")

if result['success']:
    print("\n✅ 智谱GLM连接正常，可以开启AI增强生成")
else:
    print(f"\n❌ 连接失败: {result.get('error', '')}")
